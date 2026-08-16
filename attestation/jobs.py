"""Оркестрация обработки батча (Django-«клей» между сервисами и моделью).

Здесь — функции, которые загружают ``Batch``, гоняют пайплайн/рендер и
обновляют статус. Их вызывает либо Celery-задача (прод), либо фоновый поток
(локальная отладка) — см. ``ATTESTATION_TASK_RUNNER`` и ``start_*`` ниже.
"""
from __future__ import annotations

import logging
import shutil
import threading
from pathlib import Path

from django.conf import settings
from django.db import connection
from django.utils import timezone

from .models import Batch, SourceFile
from .services import render
from .services.pipeline import run_pipeline

logger = logging.getLogger(__name__)

PROCESS_QUEUED_STAGE = "Ожидание обработки…"
PROCESS_RUNNING_STAGE = "Подготовка…"
GENERATION_QUEUED_STAGE = "Ожидание формирования документов…"
GENERATION_RUNNING_STAGE = "Формирование документов…"
DELETE_WAITING_STAGE = "Ожидание остановки обработки…"
DELETE_WORKER_STAGE = "Удаление worker-ом…"
DELETE_RUNNING_STAGE = "Удаление файлов…"
DELETE_FAILED_STAGE = "Ошибка удаления"


class BatchCancelled(Exception):
    """Работа была удалена или перестала принадлежать текущему worker-у."""


def _batch_dir(batch_id: int) -> Path:
    return Path(settings.MEDIA_ROOT) / "batches" / str(batch_id)


def _safe_batch_dir(batch_id: int) -> Path:
    """Вернуть каталог батча, не разрешая symlink/path traversal наружу."""
    raw_media_root = Path(settings.MEDIA_ROOT)
    media_root = raw_media_root.resolve()
    raw_batches_root = raw_media_root / "batches"
    if raw_batches_root.is_symlink():
        raise OSError("Корневой каталог batches не должен быть символической ссылкой")
    batches_root = raw_batches_root.resolve()
    if batches_root.parent != media_root:
        raise OSError("Каталог batches находится вне MEDIA_ROOT")

    raw_batch_dir = raw_batches_root / str(int(batch_id))
    if raw_batch_dir.is_symlink():
        raise OSError("Каталог батча не должен быть символической ссылкой")
    batch_dir = raw_batch_dir.resolve()
    if batch_dir.parent != batches_root:
        raise OSError("Каталог батча находится вне MEDIA_ROOT/batches")
    return batch_dir


def _remove_batch_dir(batch_id: int) -> None:
    batch_dir = _safe_batch_dir(batch_id)
    if batch_dir.exists():
        shutil.rmtree(batch_dir)


def delete_batch_artifacts(batch: Batch) -> None:
    """Удалить все физические файлы батча.

    ``Batch.delete()`` каскадно удаляет ``SourceFile`` из БД, но Django не
    удаляет ни ``FileField``, ни рабочую директорию автоматически. Путь к
    директории строим только из числового PK и дополнительно ограничиваем
    ``MEDIA_ROOT/batches`` — значения output-полей для рекурсивного удаления не
    используем.

    Файлы удаляются до строки БД: если файловая система недоступна, view оставит
    запись на месте и оператор сможет безопасно повторить удаление.
    """
    if batch.pk is None:
        raise ValueError("Нельзя удалить файлы несохранённого батча")

    _remove_batch_dir(batch.pk)

    archive_name = batch.archive.name
    if archive_name:
        batch.archive.storage.delete(archive_name)


def _set(batch_id: int, **fields) -> int:
    """Точечно обновить поля батча (не затирая остальные, видно из других потоков)."""
    return Batch.objects.filter(pk=batch_id).update(**fields)


def _set_processing(batch_id: int, **fields) -> None:
    """Обновить только живую работу текущего worker-а или остановить его."""
    updated = Batch.objects.filter(
        pk=batch_id, status=Batch.Status.PROCESSING
    ).update(**fields)
    if updated != 1:
        raise BatchCancelled


def _assert_processing(batch_id: int) -> None:
    if not Batch.objects.filter(
        pk=batch_id, status=Batch.Status.PROCESSING
    ).exists():
        raise BatchCancelled


def _claim_queued_job(batch_id: int, queued_stage: str, running_stage: str) -> bool:
    """Разрешить выполнение только одному worker-у даже при повторной доставке."""
    return Batch.objects.filter(
        pk=batch_id,
        status=Batch.Status.PROCESSING,
        stage=queued_stage,
    ).update(stage=running_stage) == 1


def _cleanup_deleted_batch_dir(batch_id: int) -> None:
    """Финализировать отмену worker-а и убрать возможные поздние файлы.

    Для активной работы delete-view оставляет tombstone ``DELETING``. Worker
    удаляет её только в ``finally`` после последней возможной записи на диск.
    Если строку уже удалили (обычная завершённая работа), дочищаем один каталог.
    """
    try:
        batch = Batch.objects.filter(pk=batch_id).first()
        if batch is None:
            _remove_batch_dir(batch_id)
            return
        # CAS ownership: только один worker может забрать WAITING-tombstone;
        # view после stale takeover меняет stage на DELETE_RUNNING_STAGE, и этот
        # finally тогда не касается его файлов.
        claimed = Batch.objects.filter(
            pk=batch_id,
            status=Batch.Status.DELETING,
            stage=DELETE_WAITING_STAGE,
        ).update(stage=DELETE_WORKER_STAGE, updated_at=timezone.now())
        if claimed != 1:
            return
        batch = Batch.objects.get(pk=batch_id)
        delete_batch_artifacts(batch)
        Batch.objects.filter(
            pk=batch_id,
            status=Batch.Status.DELETING,
            stage=DELETE_WORKER_STAGE,
        ).delete()
    except Exception as exc:  # noqa: BLE001 — worker уже остановлен, сохраняем retry
        logger.exception("Не удалось финализировать удаление Batch #%s", batch_id)
        try:
            Batch.objects.filter(
                pk=batch_id,
                status=Batch.Status.DELETING,
                stage=DELETE_WORKER_STAGE,
            ).update(
                status=Batch.Status.FAILED,
                stage=DELETE_FAILED_STAGE,
                error=f"Не удалось удалить файлы работы: {exc}",
                updated_at=timezone.now(),
            )
        except Exception:  # noqa: BLE001 — соединение уже может быть недоступно
            logger.exception("Не удалось сохранить ошибку удаления Batch #%s", batch_id)


# ---------------------------------------------------------------------------
# Этап 1: распаковка → конвертация → извлечение
# ---------------------------------------------------------------------------
def process_batch(batch_id: int) -> None:
    try:
        if not _claim_queued_job(batch_id, PROCESS_QUEUED_STAGE, PROCESS_RUNNING_STAGE):
            return
        try:
            batch = Batch.objects.get(pk=batch_id)
        except Batch.DoesNotExist as exc:
            raise BatchCancelled from exc
        work_dir = _batch_dir(batch_id) / "work"

        def progress(stage: str) -> None:
            _set_processing(batch_id, stage=stage)

        result = run_pipeline(batch.archive.path, work_dir, progress=progress)
        _assert_processing(batch_id)

        # Сводка по файлам архива
        SourceFile.objects.filter(batch=batch).delete()
        SourceFile.objects.bulk_create(
            [SourceFile(batch=batch, path=f["path"], kind=f["kind"]) for f in result.files]
        )

        _set_processing(
            batch_id,
            status=Batch.Status.EXTRACTED,
            stage=f"Извлечено рабочих мест: {len(result.workplaces)}",
            company_data=result.company_data,
            extracted_data=result.workplaces,
            error="\n".join(result.warnings),
        )
    except BatchCancelled:
        pass
    except Exception as exc:  # noqa: BLE001 — фиксируем ошибку в модели
        Batch.objects.filter(
            pk=batch_id, status=Batch.Status.PROCESSING
        ).update(status=Batch.Status.FAILED, stage="Ошибка", error=str(exc))
    finally:
        _cleanup_deleted_batch_dir(batch_id)
        connection.close()  # закрыть соединение потока (важно для thread-раннера)


# ---------------------------------------------------------------------------
# Этап 2: генерация документов из (возможно отредактированного) датасета
# ---------------------------------------------------------------------------
def generate_documents(batch_id: int) -> None:
    try:
        if not _claim_queued_job(
            batch_id, GENERATION_QUEUED_STAGE, GENERATION_RUNNING_STAGE
        ):
            return
        try:
            batch = Batch.objects.get(pk=batch_id)
        except Batch.DoesNotExist as exc:
            raise BatchCancelled from exc
        out_dir = _batch_dir(batch_id)
        out_dir.mkdir(parents=True, exist_ok=True)
        lang = batch.output_lang
        p5 = render.render_5_1b(batch.extracted_data, out_dir / "5_1b.docx", lang=lang)
        _assert_processing(batch_id)
        p65 = render.render_6_5(
            batch.company_data, batch.extracted_data, out_dir / "6_5.docx", lang=lang
        )
        _assert_processing(batch_id)
        doc_warnings: list[str] = []
        p64 = render.render_6_4(
            batch.company_data, batch.extracted_data, out_dir / "6_4.docx", lang=lang,
            warnings=doc_warnings,
        )
        _assert_processing(batch_id)
        p66 = render.render_6_6(
            batch.company_data, batch.extracted_data, out_dir / "6_6.docx", lang=lang
        )
        _assert_processing(batch_id)
        media_root = Path(settings.MEDIA_ROOT)
        # Пять Excel-протоколов: сбой одного НЕ должен ронять весь батч и терять
        # уже готовые docx — рендерим каждый в try, ошибку кладём в warnings.
        excel_rel: dict[int, str] = {}
        for n in range(1, 6):
            _assert_processing(batch_id)
            try:
                p = getattr(render, f"render_excel_{n}")(
                    batch.company_data, batch.extracted_data,
                    out_dir / f"excel_{n}.xlsx", lang=lang,
                )
                excel_rel[n] = str(Path(p).relative_to(media_root))
            except Exception as exc:  # noqa: BLE001 — не валим docx из-за одного xlsx
                doc_warnings.append(f"Excel-протокол {n}: ошибка генерации — {exc}")
        error = batch.error or ""
        if doc_warnings:
            error = "\n".join(filter(None, [error, *doc_warnings]))
        _set_processing(
            batch_id,
            status=Batch.Status.DONE,
            stage="Готово",
            output_5_1b=str(Path(p5).relative_to(media_root)),
            output_6_5=str(Path(p65).relative_to(media_root)),
            output_6_4=str(Path(p64).relative_to(media_root)),
            output_6_6=str(Path(p66).relative_to(media_root)),
            output_excel_1=excel_rel.get(1, ""),
            output_excel_2=excel_rel.get(2, ""),
            output_excel_3=excel_rel.get(3, ""),
            output_excel_4=excel_rel.get(4, ""),
            output_excel_5=excel_rel.get(5, ""),
            error=error,
        )
    except BatchCancelled:
        pass
    except Exception as exc:  # noqa: BLE001
        Batch.objects.filter(
            pk=batch_id, status=Batch.Status.PROCESSING
        ).update(status=Batch.Status.FAILED, stage="Ошибка", error=str(exc))
    finally:
        _cleanup_deleted_batch_dir(batch_id)
        connection.close()


# ---------------------------------------------------------------------------
# Диспетчеризация: Celery или фоновый поток
# ---------------------------------------------------------------------------
def _dispatch(job_name: str, batch_id: int) -> None:
    if settings.ATTESTATION_TASK_RUNNER == "celery":
        from . import tasks
        getattr(tasks, f"{job_name}_task").delay(batch_id)
    else:
        threading.Thread(target=globals()[job_name], args=(batch_id,), daemon=True).start()


def start_processing(batch_id: int) -> bool:
    claimed = Batch.objects.filter(
        pk=batch_id, status=Batch.Status.UPLOADED
    ).update(status=Batch.Status.PROCESSING, stage=PROCESS_QUEUED_STAGE, error="")
    if claimed != 1:
        return False
    try:
        _dispatch("process_batch", batch_id)
    except Exception as exc:
        Batch.objects.filter(
            pk=batch_id,
            status=Batch.Status.PROCESSING,
            stage=PROCESS_QUEUED_STAGE,
        ).update(status=Batch.Status.FAILED, stage="Ошибка", error=str(exc))
        raise
    return True


def start_generation(batch_id: int) -> bool:
    # Атомарный claim закрывает двойной клик и повторную доставку Celery-задачи.
    claimed = Batch.objects.filter(
        pk=batch_id,
        status__in=(Batch.Status.EXTRACTED, Batch.Status.DONE),
    ).update(
        status=Batch.Status.PROCESSING,
        stage=GENERATION_QUEUED_STAGE,
    )
    if claimed != 1:
        return False
    try:
        _dispatch("generate_documents", batch_id)
    except Exception as exc:
        Batch.objects.filter(
            pk=batch_id,
            status=Batch.Status.PROCESSING,
            stage=GENERATION_QUEUED_STAGE,
        ).update(status=Batch.Status.FAILED, stage="Ошибка", error=str(exc))
        raise
    return True
