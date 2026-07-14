"""Оркестрация обработки батча (Django-«клей» между сервисами и моделью).

Здесь — функции, которые загружают ``Batch``, гоняют пайплайн/рендер и
обновляют статус. Их вызывает либо Celery-задача (прод), либо фоновый поток
(локальная отладка) — см. ``ATTESTATION_TASK_RUNNER`` и ``start_*`` ниже.
"""
from __future__ import annotations

import threading
from pathlib import Path

from django.conf import settings
from django.db import connection

from .models import Batch, SourceFile
from .services import render
from .services.pipeline import run_pipeline


def _batch_dir(batch_id: int) -> Path:
    return Path(settings.MEDIA_ROOT) / "batches" / str(batch_id)


def _set(batch_id: int, **fields) -> None:
    """Точечно обновить поля батча (не затирая остальные, видно из других потоков)."""
    Batch.objects.filter(pk=batch_id).update(**fields)


# ---------------------------------------------------------------------------
# Этап 1: распаковка → конвертация → извлечение
# ---------------------------------------------------------------------------
def process_batch(batch_id: int) -> None:
    try:
        batch = Batch.objects.get(pk=batch_id)
    except Batch.DoesNotExist:
        return
    _set(batch_id, status=Batch.Status.PROCESSING, stage="Подготовка…", error="")
    try:
        work_dir = _batch_dir(batch_id) / "work"

        def progress(stage: str) -> None:
            _set(batch_id, stage=stage)

        result = run_pipeline(batch.archive.path, work_dir, progress=progress)

        # Сводка по файлам архива
        SourceFile.objects.filter(batch=batch).delete()
        SourceFile.objects.bulk_create(
            [SourceFile(batch=batch, path=f["path"], kind=f["kind"]) for f in result.files]
        )

        _set(
            batch_id,
            status=Batch.Status.EXTRACTED,
            stage=f"Извлечено рабочих мест: {len(result.workplaces)}",
            company_data=result.company_data,
            extracted_data=result.workplaces,
            error="\n".join(result.warnings),
        )
    except Exception as exc:  # noqa: BLE001 — фиксируем ошибку в модели
        _set(batch_id, status=Batch.Status.FAILED, stage="Ошибка", error=str(exc))
    finally:
        connection.close()  # закрыть соединение потока (важно для thread-раннера)


# ---------------------------------------------------------------------------
# Этап 2: генерация документов из (возможно отредактированного) датасета
# ---------------------------------------------------------------------------
def generate_documents(batch_id: int) -> None:
    try:
        batch = Batch.objects.get(pk=batch_id)
    except Batch.DoesNotExist:
        return
    _set(batch_id, status=Batch.Status.PROCESSING, stage="Формирование документов…", error="")
    try:
        out_dir = _batch_dir(batch_id)
        out_dir.mkdir(parents=True, exist_ok=True)
        lang = batch.output_lang
        p5 = render.render_5_1b(batch.extracted_data, out_dir / "5_1b.docx", lang=lang)
        p65 = render.render_6_5(
            batch.company_data, batch.extracted_data, out_dir / "6_5.docx", lang=lang
        )
        doc_warnings: list[str] = []
        p64 = render.render_6_4(
            batch.company_data, batch.extracted_data, out_dir / "6_4.docx", lang=lang,
            warnings=doc_warnings,
        )
        media_root = Path(settings.MEDIA_ROOT)
        error = batch.error or ""
        if doc_warnings:
            error = "\n".join(filter(None, [error, *doc_warnings]))
        _set(
            batch_id,
            status=Batch.Status.DONE,
            stage="Готово",
            output_5_1b=str(Path(p5).relative_to(media_root)),
            output_6_5=str(Path(p65).relative_to(media_root)),
            output_6_4=str(Path(p64).relative_to(media_root)),
            error=error,
        )
    except Exception as exc:  # noqa: BLE001
        _set(batch_id, status=Batch.Status.FAILED, stage="Ошибка", error=str(exc))
    finally:
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


def start_processing(batch_id: int) -> None:
    _dispatch("process_batch", batch_id)


def start_generation(batch_id: int) -> None:
    _dispatch("generate_documents", batch_id)
