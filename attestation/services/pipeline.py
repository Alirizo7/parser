"""Оркестрация конвейера: unpack → convert → extract → единый датасет.

Функции не зависят от Django: принимают пути, отдают данные и зовут
``progress(stage)`` для обновления UI. Celery-задача оборачивает это,
обновляя ``Batch.status`` / ``Batch.stage``.
"""
from __future__ import annotations

import copy
import re
from collections import Counter, deque
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from . import extract as E
from . import mapping as M
from .convert import convert_to_docx
from .normalize import fold
from .unpack import UnpackedFile, unpack

ProgressFn = Callable[[str], None]


def _noop(_stage: str) -> None:
    pass


# Имя похоже на карту: число + опц. буквенный суффикс + разделитель.
# Разделитель включает ЗАПЯТУЮ: у части карт имя набрано «014,Бўлинма…» /
# «55,Бўлинма…» (запятая вместо точки — опечатка при наборе). Без запятой в
# классе такие файлы не опознавались как карты и молча выпадали из итога
# (2 рабочих места, 000014 и 000055, терялись → 112 вместо 114).
_CARD_NAME_HINT = re.compile(r"^\s*\d+\s*[а-яёА-ЯЁa-zA-Z]?\s*[.,_\-)\s]")


def _looks_like_card_filename(basename: str) -> bool:
    return bool(_CARD_NAME_HINT.match(basename))


def _folder_of(uf: UnpackedFile) -> str:
    path = uf.arc_name.replace("\\", "/")
    return path.rsplit("/", 1)[0] if "/" in path else ""


# --- Классификация файлов архива -------------------------------------------
def classify(uf: UnpackedFile) -> str:
    """Тип файла. «card» здесь — КАНДИДАТ; финальную папку карт выбирает pipeline.

    Карты не привязаны к точному имени папки: годятся и «карта база», и плоская
    «KARTA/» с файлами «000001. …». Кандидат — .doc/.docx (не «Перечень») в папке
    «карта база» ИЛИ с именем, похожим на карту (число + разделитель).
    """
    name = uf.basename.lower()
    parent = uf.arc_name.replace("\\", "/").lower()
    if uf.suffix == ".pdf":
        return "pdf"
    if uf.suffix in (".xlsx", ".xls"):
        return "xlsx"
    if any(h in name for h in M.PERECHEN_FILENAME_HINTS):
        return "perechen"
    if uf.suffix in (".doc", ".docx") and (
        "карта база/" in parent or _looks_like_card_filename(uf.basename)
    ):
        return "card"
    return "other"


@dataclass
class PipelineResult:
    company_data: dict
    workplaces: list[dict]
    files: list[dict] = field(default_factory=list)   # сводка по файлам (для SourceFile)
    warnings: list[str] = field(default_factory=list)


# Ключ сортировки и разбор номера — общие (в extract), чтобы рендер и UI
# сортировали РМ так же, как pipeline.
split_workplace_no = E.split_workplace_no
_workplace_sort_key = E.workplace_sort_key


# Поля, наследуемые «а»-суффиксным рабочим местом от базового
_SUFFIX_COPY_FIELDS = (
    "factors", "substances", "benefits", "ppe_provided",
    "plan_6_6",
    "injury_risk", "privileged_pension", "employees_count", "female_count",
    "injury_risk_class_6_4", "ppe_status_6_4", "ppe_not_envisaged_6_4",
    # Построчные замеры для Excel-протоколов (файлы 1–5)
    "physical_measurements", "microclimate_measurements",
    "lighting_measurements", "em_measurements",
)


def _fill_suffix_from_base(workplaces: list[dict], perechen_map: dict, warnings: list[str]) -> None:
    """Достроить данные «а»-суффиксных рабочих мест по базовому номеру.

    (a) карта суффикса есть, но факторы пусты → копируем данные базы;
    (b) суффикс есть в «Перечне», но карты нет → создаём РМ копией базы.
    """
    by_no = {w["workplace_no"]: w for w in workplaces}
    base_by_num = {}
    for w in workplaces:
        num, suf = split_workplace_no(w["workplace_no"])
        if num is not None and not suf:
            base_by_num[num] = w

    # (a) пустые факторы у существующей суффиксной карты
    for w in workplaces:
        num, suf = split_workplace_no(w["workplace_no"])
        if not suf or num not in base_by_num:
            continue
        overall = (w.get("factors") or {}).get("overall", "")
        if overall in ("", "-"):
            base = base_by_num[num]
            for fld in _SUFFIX_COPY_FIELDS:
                w[fld] = copy.deepcopy(base.get(fld))
            if not w.get("position"):
                w["position"] = base.get("position", "")
            if not w.get("subdivision"):
                w["subdivision"] = base.get("subdivision", "")
            w.setdefault("flags", []).append("copied_from_base")
            warnings.append(
                f"{w['workplace_no']}: раздел факторов в карте пуст — данные "
                f"скопированы из базового {base['workplace_no']}."
            )

    # (b) суффикс есть в «Перечне», но карты нет
    for pno in list(perechen_map):
        if pno in by_no:
            continue
        num, suf = split_workplace_no(pno)
        if not suf or num not in base_by_num:
            continue
        base = base_by_num[num]
        new = copy.deepcopy(base)
        new["workplace_no"] = pno
        new["source_file"] = ""
        new["flags"] = list(base.get("flags", [])) + ["copied_from_base", "card_missing"]
        workplaces.append(new)
        by_no[pno] = new
        warnings.append(
            f"{pno}: карта не найдена — рабочее место создано копированием базового "
            f"{base['workplace_no']}."
        )


def _repair_duplicate_workplace_numbers_from_perechen(
    workplaces: list[dict], perechen_map: dict[str, dict], warnings: list[str]
) -> None:
    """Исправить однозначную опечатку номера карты по должности из «Перечня».

    Источник номера в карте остаётся главным. Исключение безопасно только при
    одновременном выполнении трёх условий: номер карты продублирован, в
    «Перечне» есть ещё не занятый номер, и нормализованное название должности
    этой карты совпадает ровно с одной такой строкой. Неоднозначные случаи не
    меняем и оставляем оператору обычное предупреждение о дубле.
    """
    counts = Counter(w.get("workplace_no", "") for w in workplaces)
    duplicated = {no for no, count in counts.items() if no and count > 1}
    if not duplicated or not perechen_map:
        return

    occupied = {
        w.get("workplace_no", "") for w in workplaces
        if w.get("workplace_no") and w.get("workplace_no") not in duplicated
    }
    available = set(perechen_map) - occupied
    for old_no in sorted(duplicated, key=_workplace_sort_key):
        records = [w for w in workplaces if w.get("workplace_no") == old_no]
        assignments: list[tuple[dict, str]] = []
        used: set[str] = set()
        for rec in records:
            position_key = fold(rec.get("position", ""))
            candidates = [
                no for no in available - used
                if position_key and fold(perechen_map[no].get("position", "")) == position_key
            ]
            if len(candidates) == 1:
                assignments.append((rec, candidates[0]))
                used.add(candidates[0])
        # Исправляем только если весь набор дублей разложился однозначно и без
        # коллизий. Частичное угадывание могло бы скрыть ещё одну ошибку карты.
        if len(assignments) != len(records) or len(used) != len(records):
            continue
        for rec, new_no in assignments:
            if new_no == old_no:
                continue
            rec["workplace_no"] = new_no
            rec.setdefault("flags", []).append("workplace_no_repaired_from_perechen")
            warnings.append(
                f"{rec.get('source_file', 'карта')}: дублированный номер {old_no} "
                f"исправлен на {new_no} по точному совпадению должности в «Перечне»."
            )
        available -= used


def run_pipeline(
    zip_path: str | Path,
    work_dir: str | Path,
    *,
    progress: ProgressFn = _noop,
) -> PipelineResult:
    """Полный прогон: распаковать архив и собрать датасет рабочих мест."""
    work_dir = Path(work_dir)
    unpack_dir = work_dir / "unpacked"
    docx_dir = work_dir / "docx"
    docx_dir.mkdir(parents=True, exist_ok=True)

    # 1) Распаковка
    progress("Распаковка архива")
    files = unpack(zip_path, unpack_dir)

    card_candidates: list[UnpackedFile] = []
    perechen_candidates: list[UnpackedFile] = []
    for uf in files:
        kind = classify(uf)
        if kind == "card":
            card_candidates.append(uf)
        elif kind == "perechen":
            perechen_candidates.append(uf)

    # Папка карт = та, где БОЛЬШЕ ВСЕГО кандидатов. Так детектируем и «карта база»,
    # и плоскую «KARTA/», не цепляя одиночные файлы-имена-цифры из корня архива.
    folder_counts = Counter(_folder_of(uf) for uf in card_candidates)
    cards_folder = folder_counts.most_common(1)[0][0] if folder_counts else None
    cards = [uf for uf in card_candidates if _folder_of(uf) == cards_folder]

    file_summary = [{"path": uf.arc_name, "kind": classify(uf)} for uf in files]

    # Порядок обработки — по номеру из имени (косметика для прогресса);
    # итоговый список сортируется по номеру РМ из содержимого.
    def _name_num(u: UnpackedFile):
        n = E.parse_card_filename(u.basename)[0]
        return (n if n is not None else 10**9, u.basename)

    cards.sort(key=_name_num)
    warnings: list[str] = []
    if not cards:
        warnings.append("В архиве не найдено карт рабочих мест.")

    # 2) Конвертация + 3) извлечение карт
    company_data: dict = {}
    workplaces: list[dict] = []
    unrecognized: list[str] = []  # файлы-карты без распознанного номера РМ
    total = len(cards)
    for i, uf in enumerate(cards, start=1):
        progress(f"Конвертация и извлечение карт {i}/{total}")
        try:
            docx_path = convert_to_docx(uf.abs_path, docx_dir)
        except Exception as exc:  # noqa: BLE001
            unrecognized.append(f"{uf.basename} (ошибка конвертации: {exc})")
            continue
        record = E.extract_card(docx_path, uf.basename)
        if not record.get("workplace_no"):
            # Файл лежит в папке карт, но номер «…-сонли» не найден в содержимом
            unrecognized.append(f"{uf.basename} (не найден номер «…-сонли»)")
            continue
        workplaces.append(record)
        # Реквизиты компании — из первой удачной карты (они одинаковы во всех)
        if not company_data:
            company_data = E.extract_company_data(E.read_docx(docx_path))

    # ГРОМКОЕ предупреждение: тихая потеря рабочих мест недопустима
    if unrecognized:
        warnings.append(
            f"Карты: найдено файлов {total}, распознано рабочих мест {len(workplaces)}. "
            f"НЕ РАСПОЗНАНЫ ({len(unrecognized)}): " + "; ".join(unrecognized)
        )
    # 4) «Перечень» → коды должностей. Кандидатов может быть несколько
    # (напр. лист подписей «ИМЗО ПЕРЕЧЕН») — берём тот, что даёт больше записей.
    perechen_map: dict[str, dict] = {}
    best_perechen_docx: Path | None = None
    if perechen_candidates:
        progress("Разбор «Перечня» (коды должностей)")
        for cand in perechen_candidates:
            try:
                p_docx = convert_to_docx(cand.abs_path, docx_dir)
                candidate_map = E.parse_perechen(p_docx)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"Не удалось разобрать «{cand.basename}»: {exc}")
                continue
            if len(candidate_map) > len(perechen_map):
                perechen_map = candidate_map
                best_perechen_docx = p_docx
    if not perechen_map:
        warnings.append("«Перечень» с кодами должностей не найден — коды останутся пустыми.")

    # Исправляем только доказуемые опечатки: дублированный номер карты и ровно
    # одно совпадение её должности с ещё не занятым номером «Перечня».
    _repair_duplicate_workplace_numbers_from_perechen(workplaces, perechen_map, warnings)
    seen = Counter(rec["workplace_no"] for rec in workplaces)
    dups = [no for no, count in seen.items() if count > 1]
    if dups:
        warnings.append(
            "Повторяющиеся номера рабочих мест (проверьте карты): "
            + ", ".join(sorted(dups, key=_workplace_sort_key))
        )

    # 4.5) «а»-суффиксные рабочие места (доп. место того же типа): если в карте
    # пусты факторы или карта отсутствует — наследуем данные базового номера.
    _fill_suffix_from_base(workplaces, perechen_map, warnings)

    # 5) Слияние: код должности из «Перечня» по номеру рабочего места.
    # В 6_5 должность/код берутся из «Перечня» (как в эталонах клиента). Если
    # имя должности в «Перечне» расходится с именем в карте на одном номере —
    # помечаем флагом «проверить вручную» (код мог уйти к другой должности).
    for rec in workplaces:
        info = perechen_map.get(rec["workplace_no"])
        if info:
            rec["job_code"] = info.get("job_code", "")
            rec["position_from_perechen"] = info.get("position", "")
            card_pos, per_pos = rec.get("position", ""), info.get("position", "")
            if per_pos and card_pos and fold(card_pos) != fold(per_pos):
                rec["flags"].append("position_mismatch")
        else:
            rec["position_from_perechen"] = ""
            if "job_code_missing" not in rec["flags"]:
                rec["flags"].append("job_code_missing")
        # Нет извлечённых веществ — подсказка оператору (в карте раздел 1.1 пуст)
        if not rec.get("substances"):
            rec["flags"].append("substances_missing")

    # 5.04) 6_5 получает должность, код, подразделение и порядок из отдельного
    # логического чтения Перечня. Это изолирует исправление от остальных
    # документов: старые .doc иногда дублируют каждую физическую ячейку после
    # конвертации, из-за чего общий индексный разбор сдвигает должность/код.
    if best_perechen_docx is not None:
        try:
            perechen_positions_6_5, w = E.parse_perechen_positions_6_5(best_perechen_docx)
        except Exception as exc:  # noqa: BLE001
            perechen_positions_6_5, w = [], [f"6_5: не удалось разобрать «Перечень»: {exc}"]
        warnings.extend(w)

        queues_6_5: dict[str, deque] = {}
        first_6_5: dict[str, dict] = {}
        for pos in perechen_positions_6_5:
            no = pos["workplace_no"]
            queues_6_5.setdefault(no, deque()).append(pos)
            first_6_5.setdefault(no, pos)

        for rec in workplaces:
            no = rec["workplace_no"]
            q = queues_6_5.get(no)
            pos = q.popleft() if q else first_6_5.get(no)
            if pos is None:
                number, suffix = split_workplace_no(no)
                if number is not None and suffix:
                    pos = first_6_5.get(E.canonical_workplace_no(number))
            if pos is None:
                continue
            rec["position_from_perechen_6_5"] = pos["position"]
            rec["job_code_6_5"] = pos["job_code"]
            rec["subdivision_6_5"] = pos["subdivision"]
            rec["subdivision_headers_6_5"] = pos["subdivision_headers"]
            rec["subdivision_group_6_5"] = pos["group_index"]
            rec["perechen_order_6_5"] = pos["order"]
    elif perechen_candidates:
        warnings.append("6_5: не удалось выбрать «Перечень» — используются данные карт.")

    # 5.05) 6_4: позиции берутся из строк «Перечня» (не из карт) — Шаг 3/4/5
    # спецификации. Сопоставляем каждую строку Перечня со «своей» картой по
    # номеру рабочего места; дубли номеров (напр. две карты одной «а»-позиции
    # на разные смены) раскладываются по очереди в порядке появления карт —
    # не схлопываются в одну запись. Строка без карты — исключается из 6_4
    # и громко логируется (не подмена данными базового номера, как в 4.5:
    # там это для 6_5/5_1б, где такая подмена уже подтверждена эталонами).
    if best_perechen_docx is not None:
        try:
            perechen_positions_6_4, w = E.parse_perechen_positions_6_4(best_perechen_docx)
        except Exception as exc:  # noqa: BLE001
            perechen_positions_6_4, w = [], [f"6_4: не удалось разобрать «Перечень»: {exc}"]
        warnings.extend(w)
        queues: dict[str, deque] = {}
        for pos in perechen_positions_6_4:
            queues.setdefault(pos["workplace_no"], deque()).append(pos)
        for rec in workplaces:
            q = queues.get(rec["workplace_no"])
            if q:
                pos = q.popleft()
                rec["subdivision_6_4"] = pos["subdivision"]
                rec["subdivision_headers_6_4"] = pos["subdivision_headers"]
                rec["subdivision_group_6_4"] = pos["group_index"]
                rec["perechen_order_6_4"] = pos["order"]
                rec["employees_count_6_4"] = pos["employees_count"]
                rec["female_count_6_4"] = pos["female_count"]
        missing = [(no, len(q)) for no, q in queues.items() if q]
        if missing:
            details = "; ".join(f"{no} (не хватает карт: {n})" for no, n in missing)
            warnings.append(
                f"6_4: нет карты для {sum(n for _, n in missing)} позици(й) Перечня: {details}"
            )
    elif perechen_candidates:
        warnings.append("6_4: не удалось выбрать «Перечень» для позиций — 6_4 будет пустым.")

    # 5.1) Код должности для «а»-суффиксных РМ: «Перечень» обычно не содержит
    # отдельной строки «000012а», поэтому наследуем код (и должность) от базового
    # номера «000012» — это то же рабочее место того же типа.
    base_info = {
        split_workplace_no(r["workplace_no"])[0]: (r.get("job_code", ""), r.get("position_from_perechen", ""))
        for r in workplaces
        if not split_workplace_no(r["workplace_no"])[1] and r.get("job_code")
    }
    for rec in workplaces:
        num, suf = split_workplace_no(rec["workplace_no"])
        if suf and not rec.get("job_code") and num in base_info:
            rec["job_code"], base_pos = base_info[num]
            if not rec.get("position_from_perechen"):
                rec["position_from_perechen"] = base_pos
            if "job_code_missing" in rec["flags"]:
                rec["flags"].remove("job_code_missing")

    workplaces.sort(key=lambda r: _workplace_sort_key(r["workplace_no"]))
    progress("Готово")
    return PipelineResult(
        company_data=company_data,
        workplaces=workplaces,
        files=file_summary,
        warnings=warnings,
    )
