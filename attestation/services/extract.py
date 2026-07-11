"""Движок извлечения данных из карт и «Перечня».

Архитектурный приём: сначала каждую карту и «Перечень» раскладываем в чистые
структуры (таблицы как сетки строк/ячеек + абзацы), затем по декларативным
якорям (см. ``mapping.py``) достаём поля. Результат — одна запись на рабочее
место + блок реквизитов компании. Оба шаблона (5_1б, 6_5) потом заполняются
из этого единого набора.

Зависит только от python-docx и стандартной библиотеки.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document

from . import mapping as M
from .normalize import (
    canon_yesno,
    clean_substance_name,
    fold,
    fold_contains,
    fold_contains_all,
    is_empty,
    max_class,
    normalize_class,
    normalize_number,
    normalize_spaces,
)

# --- Имя файла карты: число + ЛЮБОЙ разделитель + название -------------------
# Допускаем точку, подчёркивание, дефис, пробел: «24а. Должность.doc»,
# «3_Умумий_...doc», «000004.Bosh mexanik.doc». Имя — лишь запасной источник;
# номер и должность берём из содержимого документа (имя бывает ненадёжным).
_CARD_FILENAME_RE = re.compile(
    r"^\s*(\d+)\s*([а-яёА-ЯЁ]?)\s*[._\-\s]\s*(.+?)\s*\.(?:doc|docx)$", re.IGNORECASE
)


def canonical_workplace_no(number: int, suffix: str = "") -> str:
    """Каноничный номер рабочего места: 6 цифр + опц. суффикс-вариант."""
    return f"{number:06d}{suffix}"


def split_workplace_no(wp: str) -> tuple[int | None, str]:
    """Разбить номер РМ на (число, буквенный суффикс): «000012а» → (12, 'а')."""
    m = re.match(r"^0*(\d+)\s*([^\d\s]*)\s*$", wp or "")
    if not m:
        return None, ""
    return int(m.group(1)), m.group(2)


def workplace_sort_key(wp: str) -> tuple[int, str]:
    """Ключ сортировки: 000011 < 000011а < 000011б < 000012 (суффикс — любая буква)."""
    num, suffix = split_workplace_no(wp)
    return (num if num is not None else 10**9, suffix)


def parse_card_filename(basename: str) -> tuple[int | None, str, str]:
    """Разобрать имя файла карты → (число, суффикс-вариант, название).

    Запасной источник: основной — содержимое документа. Возвращает
    ``(None, "", "")`` если имя не похоже на карту.
    """
    m = _CARD_FILENAME_RE.match(basename)
    if not m:
        return None, "", ""
    return int(m.group(1)), m.group(2), normalize_spaces(m.group(3))


# ---------------------------------------------------------------------------
# Чтение docx в простые структуры
# ---------------------------------------------------------------------------
@dataclass
class Doc:
    """Документ как набор таблиц-сеток и абзацев."""

    tables: list[list[list[str]]]  # таблица → строки → ячейки (текст)
    paragraphs: list[str]


def read_docx(path: str | Path) -> Doc:
    document = Document(str(path))
    tables: list[list[list[str]]] = []
    for tbl in document.tables:
        grid: list[list[str]] = []
        ncols = len(tbl.columns)
        for row in tbl.rows:
            cells = row.cells
            grid.append([normalize_spaces(cells[c].text) for c in range(min(ncols, len(cells)))])
        tables.append(grid)
    paragraphs = [normalize_spaces(p.text) for p in document.paragraphs]
    return Doc(tables=tables, paragraphs=paragraphs)


# --- Помощники поиска по якорям (двуязычные, через fold) ---------------------
def _contains(text: str, anchor: str) -> bool:
    return fold_contains(text, anchor)


def find_table(tables: list[list[list[str]]], anchor: str) -> list[list[str]] | None:
    """Первая таблица, в любой ячейке которой встречается anchor."""
    for grid in tables:
        for row in grid:
            if any(_contains(c, anchor) for c in row):
                return grid
    return None


def find_row(grid: list[list[str]], anchor: str) -> list[str] | None:
    for row in grid:
        if any(_contains(c, anchor) for c in row):
            return row
    return None


def find_row_tokens(grid: list[list[str]], tokens: list[str]) -> list[str] | None:
    """Первая строка, в тексте которой присутствуют ВСЕ токены (после fold)."""
    for row in grid:
        if fold_contains_all(" ".join(row), tokens):
            return row
    return None


def row_value_after_label(row: list[str], label: str) -> str:
    """Значение в строке «метка | значение»: первая непустая ячейка ≠ метки."""
    for cell in row:
        if not is_empty(cell) and not _contains(cell, label):
            return cell
    return ""


def last_value(row: list[str]) -> str:
    """Последняя непустая ячейка строки (аналог col -1)."""
    for cell in reversed(row):
        if not is_empty(cell):
            return cell
    return ""


# --- Сопоставление номеров разделов -----------------------------------------
_SECTION_RE = re.compile(r"^\s*(\d+(?:\.\d+)*)")


def _section_of(cell0: str) -> str | None:
    """Если ячейка col0 начинается с номера раздела (1.1, 1.3.2 …) — вернуть его."""
    m = _SECTION_RE.match(cell0)
    return m.group(1) if m else None


def section_matches(section: str | None, prefix: str) -> bool:
    """Раздел принадлежит префиксу? (1.1.3 ∈ 1.1, но 1.11 ∉ 1.1)."""
    if not section:
        return False
    return section == prefix or section.startswith(prefix + ".")


# ---------------------------------------------------------------------------
# Реквизиты компании
# ---------------------------------------------------------------------------
def _reqs_value(row: list[str], anchor: str) -> str:
    """Значение реквизита: первая непустая ячейка ≠ метки. Сохраняет «-»."""
    flabel = fold(anchor)
    for cell in row:
        raw = normalize_spaces(cell)
        if not raw:
            continue
        if flabel and flabel in fold(cell):  # это ячейка-метка
            continue
        return raw
    return ""


# Метки кодов в двух письмах; берём цифры (даже с пробелами) после метки
_CODE_LABELS = (
    ("stir", r"(?:СТИР|STIR)"),
    ("ifut", r"(?:ИФУТ|IFUT)"),
    ("mxbt", r"(?:МХБТ|MX[БB]T)"),
)


def _extract_codes(grid: list[list[str]]) -> dict:
    text = normalize_spaces(" ".join(" ".join(row) for row in grid))
    out = {}
    for key, lab in _CODE_LABELS:
        m = re.search(lab + r"\s*[:\-]?\s*([0-9][0-9 ]*)", text, re.IGNORECASE)
        out[key] = re.sub(r"\s+", "", m.group(1)) if m else ""
    return out


def extract_company_data(doc: Doc) -> dict:
    grid = find_table(doc.tables, M.COMPANY_FIELDS["name"])
    data = {k: "" for k in (*M.COMPANY_FIELDS, "stir", "ifut", "mxbt")}
    if grid is None:
        return data

    for key, anchor in M.COMPANY_FIELDS.items():
        row = find_row(grid, anchor)
        if row:
            data[key] = _reqs_value(row, anchor)

    data.update(_extract_codes(grid))
    return data


# ---------------------------------------------------------------------------
# Разбор большой таблицы факторов
# ---------------------------------------------------------------------------
@dataclass
class FactorRow:
    section: str | None  # номер раздела, «перенесённый» сверху
    c0: str              # исходный col0 (пусто у под-строк)
    name: str            # col1 — фактор/вещество
    actual: str          # «Ҳақиқий даражаси»
    duration: str        # «Таъсир этиш давомийлиги (соат/%)»
    cls: str             # класс (последняя колонка)


def _find_factor_table(doc: Doc) -> list[list[str]] | None:
    """Большая таблица факторов: по якорю заголовка, иначе самая длинная."""
    for grid in doc.tables:
        for row in grid[:3]:
            joined = " ".join(row).lower()
            if all(a in joined for a in (M.FACTOR_HEADER_CLASS,)) and any(
                a in joined for a in M.FACTOR_TABLE_HEADER_ANCHORS
            ):
                return grid
    return max(doc.tables, key=len) if doc.tables else None


def _factor_columns(grid: list[list[str]]) -> tuple[int, int, int, int]:
    """Индексы колонок (name, actual, duration, class) по строке-заголовку."""
    name_col, actual_col, duration_col = 1, 3, 4  # дефолты по подтверждённой раскладке
    class_col = max(len(r) for r in grid) - 1
    for row in grid[:3]:
        for i, c in enumerate(row):
            if _contains(c, M.FACTOR_HEADER_ACTUAL):
                actual_col = i
            if _contains(c, M.FACTOR_HEADER_DURATION):
                duration_col = i
            if _contains(c, "омиллари"):
                name_col = i
        if any(_contains(c, M.FACTOR_HEADER_CLASS) for c in row):
            class_col = len(row) - 1
    return name_col, actual_col, duration_col, class_col


def _parse_factor_rows(grid: list[list[str]]) -> list[FactorRow]:
    name_col, actual_col, duration_col, class_col = _factor_columns(grid)
    rows: list[FactorRow] = []
    current_section: str | None = None
    for raw in grid:
        if not raw:
            continue
        c0 = raw[0] if len(raw) > 0 else ""
        sec = _section_of(c0)
        if sec:
            current_section = sec

        def cell(i: int) -> str:
            return raw[i] if i < len(raw) else ""

        rows.append(
            FactorRow(
                section=current_section,
                c0=c0,
                name=cell(name_col),
                actual=cell(actual_col),
                duration=cell(duration_col),
                cls=last_value(raw) if class_col >= len(raw) else cell(class_col),
            )
        )
    return rows


def _factor_values(grid: list[list[str]], rows: list[FactorRow]) -> dict:
    """Классы по всем факторным колонкам 6_5."""
    factors: dict[str, str] = {}
    class_col = _factor_columns(grid)[3]

    # Подытоги/общий класс — по набору токенов, значение в КОЛОНКЕ КЛАССА.
    # Берём именно класс-колонку (а не последнюю непустую): если класс пуст,
    # последней непустой была бы объединённая ячейка-метка («Zararli
    # moddalarni umumiy baholash») — тогда вместо «-» протекал бы текст.
    for key, tokens in M.FACTOR_SUBTOTALS.items():
        row = find_row_tokens(grid, tokens)
        if not row:
            factors[key] = "-"
        else:
            val = row[class_col] if class_col < len(row) else last_value(row)
            factors[key] = normalize_class(val)

    # Факторы без подытога — максимум класса среди строк раздела
    for key, prefixes in M.FACTOR_SECTIONS.items():
        classes = [
            fr.cls
            for fr in rows
            if any(section_matches(fr.section, p) for p in prefixes)
            and not is_empty(fr.cls)
        ]
        factors[key] = max_class(classes)
    return factors


def _extract_substances(rows: list[FactorRow]) -> list[dict]:
    """Вредные вещества для 5_1б: под-строки разделов 1.1.* с реальным значением.

    Возвращаются в каноничном порядке эталона (см. ``SUBSTANCE_ORDER``):
    углерод оксиди первым, затем по приоритету; неизвестные — в порядке карты.
    """
    substances: list[dict] = []
    for fr in rows:
        if not section_matches(fr.section, M.SUBSTANCE_SECTION_PREFIX):
            continue
        # Вещество — под-строка (col0 пуст) с измеренным «Ҳақиқий даражаси»
        if is_empty(fr.c0) and not is_empty(fr.actual) and not is_empty(fr.name):
            name = clean_substance_name(fr.name)
            pct = normalize_number(fr.duration)
            if name:
                substances.append({"name": name, "pct": pct})

    def order_key(item: tuple[int, dict]) -> tuple[int, int]:
        idx, s = item
        rank = M.SUBSTANCE_ORDER.get(s["name"].lower(), M.SUBSTANCE_ORDER_UNKNOWN)
        return (rank, idx)

    return [s for _, s in sorted(enumerate(substances), key=order_key)]


# ---------------------------------------------------------------------------
# Подразделение, должность, номер
# ---------------------------------------------------------------------------
def _extract_position_from_doc(doc: Doc) -> str:
    """Должность — следующий непустой абзац после якоря «ИШ ЖОЙИ НОМИ».

    Работает и для кириллицы (``…ИШ ЖОЙИ НОМИ`` / ``Бош бухгалтер``), и для
    латиницы (``…ISH JOYI NOMI`` / ``Stropalchi``) — якорь матчится через fold.
    """
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        if _contains(p, M.POSITION_ANCHOR):
            for q in paras[i + 1:]:
                if not is_empty(q):
                    return q
    return ""


# Номер РМ из содержимого: «000009а-сонли» / «000012-сонли» / «000001-sonli».
# Берём по СЫРОМУ тексту (не fold), чтобы сохранить кириллический суффикс «а».
_WORKPLACE_CONTENT_RE = re.compile(
    r"([0-9]{4,6})\s*([а-яёА-ЯЁa-zA-Z]?)\s*[-—–]?\s*(?:сонли|sonli)", re.IGNORECASE
)


def _extract_workplace_no_from_doc(doc: Doc) -> tuple[str, str]:
    """Вернуть (число, суффикс-вариант) из строки «…-сонли» документа."""
    for p in doc.paragraphs:
        m = _WORKPLACE_CONTENT_RE.search(p)
        if m:
            return m.group(1), m.group(2).lower()
    for grid in doc.tables:
        for row in grid:
            m = _WORKPLACE_CONTENT_RE.search(" ".join(row))
            if m:
                return m.group(1), m.group(2).lower()
    return "", ""


def _extract_subdivision(doc: Doc) -> str:
    grid = find_table(doc.tables, M.SUBDIVISION_TABLE_ANCHOR)
    if grid is None:
        return ""
    row = find_row(grid, M.SUBDIVISION_ROW_ANCHOR)
    return row_value_after_label(row, M.SUBDIVISION_ROW_ANCHOR) if row else ""


def _extract_employee_counts(doc: Doc) -> tuple[str, str]:
    """(ишловчилар сони, шу жумладан аёллар) из таблицы «Таркибий бўлинма».

    Нужно для 6_4 (сводная қайднома по подразделениям): там строки не только
    «сколько рабочих мест», но и «сколько на них занято людей» / «из них
    женщин». Строка гендерной разбивки есть не во всех вариантах карты —
    тогда возвращаем "" (НЕ "0"), чтобы при агрегации в render_6_4 отличить
    «неизвестно» от «действительно ноль».
    """
    grid = find_table(doc.tables, M.SUBDIVISION_TABLE_ANCHOR)
    if grid is None:
        return "", ""
    workers_row = find_row(grid, M.WORKERS_ROW_ANCHOR)
    workers = normalize_number(row_value_after_label(workers_row, M.WORKERS_ROW_ANCHOR)) if workers_row else ""
    female_row = find_row(grid, M.WORKERS_FEMALE_ROW_ANCHOR)
    female = normalize_number(row_value_after_label(female_row, M.WORKERS_FEMALE_ROW_ANCHOR)) if female_row else ""
    return workers, female


# ---------------------------------------------------------------------------
# СИЗ (ЯТҲВ) и льготы
# ---------------------------------------------------------------------------
def _extract_ppe(doc: Doc) -> str:
    """ЯТҲВ обеспеченность: «ҳа», если хоть где-то отмечено наличие (бор/ҳа).

    Внимание: в шапке таблицы 6 подстрока «мавжудлиги» встречается дважды —
    в нужной колонке «Ходимда ЯТҲВ мавжудлиги (йўқ/бор)» и в последней
    «Сертификат … мавжудлиги». Берём ПЕРВОЕ совпадение (нужная колонка левее).
    """
    grid = find_table(doc.tables, M.PPE_TABLE_ANCHOR)
    if grid is None:
        return "йўқ"
    value_col = None
    for row in grid[:2]:
        for i, c in enumerate(row):
            if _contains(c, M.PPE_ROW_VALUE_ANCHOR):
                value_col = i
                break
        if value_col is not None:
            break
    has_ppe = False
    for row in grid[1:]:  # пропускаем шапку
        candidates = [row[value_col]] if value_col is not None and value_col < len(row) else row
        for c in candidates:
            if not is_empty(c) and canon_yesno(c) == "ҳа":
                has_ppe = True
    return "ҳа" if has_ppe else "йўқ"


def _extract_benefits(doc: Doc) -> dict:
    grid = find_table(doc.tables, M.BENEFITS_TABLE_ANCHOR)
    result = {k: "йўқ" for k in M.BENEFITS_ROWS}
    if grid is None:
        return result
    # Колонка фактического наличия: «Ҳақиқатда мавжудлиги» ИЛИ «Ҳақиқий
    # мавжудлиги» (формулировка различается у клиентов) — берём ПЕРВУЮ колонку
    # с «мавжудлиги» (она левее «ўрнатиш зарурлиги»/«асос»).
    fact_col = None
    for row in grid[:2]:
        for i, c in enumerate(row):
            if _contains(c, M.BENEFITS_FACT_COL_ANCHOR):
                fact_col = i
                break
        if fact_col is not None:
            break
    for key, anchor in M.BENEFITS_ROWS.items():
        for row in grid:
            if any(_contains(c, anchor) for c in row):
                val = row[fact_col] if fact_col is not None and fact_col < len(row) else ""
                result[key] = canon_yesno(val)
                break
    return result


# ---------------------------------------------------------------------------
# Эвристики для полей без детерминированного источника
# ---------------------------------------------------------------------------
def is_medical(position: str, subdivision: str) -> bool:
    """Медицинское/врачебное рабочее место?

    Опираемся и на должность, и на подразделение: группа
    «Врачлар, ўрта ва кичик тиббиёт ходимлари» содержит «тиббиёт»/«врач»,
    что относит к медицине даже немедицинские должности внутри неё
    (напр. «Хўжалик бекаси»), а директор-«Шифокор» ловится по должности.
    """
    text = f"{position} {subdivision}".lower()
    return any(kw in text for kw in M.MEDICAL_KEYWORDS)


def _injury_risk(medical: bool) -> str:
    return M.INJURY_RISK_MEDICAL if medical else M.INJURY_RISK_DEFAULT


def _extract_pension(doc: Doc) -> str:
    """Льготная пенсия — из раздела 4.2 карты: «…пенсия таъминоти ҳуқуқи __ ҳа __».

    Значение вписано в пропуск между «ҳуқуқи»/«huquqi» и скобкой «(агар…)».
    """
    for p in doc.paragraphs:
        if not (_contains(p, "пенсия таъминоти") or _contains(p, "пенсия таминоти")):
            continue
        m = re.search(r"(?:ҳуқуқи|хуқуқи|huquqi)(.*?)(?:\(|$)", p, re.IGNORECASE)
        if m:
            val = re.sub(r"[_‎\s]+", " ", m.group(1)).strip()
            return canon_yesno(val)
    return "йўқ"


# ---------------------------------------------------------------------------
# Сборка записи по одной карте
# ---------------------------------------------------------------------------
def extract_card(docx_path: str | Path, basename: str) -> dict:
    """Извлечь запись рабочего места из одной карты (.docx)."""
    doc = read_docx(docx_path)

    # Номер и должность — ИЗ СОДЕРЖИМОГО документа (имя файла ненадёжно:
    # бывает через подчёркивание «3_…», или вовсе не совпадает — «Energetik»
    # в имени при «Montajchi» внутри). Имя файла — лишь запасной источник.
    _f_num, f_suffix, f_title = parse_card_filename(basename)
    doc_num, doc_suffix = _extract_workplace_no_from_doc(doc)

    # Номер — ТОЛЬКО из содержимого: имя файла ненадёжно (подчёркивания, чужое
    # название). Нет строки «…-сонли» → номер пуст → карта уйдёт в «НЕ РАСПОЗНАНЫ»
    # (громкое предупреждение в pipeline), а не получит случайный номер из имени.
    suffix = doc_suffix or f_suffix  # суффикс-вариант обычно есть и в содержимом
    workplace_no = canonical_workplace_no(int(doc_num), suffix) if doc_num else ""

    position = _extract_position_from_doc(doc) or f_title
    subdivision = _extract_subdivision(doc)

    grid = _find_factor_table(doc)
    factor_rows = _parse_factor_rows(grid) if grid else []
    factors = _factor_values(grid, factor_rows) if grid else {}
    substances = _extract_substances(factor_rows)

    # Льготы (колонки 21–25 в 6_5) — ПРЯМО ИЗ КАРТЫ, раздел 4 (НЕ эвристики):
    #   отпуск/сокр.день/питание/молоко — таблица 4.1, колонка «Ҳақиқатда мавжудлиги»;
    #   льготная пенсия — текст 4.2 («…пенсия таъминоти ҳуқуқи __ ҳа __»).
    # Прим.: эталон Бухоро вручную ставит молоко=ҳа всем медикам (20 РМ) вопреки
    # форме карты — эти ячейки в самопроверке Бухоро идут в «заметки». См. README.
    benefits = _extract_benefits(doc)
    ppe = _extract_ppe(doc)
    medical = is_medical(position, subdivision)
    injury_risk = _injury_risk(medical)
    pension = _extract_pension(doc)
    employees_count, female_count = _extract_employee_counts(doc)

    # Флаги для подсветки в UI (поля, требующие проверки оператором)
    flags: list[str] = []
    if not factors.get("overall") or factors.get("overall") == "-":
        flags.append("overall_missing")
    flags.append("injury_risk_heuristic")     # травмоопасность (c18) — эвристика
    if not employees_count:
        # Нет строки «ишловчилар сони» в карте — считаем минимум 1 (карта
        # описывает занятое рабочее место), но подсвечиваем для проверки.
        employees_count = "1"
        flags.append("employees_count_missing")
    if not female_count:
        flags.append("female_count_missing")  # нужно для 6_4; см. mapping.WORKERS_FEMALE_ROW_ANCHOR

    return {
        "workplace_no": workplace_no,
        "source_file": basename,
        "workplace_no_in_doc": (doc_num + doc_suffix) if doc_num else "",
        "position": position,
        "position_from_filename": f_title,
        "subdivision": subdivision,
        "job_code": "",  # заполняется из «Перечня» на этапе слияния
        "factors": factors,
        "substances": substances,
        "ppe_provided": ppe,
        "benefits": benefits,
        "injury_risk": injury_risk,
        "privileged_pension": pension,
        "employees_count": employees_count,
        "female_count": female_count,
        "flags": flags,
    }


# ---------------------------------------------------------------------------
# Разбор «Перечня» — источник кода должности (и сверка должности)
# ---------------------------------------------------------------------------
def parse_perechen(docx_path: str | Path) -> dict[str, dict]:
    """Вернуть {workplace_no: {'job_code', 'position'}} из «Перечня»."""
    doc = read_docx(docx_path)
    grid = max(doc.tables, key=len) if doc.tables else None
    if not grid:
        return {}

    # Определяем колонки по объединённому заголовку (строки 0–3)
    header = [" ".join(grid[r][c] if c < len(grid[r]) else "" for r in range(min(4, len(grid))))
              for c in range(max(len(r) for r in grid))]
    code_col = position_col = workplace_col = None
    for i, h in enumerate(header):
        hl = h.lower()
        if code_col is None and M.PERECHEN_COL_CODE in hl:
            code_col = i
        if workplace_col is None and M.PERECHEN_COL_WORKPLACE in hl:
            workplace_col = i
        if position_col is None and M.PERECHEN_COL_POSITION in hl and M.PERECHEN_COL_CODE not in hl:
            position_col = i
    if workplace_col is None:
        workplace_col = 0
    if position_col is None:
        position_col = 1
    if code_col is None:
        code_col = 2

    result: dict[str, dict] = {}
    for row in grid:
        if workplace_col >= len(row):
            continue
        wp = normalize_spaces(row[workplace_col])
        if not re.match(r"^\d{4,6}[а-яёА-ЯЁ]?$", wp):
            continue
        result[wp] = {
            "job_code": normalize_spaces(row[code_col]) if code_col < len(row) else "",
            "position": normalize_spaces(row[position_col]) if position_col < len(row) else "",
        }
    return result
