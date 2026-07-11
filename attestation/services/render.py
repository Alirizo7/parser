"""Заполнение шаблонов 5_1б и 6_5 через python-docx.

Подход: берём фиксированный пустой шаблон-ассет, очищаем строки тела ниже
шапки и генерируем тело заново из единого датасета (см. ``pipeline``).
Форматирование сохраняем, клонируя «строку-прототип» тела шаблона.
"""
from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.table import Table

from . import mapping as M
from .extract import workplace_sort_key
from .normalize import fold, fold_contains, to_latin

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
TEMPLATE_5_1B = ASSETS_DIR / "template_5_1b.docx"
TEMPLATE_6_5 = ASSETS_DIR / "template_6_5.docx"
TEMPLATE_6_4 = ASSETS_DIR / "template_6_4.docx"


def _transliterate_doc(doc, lang: str) -> None:
    """Привести ВЕСЬ текст документа к выбранному письму перед сохранением.

    ``lang == 'lat'`` → транслитерируем кириллицу в латиницу (заголовки шаблона
    и данные); ``'cyr'`` → ничего не делаем (текущее поведение, документ уже на
    кириллице). Цифры/коды/пунктуацию ``to_latin`` не трогает.
    """
    if lang != "lat":
        return

    def fix_paragraph(p):
        for run in p.runs:
            if run.text:
                run.text = to_latin(run.text)

    def fix_table(tbl):
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    fix_paragraph(p)
                for inner in cell.tables:
                    fix_table(inner)

    for p in doc.paragraphs:
        fix_paragraph(p)
    for tbl in doc.tables:
        fix_table(tbl)


# --- Низкоуровневые помощники работы с ячейками -----------------------------
def set_cell_text(cell, text: str) -> None:
    """Записать текст в ячейку, сохранив форматирование первого run-а."""
    text = "" if text is None else str(text)
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def _clear_body(table: Table, keep_header_rows: int):
    """Удалить строки тела ниже шапки, вернув копию строки-прототипа."""
    rows = list(table.rows)
    prototype = deepcopy(rows[keep_header_rows]._tr) if len(rows) > keep_header_rows else None
    for row in rows[keep_header_rows:]:
        table._tbl.remove(row._tr)
    return prototype


def _append_row(table: Table, prototype):
    """Добавить новую строку тела по прототипу; вернуть объект строки."""
    tr = deepcopy(prototype)
    table._tbl.append(tr)
    return table.rows[-1]


def _pct_num(value: str) -> str:
    m = re.search(r"\d+", value or "")
    return m.group(0) if m else ""


# --- 5_1б: вредные вещества по рабочим местам -------------------------------
def group_substances(substances: list[dict]) -> list[tuple[str, str]]:
    """Сгруппировать вещества по проценту воздействия.

    Возвращает список (имена_через_запятую, процент) в порядке появления
    процентов. Вещества внутри группы уже упорядочены на этапе извлечения.
    """
    groups: dict[str, list[str]] = {}
    for s in substances:
        pct = _pct_num(s.get("pct", ""))
        groups.setdefault(pct, []).append(s["name"])

    # Как в эталоне: первое вещество в ячейке с заглавной, остальные — строчными
    # («Углерод оксиди, азот оксиди, силикат чанги (лой)»).
    def join(names: list[str]) -> str:
        if not names:
            return ""
        return ", ".join([names[0]] + [n[:1].lower() + n[1:] for n in names[1:]])

    return [(join(names), pct) for pct, names in groups.items()]


def render_5_1b(workplaces: list[dict], out_path: str | Path,
                *, template_path: str | Path = TEMPLATE_5_1B, lang: str = "cyr") -> Path:
    """Сформировать документ 5_1б из датасета рабочих мест."""
    doc = Document(str(template_path))
    table = doc.tables[0]
    # Шапка: R0 (заголовки) + R1 («1|2|3»). Тело — с R2.
    prototype = _clear_body(table, keep_header_rows=2)

    # Гарантируем порядок: 000011 < 000011а < 000012 (даже если вход не отсортирован)
    workplaces = sorted(workplaces, key=lambda w: workplace_sort_key(w.get("workplace_no", "")))
    for wp in workplaces:
        groups = group_substances(wp.get("substances", []))
        if not groups:
            # Веществ в карте нет (раздел 1.1 пуст) — выводим РАБОЧЕЕ МЕСТО ВСЁ
            # РАВНО, оставляя «Модданинг номи» пустым (флаг substances_missing).
            # Так число строк совпадает с ожиданием, а пропуск виден в документе.
            row = _append_row(table, prototype)
            set_cell_text(row.cells[0], wp["workplace_no"])
            set_cell_text(row.cells[1], "")
            set_cell_text(row.cells[2], "")
            continue
        for gi, (names, pct) in enumerate(groups):
            row = _append_row(table, prototype)
            set_cell_text(row.cells[0], wp["workplace_no"] if gi == 0 else "")
            set_cell_text(row.cells[1], names)
            set_cell_text(row.cells[2], f"{pct} %" if pct else "")

    _transliterate_doc(doc, lang)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


# --- 6_5: большая сводная таблица (25 логических колонок = 26 grid) ----------
# Соответствие grid-колонок таблицы 1 полям записи рабочего места.
# «Сут» (молоко) занимает 2 объединённые grid-колонки c22–c23.
_FACTOR_COLS = [
    (3, "chem"), (4, "biological"), (5, "aerosols"), (6, "noise"),
    (7, "infrasound"), (8, "ultrasound_air"), (9, "vibration_general"),
    (10, "vibration_local"), (11, "em_field"), (12, "ionizing"),
    (13, "microclimate"), (14, "lighting"), (15, "severity"),
    (16, "intensity"), (17, "overall"),
]

_REQS_LABELS = {
    "name": "корхонанинг номи",
    "parent": "юқори турувчи",
    "address": "юридик манзили",
    "product": "асосий маҳсулот",
}
_REQS_CODES = (("stir", "стир"), ("ifut", "ифут"), ("mxbt", "мхбт"), ("mxbt", "mxбt"))


def row_values_6_5(rec: dict) -> list[str]:
    """26 значений grid-строки сводной таблицы для одной записи."""
    f = rec.get("factors", {}) or {}
    b = rec.get("benefits", {}) or {}
    vals = [""] * 26
    vals[0] = rec.get("workplace_no", "")
    # Должность в 6_5 — из «Перечня» (как в эталонах клиента); иначе из карты
    vals[1] = rec.get("position_from_perechen") or rec.get("position", "")
    vals[2] = rec.get("job_code", "")
    for ci, key in _FACTOR_COLS:
        vals[ci] = f.get(key, "-") or "-"
    vals[18] = rec.get("injury_risk", "")
    vals[19] = rec.get("ppe_provided", "")
    vals[20] = b.get("extra_leave", "")
    vals[21] = b.get("reduced_hours", "")
    vals[22] = b.get("milk", "")
    vals[23] = b.get("milk", "")  # «Сут» объединена на 2 grid-колонки
    vals[24] = b.get("therapeutic_food", "")
    vals[25] = rec.get("privileged_pension", "")
    return vals


def _norm(s: str) -> str:
    return " ".join((s or "").split())


_CODE_CELL_KEYS = {"стир": "stir", "ифут": "ifut", "мхбт": "mxbt", "mxбt": "mxbt"}


def _fill_reqs(table, company: dict) -> None:
    """Подставить реквизиты компании в таблицу 0.

    ВСЕГДА перезаписываем все ячейки значений из источника; если значение не
    извлеклось — ОЧИЩАЕМ ячейку (пусто), чтобы данные компании-примера из
    шаблона (СТИР/ИФУТ/МХБТ Бухоро) не «протекали» к другому клиенту.
    """
    for row in table.rows:
        cells = row.cells
        c0 = cells[0].text
        # Текстовые поля «метка | значение»: значение в объединённой ячейке cells[2]
        for key, anchor in _REQS_LABELS.items():
            if fold_contains(c0, anchor) and len(cells) > 2:
                set_cell_text(cells[2], company.get(key, "") or "")
        # Строка кодов: ячейка-метка СТИР/ИФУТ/МХБТ → следующая ячейка
        for i, cell in enumerate(cells):
            fc = fold(cell.text)
            for cyr, key in _CODE_CELL_KEYS.items():
                if fc == fold(cyr) and i + 1 < len(cells):
                    set_cell_text(cells[i + 1], company.get(key, "") or "")


def _group_by_subdivision(workplaces: list[dict]) -> list[tuple[str, list[dict]]]:
    """Сгруппировать соседние рабочие места по подразделению (как в эталоне)."""
    groups: list[tuple[str, list[dict]]] = []
    for wp in workplaces:
        sub = wp.get("subdivision", "") or "—"
        if not groups or groups[-1][0] != sub:
            groups.append((sub, []))
        groups[-1][1].append(wp)
    return groups


def render_6_5(company_data: dict, workplaces: list[dict], out_path: str | Path,
               *, template_path: str | Path = TEMPLATE_6_5, lang: str = "cyr") -> Path:
    """Сформировать сводный документ 6_5 из датасета.

    Прототипы строк:
    * R3 — строка-заголовок группы (одна объединённая на всю ширину ячейка);
    * R2 — строка нумерации «1|2|…|25»: у неё объединение «Сут» стоит на ВЕРНОЙ
      позиции (grid-колонки 22–23). У пустых строк данных шаблона (R4) объединение
      смещено на 21–22, из-за чего «сокр.день» и «Сут» схлопывались в одну ячейку
      (баг проявлялся, когда их значения различались). Поэтому строки данных
      клонируем из R2 и перезаписываем числа-плейсхолдеры реальными значениями.
    """
    doc = Document(str(template_path))
    _fill_reqs(doc.tables[0], company_data)
    summary = doc.tables[1]

    # Гарантируем порядок РМ (а-суффикс сразу после базового), независимо от входа
    workplaces = sorted(workplaces, key=lambda w: workplace_sort_key(w.get("workplace_no", "")))

    rows = list(summary.rows)
    group_proto = deepcopy(rows[3]._tr)  # заголовок группы (спанящая ячейка)
    data_proto = deepcopy(rows[2]._tr)   # строка данных (объединение «Сут» на 22–23)
    for row in rows[3:]:                 # очищаем тело ниже 3-уровневой шапки
        summary._tbl.remove(row._tr)

    for sub, members in _group_by_subdivision(workplaces):
        header = _append_row(summary, group_proto)
        set_cell_text(header.cells[0], sub)
        for wp in members:
            row = _append_row(summary, data_proto)
            cells = row.cells
            for ci, val in enumerate(row_values_6_5(wp)):
                if ci < len(cells):
                    set_cell_text(cells[ci], val)

    _transliterate_doc(doc, lang)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


# --- 6_4: сводная қайднома (итоги по подразделениям) ------------------------
# Таблица 1 шаблона (13 grid-колонок):
#   c0  — название строки/подразделения
#   c1  — итого (аттестовано рабочих мест / занято людей / из них женщин)
#   c2..c5  — по классу условий труда: 1, 2, 3, 4
#   c6..c8  — по классу травмоопасности: 1, 2, 3
#   c9..c11 — обеспеченность ЯТҲВ: Мос / Мос эмас / Ятҳв кўзда тутилмаган
#   c12 — попадает под «3-4 класс и/или ЯТҲВ не соответствует»
#
# Каждый блок (итог по компании, итог по подразделению) — 4 строки: строка-
# заголовок (название) + 3 строки данных (иш ўринлари / ходимлар / аёллар).
_ROW_KINDS = ("units", "employees", "female")


def _to_int(value) -> int:
    m = re.search(r"\d+", str(value or ""))
    return int(m.group(0)) if m else 0


def _row_weight(wp: dict, row_kind: str) -> int:
    """«Вес» одного рабочего места в данной строке-категории."""
    if row_kind == "units":
        return 1
    if row_kind == "employees":
        return _to_int(wp.get("employees_count"))
    return _to_int(wp.get("female_count"))  # "female"


def _overall_degree(wp: dict) -> str:
    """Старшая цифра общего класса («3.2» → «3»); "" если класс не извлечён."""
    overall = (wp.get("factors") or {}).get("overall", "") or ""
    return overall[0] if overall[:1] in ("1", "2", "3", "4") else ""


def _aggregate_group(workplaces: list[dict]) -> dict[str, list[int]]:
    """3 строки (units/employees/female) × 13 grid-значений (c0 не считаем)."""
    out: dict[str, list[int]] = {}
    for row_kind in _ROW_KINDS:
        vals = [0] * 13
        for wp in workplaces:
            w = _row_weight(wp, row_kind)
            if not w:
                continue
            degree = _overall_degree(wp)
            risk = wp.get("injury_risk", "")
            ppe = wp.get("ppe_provided", "")

            vals[1] += w
            if degree == "1":
                vals[2] += w
            elif degree == "2":
                vals[3] += w
            elif degree == "3":
                vals[4] += w
            elif degree == "4":
                vals[5] += w
            if risk == "1":
                vals[6] += w
            elif risk == "2":
                vals[7] += w
            elif risk == "3":
                vals[8] += w
            if ppe == "ҳа":
                vals[9] += w
            elif ppe == "йўқ":
                vals[10] += w
            else:
                vals[11] += w
            if degree in ("3", "4") or ppe == "йўқ":
                vals[12] += w
        out[row_kind] = vals
    return out


def _group_by_subdivision_full(workplaces: list[dict]) -> list[tuple[str, list[dict]]]:
    """Группировка ПО ВСЕМ рабочим местам подразделения (не только соседним).

    В отличие от ``_group_by_subdivision`` (для 6_5, где повтор заголовка
    подразделения между несмежными группами — норма визуального списка),
    6_4 — сводная таблица ИТОГОВ: каждое подразделение должно быть ровно
    одним блоком, иначе итоги по нему разъедутся на две строки.
    """
    order: list[str] = []
    buckets: dict[str, list[dict]] = {}
    for wp in workplaces:
        sub = wp.get("subdivision", "") or "—"
        if sub not in buckets:
            buckets[sub] = []
            order.append(sub)
        buckets[sub].append(wp)
    return [(sub, buckets[sub]) for sub in order]


def _fill_numeric_row(row, vals: list[int], *, unknown: bool = False) -> None:
    """Заполнить строку данных (c1..c12); c0 — метка, не трогаем."""
    cells = row.cells
    for i in range(1, min(13, len(cells))):
        set_cell_text(cells[i], "-" if unknown else str(vals[i]))


def render_6_4(company_data: dict, workplaces: list[dict], out_path: str | Path,
               *, template_path: str | Path = TEMPLATE_6_4, lang: str = "cyr") -> Path:
    """Сформировать сводную қайднома 6_4 (итоги по подразделениям) из датасета.

    Шаблон-ассет несёт только скелет: шапку (R0-2), блок «Корхона бўйича
    жами» (R3 заголовок + R4-6 данные) и ОДИН прототип блока подразделения
    (R10 заголовок + R11-13 данные) — остальные строки шаблона (в т.ч.
    повтор шапки R7-9 и жёстко перечисленные отделы конкретного клиента, из
    чьего архива был собран этот файл) отбрасываются: список подразделений
    и их итоги строятся заново из датасета батча, поэтому документ одинаково
    верно собирается для ЛЮБОГО клиента, а не только для того, чей набор
    отделов сейчас «зашит» в файле шаблона.

    «Шу жумладан аёллар»: если ни для одного рабочего места в батче не
    извлечена гендерная разбивка (см. ``extract._extract_employee_counts``),
    строка «аёллар» заполняется прочерками, а не нулями — «нет данных» и
    «действительно ноль женщин» не одно и то же.
    """
    doc = Document(str(template_path))
    _fill_reqs(doc.tables[0], company_data)
    summary = doc.tables[1]

    workplaces = sorted(workplaces, key=lambda w: workplace_sort_key(w.get("workplace_no", "")))
    has_female_data = any((wp.get("female_count") or "") != "" for wp in workplaces)

    rows = list(summary.rows)
    total_data_rows = rows[4:7]                     # «Корхона бўйича жами»: 3 строки данных
    dept_label_proto = deepcopy(rows[10]._tr)        # прототип строки-заголовка подразделения
    dept_data_protos = [deepcopy(r._tr) for r in rows[11:14]]  # прототип 3 строк данных

    # Итоговый блок по компании — заполняем на месте (заголовок R3 не трогаем)
    totals = _aggregate_group(workplaces)
    for row_obj, kind in zip(total_data_rows, _ROW_KINDS):
        _fill_numeric_row(row_obj, totals[kind], unknown=(kind == "female" and not has_female_data))

    # Всё остальное тело (повтор шапки + захардкоженные отделы шаблона) — долой
    for row_obj in rows[7:]:
        summary._tbl.remove(row_obj._tr)

    for sub, members in _group_by_subdivision_full(workplaces):
        agg = _aggregate_group(members)
        label_tr = deepcopy(dept_label_proto)
        summary._tbl.append(label_tr)
        # Название подразделения — объединённая на всю ширину ячейка (одна и
        # та же Cell под всеми индексами row.cells), достаточно записать раз.
        set_cell_text(summary.rows[-1].cells[0], sub)
        for proto, kind in zip(dept_data_protos, _ROW_KINDS):
            tr = deepcopy(proto)
            summary._tbl.append(tr)
            _fill_numeric_row(summary.rows[-1], agg[kind], unknown=(kind == "female" and not has_female_data))

    _transliterate_doc(doc, lang)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path
