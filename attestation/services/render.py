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

from openpyxl import load_workbook

from . import mapping as M
from . import xlsx
from .extract import split_workplace_no, workplace_sort_key
from .normalize import _to_int, fold, fold_contains, max_class, normalize_spaces, to_latin

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
TEMPLATE_5_1B = ASSETS_DIR / "template_5_1b.docx"
TEMPLATE_6_5 = ASSETS_DIR / "template_6_5.docx"
TEMPLATE_6_4 = ASSETS_DIR / "template_6_4.docx"
TEMPLATE_EXCEL = {n: ASSETS_DIR / f"template_excel_{n}.xlsx" for n in range(1, 6)}


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

# Позиция считается «учтённой в 6_4», только если пайплайн сопоставил её со
# строкой Перечня (см. pipeline._pair_perechen_positions_6_4) — иначе у неё
# нет ключа subdivision_6_4, и включать её в подсчёт нельзя (см. mapping/extract).


def _overall_degree(wp: dict) -> str:
    """Старшая цифра общего класса («3.2» → «3»); "" если класс не извлечён."""
    overall = (wp.get("factors") or {}).get("overall", "") or ""
    return overall[0] if overall[:1] in ("1", "2", "3", "4") else ""


def _aggregate_group_6_4(positions: list[dict], warnings: list[str]) -> dict[str, list[int]]:
    """3 строки (units/employees/female) × 13 grid-значений (c0 не считаем).

    Источник числа работников/женщин — Перечень (``employees_count_6_4``/
    ``female_count_6_4``, см. Шаг 4 спецификации), НЕ карта: карта даёт лишь
    3 итоговых показателя (класс условий труда, класс травмоопасности,
    статус ЯТҲВ) для распределения этих чисел по колонкам.
    """
    out: dict[str, list[int]] = {}
    for row_kind in _ROW_KINDS:
        vals = [0] * 13
        for wp in positions:
            if row_kind == "units":
                # «а»-суффиксные строки Перечня — доп. смена/условие ТОГО ЖЕ
                # рабочего места, не новая единица (сверено с эталоном: сумма
                # строк БЕЗ суффикса == «Иш ўринлари, бирлик» итога компании).
                _, suffix = split_workplace_no(wp.get("workplace_no", ""))
                w = 0 if suffix else 1
            elif row_kind == "employees":
                w = _to_int(wp.get("employees_count_6_4"))
            else:
                w = _to_int(wp.get("female_count_6_4"))
            if not w:
                continue
            degree = _overall_degree(wp)
            risk = wp.get("injury_risk_class_6_4", "")
            ppe = wp.get("ppe_status_6_4", "")

            vals[1] += w
            if degree == "1":
                vals[2] += w
            elif degree == "2":
                vals[3] += w
            elif degree == "3":
                vals[4] += w
            elif degree == "4":
                vals[5] += w
            elif row_kind == "units":
                warnings.append(
                    f"{wp.get('workplace_no', '?')}: класс условий труда не извлечён из карты — "
                    "не учтён в разбивке по классам 6_4."
                )
            if risk == "1":
                vals[6] += w
            elif risk == "2":
                vals[7] += w
            elif risk == "3":
                vals[8] += w
            elif row_kind == "units":
                warnings.append(
                    f"{wp.get('workplace_no', '?')}: класс травмоопасности не извлечён из карты — "
                    "не учтён в разбивке по классам 6_4."
                )
            # c9/c10 — соответствие требованиям ЯТҲВ (п.3.3); c11 — «СИЗ не
            # предусмотрены» (таблица СИЗ п.3.2). Это НЕЗАВИСИМЫЕ величины:
            # рабочее место одновременно «Мос» (c9) и «кўзда тутилмаган» (c11).
            # В эталоне Мос=114 (все) и кўзда=93 (подмножество) — суммы намеренно
            # пересекаются, поэтому НЕ раскладываем по взаимоисключающим колонкам.
            if ppe == "mos_emas":
                vals[10] += w
            else:
                vals[9] += w
            if wp.get("ppe_not_envisaged_6_4"):
                vals[11] += w
            # c12 («3-4 даража ва/ёки Ятҳв мос эмас») — в эталоне ПУСТОЙ по всем
            # подразделениям (автор оставил его незаполненным), поэтому c12 не
            # заполняем (vals[12] остаётся 0 → рендерится «-»).
        out[row_kind] = vals
    return out


def _group_positions_by_subdivision_6_4(positions: list[dict]) -> list[tuple[str, list[dict]]]:
    """Группировка учтённых позиций по подразделению Перечня (одна группа = один блок).

    Ключ группировки — ``fold(subdivision_6_4)`` (не учитывает регистр/письмо/
    ъ-ь опечатки); заголовок группы — сам текст разделительной строки Перечня.
    """
    order: list[str] = []
    buckets: dict[str, list[dict]] = {}
    labels: dict[str, str] = {}
    for wp in positions:
        sub = wp.get("subdivision_6_4", "") or "—"
        key = fold(sub)
        if key not in buckets:
            buckets[key] = []
            labels[key] = sub
            order.append(key)
        buckets[key].append(wp)
    return [(labels[key], buckets[key]) for key in order]


def _fill_numeric_row(row, vals: list[int]) -> None:
    """Заполнить строку данных (c1..c12); c0 — метка, не трогаем.

    Ноль отображаем как «-» (эталон использует прочерк для отсутствующих
    значений в каждой колонке-категории, не «0»).
    """
    cells = row.cells
    for i in range(1, min(13, len(cells))):
        set_cell_text(cells[i], "-" if not vals[i] else str(vals[i]))


def render_6_4(company_data: dict, workplaces: list[dict], out_path: str | Path,
               *, template_path: str | Path = TEMPLATE_6_4, lang: str = "cyr",
               warnings: list[str] | None = None) -> Path:
    """Сформировать сводную қайднома 6_4 (итоги по подразделениям) из датасета.

    Шаблон-ассет несёт готовые блоки подразделений (R10, R14, R18, … по 4
    строки: заголовок + 3 строки данных). Логика на любой архив/«Перечень»:
    * блок шаблона, чьё название совпало (после ``fold``) с подразделением из
      «Перечня», — заполняется НА МЕСТЕ (сохраняя форматирование шаблона);
    * подразделение «Перечня» без блока в шаблоне — добавляется новым блоком
      в конец (сообщение в ``warnings``, данные не теряются);
    * блок шаблона, которого НЕТ в «Перечне» этого архива (заготовка из чужого
      шаблона), — УДАЛЯЕТСЯ, чтобы в итоге остались только нужные подразделения.
    Итог: набор блоков всегда равен набору подразделений «Перечня» — независимо
    от того, под какое учреждение изначально свёрстан шаблон.
    """
    if warnings is None:
        warnings = []
    doc = Document(str(template_path))
    _fill_reqs(doc.tables[0], company_data)
    summary = doc.tables[1]

    positions = [wp for wp in workplaces if "subdivision_6_4" in wp]
    positions.sort(key=lambda w: workplace_sort_key(w.get("workplace_no", "")))

    rows = list(summary.rows)
    total_data_rows = rows[4:7]  # «Корхона бўйича жами»: 3 строки данных

    # Итоговый блок по компании — заполняем на месте (заголовок R3 не трогаем)
    totals = _aggregate_group_6_4(positions, warnings)
    for row_obj, kind in zip(total_data_rows, _ROW_KINDS):
        _fill_numeric_row(row_obj, totals[kind])

    # Блоки подразделений шаблона: R10, R14, R18, … по 4 строки каждый.
    block_starts = list(range(10, len(rows), 4))
    groups = _group_positions_by_subdivision_6_4(positions)
    used: set[int] = set()
    unused_block_rows: list = []  # строки пустых блоков шаблона — на удаление
    # Прототип нового блока берём ДО удаления (deepcopy — не зависит от того,
    # окажется ли исходный блок сматченным или удалённым).
    header_proto = deepcopy(rows[block_starts[0]]._tr) if block_starts else None
    data_protos = [deepcopy(r._tr) for r in rows[block_starts[0] + 1:block_starts[0] + 4]] if block_starts else []
    for start in block_starts:
        if start + 3 >= len(rows):
            break
        label = normalize_spaces(rows[start].cells[0].text)
        match_idx = next(
            (i for i, (name, _) in enumerate(groups) if i not in used and fold(name) == fold(label)),
            None,
        )
        if match_idx is None:
            # Такого подразделения нет в «Перечне» этого архива → блок на удаление.
            unused_block_rows.extend(rows[start:start + 4])
            continue
        used.add(match_idx)
        agg = _aggregate_group_6_4(groups[match_idx][1], warnings)
        for row_obj, kind in zip(rows[start + 1:start + 4], _ROW_KINDS):
            _fill_numeric_row(row_obj, agg[kind])

    # Подразделения Перечня без блока в шаблоне — добавляем блок в конец,
    # а не пропускаем молча (Шаг 2/7 спецификации).
    if header_proto is not None:
        for i, (name, members) in enumerate(groups):
            if i in used:
                continue
            warnings.append(
                f"Подразделение «{name}» из Перечня не найдено ни в одном блоке шаблона 6_4 — добавлен новый блок."
            )
            summary._tbl.append(deepcopy(header_proto))
            set_cell_text(summary.rows[-1].cells[0], name)
            agg = _aggregate_group_6_4(members, warnings)
            for proto, kind in zip(data_protos, _ROW_KINDS):
                summary._tbl.append(deepcopy(proto))
                _fill_numeric_row(summary.rows[-1], agg[kind])

    # Удаляем пустые блоки чужого шаблона (подразделения, которых нет в «Перечне»
    # этого архива) — делаем это ПОСЛЕ добавления новых, чтобы не сбить индексы.
    for row_obj in unused_block_rows:
        summary._tbl.remove(row_obj._tr)

    _transliterate_doc(doc, lang)
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


# ===========================================================================
# Excel-протоколы лабораторных замеров (5 файлов) — render_excel_1..5
# ===========================================================================
# Тело каждого протокола генерируется из единого датасета клонированием прото-
# блока рабочего места (см. services/xlsx.py). Управляемая часть ассета держится
# на письме эталона (латиница); в конце приводится к output_lang «к целевому
# письму» (xlsx.transliterate_region). Статичная шапка (приборы/НД/лаборатория) —
# «как есть», меняется только заказчик (company_data).

# Геометрия ассетов (совпадает с build_excel_assets.CFG). ncols — число колонок.
EXCEL_CFG = {
    1: dict(ncols=12, managed_start=27, group_row=30, block_start=31, block_len=1),
    2: dict(ncols=11, managed_start=27, group_row=30, block_start=31, block_len=4),
    3: dict(ncols=11, managed_start=24, group_row=27, block_start=28, block_len=9),
    4: dict(ncols=11, managed_start=24, group_row=27, block_start=28, block_len=5),
    5: dict(ncols=11, managed_start=24, group_row=None, block_start=28, block_len=16),
}

# Метки строки-заказчика в шапке (двуязычные: файл 1 — узб. латиница, 2–5 — рус.)
_EXCEL_REQS_NAME = ("buyurtmachi nomi", "наименование заказчик")
_EXCEL_REQS_ADDR = ("buyurtmachi manzil", "адрес заказчик")


def _fill_excel_reqs(ws, company: dict) -> None:
    """Вписать имя/адрес заказчика в шапку (перезаписать метку-строку).

    Значение идёт ИНЛАЙН после метки в той же ячейке («Наименование заказчика:
    <имя>»); длинное имя в эталоне переносится на следующую строку — её очищаем,
    чтобы имя компании-образца не «протекло» (анти-утечка, как в docx _fill_reqs).
    """
    company = company or {}
    name_row = None
    for r in range(1, EXCEL_HEADER_SCAN + 1):
        cell = ws.cell(row=r, column=1)
        v = cell.value
        if not isinstance(v, str):
            continue
        if any(fold_contains(v, a) for a in _EXCEL_REQS_NAME):
            cell.value = _reqs_prefix(v) + (company.get("name", "") or "")
            name_row = r
        elif any(fold_contains(v, a) for a in _EXCEL_REQS_ADDR):
            cell.value = _reqs_prefix(v) + (company.get("address", "") or "")
    # Очистить строку-продолжение длинного имени образца (анти-утечка), но НЕ
    # трогать её, если это уже другая метка (адрес и т.п. — имя было в одну строку).
    if name_row is not None:
        cont = ws.cell(row=name_row + 1, column=1)
        cv = cont.value
        is_label = isinstance(cv, str) and any(
            fold_contains(cv, a) for a in (*_EXCEL_REQS_NAME, *_EXCEL_REQS_ADDR)
        )
        if not is_label:
            cont.value = None


EXCEL_HEADER_SCAN = 26  # в пределах скольких строк искать метки заказчика


def _reqs_prefix(text: str) -> str:
    """Префикс метки до двоеточия включительно («Наименование заказчика: »)."""
    return (text.split(":", 1)[0] + ": ") if ":" in text else (text + " ")


# --- Преобразование значений замеров в ячейки -------------------------------
def _d_norma(value):
    """Норма в колонку D: число / диапазон-строка / «-»; пусто → None."""
    if value is None:
        return None
    s = str(value).strip()
    if s == "":
        return None
    if s == "-":
        return "-"
    return xlsx.to_number(s)


def _cls_cell(value):
    """Класс в ячейку: число (2.0→2, 3.1→3.1); «-»/пусто → None."""
    if value is None:
        return None
    s = str(value).strip()
    if s in ("", "-"):
        return None
    return xlsx.to_number(s)


# Колонки D..J (норма, 1/2/3-фаолият, факт, время, класс) для файлов 2–5
_DJ_COLS = range(4, 11)


def _measure_fill(m: dict | None, *, with_class: bool = True) -> dict:
    """Заполнение колонок D..J под-строки по замеру {norma,actual,time,cls}.

    * замера нет → пустая рамка (D..J очищены);
    * факт-текст «йўқ» → псевдозамеры E/F/G и факт H = тот же текст, время/класс пусты;
    * числовой факт → норма в D, факт в H, формулы E/F/G из прото (KEEP), время/класс.

    ``with_class=False`` — не выводить класс (J): в некоторых под-строках эталон
    его не сообщает (напр. естественная освещённость КЕО в файле 4).
    """
    if not m:
        return {c: xlsx.CLEAR for c in _DJ_COLS}
    actual = xlsx.to_number(m.get("actual"))
    if actual is None:
        return {c: xlsx.CLEAR for c in _DJ_COLS}
    d = _d_norma(m.get("norma"))
    if isinstance(actual, str):  # факт — текст «йўқ»
        return {4: d, 5: actual, 6: actual, 7: actual, 8: actual, 9: xlsx.CLEAR, 10: xlsx.CLEAR}
    cls = _cls_cell(m.get("cls")) if with_class else xlsx.CLEAR
    return {4: d, 5: xlsx.KEEP, 6: xlsx.KEEP, 7: xlsx.KEEP, 8: actual,
            9: xlsx.to_number(m.get("time")), 10: cls}


def _label_d_fill(text: str) -> dict:
    """Под-строка «метка в D, остальное пусто» (Ishlar toifasi / Vizual ish / энергия)."""
    fill = {c: xlsx.CLEAR for c in _DJ_COLS}
    if text:
        fill[4] = text
    return fill


# --- Построители под-строк блока (файлы 2–5) --------------------------------
def _subrows_file2(wp: dict) -> list[dict]:
    pm = wp.get("physical_measurements") or {}
    return [
        _measure_fill(pm.get("noise")),
        _measure_fill(pm.get("vibration_local")),
        _measure_fill(pm.get("vibration_general")),
        _measure_fill(pm.get("infrasound")),
    ]


def _subrows_file3(wp: dict) -> list[dict]:
    mc = wp.get("microclimate_measurements") or {}
    return [
        _label_d_fill(""),                       # 0 энергозатраты (только K=итог)
        _label_d_fill(mc.get("category_label") or ""),  # 1 Ishlar toifasi
        _measure_fill(mc.get("temp")),           # 2 температура
        _measure_fill(mc.get("air_speed")),      # 3 скорость воздуха
        _measure_fill(mc.get("humidity")),       # 4 влажность
        _measure_fill(mc.get("heat_radiation")), # 5 теплоизлучение («йўқ»)
        _label_d_fill(""),                       # 6 температура (откр. территория)
        _label_d_fill(""),                       # 7 WBGT
        _label_d_fill(""),                       # 8 средняя тепл. нагрузка TNS
    ]


def _subrows_file4(wp: dict) -> list[dict]:
    lg = wp.get("lighting_measurements") or {}
    return [
        _label_d_fill(lg.get("discharge") or ""),          # 0 разряд зрит. работ
        _measure_fill(lg.get("natural"), with_class=False),  # 1 естественная (КЕО): без класса
        _measure_fill(lg.get("combined")),                 # 2 смешанная (КЕО)
        _measure_fill(lg.get("artificial")),               # 3 искусственная (лк)
        # 4 пульсация: эталон систематически НЕ сообщает её в этом протоколе
        # (0/52 РМ), хотя в картах раздел 1.11.6 заполнен — оставляем пустой рамкой,
        # чтобы вывод совпал с эталоном по данным.
        _measure_fill(None),
    ]


# Файл 5: 16 под-строк; данные конкретной 1.4.x лежат в верхней строке пары.
# ⚠️ Прото-формулы псевдозамеров (E/F/G = H±0.01) есть в эталоне ТОЛЬКО на строке
# 1.4.10 (единственной заполненной у медиков). Если у клиента заполнен другой
# раздел (1.4.1–1.4.9), его норма/факт/время/класс выводятся корректно, но
# декоративные колонки «1/2/3-фаолият» останутся пустыми (в прото-строке формулы
# нет). Данные при этом не теряются; при появлении таких клиентов формулы можно
# до-синтезировать. То же касается пульсации в файле 4.
_EM_SUBROW_SECTION = {
    0: "1.4.1", 1: "1.4.2", 2: "1.4.3", 3: "1.4.4", 4: "1.4.5",
    6: "1.4.6", 8: "1.4.7", 10: "1.4.8", 12: "1.4.9", 14: "1.4.10",
}


def _subrows_file5(wp: dict) -> list[dict]:
    em = wp.get("em_measurements") or {}
    fills = []
    for i in range(16):
        sec = _EM_SUBROW_SECTION.get(i)
        fills.append(_measure_fill(em.get(sec)) if sec else _label_d_fill(""))
    return fills


# --- Итоговый класс (K/L) по файлу ------------------------------------------
def _factor(wp: dict, key: str) -> str:
    return (wp.get("factors") or {}).get(key, "-") or "-"


def _final_file1(wp):
    return _cls_cell(_factor(wp, "chem"))


def _final_file2(wp):
    phys = [_factor(wp, k) for k in
            ("noise", "infrasound", "ultrasound_air", "vibration_general", "vibration_local")]
    return _cls_cell(max_class(phys))


def _final_file3(wp):
    return _cls_cell(_factor(wp, "microclimate"))


def _final_file4(wp):
    return _cls_cell(_factor(wp, "lighting"))


def _final_file5(wp):
    return _cls_cell(_factor(wp, "em_field"))


def _has_em(wp: dict) -> bool:
    return any(v for v in (wp.get("em_measurements") or {}).values())


# --- Общий рендер файлов 2–5 (фикс. блок из N под-строк) --------------------
def _render_excel_blocks(company_data, workplaces, out_path, *, idx, subrows_of,
                         final_of, include, grouped, lang):
    cfg = EXCEL_CFG[idx]
    ncols, gr, bs, blen = cfg["ncols"], cfg["group_row"], cfg["block_start"], cfg["block_len"]
    wb = load_workbook(str(TEMPLATE_EXCEL[idx]))
    ws = wb["complete"]

    # Снять прототипы ДО очистки тела
    group_proto = xlsx.capture_row(ws, gr, ncols) if gr else None
    group_merges = xlsx.capture_merges(ws, gr, gr) if gr else []
    block_proto = [xlsx.capture_row(ws, bs + i, ncols) for i in range(blen)]
    block_merges = xlsx.capture_merges(ws, bs, bs + blen - 1)

    body_start = gr if gr else bs
    xlsx.clear_body(ws, body_start)
    _fill_excel_reqs(ws, company_data)

    wps = [w for w in sorted(workplaces, key=lambda w: workplace_sort_key(w.get("workplace_no", "")))
           if include(w)]

    cur = body_start

    def emit_block(wp):
        nonlocal cur
        fills = subrows_of(wp)
        start = cur
        for i in range(blen):
            f = dict(fills[i])
            if i == 0:  # № РМ, должность, итоговый класс — только 1-я строка (merge)
                f[1] = wp.get("workplace_no", "")
                f[2] = wp.get("position_from_perechen") or wp.get("position", "")
                f[ncols] = final_of(wp)
            xlsx.emit_row(ws, cur, block_proto[i], ncols, f)
            cur += 1
        xlsx.apply_merges(ws, block_merges, start)  # A/B/K спаны + внутренние объединения

    if grouped:
        for sub, members in _group_by_subdivision(wps):
            xlsx.emit_row(ws, cur, group_proto, ncols, {1: sub})
            grow = cur
            cur += 1
            xlsx.apply_merges(ws, group_merges, grow)
            for wp in members:
                emit_block(wp)
    else:
        for wp in wps:
            emit_block(wp)

    xlsx.transliterate_region(ws, lang, cfg["managed_start"])
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def render_excel_1(company_data, workplaces, out_path, *,
                   template_path=None, lang: str = "cyr") -> Path:
    """Файл 1 — вредные вещества (одна строка на вещество, блок переменной высоты)."""
    cfg = EXCEL_CFG[1]
    ncols, gr, bs = cfg["ncols"], cfg["group_row"], cfg["block_start"]
    wb = load_workbook(str(TEMPLATE_EXCEL[1]))
    ws = wb["complete"]

    group_proto = xlsx.capture_row(ws, gr, ncols)
    group_merges = xlsx.capture_merges(ws, gr, gr)      # A30:L30
    sub_proto = xlsx.capture_row(ws, bs, ncols)         # одна прото-строка вещества

    xlsx.clear_body(ws, gr)
    _fill_excel_reqs(ws, company_data)

    wps = sorted(workplaces, key=lambda w: workplace_sort_key(w.get("workplace_no", "")))
    cur = gr

    for sub, members in _group_by_subdivision(wps):
        xlsx.emit_row(ws, cur, group_proto, ncols, {1: sub})
        grow = cur
        cur += 1
        xlsx.apply_merges(ws, group_merges, grow)
        for wp in members:
            subs = wp.get("substances") or []
            position = wp.get("position_from_perechen") or wp.get("position", "")
            final = _final_file1(wp)
            start = cur
            if not subs:  # РМ без веществ — одна пустая строка (не теряем место)
                xlsx.emit_row(ws, cur, sub_proto, ncols, {
                    1: wp.get("workplace_no", ""), 2: position, 3: xlsx.CLEAR,
                    4: xlsx.CLEAR, 5: xlsx.CLEAR, 6: xlsx.CLEAR, 7: xlsx.CLEAR,
                    8: xlsx.CLEAR, 9: xlsx.CLEAR, 10: xlsx.CLEAR, 11: xlsx.CLEAR, 12: final,
                })
                cur += 1
            for si, s in enumerate(subs):
                actual = xlsx.to_number(s.get("actual"))  # I — факт (число или текст)
                if isinstance(actual, str):
                    # Факт — текст («йўқ» и т.п.): НЕ оставляем формулы «=I±шаг» на
                    # текстовой ячейке (иначе Excel даст #VALUE!) — как _measure_fill
                    # для файлов 2–5: псевдозамеры F/G/H = тот же текст.
                    pseudo = {6: actual, 7: actual, 8: actual}
                else:
                    pseudo = {6: xlsx.KEEP, 7: xlsx.KEEP, 8: xlsx.KEEP}  # формулы =I±шаг
                f = {
                    3: s.get("name", ""),           # C — вещество
                    4: xlsx.CLEAR,                  # D — класс опасности (нет в карте)
                    5: _d_norma(s.get("norma")),    # E — ПДК/норма
                    **pseudo,                       # F/G/H
                    9: actual,                      # I — факт
                    10: xlsx.to_number(s.get("pct")),  # J — время
                    11: _cls_cell(s.get("cls")),       # K — класс вещества
                }
                if si == 0:
                    f[1] = wp.get("workplace_no", "")
                    f[2] = position
                    f[12] = final              # L — итог по РМ
                else:
                    f[1] = xlsx.CLEAR
                    f[2] = xlsx.CLEAR
                    f[12] = xlsx.CLEAR
                xlsx.emit_row(ws, cur, sub_proto, ncols, f)
                cur += 1
            # Вертикальные объединения № РМ / должность / итог по высоте блока
            xlsx.merge_span(ws, start, cur - 1, 1)
            xlsx.merge_span(ws, start, cur - 1, 2)
            xlsx.merge_span(ws, start, cur - 1, 12)

    xlsx.transliterate_region(ws, lang, cfg["managed_start"])
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def render_excel_2(company_data, workplaces, out_path, *, template_path=None, lang="cyr") -> Path:
    """Файл 2 — физические факторы (шум/вибрация/инфразвук)."""
    return _render_excel_blocks(company_data, workplaces, out_path, idx=2,
                                subrows_of=_subrows_file2, final_of=_final_file2,
                                include=lambda w: True, grouped=True, lang=lang)


def render_excel_3(company_data, workplaces, out_path, *, template_path=None, lang="cyr") -> Path:
    """Файл 3 — микроклимат."""
    return _render_excel_blocks(company_data, workplaces, out_path, idx=3,
                                subrows_of=_subrows_file3, final_of=_final_file3,
                                include=lambda w: True, grouped=True, lang=lang)


def render_excel_4(company_data, workplaces, out_path, *, template_path=None, lang="cyr") -> Path:
    """Файл 4 — освещённость."""
    return _render_excel_blocks(company_data, workplaces, out_path, idx=4,
                                subrows_of=_subrows_file4, final_of=_final_file4,
                                include=lambda w: True, grouped=True, lang=lang)


def render_excel_5(company_data, workplaces, out_path, *, template_path=None, lang="cyr") -> Path:
    """Файл 5 — магнитные поля / ЭМИ (без группировки; только РМ с замером 1.4.x)."""
    return _render_excel_blocks(company_data, workplaces, out_path, idx=5,
                                subrows_of=_subrows_file5, final_of=_final_file5,
                                include=_has_em, grouped=False, lang=lang)
