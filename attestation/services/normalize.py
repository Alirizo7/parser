"""Нормализация «грязных» значений из карт.

В исходных .doc много шума от OCR/раскладок: пробелы внутри чисел (``3. 3``,
``8 0``, ``0 , 7``), латиница вперемешку с кириллицей (``yo'q``), разные тире.
Здесь — функции приведения к канону.
"""
from __future__ import annotations

import re

# Прочерк/пустое значение → отсутствие данных
EMPTY_TOKENS = {"", "-", "—", "–", "−", "нет", "yoq"}

# Канонические варианты «да» и «нет» (узб. кириллица)
YES = "ҳа"
NO = "йўқ"

# Что считаем «да»: ҳа/ха/бор/+/да/yes/ha/bor. «xa»/«Xa» (лат) тоже = ha.
_YES_TOKENS = {"ҳа", "ха", "бор", "+", "да", "yes", "ha", "bor", "xa"}
# Что считаем «нет»: йўқ/yo'q/yoq/нет/-/—
_NO_TOKENS = {"йўқ", "йуқ", "йук", "yo'q", "yo`q", "yoq", "yo'q.", "нет", "no", "yuq"}


# Невидимые маркеры направления/нулевой ширины (часто «прилипают» к ҳа/йўқ)
_INVISIBLE = "‎‏​⁠﻿"


def normalize_spaces(value: str) -> str:
    """Сжать пробелы (вкл. неразрывные), убрать невидимые маркеры, обрезать края."""
    if value is None:
        return ""
    s = value.replace("\xa0", " ")
    s = "".join(ch for ch in s if ch not in _INVISIBLE)
    return re.sub(r"\s+", " ", s).strip()


# --- Двуязычие: приведение к одному письму (узб. кириллица → латиница) -------
# Карты бывают на узбекской кириллице (Бухоро) и латинице (RIZOYEV). Чтобы ОДИН
# набор якорей (в mapping.py — кириллицей) работал для обоих писем, и текст
# документа, и якоря приводим к общей «свёрнутой» форме через транслитерацию.
_CYR2LAT = {
    "ё": "yo", "ю": "yu", "я": "ya", "ў": "o", "қ": "q", "ғ": "g", "ҳ": "h",
    "ч": "ch", "ш": "sh", "ъ": "", "ь": "", "э": "e", "е": "e",
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ж": "j", "з": "z",
    "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "x",
    "ц": "ts", "ы": "i",
}
# Все варианты апострофа (oʻ/gʻ/tutuq) — убираем, чтобы o`≡ў, g`≡ғ
_APOSTROPHES = "`'’ʻ‘ʼ´ʹ"


def fold(value: str) -> str:
    """Свернуть текст к общей форме для двуязычного сопоставления.

    Узб. кириллицу транслитерируем в латиницу, латиницу оставляем как есть,
    убираем апострофы и прочую пунктуацию, приводим к нижнему регистру.
    Тогда ``fold('баҳолаш') == fold('baholash')`` и т.п.
    """
    s = (value or "").lower().replace("\xa0", " ")
    s = "".join(_CYR2LAT.get(ch, ch) for ch in s)
    s = "".join(ch for ch in s if ch not in _APOSTROPHES)
    s = re.sub(r"[^0-9a-z ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# --- Полная транслитерация для ВЫВОДА документов (узб. кир ↔ лат) ------------
# Используется при рендере, когда выбран язык вывода. В отличие от fold() —
# сохраняет апострофы (oʻ/gʻ/ъ) и регистр, даёт читаемый текст.
_APOS = "'"  # стандартный узбекский апостроф для oʻ/gʻ/tutuq (ASCII — совместимо)

# Кириллица → латиница (нижний регистр; «е» обрабатывается позиционно отдельно)
_TO_LAT = {
    "ё": "yo", "ю": "yu", "я": "ya", "ў": "o" + _APOS, "ғ": "g" + _APOS,
    "қ": "q", "ҳ": "h", "ч": "ch", "ш": "sh", "щ": "sh", "ъ": _APOS, "ь": "",
    "э": "e", "ы": "i", "ц": "ts",
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "ж": "j", "з": "z",
    "и": "i", "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o",
    "п": "p", "р": "r", "с": "s", "т": "t", "у": "u", "ф": "f", "х": "x",
}
_VOWELS_CYR = "аеёиоуўэюяы"


def _apply_case(src_ch: str, mapped: str) -> str:
    """Перенести регистр исходной буквы на результат транслитерации."""
    if src_ch.isupper() and mapped:
        return mapped[0].upper() + mapped[1:]
    return mapped


def to_latin(text: str) -> str:
    """Узб. кириллица → латиница. Цифры/коды/пунктуацию не трогает."""
    out: list[str] = []
    for i, ch in enumerate(text or ""):
        low = ch.lower()
        if low == "е":  # позиционно: «ye» в начале слова/после гласной, иначе «e»
            prev = text[i - 1] if i > 0 else ""
            mapped = "ye" if (not prev.isalpha() or prev.lower() in _VOWELS_CYR) else "e"
            out.append(_apply_case(ch, mapped))
        elif low in _TO_LAT:
            out.append(_apply_case(ch, _TO_LAT[low]))
        else:
            out.append(ch)
    return "".join(out)


# Латиница → кириллица: диграфы (длинное совпадение) раньше одиночных
_TO_CYR_DIGRAPHS = [
    ("o'", "ў"), ("o`", "ў"), ("oʻ", "ў"), ("g'", "ғ"), ("g`", "ғ"), ("gʻ", "ғ"),
    ("yo", "ё"), ("yu", "ю"), ("ya", "я"), ("ye", "е"), ("ch", "ч"),
    ("sh", "ш"), ("ts", "ц"),
]
_TO_CYR_SINGLE = {
    "a": "а", "b": "б", "v": "в", "g": "г", "d": "д", "e": "е", "z": "з",
    "i": "и", "j": "ж", "k": "к", "l": "л", "m": "м", "n": "н", "o": "о",
    "p": "п", "r": "р", "s": "с", "t": "т", "u": "у", "f": "ф", "x": "х",
    "y": "й", "q": "қ", "h": "ҳ", "c": "к", "w": "в",
}


def to_cyrillic(text: str) -> str:
    """Узб. латиница → кириллица. Цифры/коды/пунктуацию не трогает."""
    s = text or ""
    out: list[str] = []
    i = 0
    while i < len(s):
        ch = s[i]
        two = s[i:i + 2].lower()
        matched = False
        for lat, cyr in _TO_CYR_DIGRAPHS:
            if two == lat.lower():
                # «yo'» — это y + oʻ (й+ў), а не ё: пропускаем диграф перед апострофом
                if lat == "yo" and s[i + 2:i + 3] in _APOSTROPHES:
                    break
                out.append(cyr.upper() if ch.isupper() else cyr)
                i += 2
                matched = True
                break
        if matched:
            continue
        low = ch.lower()
        if low in _TO_CYR_SINGLE:
            cyr = _TO_CYR_SINGLE[low]
            out.append(cyr.upper() if ch.isupper() else cyr)
        elif ch in _APOSTROPHES:
            pass  # tutuq между буквами — опускаем (после oʻ/gʻ уже обработан)
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def fold_contains(text: str, anchor: str) -> bool:
    """Подстрочное вхождение с учётом письма (через ``fold``)."""
    return fold(anchor) in fold(text)


def fold_contains_all(text: str, tokens) -> bool:
    """Все токены присутствуют в тексте (как подстроки) после ``fold``."""
    ft = fold(text)
    return all(fold(t) in ft for t in tokens)


def is_empty(value: str) -> bool:
    """Значение трактуется как «отсутствует»?"""
    return normalize_spaces(value).lower() in EMPTY_TOKENS


def normalize_class(value: str) -> str:
    """Привести класс условий труда к виду ``X.Y`` (``3. 3`` → ``3.3``).

    Возвращает ``-``, если класс отсутствует.
    """
    s = normalize_spaces(value)
    if is_empty(s):
        return "-"
    # Класс вида «3.3» / «3, 3» / «3. 1» (возможны пробелы вокруг разделителя)
    m = re.search(r"([1-4])\s*[.,]\s*([0-9])", s)
    if m:
        return f"{m.group(1)}.{m.group(2)}"
    # Голое целое («2», «3») — приводим к «2.0»
    m = re.fullmatch(r"([1-4])", s)
    if m:
        return f"{m.group(1)}.0"
    return s


def class_rank(value: str) -> float:
    """Числовой ранг класса для сравнения (чем хуже условия — тем больше)."""
    s = normalize_class(value)
    m = re.fullmatch(r"([1-4])\.([0-9])", s)
    if not m:
        return -1.0
    return int(m.group(1)) * 10 + int(m.group(2))


def max_class(values) -> str:
    """Максимальный (худший) класс среди значений; ``-`` если нет ни одного."""
    best, best_rank = "-", -1.0
    for v in values:
        r = class_rank(v)
        if r > best_rank:
            best, best_rank = normalize_class(v), r
    return best


def normalize_number(value: str) -> str:
    """Убрать лишние пробелы внутри числа: ``8 0`` → ``80``, ``0 , 7`` → ``0,7``."""
    s = normalize_spaces(value)
    # Склеиваем пробелы между цифрами и вокруг разделителей
    s = re.sub(r"(?<=\d)\s+(?=\d)", "", s)
    s = re.sub(r"\s*([.,])\s*", r"\1", s)
    return s


def canon_yesno(value: str) -> str:
    """Привести значение к канону ``ҳа`` / ``йўқ``.

    ``бор``/``+``/``ha`` → ``ҳа``; ``yo'q``/``-``/пусто → ``йўқ``.
    """
    s = normalize_spaces(value).lower().rstrip(".")
    if not s:
        return NO
    if s in _YES_TOKENS:
        return YES
    if s in _NO_TOKENS:
        return NO
    # Иногда в ячейке текст вида «ҳа, 3-илова …» / «ha 4-6 kun» — ведущее слово.
    # «йўқ» проверяем раньше «yo»-да, чтобы не спутать.
    if s.startswith(("йў", "йу", "yo'q", "yo`q", "yoq", "yo q")):
        return NO
    if s.startswith(("ҳа", "ха", "бор", "ha", "xa", "bor", "да")):
        return YES
    return NO


# Единицы измерения, которые срезаем из названия вещества (кириллица + латиница)
_UNIT_RE = re.compile(
    r"\s*,?\s*("
    r"мг\s*/\s*м\s*3|мг\s*/\s*м³|мг\s*/\s*куб\.?\s*м|мкг\s*/\s*м3|мг\s*/\s*куб|г\s*/\s*м3|"
    r"mg\s*/\s*m\s*3|mg\s*/\s*m³|mkg\s*/\s*m3|mg\s*/\s*kub\.?\s*m|g\s*/\s*m3"
    r")\s*\.?$",
    re.IGNORECASE,
)


def clean_substance_name(value: str) -> str:
    """Очистить название вещества: убрать единицы (``мг/м3``/``mg/m3``), лишние пробелы."""
    s = normalize_spaces(value)
    s = _UNIT_RE.sub("", s)
    return s.strip(" .;")
