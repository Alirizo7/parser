"""Распаковка zip-архива клиента.

Особенности:
* кириллические/узбекские имена внутри zip иногда хранятся без флага UTF-8 —
  декодируем robust-методом (cp437 → cp866/cp1251), сохраняя корректные имена;
* мусорные файлы (``~$...``, ``~WRL*.tmp``, ``.tmp``) игнорируем.
"""
from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path

# Расширения/префиксы, которые считаем мусором
# ``._*`` и ``.DS_Store`` — служебные файлы macOS (AppleDouble), ``~$``/``~WRL`` — Word
_JUNK_PREFIXES = ("~$", "~wrl", "~wrd", ".~", "._")
_JUNK_SUFFIXES = (".tmp", ".ds_store")


@dataclass
class UnpackedFile:
    """Один извлечённый файл."""

    arc_name: str        # имя внутри архива (с подпапками), уже в корректной кодировке
    abs_path: Path       # путь на диске после распаковки

    @property
    def basename(self) -> str:
        return self.arc_name.replace("\\", "/").split("/")[-1]

    @property
    def suffix(self) -> str:
        return Path(self.basename).suffix.lower()


def _decode_name(info: zipfile.ZipInfo) -> str:
    """Получить корректное имя файла из ZipInfo.

    Случаи в реальных архивах:
    * флаг UTF-8 (0x800) установлен — zipfile уже декодировал имя правильно;
    * флаг не установлен, но байты на самом деле UTF-8/cp866/cp1251 — zipfile
      декодирует их как cp437 (получается «мусор»), нужно восстановить.

    Хитрость: если имя содержит кириллицу, его НЕЛЬЗЯ закодировать в cp437 —
    значит оно уже корректно (некоторые архиваторы пишут UTF-8 без флага, а
    zipfile отдаёт верную строку). Перекодируем только когда cp437-обратная
    кодировка проходит без потерь (errors="strict").
    """
    name = info.filename
    if info.flag_bits & 0x800:
        return name
    try:
        raw = name.encode("cp437")  # строго: если кириллица — бросит исключение
    except UnicodeEncodeError:
        return name  # имя уже в правильной кодировке
    for enc in ("utf-8", "cp866", "cp1251"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return name


def is_junk(basename: str) -> bool:
    """Временный/локовый/служебный файл?"""
    low = basename.lower()
    if low.startswith(_JUNK_PREFIXES):
        return True
    if low.endswith(_JUNK_SUFFIXES):
        return True
    return False


def unpack(zip_path: str | Path, dest_dir: str | Path) -> list[UnpackedFile]:
    """Распаковать архив в ``dest_dir``. Вернуть список значимых файлов."""
    zip_path = Path(zip_path)
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)

    result: list[UnpackedFile] = []
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = _decode_name(info)
            # Папка ресурсов macOS — целиком мусор
            if "__macosx" in name.lower():
                continue
            basename = name.replace("\\", "/").split("/")[-1]
            if not basename or is_junk(basename):
                continue

            # Безопасный путь назначения (защита от обхода каталогов)
            rel = Path(name.replace("\\", "/"))
            parts = [p for p in rel.parts if p not in ("", ".", "..")]
            target = dest_dir.joinpath(*parts)
            target.parent.mkdir(parents=True, exist_ok=True)

            with zf.open(info) as src, open(target, "wb") as out:
                out.write(src.read())

            result.append(UnpackedFile(arc_name=name, abs_path=target))

    return result
