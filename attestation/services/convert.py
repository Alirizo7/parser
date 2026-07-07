"""Конвертация документов через LibreOffice (headless).

``python-docx`` НЕ читает старый бинарный ``.doc`` — поэтому карты и шаблоны
сначала конвертируем в ``.docx`` через ``soffice``.

Грабли (учтено):
* ``soffice`` не любит параллельные запуски с общим профилем — на каждый вызов
  даём уникальный ``-env:UserInstallation`` и сериализуем вызовы глобальным
  блокировщиком (фактически concurrency=1 на процесс воркера);
* бинарь ищем по настройке ``SOFFICE_BIN`` → в PATH → стандартный путь macOS.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import threading
import uuid
from pathlib import Path

# Сериализация вызовов soffice внутри процесса
_SOFFICE_LOCK = threading.Lock()

# Кандидаты на бинарь LibreOffice (по убыванию приоритета)
_CANDIDATES = (
    "soffice",
    "libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
)


class ConversionError(RuntimeError):
    """Ошибка конвертации документа."""


def find_soffice() -> str:
    """Найти исполняемый файл LibreOffice."""
    env = os.environ.get("SOFFICE_BIN")
    if env:
        return env
    for cand in _CANDIDATES:
        path = shutil.which(cand) or (cand if Path(cand).exists() else None)
        if path:
            return path
    raise ConversionError(
        "Не найден LibreOffice (soffice). Укажите путь через переменную SOFFICE_BIN."
    )


def convert_to_docx(src_path: str | Path, out_dir: str | Path, *, timeout: int = 180) -> Path:
    """Сконвертировать ``src_path`` в ``.docx`` и вернуть путь к результату."""
    return _convert(src_path, out_dir, "docx", timeout=timeout)


def convert_to_doc(src_path: str | Path, out_dir: str | Path, *, timeout: int = 180) -> Path:
    """Сконвертировать в старый ``.doc`` (для выгрузки, если требуется)."""
    return _convert(src_path, out_dir, "doc", timeout=timeout)


def _convert(src_path: str | Path, out_dir: str | Path, target: str, *, timeout: int) -> Path:
    src_path = Path(src_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if not src_path.exists():
        raise ConversionError(f"Файл не найден: {src_path}")

    soffice = find_soffice()
    # Уникальный профиль на каждый вызов — защита от конкурентности
    profile = Path("/tmp") / f"lo_{uuid.uuid4().hex}"
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to",
        target,
        "--outdir",
        str(out_dir),
        str(src_path),
    ]

    expected = out_dir / (src_path.stem + f".{target}")
    with _SOFFICE_LOCK:
        try:
            proc = subprocess.run(
                cmd, capture_output=True, text=True, timeout=timeout, check=False
            )
        finally:
            shutil.rmtree(profile, ignore_errors=True)

    if proc.returncode != 0 or not expected.exists():
        raise ConversionError(
            f"Не удалось сконвертировать {src_path.name} → {target}.\n"
            f"stdout: {proc.stdout}\nstderr: {proc.stderr}"
        )
    return expected
