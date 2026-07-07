"""Прогнать конвейер извлечения на zip-архиве и сохранить датасет в JSON.

Используется для отладки этапа извлечения (без БД, Celery и UI):

    python manage.py extract_demo "Бухоро болалар стоматологияси.zip" \
        --out media/demo_extracted.json
"""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

from django.core.management.base import BaseCommand, CommandError

from attestation.services.pipeline import run_pipeline


class Command(BaseCommand):
    help = "Извлечь датасет из примера zip и сохранить JSON (отладка)."

    def add_arguments(self, parser) -> None:
        parser.add_argument("zip_path", help="Путь к zip-архиву")
        parser.add_argument("--out", default="media/demo_extracted.json", help="Куда сохранить JSON")
        parser.add_argument("--workdir", default="", help="Рабочая директория (по умолчанию tmp)")

    def handle(self, *args, **opts) -> None:
        zip_path = Path(opts["zip_path"])
        if not zip_path.exists():
            raise CommandError(f"Архив не найден: {zip_path}")

        work_dir = Path(opts["workdir"]) if opts["workdir"] else Path(tempfile.mkdtemp(prefix="att_"))
        self.stdout.write(f"Рабочая директория: {work_dir}")

        def progress(stage: str) -> None:
            self.stdout.write(self.style.HTTP_INFO(f"  · {stage}"))

        result = run_pipeline(zip_path, work_dir, progress=progress)

        payload = {
            "company_data": result.company_data,
            "workplaces": result.workplaces,
            "warnings": result.warnings,
        }
        out = Path(opts["out"])
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

        self.stdout.write("")
        self.stdout.write(self.style.SUCCESS(f"Извлечено рабочих мест: {len(result.workplaces)}"))
        self.stdout.write(f"Реквизиты компании: {result.company_data.get('name', '—')}")
        if result.warnings:
            self.stdout.write(self.style.WARNING("Предупреждения:"))
            for w in result.warnings:
                self.stdout.write(f"  ! {w}")
        self.stdout.write(self.style.SUCCESS(f"JSON сохранён: {out}"))
