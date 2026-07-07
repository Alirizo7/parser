"""Самопроверка движка на примере zip: извлечь → заполнить → сверить с эталоном.

    # Кириллический клиент (эталоны лежат внутри zip):
    python manage.py selfcheck "Бухоро болалар стоматологияси.zip"

    # Латинский клиент (правильные документы — отдельные файлы, сверка по письму):
    python manage.py selfcheck KARTA.zip --bilingual \
        --ref-5-1b 5_1б_правилный.doc --ref-6-5 6_5-правилный.doc

Прогоняет конвейер, генерирует документы и сравнивает их с эталонными.
В режиме ``--bilingual`` обе стороны приводятся к одному письму (транслитерация,
``ҳа≡ha``, ``йўқ≡yo'q``); различие языка ярлыков расхождением не считается.
Завершается ненулевым кодом при расхождении ключевых полей.
"""
from __future__ import annotations

import tempfile
from pathlib import Path

from django.core.management.base import BaseCommand, CommandError

from attestation.services import render, selfcheck
from attestation.services.convert import convert_to_docx
from attestation.services.pipeline import run_pipeline


class Command(BaseCommand):
    help = "Сверить сгенерированные документы с эталонами из примера zip."

    def add_arguments(self, parser) -> None:
        parser.add_argument("zip_path")
        parser.add_argument("--only", choices=["5_1b", "6_5"], default=None,
                            help="Проверить только один шаблон")
        parser.add_argument("--bilingual", action="store_true",
                            help="Сверять по содержанию между письмами (для латинского клиента)")
        parser.add_argument("--ref-5-1b", default="", help="Путь к правильному 5_1б (вне архива)")
        parser.add_argument("--ref-6-5", default="", help="Путь к правильному 6_5 (вне архива)")
        parser.add_argument("--workdir", default="")

    def handle(self, *args, **opts) -> None:
        zip_path = Path(opts["zip_path"])
        if not zip_path.exists():
            raise CommandError(f"Архив не найден: {zip_path}")
        work_dir = Path(opts["workdir"]) if opts["workdir"] else Path(tempfile.mkdtemp(prefix="att_sc_"))
        self.out_dir = work_dir / "out"
        self.out_dir.mkdir(parents=True, exist_ok=True)
        self.bilingual = opts["bilingual"]

        self.stdout.write("Извлечение данных…")
        result = run_pipeline(zip_path, work_dir)
        self.stdout.write(f"  рабочих мест: {len(result.workplaces)}")
        unpacked = work_dir / "unpacked"

        # Источник эталона: явный путь (вне архива) либо поиск внутри архива
        ref_5 = Path(opts["ref_5_1b"]) if opts["ref_5_1b"] else selfcheck.find_reference(unpacked, "5_1b")
        ref_65 = Path(opts["ref_6_5"]) if opts["ref_6_5"] else selfcheck.find_reference(unpacked, "6_5")

        failed = False
        targets = [opts["only"]] if opts["only"] else ["5_1b", "6_5"]
        if "5_1b" in targets:
            failed |= not self._check_5_1b(result, ref_5)
        if "6_5" in targets:
            failed |= not self._check_6_5(result, ref_65)

        if failed:
            raise CommandError("Самопроверка НЕ пройдена (см. расхождения выше).")
        self.stdout.write(self.style.SUCCESS("Самопроверка пройдена."))

    def _to_docx(self, ref: Path) -> Path:
        return convert_to_docx(ref, self.out_dir) if ref.suffix.lower() == ".doc" else ref

    def _report(self, check, unit: str) -> bool:
        style = self.style.SUCCESS if check.ok else self.style.ERROR
        self.stdout.write(style(f"  совпало {unit}: {check.matched}/{check.total}"))
        for m in check.mismatches[:40]:
            self.stdout.write(self.style.ERROR(f"  ✗ {m}"))
        if check.notes:
            self.stdout.write(self.style.WARNING(
                f"  к ручной проверке / расхождения карта↔эталон клиента ({len(check.notes)}):"))
            for n in check.notes[:40]:
                self.stdout.write(self.style.WARNING(f"  • {n}"))
        return check.ok

    def _check_5_1b(self, result, ref) -> bool:
        self.stdout.write("\n=== Шаблон 5_1б ===")
        if ref is None or not ref.exists():
            self.stdout.write(self.style.WARNING("Эталон 5_1б не найден — пропуск."))
            return True
        generated = render.render_5_1b(result.workplaces, self.out_dir / "generated_5_1b.docx")
        check = selfcheck.compare_5_1b(generated, self._to_docx(ref), bilingual=self.bilingual)
        return self._report(check, "рабочих мест")

    def _check_6_5(self, result, ref) -> bool:
        self.stdout.write("\n=== Шаблон 6_5 ===")
        if ref is None or not ref.exists():
            self.stdout.write(self.style.WARNING("Эталон 6_5 не найден — пропуск."))
            return True
        generated = render.render_6_5(
            result.company_data, result.workplaces, self.out_dir / "generated_6_5.docx"
        )
        if self.bilingual:
            check = selfcheck.compare_6_5_bilingual(generated, self._to_docx(ref))
        else:
            check = selfcheck.compare_6_5(generated, self._to_docx(ref))
        return self._report(check, "ячеек")
