"""Модели домена аттестации.

Главная сущность — Batch (одна загрузка zip). Извлечённый датасет хранится
прямо в JSON-полях, чтобы правки оператора сохранялись и документы можно было
перегенерировать без повторного разбора архива.
"""
from __future__ import annotations

from django.db import models


class Batch(models.Model):
    """Одна загрузка/обработка zip-архива клиента."""

    class Status(models.TextChoices):
        UPLOADED = "uploaded", "Загружен"
        PROCESSING = "processing", "Обрабатывается"
        EXTRACTED = "extracted", "Извлечён"
        DONE = "done", "Готов"
        FAILED = "failed", "Ошибка"

    created_at = models.DateTimeField("Создан", auto_now_add=True)
    updated_at = models.DateTimeField("Обновлён", auto_now=True)
    original_filename = models.CharField("Имя файла", max_length=512)

    status = models.CharField(
        "Статус", max_length=16, choices=Status.choices, default=Status.UPLOADED
    )
    # Текущий этап обработки (для прогресс-экрана), напр. "Конвертация 12/41"
    stage = models.CharField("Этап", max_length=255, blank=True, default="")
    error = models.TextField("Ошибка", blank=True, default="")

    # Загруженный архив и распакованная директория
    archive = models.FileField("Архив", upload_to="uploads/%Y/%m/%d/", blank=True)

    class OutputLang(models.TextChoices):
        CYRILLIC = "cyr", "Кириллица"
        LATIN = "lat", "Lotin (латиница)"

    # Язык выходных документов (5_1б, 6_5)
    output_lang = models.CharField(
        "Язык документов", max_length=3, choices=OutputLang.choices, default=OutputLang.CYRILLIC
    )

    # Реквизиты компании (общие на батч) и список записей рабочих мест
    company_data = models.JSONField("Реквизиты компании", default=dict, blank=True)
    extracted_data = models.JSONField("Извлечённые рабочие места", default=list, blank=True)

    # Пути к сгенерированным документам (относительно MEDIA_ROOT)
    output_5_1b = models.CharField("Документ 5_1б", max_length=512, blank=True, default="")
    output_6_5 = models.CharField("Документ 6_5", max_length=512, blank=True, default="")

    class Meta:
        verbose_name = "Батч"
        verbose_name_plural = "Батчи"
        ordering = ["-created_at"]

    def __str__(self) -> str:  # pragma: no cover - служебное
        return f"Batch #{self.pk}: {self.original_filename} ({self.status})"

    @property
    def workplaces_count(self) -> int:
        return len(self.extracted_data or [])


class SourceFile(models.Model):
    """Значимый файл, найденный внутри архива."""

    class Kind(models.TextChoices):
        CARD = "card", "Карта рабочего места"
        PERECHEN = "perechen", "Перечень"
        TEMPLATE = "template", "Шаблон (заполненный эталон)"
        PDF = "pdf", "PDF"
        XLSX = "xlsx", "Excel"
        OTHER = "other", "Прочее"

    batch = models.ForeignKey(Batch, related_name="files", on_delete=models.CASCADE)
    path = models.CharField("Путь в архиве", max_length=1024)
    kind = models.CharField("Тип", max_length=16, choices=Kind.choices, default=Kind.OTHER)
    converted_docx_path = models.CharField(
        "Путь к .docx", max_length=1024, blank=True, default=""
    )
    parsed = models.BooleanField("Разобран", default=False)

    class Meta:
        verbose_name = "Файл архива"
        verbose_name_plural = "Файлы архива"
        ordering = ["path"]

    def __str__(self) -> str:  # pragma: no cover - служебное
        return f"{self.get_kind_display()}: {self.path}"
