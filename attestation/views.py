"""Представления приложения «Авто-аттестация».

Экраны: дашборд → загрузка → прогресс (HTMX-поллинг) → ревью (инлайн-правка
с подсветкой флагов) → скачивание двух документов.
"""
from __future__ import annotations

import re
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.http import FileResponse, Http404, HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.views.decorators.http import require_POST

from . import jobs
from .models import Batch
from .services.extract import workplace_sort_key

YESNO = ["ҳа", "йўқ"]

# Описание колонок таблицы ревью: (путь, заголовок, тип, варианты, флаг-подсветки)
# path — точечный путь в записи рабочего места; flag — имя флага, подсвечивающего колонку.
REVIEW_COLUMNS = [
    ("workplace_no", "№", "ro", None, None),
    ("position", "Касб/лавозим", "text", None, "position_mismatch"),
    ("job_code", "Коди", "text", None, "job_code_missing"),
    ("factors.chem", "Кимёвий", "text", None, None),
    ("factors.biological", "Биологик", "text", None, None),
    ("factors.aerosols", "Аэрозол", "text", None, None),
    ("factors.noise", "Шовқин", "text", None, None),
    ("factors.infrasound", "Инфра", "text", None, None),
    ("factors.ultrasound_air", "Ультра", "text", None, None),
    ("factors.vibration_general", "Умумий вибр.", "text", None, None),
    ("factors.vibration_local", "Маҳаллий вибр.", "text", None, None),
    ("factors.em_field", "ЭМ майдон", "text", None, None),
    ("factors.ionizing", "Ион нурл.", "text", None, None),
    ("factors.microclimate", "Микроиқлим", "text", None, None),
    ("factors.lighting", "Ёруғлик", "text", None, None),
    ("factors.severity", "Оғирлик", "text", None, None),
    ("factors.intensity", "Тиғизлик", "text", None, None),
    ("factors.overall", "Умумий", "text", None, "overall_missing"),
    ("injury_risk", "Шикастланиш", "select", ["1", "2", "3"], "injury_risk_heuristic"),
    ("ppe_provided", "ЯТҲВ", "select", YESNO, None),
    ("benefits.extra_leave", "Қўш. таътил", "select", YESNO, None),
    ("benefits.reduced_hours", "Қисқарт. вақт", "select", YESNO, None),
    ("benefits.milk", "Сут", "select", YESNO, None),
    ("benefits.therapeutic_food", "Даво-проф.", "select", YESNO, None),
    ("privileged_pension", "Имтиёзли пенсия", "select", YESNO, None),
    ("employees_count", "Ходимлар сони", "text", None, "employees_count_missing"),
    ("female_count", "Шу жумладан аёллар", "text", None, "female_count_missing"),
]

# Пояснения к флагам (для легенды и подсказок)
FLAG_HINTS = {
    "injury_risk_heuristic": "Травмоопасность определена эвристикой (медик→1, иначе→2) — проверьте.",
    "job_code_missing": "Код должности не найден в «Перечне» — введите вручную.",
    "overall_missing": "Не извлечён общий класс условий труда — проверьте карту.",
    "position_mismatch": "Должность в «Перечне» и в карте на этом номере различаются — проверьте.",
    "substances_missing": "В карте не найдены вредные вещества — при необходимости добавьте вручную.",
    "employees_count_missing": "В карте не найдена строка «Ишловчилар сони» — принято 1, проверьте (влияет на документ 6_4).",
    "female_count_missing": "В карте не найдена гендерная разбивка — «Шу жумладан аёллар» пусто (влияет на документ 6_4).",
}


# --- Помощники работы с вложенными путями -----------------------------------
def get_nested(record: dict, path: str) -> str:
    cur = record
    for part in path.split("."):
        if not isinstance(cur, dict):
            return ""
        cur = cur.get(part, "")
    return "" if cur is None else str(cur)


def set_nested(record: dict, path: str, value: str) -> None:
    parts = path.split(".")
    cur = record
    for part in parts[:-1]:
        cur = cur.setdefault(part, {})
    cur[parts[-1]] = value


# --- Экраны -----------------------------------------------------------------
def dashboard(request):
    batches = Batch.objects.all()[:100]
    return render(request, "attestation/dashboard.html", {"batches": batches})


def upload(request):
    if request.method == "POST":
        f = request.FILES.get("archive")
        if not f:
            messages.error(request, "Выберите zip-архив.")
            return redirect("attestation:upload")
        if not f.name.lower().endswith(".zip"):
            messages.error(request, "Нужен файл .zip.")
            return redirect("attestation:upload")
        lang = request.POST.get("output_lang", Batch.OutputLang.CYRILLIC)
        if lang not in Batch.OutputLang.values:
            lang = Batch.OutputLang.CYRILLIC
        batch = Batch.objects.create(
            original_filename=f.name, archive=f, status=Batch.Status.UPLOADED,
            output_lang=lang,
        )
        jobs.start_processing(batch.pk)
        return redirect("attestation:detail", pk=batch.pk)
    return render(request, "attestation/upload.html", {"lang_choices": Batch.OutputLang.choices})


# Пять Excel-протоколов: (which для download, поле модели, заголовок карточки)
EXCEL_DOCS = [
    ("excel_1", "output_excel_1", "1 — Зарарли моддалар (вредные вещества)"),
    ("excel_2", "output_excel_2", "2 — Физик омиллар (шум/вибрация/инфразвук)"),
    ("excel_3", "output_excel_3", "3 — Микроиқлим"),
    ("excel_4", "output_excel_4", "4 — Ёруғлик (освещённость)"),
    ("excel_5", "output_excel_5", "5 — Электромагнит майдонлар (ЭМИ)"),
]


def detail(request, pk):
    batch = get_object_or_404(Batch, pk=pk)
    # Показываем карточку только для реально сформированных протоколов (путь непуст),
    # чтобы сбой одного рендера не давал битую ссылку-404.
    excel_docs = [(which, title) for which, attr, title in EXCEL_DOCS if getattr(batch, attr, "")]
    context = {"batch": batch, "lang_choices": Batch.OutputLang.choices,
               "excel_docs": excel_docs}
    if batch.status in (Batch.Status.EXTRACTED, Batch.Status.DONE):
        context.update(_review_context(batch))
    return render(request, "attestation/detail.html", context)


def _progress_percent(stage: str) -> int | None:
    m = re.search(r"(\d+)\s*/\s*(\d+)", stage or "")
    if not m:
        return None
    done, total = int(m.group(1)), int(m.group(2))
    return int(done * 100 / total) if total else None


def status(request, pk):
    """HTMX-поллинг статуса. Пока идёт обработка — отдаём прогресс-фрагмент;
    как только готово/ошибка — просим HTMX перезагрузить страницу."""
    batch = get_object_or_404(Batch, pk=pk)
    if batch.status in (Batch.Status.UPLOADED, Batch.Status.PROCESSING):
        return render(
            request,
            "attestation/_progress.html",
            {"batch": batch, "percent": _progress_percent(batch.stage)},
        )
    resp = HttpResponse(status=204)
    resp["HX-Redirect"] = reverse("attestation:detail", args=[pk])
    return resp


def _review_context(batch: Batch) -> dict:
    """Подготовить строки таблицы ревью с метаданными ячеек."""
    rows = []
    # Тот же порядок, что и в документах (а-суффикс сразу после базового номера)
    workplaces = sorted(
        batch.extracted_data, key=lambda w: workplace_sort_key(w.get("workplace_no", ""))
    )
    for wp in workplaces:
        flags = set(wp.get("flags", []))
        cells = []
        for path, label, ctype, choices, flag in REVIEW_COLUMNS:
            cells.append(
                {
                    "path": path,
                    "value": get_nested(wp, path),
                    "type": ctype,
                    "choices": choices,
                    "flagged": flag in flags if flag else False,
                    "hint": FLAG_HINTS.get(flag, "") if flag else "",
                }
            )
        rows.append({"workplace_no": wp.get("workplace_no", ""), "cells": cells})
    columns = [{"label": c[1]} for c in REVIEW_COLUMNS]
    cd = batch.company_data or {}
    company_rows = [
        ("Наименование", cd.get("name", "")),
        ("Вышестоящая организация", cd.get("parent", "")),
        ("Юридический адрес", cd.get("address", "")),
        ("Основной вид продукции", cd.get("product", "")),
        ("СТИР", cd.get("stir", "")),
        ("ИФУТ", cd.get("ifut", "")),
        ("МХБТ", cd.get("mxbt", "")),
    ]
    return {"columns": columns, "rows": rows, "company_rows": company_rows}


@require_POST
def edit_cell(request, pk):
    """Инлайн-сохранение одной ячейки в Batch.extracted_data."""
    batch = get_object_or_404(Batch, pk=pk)
    no = request.POST.get("no", "")
    field = request.POST.get("field", "")
    value = request.POST.get("value", "")
    valid_paths = {c[0] for c in REVIEW_COLUMNS if c[2] != "ro"}
    if field not in valid_paths:
        return HttpResponse("Недопустимое поле", status=400)

    data = batch.extracted_data
    for wp in data:
        if wp.get("workplace_no") == no:
            set_nested(wp, field, value)
            batch.extracted_data = data
            batch.save(update_fields=["extracted_data", "updated_at"])
            return HttpResponse(status=204)
    raise Http404("Рабочее место не найдено")


@require_POST
def generate(request, pk):
    batch = get_object_or_404(Batch, pk=pk)
    if not batch.extracted_data:
        messages.error(request, "Нет данных для формирования документов.")
        return redirect("attestation:detail", pk=pk)
    lang = request.POST.get("output_lang")
    if lang in Batch.OutputLang.values and lang != batch.output_lang:
        batch.output_lang = lang
        batch.save(update_fields=["output_lang", "updated_at"])
    jobs.start_generation(batch.pk)
    return redirect("attestation:detail", pk=pk)


def download(request, pk, which):
    batch = get_object_or_404(Batch, pk=pk)
    rel = {
        "5_1b": batch.output_5_1b,
        "6_5": batch.output_6_5,
        "6_4": batch.output_6_4,
        "excel_1": batch.output_excel_1,
        "excel_2": batch.output_excel_2,
        "excel_3": batch.output_excel_3,
        "excel_4": batch.output_excel_4,
        "excel_5": batch.output_excel_5,
    }.get(which)
    if not rel:
        raise Http404("Документ ещё не сформирован")
    path = Path(settings.MEDIA_ROOT) / rel
    if not path.exists():
        raise Http404("Файл не найден")
    nice = {
        "5_1b": "5_1б.docx",
        "6_5": "6_5_заключение.docx",
        "6_4": "6_4_йиғма_қайднома.docx",
        "excel_1": "1_Зарарли_моддалар.xlsx",
        "excel_2": "2_Физик_омиллар.xlsx",
        "excel_3": "3_Микроиқлим.xlsx",
        "excel_4": "4_Ёруғлик.xlsx",
        "excel_5": "5_Электромагнит_майдонлар.xlsx",
    }[which]
    return FileResponse(open(path, "rb"), as_attachment=True, filename=nice)
