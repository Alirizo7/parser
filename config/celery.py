"""Конфигурация Celery."""
import os

from celery import Celery

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")

app = Celery("attestation")
# Читаем настройки из Django settings (префикс CELERY_)
app.config_from_object("django.conf:settings", namespace="CELERY")
# Автопоиск задач в приложениях (attestation/tasks.py)
app.autodiscover_tasks()
