"""Django-проект «Авто-аттестация».

Импортируем Celery-приложение, чтобы оно инициализировалось при старте Django
(нужно для регистрации задач через декоратор @shared_task).
"""
from .celery import app as celery_app

__all__ = ("celery_app",)
