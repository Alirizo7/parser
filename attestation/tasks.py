"""Celery-задачи: тонкие обёртки над функциями оркестрации из jobs.py."""
from __future__ import annotations

from celery import shared_task

from . import jobs


@shared_task
def process_batch_task(batch_id: int) -> None:
    jobs.process_batch(batch_id)


@shared_task
def generate_documents_task(batch_id: int) -> None:
    jobs.generate_documents(batch_id)
