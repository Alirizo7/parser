from __future__ import annotations

import shutil
import tempfile
from datetime import timedelta
from pathlib import Path
from unittest import mock

from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.messages import get_messages
from django.core.files.uploadedfile import SimpleUploadedFile
from django.db import connection
from django.test import Client, TestCase, TransactionTestCase, override_settings
from django.urls import reverse
from django.utils import timezone

from attestation import jobs, views
from attestation.models import Batch, SourceFile


class BatchDeleteTests(TestCase):
    """Deletion must reclaim the complete persisted footprint of one job."""

    def setUp(self) -> None:
        super().setUp()
        self._media_tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._media_tmp.cleanup)
        self.media_root = Path(self._media_tmp.name)
        self._settings = override_settings(MEDIA_ROOT=self.media_root)
        self._settings.enable()
        self.addCleanup(self._settings.disable)

        self.user = get_user_model().objects.create_user(
            username="operator", password="test-password"
        )

    def _create_batch_with_files(
        self, filename: str = "client-work.zip"
    ) -> tuple[Batch, list[Path]]:
        batch = Batch.objects.create(
            original_filename=filename,
            status=Batch.Status.DONE,
            archive=SimpleUploadedFile(filename, b"source archive"),
            extracted_data=[{"workplace_no": "000001"}],
        )

        batch_dir = self.media_root / "batches" / str(batch.pk)
        generated = {
            "output_5_1b": batch_dir / "5_1b.docx",
            "output_6_5": batch_dir / "6_5.docx",
            "output_6_4": batch_dir / "6_4.docx",
            "output_6_6": batch_dir / "6_6.docx",
            "output_excel_1": batch_dir / "excel_1.xlsx",
            "output_excel_2": batch_dir / "excel_2.xlsx",
            "output_excel_3": batch_dir / "excel_3.xlsx",
            "output_excel_4": batch_dir / "excel_4.xlsx",
            "output_excel_5": batch_dir / "excel_5.xlsx",
        }
        for path in generated.values():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b"generated output")

        source_doc = batch_dir / "work" / "unpacked" / "card.doc"
        converted_doc = batch_dir / "work" / "docx" / "card.docx"
        source_doc.parent.mkdir(parents=True, exist_ok=True)
        converted_doc.parent.mkdir(parents=True, exist_ok=True)
        source_doc.write_bytes(b"source")
        converted_doc.write_bytes(b"converted")

        media_root = self.media_root.resolve()
        for field, path in generated.items():
            setattr(batch, field, str(path.resolve().relative_to(media_root)))
        batch.save(update_fields=[*generated, "updated_at"])

        SourceFile.objects.create(
            batch=batch,
            path="cards/card.doc",
            kind=SourceFile.Kind.CARD,
            converted_docx_path=str(converted_doc),
            parsed=True,
        )
        SourceFile.objects.create(
            batch=batch,
            path="perechen.xlsx",
            kind=SourceFile.Kind.PERECHEN,
            parsed=True,
        )

        return batch, [
            Path(batch.archive.path),
            batch_dir,
            source_doc,
            converted_doc,
            *generated.values(),
        ]

    @staticmethod
    def _delete_url(batch: Batch) -> str:
        return reverse("attestation:delete_batch", args=[batch.pk])

    def test_delete_is_post_only(self):
        batch, paths = self._create_batch_with_files()
        self.client.force_login(self.user)

        response = self.client.get(self._delete_url(batch))

        self.assertEqual(response.status_code, 405)
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))

    def test_delete_requires_authentication_and_preserves_job_on_redirect(self):
        batch, paths = self._create_batch_with_files()
        url = self._delete_url(batch)

        response = self.client.post(url)

        self.assertRedirects(
            response,
            f"{reverse('login')}?next={url}",
            fetch_redirect_response=False,
        )
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))

    def test_delete_rejects_post_without_csrf_token(self):
        batch, paths = self._create_batch_with_files()
        csrf_client = Client(enforce_csrf_checks=True)
        csrf_client.force_login(self.user)

        response = csrf_client.post(self._delete_url(batch))

        self.assertEqual(response.status_code, 403)
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))

    def test_delete_immediately_removes_uploaded_job_without_worker(self):
        batch, paths = self._create_batch_with_files("uploaded.zip")
        batch.status = Batch.Status.UPLOADED
        batch.save(update_fields=["status", "updated_at"])
        batch_pk = batch.pk
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertFalse(Batch.objects.filter(pk=batch_pk).exists())
        self.assertTrue(all(not path.exists() for path in paths))
        queued = list(get_messages(response.wsgi_request))
        self.assertEqual(queued[0].level, messages.SUCCESS)

    def test_worker_finalizer_removes_files_recreated_after_active_delete(self):
        batch, paths = self._create_batch_with_files("active.zip")
        batch.status = Batch.Status.PROCESSING
        batch.save(update_fields=["status", "updated_at"])
        batch_pk = batch.pk
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        batch.refresh_from_db()
        self.assertEqual(batch.status, Batch.Status.DELETING)
        self.assertEqual(batch.stage, "Ожидание остановки обработки…")
        self.assertTrue(all(path.exists() for path in paths))
        queued = list(get_messages(response.wsgi_request))
        self.assertEqual(queued[0].level, messages.INFO)
        self.assertIn("Файлы будут удалены после остановки", str(queued[0]))

        # Worker мог закончить текущий renderer и дописать ещё один файл перед
        # тем, как увидел DELETING. Его finally обязан убрать и этот файл.
        recreated_dir = self.media_root / "batches" / str(batch_pk)
        (recreated_dir / "late-output.xlsx").write_bytes(b"late worker output")

        jobs._cleanup_deleted_batch_dir(batch_pk)  # noqa: SLF001 — worker finally guard

        self.assertFalse(Batch.objects.filter(pk=batch_pk).exists())
        self.assertFalse(recreated_dir.exists())
        self.assertTrue(all(not path.exists() for path in paths))

    def test_worker_finalizer_does_not_touch_view_owned_cleanup(self):
        batch, paths = self._create_batch_with_files("view-owned.zip")
        Batch.objects.filter(pk=batch.pk).update(
            status=Batch.Status.DELETING,
            stage=jobs.DELETE_RUNNING_STAGE,
            updated_at=timezone.now(),
        )

        jobs._cleanup_deleted_batch_dir(batch.pk)  # noqa: SLF001

        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))

    def test_delete_removes_batch_related_rows_archive_and_work_directory(self):
        batch, paths = self._create_batch_with_files()
        batch_pk = batch.pk
        source_file_pks = list(batch.files.values_list("pk", flat=True))

        # A different job is the storage-scope guard: recursive cleanup must not
        # reach a sibling batch or its separately uploaded archive.
        other, other_paths = self._create_batch_with_files("other-work.zip")
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertFalse(Batch.objects.filter(pk=batch_pk).exists())
        self.assertFalse(SourceFile.objects.filter(pk__in=source_file_pks).exists())
        self.assertTrue(Batch.objects.filter(pk=other.pk).exists())
        self.assertTrue(all(path.exists() for path in other_paths))
        self.assertTrue(all(not path.exists() for path in paths))

        queued = list(get_messages(response.wsgi_request))
        self.assertEqual(len(queued), 1)
        self.assertEqual(queued[0].level, messages.SUCCESS)
        self.assertIn("client-work.zip", str(queued[0]))

    def test_delete_succeeds_when_physical_files_are_already_missing(self):
        batch, paths = self._create_batch_with_files()
        batch_pk = batch.pk
        Path(batch.archive.path).unlink()
        shutil.rmtree(self.media_root / "batches" / str(batch.pk))
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertFalse(Batch.objects.filter(pk=batch_pk).exists())
        self.assertTrue(all(not path.exists() for path in paths))

    def test_delete_unknown_batch_returns_404_without_touching_other_jobs(self):
        existing, existing_paths = self._create_batch_with_files()
        self.client.force_login(self.user)

        response = self.client.post(reverse("attestation:delete_batch", args=[999_999]))

        self.assertEqual(response.status_code, 404)
        self.assertTrue(Batch.objects.filter(pk=existing.pk).exists())
        self.assertTrue(all(path.exists() for path in existing_paths))

    @mock.patch(
        "attestation.views.jobs.delete_batch_artifacts",
        side_effect=OSError("storage unavailable"),
    )
    def test_cleanup_failure_keeps_database_row_for_retry(self, _delete_artifacts):
        batch, paths = self._create_batch_with_files()
        self.client.force_login(self.user)

        with mock.patch("attestation.views.logger.exception"):
            response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))
        batch.refresh_from_db()
        self.assertEqual(batch.status, Batch.Status.FAILED)
        self.assertEqual(batch.stage, "Ошибка удаления")
        queued = list(get_messages(response.wsgi_request))
        self.assertEqual(queued[0].level, messages.ERROR)
        self.assertEqual(
            str(queued[0]),
            "Не удалось удалить все файлы работы. Повторите попытку.",
        )

    def test_cleanup_refuses_symlinked_batch_directory(self):
        batch, _paths = self._create_batch_with_files()
        batch_dir = self.media_root / "batches" / str(batch.pk)
        shutil.rmtree(batch_dir)
        protected_dir = self.media_root / "must-not-delete"
        protected_dir.mkdir()
        protected_file = protected_dir / "keep.txt"
        protected_file.write_text("keep", encoding="utf-8")
        batch_dir.symlink_to(protected_dir, target_is_directory=True)
        self.client.force_login(self.user)

        with mock.patch("attestation.views.logger.exception"):
            response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(protected_file.exists())

    def test_cleanup_refuses_symlinked_batches_root(self):
        batch = Batch.objects.create(
            original_filename="root-symlink.zip",
            status=Batch.Status.DONE,
            archive=SimpleUploadedFile("root-symlink.zip", b"archive"),
        )
        protected_dir = self.media_root / "must-not-delete"
        protected_dir.mkdir()
        protected_file = protected_dir / str(batch.pk)
        protected_file.mkdir()
        (protected_file / "keep.txt").write_text("keep", encoding="utf-8")
        (self.media_root / "batches").symlink_to(protected_dir, target_is_directory=True)
        self.client.force_login(self.user)

        with mock.patch("attestation.views.logger.exception"):
            response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue((protected_file / "keep.txt").exists())

    def test_archive_delete_failure_is_retryable(self):
        batch, paths = self._create_batch_with_files()
        batch_pk = batch.pk
        url = self._delete_url(batch)
        archive_path = Path(batch.archive.path)
        batch_dir = self.media_root / "batches" / str(batch.pk)
        self.client.force_login(self.user)

        with mock.patch.object(
            batch.archive.storage, "delete", side_effect=OSError("storage unavailable")
        ), mock.patch("attestation.views.logger.exception"):
            failed_response = self.client.post(url)

        self.assertRedirects(
            failed_response,
            reverse("attestation:dashboard"),
            fetch_redirect_response=False,
        )
        self.assertTrue(Batch.objects.filter(pk=batch_pk).exists())
        self.assertTrue(archive_path.exists())
        self.assertFalse(batch_dir.exists())
        batch.refresh_from_db()
        self.assertEqual(batch.status, Batch.Status.FAILED)
        self.assertEqual(batch.stage, "Ошибка удаления")

        retry_response = self.client.post(url)

        self.assertRedirects(
            retry_response,
            reverse("attestation:dashboard"),
            fetch_redirect_response=False,
        )
        self.assertFalse(Batch.objects.filter(pk=batch_pk).exists())
        self.assertTrue(all(not path.exists() for path in paths))

    def test_second_delete_request_does_not_race_active_cleanup(self):
        batch, paths = self._create_batch_with_files()
        Batch.objects.filter(pk=batch.pk).update(
            status=Batch.Status.DELETING,
            stage="Удаление файлов…",
            updated_at=timezone.now(),
        )
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))
        queued = list(get_messages(response.wsgi_request))
        self.assertEqual(queued[0].level, messages.INFO)
        self.assertEqual(str(queued[0]), "Эта работа уже удаляется.")

    def test_deleting_job_detail_shows_stopping_notice(self):
        batch = Batch.objects.create(
            original_filename="stopping.zip",
            status=Batch.Status.DELETING,
            stage="Ожидание остановки обработки…",
        )
        self.client.force_login(self.user)

        response = self.client.get(reverse("attestation:detail", args=[batch.pk]))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Работа удаляется")
        self.assertContains(response, "Текущая обработка останавливается")

    def test_stale_delete_claim_can_be_retried(self):
        batch, paths = self._create_batch_with_files()
        batch_pk = batch.pk
        Batch.objects.filter(pk=batch.pk).update(
            status=Batch.Status.DELETING,
            stage="Удаление файлов…",
            updated_at=timezone.now() - timedelta(minutes=16),
        )
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertFalse(Batch.objects.filter(pk=batch_pk).exists())
        self.assertTrue(all(not path.exists() for path in paths))

    def test_worker_owned_claim_uses_longer_stale_timeout(self):
        batch, paths = self._create_batch_with_files()
        Batch.objects.filter(pk=batch.pk).update(
            status=Batch.Status.DELETING,
            stage=jobs.DELETE_WORKER_STAGE,
            updated_at=timezone.now() - timedelta(minutes=16),
        )
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch))

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertTrue(Batch.objects.filter(pk=batch.pk).exists())
        self.assertTrue(all(path.exists() for path in paths))

    def test_dashboard_shows_delete_action_and_paginates_all_jobs(self):
        done = Batch.objects.create(original_filename="done.zip", status=Batch.Status.DONE)
        Batch.objects.create(
            original_filename="running.zip", status=Batch.Status.PROCESSING
        )
        Batch.objects.bulk_create(
            [
                Batch(original_filename=f"old-{index}.zip", status=Batch.Status.DONE)
                for index in range(24)
            ]
        )
        self.client.force_login(self.user)

        first_page = self.client.get(reverse("attestation:dashboard"))
        second_page = self.client.get(reverse("attestation:dashboard"), {"page": 2})

        self.assertEqual(first_page.status_code, 200)
        self.assertEqual(len(first_page.context["batches"]), 25)
        self.assertEqual(len(second_page.context["batches"]), 1)
        self.assertContains(first_page, "Удалить")
        self.assertContains(
            first_page,
            "Удалить работу «",
        )
        # Самая старая запись остаётся доступна на следующей странице.
        self.assertContains(second_page, self._delete_url(done))

    def test_delete_redirects_back_to_requested_dashboard_page(self):
        batch, _paths = self._create_batch_with_files()
        self.client.force_login(self.user)

        response = self.client.post(self._delete_url(batch), {"page": "3"})

        self.assertRedirects(
            response,
            f"{reverse('attestation:dashboard')}?page=3",
            fetch_redirect_response=False,
        )


class BatchGenerationClaimTests(TestCase):
    @mock.patch("attestation.jobs._dispatch")
    def test_start_generation_claims_batch_before_dispatch(self, dispatch):
        batch = Batch.objects.create(
            original_filename="regenerate.zip",
            status=Batch.Status.DONE,
            error="old warning",
        )

        def assert_claimed(job_name, batch_id):
            current = Batch.objects.get(pk=batch_id)
            self.assertEqual(job_name, "generate_documents")
            self.assertEqual(current.status, Batch.Status.PROCESSING)
            self.assertEqual(current.stage, jobs.GENERATION_QUEUED_STAGE)
            self.assertEqual(current.error, "old warning")

        dispatch.side_effect = assert_claimed

        started = jobs.start_generation(batch.pk)

        self.assertTrue(started)
        dispatch.assert_called_once_with("generate_documents", batch.pk)

    @mock.patch("attestation.jobs._dispatch")
    def test_second_generation_start_is_not_dispatched(self, dispatch):
        batch = Batch.objects.create(
            original_filename="regenerate.zip", status=Batch.Status.DONE
        )

        self.assertTrue(jobs.start_generation(batch.pk))
        self.assertFalse(jobs.start_generation(batch.pk))

        dispatch.assert_called_once_with("generate_documents", batch.pk)

    def test_duplicate_worker_can_claim_queue_only_once(self):
        batch = Batch.objects.create(
            original_filename="duplicate-delivery.zip",
            status=Batch.Status.PROCESSING,
            stage=jobs.GENERATION_QUEUED_STAGE,
        )

        first = jobs._claim_queued_job(  # noqa: SLF001 — проверяем concurrency guard
            batch.pk, jobs.GENERATION_QUEUED_STAGE, jobs.GENERATION_RUNNING_STAGE
        )
        second = jobs._claim_queued_job(  # noqa: SLF001
            batch.pk, jobs.GENERATION_QUEUED_STAGE, jobs.GENERATION_RUNNING_STAGE
        )

        self.assertTrue(first)
        self.assertFalse(second)

    @mock.patch("attestation.views.Batch.objects.filter")
    def test_delete_cas_retries_if_generation_starts_between_read_and_claim(
        self, filter_mock
    ):
        batch = Batch.objects.create(
            original_filename="race.zip", status=Batch.Status.DONE
        )
        original_filter = Batch.objects.__class__.filter.__get__(
            Batch.objects, Batch.objects.__class__
        )
        calls = 0

        def racing_filter(*args, **kwargs):
            nonlocal calls
            calls += 1
            if calls == 2:
                original_filter(pk=batch.pk, status=Batch.Status.DONE).update(
                    status=Batch.Status.PROCESSING,
                    stage=jobs.GENERATION_QUEUED_STAGE,
                )
            return original_filter(*args, **kwargs)

        filter_mock.side_effect = racing_filter

        claim = views._claim_batch_deletion(batch.pk)  # noqa: SLF001

        self.assertEqual(claim, "waiting")
        batch.refresh_from_db()
        self.assertEqual(batch.status, Batch.Status.DELETING)
        self.assertEqual(batch.stage, jobs.DELETE_WAITING_STAGE)

    @mock.patch("attestation.jobs._dispatch", side_effect=RuntimeError("queue down"))
    def test_start_generation_marks_failed_when_dispatch_fails(self, _dispatch):
        batch = Batch.objects.create(
            original_filename="regenerate.zip", status=Batch.Status.DONE
        )

        with self.assertRaisesRegex(RuntimeError, "queue down"):
            jobs.start_generation(batch.pk)

        batch.refresh_from_db()
        self.assertEqual(batch.status, Batch.Status.FAILED)
        self.assertEqual(batch.stage, "Ошибка")
        self.assertEqual(batch.error, "queue down")

    @mock.patch("attestation.jobs._dispatch", side_effect=RuntimeError("queue down"))
    def test_start_processing_marks_failed_when_dispatch_fails(self, _dispatch):
        batch = Batch.objects.create(
            original_filename="upload.zip", status=Batch.Status.UPLOADED
        )

        with self.assertRaisesRegex(RuntimeError, "queue down"):
            jobs.start_processing(batch.pk)

        batch.refresh_from_db()
        self.assertEqual(batch.status, Batch.Status.FAILED)
        self.assertEqual(batch.stage, "Ошибка")
        self.assertEqual(batch.error, "queue down")


class BatchDeleteTransactionBoundaryTests(TransactionTestCase):
    """Filesystem cleanup must not hold SQLite's long write transaction."""

    def setUp(self) -> None:
        super().setUp()
        self._media_tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._media_tmp.cleanup)
        self.media_root = Path(self._media_tmp.name)
        self._settings = override_settings(MEDIA_ROOT=self.media_root)
        self._settings.enable()
        self.addCleanup(self._settings.disable)
        self.user = get_user_model().objects.create_user(
            username="transaction-operator", password="test-password"
        )

    def test_artifact_cleanup_runs_outside_database_transaction(self):
        batch = Batch.objects.create(
            original_filename="sqlite-safe.zip",
            status=Batch.Status.DONE,
            archive=SimpleUploadedFile("sqlite-safe.zip", b"archive"),
        )
        batch_dir = self.media_root / "batches" / str(batch.pk)
        batch_dir.mkdir(parents=True)
        (batch_dir / "output.docx").write_bytes(b"output")
        original_cleanup = jobs.delete_batch_artifacts

        def assert_no_transaction(batch_to_delete):
            self.assertFalse(connection.in_atomic_block)
            return original_cleanup(batch_to_delete)

        client = Client()
        client.force_login(self.user)
        with mock.patch(
            "attestation.views.jobs.delete_batch_artifacts",
            side_effect=assert_no_transaction,
        ):
            response = client.post(
                reverse("attestation:delete_batch", args=[batch.pk])
            )

        self.assertRedirects(
            response, reverse("attestation:dashboard"), fetch_redirect_response=False
        )
        self.assertFalse(Batch.objects.filter(pk=batch.pk).exists())
        self.assertFalse(batch_dir.exists())
