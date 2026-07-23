# Веб-слой: экраны, HTMX-поток, настройки

## Маршруты (attestation/urls.py, app_name="attestation")

| URL | view | Назначение |
|---|---|---|
| `/` | `dashboard` | Список батчей (последние 100) |
| `/upload/` | `upload` | Форма загрузки zip (drag-and-drop) |
| `/batch/<pk>/` | `detail` | Один шаблон на все состояния: прогресс / ошибка / ревью / скачивание |
| `/batch/<pk>/status/` | `status` | HTMX-поллинг прогресса |
| `/batch/<pk>/cell/` | `edit_cell` (POST) | Инлайн-сохранение одной ячейки ревью |
| `/batch/<pk>/generate/` | `generate` (POST) | Запуск генерации документов (+смена языка) |
| `/batch/<pk>/download/<which>/` | `download` | which ∈ 5_1b / 6_5 / 6_4 / excel_1…excel_5 |

Плюс `config/urls.py`: `/admin/`, `/accounts/login/` (обёрнут в `login_not_required`),
`/accounts/logout/`; при DEBUG — раздача media.

**Аутентификация глобальная**: `django.contrib.auth.middleware.LoginRequiredMiddleware`
(Django 5.1+) закрывает ВСЁ, кроме страницы логина. Ролей нет — любой залогиненный
видит все батчи.

## Поток обработки глазами UI

1. `upload` POST: валидация только по расширению `.zip`; создаёт `Batch(UPLOADED)`,
   зовёт `jobs.start_processing`, редирект на detail.
2. detail при UPLOADED/PROCESSING включает фрагмент `_progress.html`, который **сам
   себя опрашивает**: `hx-get=status hx-trigger="load delay:1500ms"
   hx-swap="outerHTML"` — каждые 1.5 с приходит свежий фрагмент (stage из БД).
3. `status`: пока UPLOADED/PROCESSING — рендерит фрагмент; percent считается из
   строки stage по regex `(\d+)\s*/\s*(\d+)` («Конвертация … 12/41» → 29%). Иначе —
   `204` + заголовок **`HX-Redirect`** → HTMX перезагружает detail целиком.
4. При EXTRACTED/DONE detail показывает ревью-таблицу; при DONE — карточки
   скачивания трёх docx (эмеральд) + **пяти xlsx-протоколов (синие)**. Имена docx:
   `5_1б.docx`, `6_5_заключение.docx`, `6_4_йиғма_қайднома.docx`; xlsx:
   `1_Зарарли_моддалар.xlsx` … `5_Электромагнит_майдонлар.xlsx` (`views.EXCEL_DOCS`
   задаёт заголовки карточек, `download` — имена файлов). Модель хранит пути
   `output_excel_1…output_excel_5` (миграция `0004`), заполняет
   `jobs.generate_documents` (пять `render_excel_*` с `lang=batch.output_lang`).
5. Кнопка «Сформировать документы» (radio выбора языка Кириллица/Lotin рядом) →
   `generate` → снова PROCESSING → поллинг → DONE.

CSRF для HTMX: `<body hx-headers='{"X-CSRFToken": "…"}'>` в base.html — токен
наследуется всеми htmx-запросами.

## Экран ревью

`_review_context` строит таблицу: РМ сортируются `workplace_sort_key` (тот же
порядок, что в документах). Колонки — `REVIEW_COLUMNS` (27 шт.): №(ro), должность,
код, 15 факторов, травмоопасность(select 1/2/3), ЯТҲВ(select ҳа/йўқ), 4 льготы
(select), пенсия(select), ходимлар сони, аёллар. У колонки может быть флаг
подсветки; подсказки — `FLAG_HINTS` (см. `02-DATASET.md`, флаги).

Инлайн-правка: каждая ячейка — input/select с
`hx-post=edit_cell hx-vals='{"no", "field"}' hx-trigger="change" hx-swap="none"`;
успех → зелёная вспышка (JS-хук на `htmx:afterRequest` в base.html). `edit_cell`
принимает только пути из `REVIEW_COLUMNS` (кроме ro) — иначе 400; пишет по точечному
пути (`set_nested`) и сохраняет ВЕСЬ JSON. ⚠️ Race при одновременной правке двумя
операторами (последний перетрёт); значения select не валидируются на сервере.

Реквизиты компании показываются read-only блоком (7 строк). Предупреждения
извлечения (`batch.error` при не-failed статусе) — янтарная панель.

⚠️ Мелочи UI: на дашборде кнопки скачивания только 5_1б и 6_5 (6_4 — только на
detail); степпер этапов в `_progress.html` подсвечивается подстрочным матчем
stage-строк БЕЗ первой буквы (`'аспаковка'`, `'онвертац'`, `'звлеч'`, `'еречн'`) —
переименование stage в pipeline молча сломает подсветку.

## Шаблоны (templates/)

- `base.html` — Tailwind (play-CDN), HTMX 1.9.12, Alpine.js 3.13.10 (unpkg), шрифт
  Inter (rsms.me) — **всё с внешних CDN, офлайн UI деградирует**. CSS дата-грида:
  `.cell-flagged` (янтарная подсветка), `.cell-saved` (зелёная вспышка), sticky-колонки.
- `upload.html` — Alpine drag-and-drop, `accept=".zip"`, подпись «до 200 МБ».
- `detail.html`, `dashboard.html`, `_progress.html`, `_badge.html` (цвета статусов),
  `registration/login.html`.

## config/settings.py — все переменные окружения

| Переменная | Дефолт | Смысл |
|---|---|---|
| `DJANGO_SECRET_KEY` | dev-ключ | |
| `DJANGO_DEBUG` | **True** | дев-дефолт |
| `DJANGO_ALLOWED_HOSTS` | `*` | CSV |
| `DJANGO_CSRF_TRUSTED_ORIGINS` | пусто | CSV |
| `POSTGRES_DB` | — | **само наличие переключает БД на Postgres** (иначе SQLite) |
| `POSTGRES_USER/PASSWORD/HOST/PORT` | postgres / "" / `db` / 5432 | |
| `DJANGO_SQLITE_PATH` | `BASE_DIR/db.sqlite3` | |
| `DJANGO_TIME_ZONE` | `Asia/Tashkent` | `LANGUAGE_CODE="ru-ru"` |
| `DJANGO_MEDIA_ROOT` | `BASE_DIR/media` | |
| `CELERY_BROKER_URL` / `CELERY_RESULT_BACKEND` | redis://localhost:6379/0 и /1 | |
| `CELERY_TASK_ALWAYS_EAGER` | False | синхронные задачи (CI) |
| `SOFFICE_BIN` | `""` | путь LibreOffice (читает convert.py из os.environ) |
| `ATTESTATION_LLM_FALLBACK` | False | хук LLM-fallback, выключен |
| `ATTESTATION_TASK_RUNNER` | `thread` при DEBUG, иначе `celery` | ключевой переключатель раннера |
| `DATA_UPLOAD_MAX_MEMORY_SIZE` | 200 МБ | согласован с nginx `client_max_body_size 210m` |

Статика: WhiteNoise (`CompressedStaticFilesStorage`); в проде статику раздаёт nginx,
whitenoise — дублирующий путь (полезен без nginx).

## Админка

`BatchAdmin`: list = id, имя файла, статус, workplaces_count, created_at; фильтр по
статусу; JSON-поля редактируются сырым текстом. Inline `SourceFile` (read-only).
