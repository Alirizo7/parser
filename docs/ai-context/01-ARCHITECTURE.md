# Архитектура: конвейер и Django-клей

## Конвейер целиком

```
zip ──▶ unpack ──▶ convert (.doc→.docx) ──▶ extract (карты + «Перечень») ──▶ единый датасет
                                                                                   │
                                            ┌──────────────────────────────────────┼─────────────────────┐
                                     render_5_1b                            render_6_5             render_6_4
                                 (вредные вещества)                  (сводная, 25 колонок)  (итоги по подразделениям)
```

Сервисный слой (`attestation/services/`) не импортирует Django — функции принимают
пути и данные, возвращают данные, опционально зовут колбэк `progress(stage: str)`.
Это позволяет гонять конвейер из management-команд, тестов и Celery одинаково.

## `run_pipeline(zip_path, work_dir, *, progress)` — pipeline.py

Возвращает `PipelineResult` (dataclass): `company_data: dict`, `workplaces: list[dict]`,
`files: list[dict]` (сводка `{"path", "kind"}` для модели SourceFile), `warnings: list[str]`.

Шаги (нумерация из комментариев кода):

1. **Распаковка** → `work_dir/unpacked` (см. `04-NORMALIZE.md` про кодировки имён).
2. **Классификация файлов** — `classify()`: `.pdf`→`pdf`, `.xlsx/.xls`→`xlsx`,
   имя содержит `перечень|перечен|perechen` → `perechen`, `.doc/.docx` из папки
   `карта база/` ИЛИ с именем «число + суффикс + разделитель» → кандидат `card`,
   иначе `other`. **Папка карт** выбирается как та, где больше всего кандидатов
   (`Counter` по родительским папкам) — работает и для `карта база/`, и для плоской
   `KARTA/`, не цепляя одиночные файлы с цифровыми именами из корня.
   ⚠️ Regex кандидата `_CARD_NAME_HINT` намеренно включает **запятую** в разделители:
   имена вида `014,Бўлинма….doc` — реальная опечатка клиента; без запятой терялись
   2 РМ (000014, 000055 — было 112 вместо 114).
3. **Конвертация + извлечение** — по каждой карте: `convert_to_docx` →
   `extract_card()`. Прогресс: `"Конвертация и извлечение карт {i}/{total}"`
   (эту строку UI парсит в процент). Ошибка конвертации или отсутствие в содержимом
   строки `…-сонли` → карта попадает в список «НЕ РАСПОЗНАНЫ». Реквизиты компании
   берутся из **первой удачной** карты (они одинаковы во всех).
4. **Громкие предупреждения**: `«Карты: найдено файлов N, распознано рабочих мест M.
   НЕ РАСПОЗНАНЫ (k): …»` и отдельно — дубликаты номеров РМ.
5. **«Перечень»** — кандидатов может быть несколько (напр. лист подписей «ИМЗО
   ПЕРЕЧЕН»); каждый парсится, **выигрывает давший больше записей**. Результат:
   `{workplace_no: {job_code, position}}`.
6. **Шаг 4.5 — наследование «а»-суффиксов** (`_fill_suffix_from_base`): если у
   суффиксной карты пуст раздел факторов — копируются поля из базового номера
   (`_SUFFIX_COPY_FIELDS`: factors, substances, benefits, ppe_provided, injury_risk,
   privileged_pension, employees_count, female_count, injury_risk_class_6_4,
   ppe_status_6_4, ppe_not_envisaged_6_4) + флаг `copied_from_base`; если карты нет
   вовсе, но номер есть в «Перечне» — запись синтезируется deepcopy базовой
   (+ флаги `copied_from_base`, `card_missing`). Оба случая — с warning.
7. **Слияние с «Перечнем»**: `rec["job_code"]`, `rec["position_from_perechen"]`;
   при `fold(должность_карты) != fold(должность_перечня)` → флаг `position_mismatch`
   (код НЕ подменяется молча); нет записи в Перечне → флаг `job_code_missing`.
   Заодно: пустой `substances` → флаг `substances_missing`.
8. **Шаг 5.05 — позиции для 6_4**: `parse_perechen_positions_6_4()` читает строки
   «Перечня» ПО ПОРЯДКУ документа с привязкой к подразделению (строки-разделители =
   настоящие горизонтальные слияния ячеек на всю ширину). Раздача картам через
   `dict[str, deque]` — дубли номеров (две карты одной «а»-позиции на разные смены)
   раскладываются по очереди, не схлопываются. Карта получает `subdivision_6_4`,
   `employees_count_6_4`, `female_count_6_4`. Строка Перечня без карты — исключается
   из 6_4 с громким warning (сознательно НЕ синтезируется, в отличие от шага 4.5:
   там подмена подтверждена эталонами для 6_5/5_1б).
9. **Шаг 5.1 — коды для «а»-суффиксов**: «Перечень» обычно не содержит строку
   `000012а` → код и должность наследуются от базового `000012`, флаг
   `job_code_missing` снимается.
10. **Финал**: сортировка `workplace_sort_key` (`000011 < 000011а < 000011б < 000012`).

## Django-клей

### Модель `Batch` (attestation/models.py)

Одна загрузка zip. Статусы: `uploaded → processing → extracted → done` / `failed`.
Датасет хранится прямо в JSON-полях — правки оператора сохраняются, документы можно
перегенерировать без повторного разбора архива:

- `company_data: JSONField(dict)`, `extracted_data: JSONField(list)`
- `output_lang`: `cyr` (дефолт) | `lat` — язык выходных документов
- `output_5_1b / output_6_5 / output_6_4`: пути к готовым .docx **относительно MEDIA_ROOT**
- `stage`: текстовый этап для прогресс-бара (напр. `"Конвертация 12/41"`)
- `error`: **двойное назначение** — при `failed` текст исключения, при успешном
  извлечении сюда пишутся warnings пайплайна через `"\n"` (UI показывает их янтарной
  панелью «Предупреждения извлечения»)

`SourceFile` — информационная сводка файлов архива (path, kind); поля
`converted_docx_path`/`parsed` и kind `template` — заготовки, не заполняются.

### jobs.py — два этапа, один переключатель

- `process_batch(batch_id)`: статус PROCESSING → `run_pipeline` (колбэк progress
  пишет `stage` в БД точечным `queryset.update()`) → статус EXTRACTED, датасет и
  warnings в модель. Исключение → FAILED + текст в `error`. В `finally` —
  `connection.close()` (обязательно для thread-раннера).
- `generate_documents(batch_id)`: статус PROCESSING → три `render_*` с
  `lang=batch.output_lang` → статус DONE + относительные пути. Только `render_6_4`
  принимает мутируемый список `warnings` — они доклеиваются к `batch.error`.
- `_dispatch(job_name, batch_id)`: если `settings.ATTESTATION_TASK_RUNNER == "celery"` —
  `tasks.<job_name>_task.delay()`; иначе `threading.Thread(daemon=True)`. Дефолт
  раннера: `"thread"` при DEBUG, `"celery"` иначе. Daemon-поток умирает с процессом —
  при рестарте runserver батч навсегда зависает в PROCESSING (известное ограничение;
  у Celery-задач тоже нет retry).

`tasks.py` — две тонкие `@shared_task`-обёртки: `process_batch_task`,
`generate_documents_task`.

### Раскладка файлов батча

```
media/
  uploads/YYYY/MM/DD/<архив>.zip          # Batch.archive
  batches/<id>/
    work/unpacked/…                        # распакованный архив (структура сохранена)
    work/docx/…                            # результаты конвертации .doc→.docx
    5_1b.docx  6_5.docx  6_4.docx          # готовые документы
```

## Поток статусов и UI

```
upload POST → Batch(UPLOADED) → start_processing → detail-страница
  → _progress.html поллит GET /batch/<pk>/status/ каждые 1.5 с (HTMX)
  → пока PROCESSING: свежий фрагмент (stage + percent из "N/M")
  → иначе: 204 + заголовок HX-Redirect → полная перезагрузка detail
EXTRACTED → экран ревью (инлайн-правка) → POST generate → снова PROCESSING → DONE
DONE → карточки скачивания трёх документов (download отдаёт FileResponse)
```

Подробности экранов — в `06-WEB-UI.md`.
