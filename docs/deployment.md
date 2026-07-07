# Прод-развёртывание «Авто-аттестация»

Развёртывание в Docker на одном сервере, доступ по IP (без домена, HTTP :80).

## Архитектура

Пять сервисов в одной внутренней Docker-сети (`docker-compose.prod.yml`):

| Сервис | Образ | Роль | Наружу |
|--------|-------|------|--------|
| `db` | postgres:16-alpine | БД (том `pgdata`) | нет |
| `redis` | redis:7-alpine | брокер Celery, пароль (том `redisdata`) | нет |
| `web` | build `.` | gunicorn, 3 воркера, порт 8000 | только внутри сети |
| `worker` | build `.` | Celery + LibreOffice, concurrency=1 | нет |
| `nginx` | nginx:1.27-alpine | reverse proxy, статика/медиа | **`${HTTP_PORT}` → 80** |

Наружу проброшен только один host-порт (nginx), задаётся `HTTP_PORT` (по умолчанию 80).
Postgres и Redis портов наружу не имеют. SSH (22) — на уровне ОС/фаервола.

### Co-hosting (несколько проектов на одном сервере)

Compose-проект назван `attestation` (`name:` в файле), поэтому его контейнеры, тома и
сеть изолированы и не конфликтуют с другими стеками на том же сервере. Если host-порт 80
уже занят чужим reverse-proxy, задай свободный порт через `HTTP_PORT` — тогда приложение
доступно по `http://<IP>:<HTTP_PORT>/`, а чужой стек не затрагивается.

Единый источник правды — переменная `HTTP_PORT` в `.env.prod`: её читает и compose
(интерполяция `${HTTP_PORT}` через `--env-file`), и `scripts/deploy.sh`, и job `deploy`
в CI. Поэтому автодеплой из GitHub Actions поднимает контейнеры на том же порту без
правок workflow.

Пример (порт 8080) в `.env.prod`:

```
HTTP_PORT=8080
DJANGO_CSRF_TRUSTED_ORIGINS=http://<IP>:8080
```

`DJANGO_ALLOWED_HOSTS` порт не содержит (Django сверяет только хост). Переменную окружения
`HTTP_PORT=8080 ./scripts/deploy.sh` можно передать и вручную — она перекрывает значение
из файла.

Почему Postgres, а не SQLite: под тремя gunicorn-воркерами плюс Celery SQLite ловит
write-lock. Настройки (`config/settings.py`) переключаются на Postgres автоматически,
как только задана `POSTGRES_DB`.

## Секреты и окружение

- `.env.prod` — боевые секреты. **В git не коммитится** (внесён в `.gitignore`),
  создаётся отдельно на сервере. Плейсхолдер `__SERVER_IP__` в `DJANGO_ALLOWED_HOSTS`
  и `DJANGO_CSRF_TRUSTED_ORIGINS` подставляется скриптом `scripts/deploy.sh`.
- `.env.prod.example` — шаблон со значениями `CHANGE_ME_...`, коммитится в репозиторий.

Оба файла читаются compose через `env_file:` и попадают в контейнеры как переменные
окружения. Postgres берёт из них `POSTGRES_DB/USER/PASSWORD`, Django — `DJANGO_*`,
Celery — `CELERY_*`, деплой-скрипт — `DJANGO_SUPERUSER_*`.

## Подготовка сервера (Ubuntu)

Выполняется один раз, с sudo:

1. `apt update && apt upgrade -y`
2. Если RAM ≤ 2 ГБ и нет swap — создать swap 2 ГБ (страховка от OOM: LibreOffice при
   конвертации `.doc` спайкует до ~1.8 ГБ).
3. Фаервол `ufw` — строго в таком порядке, чтобы не потерять SSH:
   `ufw allow OpenSSH` → `ufw allow 80/tcp` → `ufw enable`.
4. `fail2ban`, `unattended-upgrades`.
5. Docker + плагин compose v2 (`get.docker.com`), добавить пользователя в группу `docker`.

Проверить состояние сервера, ничего не меняя: `scripts/audit.sh`.

```
ssh <server> 'bash -s' < scripts/audit.sh
```

## Первичный деплой

На сервере, в каталоге репозитория, при наличии заполненного `.env.prod`:

```
./scripts/deploy.sh
```

Если порт 80 занят другим стеком — с указанием порта:

```
HTTP_PORT=8080 ./scripts/deploy.sh
```

Скрипт идемпотентен: определяет публичный IP (env `SERVER_IP` → api.ipify.org →
`hostname -I`), подставляет его в `.env.prod`, поднимает контейнеры
(`up -d --build`), ждёт ответа на `:${HTTP_PORT}`, создаёт суперпользователя из
`DJANGO_SUPERUSER_*` (если его ещё нет) и печатает URL.

Проверка: `http://<IP>:<HTTP_PORT>/` и `http://<IP>:<HTTP_PORT>/admin/`.

## CI/CD (GitHub Actions)

`.github/workflows/deploy.yml`:

- **test** — на каждый push/PR в `main`: Postgres + Redis как сервисы, Python 3.12,
  `pip install`, `manage.py check`, `migrate`, `test` (Celery в eager-режиме).
- **deploy** — только push в `main`, после успешного test: по SSH
  (`appleboy/ssh-action`) на сервере `git reset --hard origin/main`,
  `docker compose ... up -d --build`, `docker image prune -f`.

Секреты репозитория (Settings → Secrets and variables → Actions):

| Секрет | Значение |
|--------|----------|
| `SSH_HOST` | IP сервера |
| `SSH_USER` | пользователь SSH |
| `SSH_KEY` | приватный ключ деплой-пары (ed25519) |
| `SSH_PORT` | порт SSH (обычно 22) |
| `DEPLOY_PATH` | путь к репозиторию на сервере |

Ключ для CI генерируется отдельной парой (`ssh-keygen -t ed25519`), публичная часть
добавляется в `~/.ssh/authorized_keys` деплой-пользователя на сервере, приватная —
в секрет `SSH_KEY`.

## Диагностика

- Логи web: `docker compose -f docker-compose.prod.yml --env-file .env.prod logs web`
- Логи всех сервисов: `... logs -f`
- Статус: `... ps`
- Пересобрать один сервис: `... up -d --build web`
- `:80` не отвечает: смотреть логи `web` (миграции/collectstatic) и `nginx`.
