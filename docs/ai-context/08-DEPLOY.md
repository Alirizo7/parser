# Инфраструктура и деплой

## Образ (Dockerfile)

Единый образ для web и worker: `python:3.12-slim` + `libreoffice-writer` +
шрифты (`fonts-dejavu`, `fonts-liberation` — для кириллицы/узбекского);
`SOFFICE_BIN=/usr/bin/soffice` в ENV. CMD: gunicorn `config.wsgi`, 3 воркера
(⚠️ без `--timeout` — таймаут 120 задаётся только в prod-compose).

## dev: docker-compose.yml

3 сервиса: `redis` (7-alpine, без пароля), `web` (migrate + gunicorn, порт
8000:8000), `worker` (`celery -A config worker --concurrency=1`). SQLite в
именованном томе `db` (`DJANGO_SQLITE_PATH=/app/db/db.sqlite3`), общий том `media`.
Примечательно: `DJANGO_DEBUG=0` даже в dev-compose.

## prod: docker-compose.prod.yml (`name: attestation`)

5 сервисов; 4 из них (db, redis, web, worker) читают `.env.prod` через `env_file`,
nginx получает только `${HTTP_PORT}` compose-интерполяцией (`--env-file` в CLI):

| Сервис | Детали |
|---|---|
| `db` | postgres:16-alpine, том `pgdata`, healthcheck `pg_isready` |
| `redis` | 7-alpine, `--requirepass` + `--appendonly yes`, том `redisdata`, healthcheck |
| `web` | migrate + collectstatic + `exec gunicorn … --workers 3 --timeout 120`; `expose: 8000` (наружу НЕ публикуется); depends_on service_healthy |
| `worker` | `celery … --concurrency=1` — **намеренно 1**: LibreOffice спайкует до ~1.8 ГБ RAM, параллельные конвертации убьют 2-ГБ сервер по OOM |
| `nginx` | 1.27-alpine; **единственный порт наружу**: `${HTTP_PORT:-80}:80` |

Имя проекта зафиксировано (`attestation`) для co-hosting нескольких стеков.

## nginx/nginx.conf

Полная замена главного конфига. Ключевое: `client_max_body_size 210m` (zip до
~200 МБ + запас); `/static/` и `/media/` — alias с `expires 30d`
(⚠️ **медиа, включая готовые .docx, раздаются без авторизации** — осознанное
упрощение MVP); `/` → proxy на `web:8000` с `proxy_read_timeout 120s`
(согласован с gunicorn `--timeout 120`). TLS нет — деплой по IP, HTTP :80.

## CI/CD (.github/workflows/deploy.yml)

- **test** (push и PR в main): сервисы postgres:16 + redis:7; env — полный набор
  (`POSTGRES_*` → тесты на Postgres, `CELERY_TASK_ALWAYS_EAGER=1`,
  `ATTESTATION_TASK_RUNNER=thread`); шаги: `manage.py check` → `migrate` → `test`.
- **deploy** (только push в main, после test): `appleboy/ssh-action` → на сервере
  `git fetch && git reset --hard origin/main && docker compose -f
  docker-compose.prod.yml --env-file .env.prod up -d --build && docker image prune -f`.
  Модель — git-pull на сервере (не registry). Секреты: `SSH_HOST`, `SSH_USER`,
  `SSH_KEY` (ed25519 деплой-пара), `SSH_PORT`, `DEPLOY_PATH`.
  ⚠️ CI-деплой НЕ запускает `deploy.sh` — подстановка IP и суперпользователь
  происходят только при первичном ручном деплое.

## scripts/deploy.sh — первичный ручной деплой (идемпотентный)

1. Требует `.env.prod` (из `.env.prod.example`).
2. Определяет IP: env `SERVER_IP` → api.ipify.org → `hostname -I`.
3. Подставляет `__SERVER_IP__` в `.env.prod` **на месте** (`sed -i`, GNU/Linux-only).
4. При отсутствии Docker — интерактивно предлагает get.docker.com.
5. `compose up -d --build` → ждёт готовности до 180 с (готов = HTTP 2xx/3xx;
   на 502/5xx продолжает поллинг).
6. Идемпотентно создаёт суперпользователя из `DJANGO_SUPERUSER_*`
   (⚠️ `get_or_create` — смена пароля в .env.prod на живом стенде НЕ применится).

## scripts/audit.sh — read-only аудит сервера

`ssh <server> 'bash -s' < scripts/audit.sh`. Секции: OS, таймзона, RAM/swap, диск,
docker, ufw, слушающие порты, sshd, fail2ban, unattended-upgrades. Вердикт:
RAM < 2000 МБ → «LibreOffice спайкует до ~1.8 ГБ; нужен swap 2 ГБ или апгрейд»;
swap < 512 МБ → рекомендация; требование «наружу только 22 и 80».

## .env.prod.example — все переменные

`HTTP_PORT` (единый источник правды порта: читают compose-интерполяция, deploy.sh),
`DJANGO_SECRET_KEY`, `DJANGO_DEBUG=0`, `DJANGO_ALLOWED_HOSTS`
(`__SERVER_IP__,localhost,127.0.0.1`), `DJANGO_CSRF_TRUSTED_ORIGINS`
(⚠️ в шаблоне порт `:8080` при `HTTP_PORT=80` — при копировании согласовать),
`DJANGO_TIME_ZONE=Asia/Tashkent`, `POSTGRES_DB/USER/PASSWORD/HOST/PORT`
(наличие `POSTGRES_DB` переключает settings на Postgres), `REDIS_PASSWORD`
(⚠️ дублируется литералом ещё в `CELERY_BROKER_URL` и `CELERY_RESULT_BACKEND` —
менять все три), `DJANGO_MEDIA_ROOT=/app/media`, `SOFFICE_BIN=/usr/bin/soffice`,
`ATTESTATION_TASK_RUNNER=celery`, `DJANGO_SUPERUSER_USERNAME/EMAIL/PASSWORD`.

## requirements.txt (все пины)

Django 5.1.4, python-docx 1.2.0, celery 5.4.0, redis 5.2.1, gunicorn 23.0.0,
psycopg[binary] 3.2.9, whitenoise 6.8.2.

## ignore-файлы

- `.gitignore`: `*.zip`, `*.doc` (фикстуры selfcheck — только локально!), `/media/`,
  `db.sqlite3`, `.env.prod` («НИКОГДА не коммитить»), локальные каталоги разведки
  (`/_explore*/`, `/карты/`).
- `.dockerignore`: все `*.docx` исключены, **кроме `!attestation/assets/*.docx`** —
  шаблоны Word обязаны попасть в образ. ⚠️ `.env.prod` не исключён — при `COPY . /app/`
  на сервере секреты попадут в слои образа.

## Подготовка Ubuntu-сервера (docs/deployment.md)

Swap 2 ГБ при RAM ≤ 2 ГБ; ufw строго в порядке `allow OpenSSH` → `allow 80/tcp` →
`enable` (чтобы не потерять SSH); fail2ban + unattended-upgrades; Docker через
get.docker.com. Почему Postgres: под 3 gunicorn-воркерами + Celery SQLite ловит
write-lock.
