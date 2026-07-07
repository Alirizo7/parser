#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")/.."

COMPOSE="docker compose -f docker-compose.prod.yml --env-file .env.prod"
ENV_FILE=".env.prod"
HTTP_PORT="${HTTP_PORT:-80}"
export HTTP_PORT

if [ "$HTTP_PORT" = "80" ]; then
    URL_SUFFIX=""
else
    URL_SUFFIX=":${HTTP_PORT}"
fi

log() { printf '\n\033[1;34m==>\033[0m %s\n' "$1"; }
err() { printf '\n\033[1;31mERROR:\033[0m %s\n' "$1" >&2; }

if [ ! -f "$ENV_FILE" ]; then
    err "$ENV_FILE не найден. Создай его из .env.prod.example с боевыми секретами."
    exit 1
fi

log "Определяю публичный IP сервера"
SERVER_IP="${SERVER_IP:-}"
if [ -z "$SERVER_IP" ]; then
    SERVER_IP="$(curl -fsS --max-time 10 https://api.ipify.org 2>/dev/null || true)"
fi
if [ -z "$SERVER_IP" ]; then
    SERVER_IP="$(hostname -I 2>/dev/null | awk '{print $1}')"
fi
if [ -z "$SERVER_IP" ]; then
    err "Не удалось определить IP. Задай его через переменную окружения: SERVER_IP=1.2.3.4 ./scripts/deploy.sh"
    exit 1
fi
echo "    IP: $SERVER_IP"

if grep -q '__SERVER_IP__' "$ENV_FILE"; then
    log "Подставляю IP в $ENV_FILE вместо плейсхолдера __SERVER_IP__"
    sed -i "s/__SERVER_IP__/${SERVER_IP}/g" "$ENV_FILE"
else
    echo "    Плейсхолдер уже подставлен ранее — пропускаю."
fi

if ! command -v docker >/dev/null 2>&1; then
    err "Docker не установлен."
    printf 'Установить Docker сейчас через get.docker.com? [y/N] '
    read -r ans || ans=""
    case "$ans" in
        [yY]|[yY][eE][sS])
            curl -fsSL https://get.docker.com | sh
            ;;
        *)
            echo "Отменено. Установи Docker вручную и запусти скрипт снова."
            exit 1
            ;;
    esac
fi

if ! docker compose version >/dev/null 2>&1; then
    err "docker compose v2 недоступен. Установи плагин docker-compose-plugin."
    exit 1
fi

log "Собираю и запускаю контейнеры"
$COMPOSE up -d --build

log "Жду ответа на http://localhost${URL_SUFFIX}/ (до 180с)"
ok=0
for i in $(seq 1 60); do
    code="$(curl -fsS -o /dev/null -w '%{http_code}' --max-time 5 "http://localhost${URL_SUFFIX}/" 2>/dev/null || true)"
    if [ -n "$code" ] && [ "$code" != "000" ]; then
        echo "    HTTP $code"
        ok=1
        break
    fi
    sleep 3
done
if [ "$ok" -ne 1 ]; then
    err "Сервис не ответил на :80. Логи web:"
    $COMPOSE logs --tail=80 web >&2 || true
    exit 1
fi

log "Идемпотентно создаю суперпользователя"
$COMPOSE exec -T web python manage.py shell -c "
import os
from django.contrib.auth import get_user_model
U = get_user_model()
name = os.environ['DJANGO_SUPERUSER_USERNAME']
email = os.environ.get('DJANGO_SUPERUSER_EMAIL', '')
password = os.environ['DJANGO_SUPERUSER_PASSWORD']
u, created = U.objects.get_or_create(username=name, defaults={'email': email})
if created:
    u.set_password(password)
    u.is_staff = True
    u.is_superuser = True
    u.save()
    print('superuser created:', name)
else:
    print('superuser already exists:', name)
"

log "Готово"
echo "    Приложение: http://${SERVER_IP}${URL_SUFFIX}/"
echo "    Админка:    http://${SERVER_IP}${URL_SUFFIX}/admin/"
