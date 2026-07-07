#!/usr/bin/env bash
set -u

line() { printf '%s\n' "------------------------------------------------------------"; }
have() { command -v "$1" >/dev/null 2>&1; }

line
echo "SERVER AUDIT (read-only) — $(date 2>/dev/null)"
line

echo "[OS]"
if [ -r /etc/os-release ]; then
    . /etc/os-release
    echo "  ${PRETTY_NAME:-unknown}"
fi
echo "  kernel: $(uname -r)  arch: $(uname -m)"
echo "  hostname: $(hostname 2>/dev/null)"
echo "  user: $(whoami)  uid: $(id -u)"
line

echo "[TIMEZONE]"
if have timedatectl; then
    timedatectl 2>/dev/null | grep -iE 'time zone|local time' | sed 's/^/  /'
else
    echo "  $(cat /etc/timezone 2>/dev/null || echo unknown)"
fi
line

echo "[MEMORY / SWAP]"
if have free; then
    free -h | sed 's/^/  /'
else
    grep -iE 'memtotal|swaptotal' /proc/meminfo | sed 's/^/  /'
fi
MEM_MB=$(awk '/MemTotal/ {printf "%d", $2/1024}' /proc/meminfo 2>/dev/null)
SWAP_MB=$(awk '/SwapTotal/ {printf "%d", $2/1024}' /proc/meminfo 2>/dev/null)
echo "  RAM: ${MEM_MB:-?} MB   SWAP: ${SWAP_MB:-?} MB"
line

echo "[DISK]"
df -h / 2>/dev/null | sed 's/^/  /'
line

echo "[DOCKER]"
if have docker; then
    echo "  docker: $(docker --version 2>/dev/null)"
    if docker compose version >/dev/null 2>&1; then
        echo "  compose v2: $(docker compose version 2>/dev/null | head -1)"
    else
        echo "  compose v2: NOT FOUND"
    fi
    if docker info >/dev/null 2>&1; then
        echo "  docker daemon: reachable by $(whoami)"
    else
        echo "  docker daemon: NOT reachable by $(whoami) (need root or docker group)"
    fi
else
    echo "  docker: NOT INSTALLED"
fi
line

echo "[FIREWALL]"
if have ufw; then
    ufw status verbose 2>/dev/null | sed 's/^/  /' || echo "  ufw present (status needs root)"
else
    echo "  ufw: NOT INSTALLED"
fi
if have nft; then
    echo "  nftables ruleset lines: $(nft list ruleset 2>/dev/null | wc -l) (may need root)"
fi
line

echo "[LISTENING PORTS]"
if have ss; then
    ss -tulnH 2>/dev/null | awk '{print "  "$1"  "$5}' | sort -u
else
    echo "  ss: NOT AVAILABLE"
fi
line

echo "[SSH CONFIG]"
SSHD=/etc/ssh/sshd_config
if [ -r "$SSHD" ]; then
    PRL=$(grep -iE '^\s*PermitRootLogin' "$SSHD" | tail -1)
    PWA=$(grep -iE '^\s*PasswordAuthentication' "$SSHD" | tail -1)
    echo "  PermitRootLogin: ${PRL:-default (unset)}"
    echo "  PasswordAuthentication: ${PWA:-default (unset)}"
    if [ -d /etc/ssh/sshd_config.d ]; then
        grep -RiE '^\s*(PermitRootLogin|PasswordAuthentication)' /etc/ssh/sshd_config.d 2>/dev/null | sed 's/^/  drop-in: /'
    fi
else
    echo "  $SSHD not readable without root"
fi
line

echo "[FAIL2BAN]"
if have fail2ban-client; then
    echo "  installed: yes"
    systemctl is-active fail2ban 2>/dev/null | sed 's/^/  active: /'
else
    echo "  fail2ban: NOT INSTALLED"
fi
line

echo "[UNATTENDED-UPGRADES]"
if dpkg -l unattended-upgrades 2>/dev/null | grep -q '^ii'; then
    echo "  installed: yes"
    systemctl is-enabled unattended-upgrades 2>/dev/null | sed 's/^/  enabled: /'
else
    echo "  unattended-upgrades: NOT INSTALLED"
fi
line

echo "[VERDICT]"
VERDICT_OK=1
if [ -n "${MEM_MB:-}" ] && [ "$MEM_MB" -lt 2000 ]; then
    echo "  ! RAM ${MEM_MB} MB < 2 GB — LibreOffice спайкует до ~1.8 ГБ; нужен swap 2 ГБ или апгрейд."
    VERDICT_OK=0
else
    echo "  OK RAM >= ~2 GB."
fi
if [ -n "${SWAP_MB:-}" ] && [ "$SWAP_MB" -lt 512 ]; then
    echo "  ! SWAP ${SWAP_MB} MB — при RAM ~2 ГБ рекомендуется добавить swap 2 ГБ (страховка от OOM при конвертации .doc)."
fi
if ! have docker; then
    echo "  ! Docker не установлен — deploy потребует установки."
fi
echo "  Требование прод-сети: наружу открыты только 22 и 80 (проверь [LISTENING PORTS] / [FIREWALL] выше)."
[ "$VERDICT_OK" -eq 1 ] && echo "  Базовые требования по RAM выполнены." || echo "  Есть замечания — см. выше."
line
echo "AUDIT DONE"
