#!/usr/bin/env bash
set -euo pipefail

SLOT="${1:-}"
case "$SLOT" in
  blue)
    SLOT_PORT="${PRODUCTION_BLUE_PORT:-9980}"
    ;;
  green)
    SLOT_PORT="${PRODUCTION_GREEN_PORT:-9981}"
    ;;
  *)
    echo "Usage: $0 <blue|green>" >&2
    exit 1
    ;;
esac

UPSTREAM_FILE="${UPSTREAM_FILE:-/etc/nginx/conf.d/collabora_backend_overten.conf}"
SITE_FILE="${SITE_FILE:-/etc/nginx/sites-available/docs.overtenai.com}"

as_root() {
  if [ "$(id -u)" -eq 0 ]; then
    "$@"
  else
    sudo -n "$@"
  fi
}

if ! as_root /usr/bin/grep -q '' /etc/nginx/nginx.conf >/dev/null 2>&1; then
  echo "[switch] This script needs root privileges (or passwordless sudo)." >&2
  echo "[switch] Configure NOPASSWD sudo for user 'humphry' on nginx switch commands." >&2
  exit 1
fi

as_root tee "$UPSTREAM_FILE" >/dev/null <<EOC
# http-level upstream used by docs.overtenai.com
upstream collabora_backend {
    server 127.0.0.1:${SLOT_PORT} max_fails=3 fail_timeout=10s;
    keepalive 16;
}
EOC

if [ -f "$SITE_FILE" ]; then
  if ! grep -q "proxy_pass http://collabora_backend;" "$SITE_FILE"; then
    as_root sed -i -E 's#proxy_pass http://(collabora|localhost:[0-9]+|127\.0\.0\.1:[0-9]+);#proxy_pass http://collabora_backend;#g' "$SITE_FILE"
  fi
  if ! grep -q "proxy_pass http://collabora_backend;" "$SITE_FILE"; then
    echo "[switch] Could not rewrite proxy_pass in $SITE_FILE" >&2
    exit 1
  fi
fi

as_root nginx -t
if command -v systemctl >/dev/null 2>&1; then
  as_root systemctl reload nginx
else
  as_root nginx -s reload
fi

echo "[switch] Nginx switched Collabora to ${SLOT} (port:${SLOT_PORT})"
