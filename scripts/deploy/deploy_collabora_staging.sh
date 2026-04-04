#!/usr/bin/env bash
set -euo pipefail

# Collabora staging deployment.
#
# Builds and runs coolwsd from the staging branch in a separate worktree,
# independent of the blue/green production slots.
#
# Layout:
#   - Blue (production):  /home/humphry/online        port 9980
#   - Green (production): /home/humphry/online-green   port 9981
#   - Staging:            /home/humphry/online-staging  port 9982
#
# Usage:
#   ./scripts/deploy/deploy_collabora_staging.sh /home/humphry/online
#
# Optional env vars:
#   STAGING_DEPLOY_REF          git ref to deploy (default: staging)
#   STAGING_PORT                port for coolwsd (default: 9982)
#   STAGING_PM2_NAME            PM2 app name (default: coolwsd-staging)
#   STAGING_ROOT                worktree path (default: <repo>/../online-staging)
#   STAGING_HEALTH_URL          health check URL
#   STAGING_HEALTH_MAX_ATTEMPTS max health check attempts (default: 60)
#   STAGING_HEALTH_SLEEP_SECONDS sleep between attempts (default: 2)
#   STAGING_MAKE_JOBS           parallel make jobs (default: nproc)

REPO_ROOT="${1:-$(pwd)}"
cd "$REPO_ROOT"

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "[staging] Not a git repository: $REPO_ROOT" >&2
  exit 1
fi

BLUE_ROOT="$REPO_ROOT"
STAGING_ROOT="${STAGING_ROOT:-$(dirname "$REPO_ROOT")/online-staging}"
STAGING_PORT="${STAGING_PORT:-9982}"
STAGING_PM2_NAME="${STAGING_PM2_NAME:-coolwsd-staging}"
STAGING_DEPLOY_REF="${STAGING_DEPLOY_REF:-staging}"
STAGING_HEALTH_URL="${STAGING_HEALTH_URL:-http://127.0.0.1:${STAGING_PORT}/hosting/discovery}"
HEALTH_MAX_ATTEMPTS="${STAGING_HEALTH_MAX_ATTEMPTS:-60}"
HEALTH_SLEEP_SECONDS="${STAGING_HEALTH_SLEEP_SECONDS:-2}"
MAKE_JOBS="${STAGING_MAKE_JOBS:-$(nproc)}"

# --- helpers ----------------------------------------------------------------

extract_configure_args() {
  local config_status_file="$BLUE_ROOT/config.status"
  if [[ ! -f "$config_status_file" ]]; then
    echo "[staging] Cannot bootstrap configure: missing $config_status_file" >&2
    exit 1
  fi
  sed -n "s/^ac_cs_config='\\(.*\\)'$/\\1/p" "$config_status_file" | head -n1
}

# --- resolve ref ------------------------------------------------------------

echo "[staging] Fetching refs..."
git -C "$REPO_ROOT" fetch origin --prune --tags

RESOLVED_REF=""
if git -C "$REPO_ROOT" rev-parse --verify --quiet "origin/${STAGING_DEPLOY_REF}^{commit}" >/dev/null; then
  RESOLVED_REF="$(git -C "$REPO_ROOT" rev-parse "origin/${STAGING_DEPLOY_REF}^{commit}")"
elif git -C "$REPO_ROOT" rev-parse --verify --quiet "${STAGING_DEPLOY_REF}^{commit}" >/dev/null; then
  RESOLVED_REF="$(git -C "$REPO_ROOT" rev-parse "${STAGING_DEPLOY_REF}^{commit}")"
else
  echo "[staging] Could not resolve deploy ref: ${STAGING_DEPLOY_REF}" >&2
  exit 1
fi

echo "[staging] Deploying ref: ${RESOLVED_REF} (from ${STAGING_DEPLOY_REF})"

# --- ensure worktree --------------------------------------------------------

if [[ ! -e "$STAGING_ROOT/.git" ]]; then
  echo "[staging] Creating staging worktree at $STAGING_ROOT"
  git -C "$REPO_ROOT" worktree prune
  if [[ -e "$STAGING_ROOT" ]]; then
    echo "[staging] Path exists but is not a git worktree: $STAGING_ROOT" >&2
    exit 1
  fi
  git -C "$REPO_ROOT" worktree add --detach "$STAGING_ROOT" "$RESOLVED_REF" >/dev/null
else
  git -C "$STAGING_ROOT" fetch origin --prune --tags >/dev/null 2>&1 || true
  git -C "$STAGING_ROOT" checkout --detach "$RESOLVED_REF" >/dev/null
fi

# --- sync runtime files -----------------------------------------------------

cp -f "$REPO_ROOT/start-coolwsd.sh" "$STAGING_ROOT/start-coolwsd.sh"
cp -f "$REPO_ROOT/coolwsd_prod.xml" "$STAGING_ROOT/coolwsd_prod.xml"

# Ensure coolkitconfig.xcu is in the config dir for release builds (spell check off, etc.)
CONFIGDIR="/usr/local/etc/coolwsd"
if [[ -f "$REPO_ROOT/coolkitconfig.xcu" ]]; then
  sudo -n mkdir -p "$CONFIGDIR" 2>/dev/null || true
  sudo -n cp -f "$REPO_ROOT/coolkitconfig.xcu" "$CONFIGDIR/coolkitconfig.xcu" 2>/dev/null || true
fi

# --- configure --------------------------------------------------------------

if [[ -x "$STAGING_ROOT/config.status" ]]; then
  echo "[staging] Rechecking configure in $STAGING_ROOT"
  ( cd "$STAGING_ROOT" && ./config.status --recheck && ./config.status config_version.h )
else
  configure_args="$(extract_configure_args)"
  echo "[staging] Bootstrapping configure from blue slot settings"
  (
    cd "$STAGING_ROOT"
    if [[ ! -x "./configure" ]]; then
      eval "./autogen.sh ${configure_args}"
    else
      eval "./configure ${configure_args}"
    fi
    ./config.status config_version.h
  )
fi

# --- build ------------------------------------------------------------------

echo "[staging] Building with make -j${MAKE_JOBS}"
( cd "$STAGING_ROOT" && make -j"${MAKE_JOBS}" )

# Overwrite the build-generated coolwsd.xml with our production config
# so branding, WOPI, SSL termination, etc. are always used.
if [[ -f "$STAGING_ROOT/coolwsd_prod.xml" ]]; then
  echo "[staging] Overwriting coolwsd.xml with coolwsd_prod.xml"
  cp -f "$STAGING_ROOT/coolwsd_prod.xml" "$STAGING_ROOT/coolwsd.xml"
fi

# --- file capabilities ------------------------------------------------------

setcap_bin="$(command -v setcap || true)"
if [[ -z "$setcap_bin" ]]; then
  echo "[staging] setcap not found; cannot set file capabilities" >&2
  exit 1
fi

echo "[staging] Applying runtime capabilities"
sudo -n "$setcap_bin" cap_chown,cap_fowner,cap_sys_chroot=ep "$STAGING_ROOT/coolforkit-caps"
sudo -n "$setcap_bin" cap_sys_admin=ep "$STAGING_ROOT/coolmount"

# --- stop old instance ------------------------------------------------------

if command -v pm2 >/dev/null 2>&1; then
  pm2 delete "$STAGING_PM2_NAME" >/dev/null 2>&1 || true
fi
pkill -f "^${STAGING_ROOT}/coolwsd " >/dev/null 2>&1 || true

# --- start ------------------------------------------------------------------

echo "[staging] Starting PM2 app $STAGING_PM2_NAME on port $STAGING_PORT"
env \
  COOLWSD_CONFIG_FILE="$STAGING_ROOT/coolwsd_prod.xml" \
  COOLWSD_PORT="$STAGING_PORT" \
  COOLWSD_SYS_TEMPLATE_PATH="$STAGING_ROOT/systemplate" \
  COOLWSD_CHILD_ROOT_PATH="$STAGING_ROOT/jails" \
  COOLWSD_CACHE_PATH="$STAGING_ROOT/cache" \
  pm2 start "$STAGING_ROOT/start-coolwsd.sh" \
    --name "$STAGING_PM2_NAME" \
    --interpreter bash \
    --cwd "$STAGING_ROOT" \
    --time \
    --update-env

pm2 save >/dev/null

# --- health check -----------------------------------------------------------

echo "[staging] Health check: $STAGING_HEALTH_URL"
consecutive_ok=0
for _ in $(seq 1 "$HEALTH_MAX_ATTEMPTS"); do
  if curl -fsS "$STAGING_HEALTH_URL" >/dev/null 2>&1; then
    consecutive_ok=$((consecutive_ok + 1))
    if [[ "$consecutive_ok" -ge 2 ]]; then
      echo "[staging] Healthy"
      echo "[staging] Collabora staging deployment complete (port $STAGING_PORT)."
      exit 0
    fi
  else
    consecutive_ok=0
  fi
  sleep "$HEALTH_SLEEP_SECONDS"
done

echo "[staging] Health check failed after ${HEALTH_MAX_ATTEMPTS} attempts" >&2
pm2 logs "$STAGING_PM2_NAME" --lines 200 --nostream || true
exit 1
