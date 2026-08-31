#!/usr/bin/env bash
set -euo pipefail

# Blue/green deployment for Overten's Collabora host.
#
# Current production layout assumptions:
# - blue slot root defaults to the current live repo: /home/humphry/online
# - green slot root defaults to a sibling git worktree: /home/humphry/online-green
# - each slot is started by PM2 using ecosystem.production.cjs
# - coolwsd must run with coolwsd_prod.xml, never by copying it over coolwsd.xml
#
# Modes:
# - deploy_only  : build the inactive slot, health-check it, then switch traffic
# - deploy_blue  : deploy and switch blue
# - deploy_green : deploy and switch green
# - deploy_both  : switch to the freshly built inactive slot, then refresh the other slot
# - rollback_*   : redeploy a slot from the last checkpoint and switch to it
#
# Optional env vars:
# - PRODUCTION_MODE
# - PRODUCTION_DEPLOY_REF
# - PRODUCTION_BLUE_ROOT / PRODUCTION_GREEN_ROOT
# - PRODUCTION_BLUE_PORT / PRODUCTION_GREEN_PORT
# - PRODUCTION_BLUE_PM2_NAME / PRODUCTION_GREEN_PM2_NAME
# - PRODUCTION_ACTIVE_SLOT_FILE / PRODUCTION_LAST_DEPLOYED_SLOT_FILE
# - PRODUCTION_DEPLOY_STATE_DIR
# - PRODUCTION_SWITCH_TO_BLUE_CMD / PRODUCTION_SWITCH_TO_GREEN_CMD
# - PRODUCTION_SHUTDOWN_OLD_SLOT=true|false (default: true)
# - PRODUCTION_FORCE_RETIRE_AFTER_DRAIN_TIMEOUT=true|false (default: true)
# - PRODUCTION_UPSTREAM_FILE (default: /etc/nginx/conf.d/collabora_backend_overten.conf)
# - PRODUCTION_DRAIN_MAX_ATTEMPTS (default: 450, or 15 minutes at 2 seconds)
# - PRODUCTION_DRAIN_SLEEP_SECONDS (default: 2)
# - PRODUCTION_DRAIN_CONSECUTIVE_ZERO_CHECKS (default: 3)
# - PRODUCTION_SS_BIN (default: ss; primarily useful for tests)
# - PRODUCTION_ENGINE_ASSETS_URL
# - PRODUCTION_BLUE_HEALTH_URL / PRODUCTION_GREEN_HEALTH_URL
# - PRODUCTION_HEALTH_MAX_ATTEMPTS / PRODUCTION_HEALTH_SLEEP_SECONDS / PRODUCTION_HEALTH_CONSECUTIVE_SUCCESS

REPO_ROOT="${1:-$(pwd)}"
cd "$REPO_ROOT"

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "[prod] Not a git repository: $REPO_ROOT" >&2
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DRAIN_HELPER="${PRODUCTION_DRAIN_HELPER:-$SCRIPT_DIR/deployment_drain.sh}"
if [[ ! -r "$DRAIN_HELPER" && -r "$REPO_ROOT/scripts/deploy/deployment_drain.sh" ]]; then
  DRAIN_HELPER="$REPO_ROOT/scripts/deploy/deployment_drain.sh"
fi
if [[ ! -r "$DRAIN_HELPER" ]]; then
  echo "[prod] Missing deployment drain helper: $DRAIN_HELPER" >&2
  exit 1
fi
# shellcheck source=deployment_drain.sh
source "$DRAIN_HELPER"

MODE="${PRODUCTION_MODE:-deploy_only}"
case "$MODE" in
  deploy_only|deploy_blue|deploy_green|deploy_both|rollback_blue|rollback_green|rollback_both) ;;
  deploy_and_switch)
    echo "[prod] Legacy mode deploy_and_switch detected; running deploy_only." >&2
    MODE="deploy_only"
    ;;
  *)
    echo "[prod] Invalid PRODUCTION_MODE: $MODE" >&2
    exit 1
    ;;
esac

resolve_path() {
  local path="$1"
  if [[ "$path" = /* ]]; then
    echo "$path"
  else
    echo "$REPO_ROOT/$path"
  fi
}

DEPLOY_STATE_DIR="$(resolve_path "${PRODUCTION_DEPLOY_STATE_DIR:-.deploy}")"
ACTIVE_SLOT_PATH="$(resolve_path "${PRODUCTION_ACTIVE_SLOT_FILE:-$DEPLOY_STATE_DIR/active_slot}")"
LAST_DEPLOYED_SLOT_PATH="$(resolve_path "${PRODUCTION_LAST_DEPLOYED_SLOT_FILE:-$DEPLOY_STATE_DIR/last_deployed_slot}")"
PENDING_SLOT_PATH="$(resolve_path "${PRODUCTION_PENDING_SLOT_FILE:-$DEPLOY_STATE_DIR/pending_slot}")"
mkdir -p "$DEPLOY_STATE_DIR"

BLUE_ROOT="$(resolve_path "${PRODUCTION_BLUE_ROOT:-$REPO_ROOT}")"
GREEN_ROOT="$(resolve_path "${PRODUCTION_GREEN_ROOT:-$(dirname "$REPO_ROOT")/online-green}")"
BLUE_PORT="${PRODUCTION_BLUE_PORT:-9980}"
GREEN_PORT="${PRODUCTION_GREEN_PORT:-9981}"
BLUE_PM2_NAME="${PRODUCTION_BLUE_PM2_NAME:-coolwsd-blue}"
GREEN_PM2_NAME="${PRODUCTION_GREEN_PM2_NAME:-coolwsd-green}"
UPSTREAM_FILE="${PRODUCTION_UPSTREAM_FILE:-/etc/nginx/conf.d/collabora_backend_overten.conf}"

HEALTH_MAX_ATTEMPTS="${PRODUCTION_HEALTH_MAX_ATTEMPTS:-90}"
HEALTH_SLEEP_SECONDS="${PRODUCTION_HEALTH_SLEEP_SECONDS:-2}"
HEALTH_CONSECUTIVE_SUCCESS="${PRODUCTION_HEALTH_CONSECUTIVE_SUCCESS:-2}"
MAKE_JOBS="${PRODUCTION_MAKE_JOBS:-$(nproc)}"
DEFAULT_DEPLOY_REF="${PRODUCTION_DEPLOY_REF:-prod}"
ENGINE_ASSETS_URL="${PRODUCTION_ENGINE_ASSETS_URL:-https://github.com/CollaboraOnline/online/releases/download/for-code-assets/engine-main-assets.tar.gz}"
BLUE_HEALTH_URL="${PRODUCTION_BLUE_HEALTH_URL:-http://127.0.0.1:${BLUE_PORT}/hosting/discovery}"
GREEN_HEALTH_URL="${PRODUCTION_GREEN_HEALTH_URL:-http://127.0.0.1:${GREEN_PORT}/hosting/discovery}"

slot_ref_path() {
  local slot="$1"
  echo "$DEPLOY_STATE_DIR/${slot}_ref"
}

slot_prev_ref_path() {
  local slot="$1"
  echo "$DEPLOY_STATE_DIR/${slot}_prev_ref"
}

slot_history_path() {
  local slot="$1"
  echo "$DEPLOY_STATE_DIR/${slot}_history.log"
}

read_slot_file() {
  local path="$1"
  if [[ -f "$path" ]]; then
    tr -d '[:space:]' < "$path"
  fi
}

current_active_slot() {
  local slot
  slot="$(read_slot_file "$ACTIVE_SLOT_PATH")"
  if [[ "$slot" == "blue" || "$slot" == "green" ]]; then
    echo "$slot"
    return
  fi

  slot="$(read_slot_file "$LAST_DEPLOYED_SLOT_PATH")"
  if [[ "$slot" == "blue" || "$slot" == "green" ]]; then
    echo "$slot"
  fi
}

other_slot() {
  local slot="$1"
  case "$slot" in
    blue) echo "green" ;;
    green) echo "blue" ;;
    *)
      echo "[prod] Unknown slot: $slot" >&2
      exit 1
      ;;
  esac
}

next_slot() {
  local active_slot
  active_slot="$(current_active_slot)"
  if [[ "$active_slot" == "blue" ]]; then
    echo "green"
  else
    echo "blue"
  fi
}

set_slot_context() {
  local slot="$1"
  case "$slot" in
    blue)
      SLOT_ROOT="$BLUE_ROOT"
      SLOT_PORT="$BLUE_PORT"
      SLOT_PM2_NAME="$BLUE_PM2_NAME"
      SLOT_HEALTH_URL="$BLUE_HEALTH_URL"
      ;;
    green)
      SLOT_ROOT="$GREEN_ROOT"
      SLOT_PORT="$GREEN_PORT"
      SLOT_PM2_NAME="$GREEN_PM2_NAME"
      SLOT_HEALTH_URL="$GREEN_HEALTH_URL"
      ;;
    *)
      echo "[prod] Unknown slot: $slot" >&2
      exit 1
      ;;
  esac
}

record_slot_checkpoint() {
  local slot="$1"
  local deployed_ref="$2"
  local current_ref_file prev_ref_file history_file current_ref now_utc

  current_ref_file="$(slot_ref_path "$slot")"
  prev_ref_file="$(slot_prev_ref_path "$slot")"
  history_file="$(slot_history_path "$slot")"
  current_ref=""

  if [[ -f "$current_ref_file" ]]; then
    current_ref="$(tr -d '[:space:]' < "$current_ref_file")"
  fi
  if [[ -n "$current_ref" && "$current_ref" != "$deployed_ref" ]]; then
    echo "$current_ref" > "$prev_ref_file"
  fi

  echo "$deployed_ref" > "$current_ref_file"
  now_utc="$(date -u +%Y-%m-%dT%H:%M:%SZ)"
  printf "%s %s\n" "$now_utc" "$deployed_ref" >> "$history_file"
}

checkpoint_ref_for_slot() {
  local slot="$1"
  local prev_ref_file prev_ref
  prev_ref_file="$(slot_prev_ref_path "$slot")"
  if [[ ! -f "$prev_ref_file" ]]; then
    echo "[prod] No rollback checkpoint found for $slot" >&2
    exit 1
  fi
  prev_ref="$(tr -d '[:space:]' < "$prev_ref_file")"
  if [[ -z "$prev_ref" ]]; then
    echo "[prod] Rollback checkpoint file is empty for $slot" >&2
    exit 1
  fi
  echo "$prev_ref"
}

resolve_deploy_ref() {
  local requested_ref="$1"
  local resolved_ref=""

  if [[ -z "$requested_ref" ]]; then
    requested_ref="$DEFAULT_DEPLOY_REF"
  fi

  echo "[prod] Fetching refs..." >&2
  git -C "$REPO_ROOT" fetch origin --prune --tags

  if git -C "$REPO_ROOT" rev-parse --verify --quiet "origin/${requested_ref}^{commit}" >/dev/null; then
    resolved_ref="$(git -C "$REPO_ROOT" rev-parse "origin/${requested_ref}^{commit}")"
  elif git -C "$REPO_ROOT" rev-parse --verify --quiet "${requested_ref}^{commit}" >/dev/null; then
    resolved_ref="$(git -C "$REPO_ROOT" rev-parse "${requested_ref}^{commit}")"
  else
    echo "[prod] Could not resolve deploy ref: ${requested_ref}" >&2
    exit 1
  fi

  echo "$resolved_ref"
}

ensure_slot_checkout() {
  local slot="$1"
  local resolved_ref="$2"

  set_slot_context "$slot"

  if [[ "$slot" == "green" && ! -e "$SLOT_ROOT/.git" ]]; then
    echo "[prod] Creating green worktree at $SLOT_ROOT"
    git -C "$REPO_ROOT" worktree prune
    if [[ -e "$SLOT_ROOT" ]]; then
      echo "[prod] Green slot path exists but is not a git worktree: $SLOT_ROOT" >&2
      exit 1
    fi
    git -C "$REPO_ROOT" worktree add --detach "$SLOT_ROOT" "$resolved_ref" >/dev/null
    return
  fi

  git -C "$SLOT_ROOT" fetch origin --prune --tags >/dev/null 2>&1 || true
  git -C "$SLOT_ROOT" checkout --detach "$resolved_ref" >/dev/null
}

extract_configure_args() {
  local config_status_file="$BLUE_ROOT/config.status"
  local configure_args=""

  if [[ ! -f "$config_status_file" ]]; then
    echo "[prod] Cannot bootstrap configure: missing $config_status_file" >&2
    exit 1
  fi

  configure_args="$(sed -n "s/^ac_cs_config='\\(.*\\)'$/\\1/p" "$config_status_file" | head -n1)"
  if [[ -z "$configure_args" ]]; then
    echo "[prod] Failed to extract configure arguments from $config_status_file" >&2
    exit 1
  fi

  printf '%s' "$configure_args"
}

slot_uses_monorepo_engine() {
  local slot_root="$1"
  [[ -d "$slot_root/engine/include" ]]
}

ensure_engine_assets() {
  local slot="$1"
  local archive_path=""
  local tmp_archive_path=""

  set_slot_context "$slot"
  if ! slot_uses_monorepo_engine "$SLOT_ROOT"; then
    return
  fi

  if [[ -x "$SLOT_ROOT/engine/instdir/program/soffice.bin" ]]; then
    return
  fi

  archive_path="$SLOT_ROOT/engine-main-assets.tar.gz"
  tmp_archive_path="${archive_path}.tmp"
  echo "[prod] Preparing monorepo engine assets for $slot in $SLOT_ROOT"

  if [[ ! -f "$archive_path" ]]; then
    if command -v curl >/dev/null 2>&1; then
      curl -fL --retry 3 --retry-delay 5 -o "$tmp_archive_path" "$ENGINE_ASSETS_URL"
    elif command -v wget >/dev/null 2>&1; then
      wget -O "$tmp_archive_path" "$ENGINE_ASSETS_URL"
    else
      echo "[prod] Cannot download engine assets: curl/wget not found" >&2
      exit 1
    fi
    mv -f "$tmp_archive_path" "$archive_path"
  fi

  tar xzf "$archive_path" -C "$SLOT_ROOT/engine"
  if [[ ! -x "$SLOT_ROOT/engine/instdir/program/soffice.bin" ]]; then
    echo "[prod] Engine assets did not create $SLOT_ROOT/engine/instdir/program/soffice.bin" >&2
    exit 1
  fi
}

configure_args_for_slot() {
  local slot="$1"
  local configure_args=""
  local normalized_args=()
  local arg=""

  set_slot_context "$slot"
  configure_args="$(extract_configure_args)"
  if ! slot_uses_monorepo_engine "$SLOT_ROOT"; then
    printf '%s' "$configure_args"
    return
  fi

  read -r -a normalized_args <<< "$configure_args"
  for arg in "${!normalized_args[@]}"; do
    case "${normalized_args[$arg]}" in
      --with-lokit-path=*|--with-lo-path=*)
        unset 'normalized_args[$arg]'
        ;;
    esac
  done
  normalized_args+=("--with-lokit-path=$SLOT_ROOT/engine/include")
  normalized_args+=("--with-lo-path=$SLOT_ROOT/engine/instdir")
  printf '%s ' "${normalized_args[@]}"
}

config_status_uses_slot_engine() {
  local slot="$1"
  set_slot_context "$slot"
  grep -F -- "--with-lokit-path=$SLOT_ROOT/engine/include" "$SLOT_ROOT/config.status" >/dev/null 2>&1 &&
    grep -F -- "--with-lo-path=$SLOT_ROOT/engine/instdir" "$SLOT_ROOT/config.status" >/dev/null 2>&1
}

run_configure_for_slot() {
  local slot="$1"
  local configure_args="$2"

  set_slot_context "$slot"
  (
    cd "$SLOT_ROOT"
    if slot_uses_monorepo_engine "$SLOT_ROOT"; then
      export COCOREPATH="$SLOT_ROOT/engine"
    fi
    if [[ ! -x "./configure" ]] ||
      { slot_uses_monorepo_engine "$SLOT_ROOT" && grep -q "LibreOfficeKit/LibreOfficeKit.h" "./configure" 2>/dev/null; }; then
      echo "[prod] Generating configure script in $SLOT_ROOT"
      ./autogen.sh
    fi
    # shellcheck disable=SC2086
    ./configure ${configure_args}
    ./config.status config_version.h
  )
}

configure_slot() {
  local slot="$1"
  local configure_args=""
  set_slot_context "$slot"

  ensure_engine_assets "$slot"

  if [[ -x "$SLOT_ROOT/config.status" ]]; then
    if slot_uses_monorepo_engine "$SLOT_ROOT" && ! config_status_uses_slot_engine "$slot"; then
      configure_args="$(configure_args_for_slot "$slot")"
      echo "[prod] Reconfiguring $SLOT_ROOT for monorepo engine assets"
      run_configure_for_slot "$slot" "$configure_args"
      return
    fi

    echo "[prod] Rechecking configure in $SLOT_ROOT"
    (
      cd "$SLOT_ROOT"
      if slot_uses_monorepo_engine "$SLOT_ROOT"; then
        export COCOREPATH="$SLOT_ROOT/engine"
      fi
      ./config.status --recheck
      ./config.status config_version.h
    )
    return
  fi

  configure_args="$(configure_args_for_slot "$slot")"
  echo "[prod] Bootstrapping configure in $SLOT_ROOT from blue slot settings"
  run_configure_for_slot "$slot" "$configure_args"
}

sync_slot_runtime_files() {
  local slot="$1"
  set_slot_context "$slot"

  # Keep the checked-out target slot as the source of truth. Older blue slots can
  # otherwise overwrite updated tracked runtime files during a rolling deploy.
  if [[ "$SLOT_ROOT" != "$REPO_ROOT" ]]; then
    local runtime_file=""
    for runtime_file in start-coolwsd.sh coolwsd_prod.xml; do
      if [[ ! -f "$SLOT_ROOT/$runtime_file" && -f "$REPO_ROOT/$runtime_file" ]]; then
        cp -f "$REPO_ROOT/$runtime_file" "$SLOT_ROOT/$runtime_file"
      fi
    done
  fi

  # Ensure coolkitconfig.xcu is available in the config dir so release builds
  # (ENABLE_DEBUG=0) pick up spell-check-off and other LO overrides.
  local configdir="/usr/local/etc/coolwsd"
  if [[ -f "$REPO_ROOT/coolkitconfig.xcu" ]]; then
    sudo -n mkdir -p "$configdir" 2>/dev/null || true
    sudo -n cp -f "$REPO_ROOT/coolkitconfig.xcu" "$configdir/coolkitconfig.xcu" 2>/dev/null || true
  fi
}

build_slot() {
  local slot="$1"
  set_slot_context "$slot"

  echo "[prod] Building slot $slot in $SLOT_ROOT with make -j${MAKE_JOBS}"
  (
    cd "$SLOT_ROOT"
    make -j"${MAKE_JOBS}"
  )

  # Overwrite the build-generated coolwsd.xml with our production config
  # so branding, WOPI, SSL termination, etc. are always used.
  if [[ -f "$SLOT_ROOT/coolwsd_prod.xml" ]]; then
    echo "[prod] Overwriting coolwsd.xml with coolwsd_prod.xml in $SLOT_ROOT"
    cp -f "$SLOT_ROOT/coolwsd_prod.xml" "$SLOT_ROOT/coolwsd.xml"
  fi
}

sync_slot_fonts() {
  local slot="$1"
  local font_src=""
  local font_dst=""

  set_slot_context "$slot"
  font_src="$SLOT_ROOT/engine/instdir/share/fonts/truetype"
  font_dst="$SLOT_ROOT/systemplate/usr/share/fonts/truetype/smartdocs-ms-fonts"

  if [[ ! -d "$font_src" ]]; then
    echo "[prod] Font source missing for slot $slot: $font_src"
    return
  fi

  echo "[prod] Syncing engine fonts into systemplate for slot $slot"
  mkdir -p "$font_dst"
  find "$font_src" -maxdepth 1 -type f \( -iname "*.ttf" -o -iname "*.otf" -o -iname "*.ttc" \) -exec cp -f {} "$font_dst/" \;

  if command -v fc-cache >/dev/null 2>&1; then
    fc-cache -f "$font_src" "$font_dst" >/dev/null 2>&1 || true
  fi
}

stop_slot_processes() {
  local slot="$1"
  set_slot_context "$slot"

  if command -v pm2 >/dev/null 2>&1; then
    pm2 delete "$SLOT_PM2_NAME" >/dev/null 2>&1 || true
    if [[ "$slot" == "blue" ]]; then
      pm2 delete coolwsd >/dev/null 2>&1 || true
    fi
  fi

  pkill -f "^${SLOT_ROOT}/coolwsd " >/dev/null 2>&1 || true
}

start_slot() {
  local slot="$1"
  set_slot_context "$slot"

  stop_slot_processes "$slot"

  echo "[prod] Starting PM2 app $SLOT_PM2_NAME for slot $slot"
  env \
    COOLWSD_CONFIG_FILE="$SLOT_ROOT/coolwsd_prod.xml" \
    COOLWSD_PORT="$SLOT_PORT" \
    COOLWSD_SYS_TEMPLATE_PATH="$SLOT_ROOT/systemplate" \
    COOLWSD_CHILD_ROOT_PATH="$SLOT_ROOT/jails" \
    COOLWSD_CACHE_PATH="$SLOT_ROOT/cache" \
    pm2 start "$SLOT_ROOT/start-coolwsd.sh" \
      --name "$SLOT_PM2_NAME" \
      --interpreter bash \
      --cwd "$SLOT_ROOT" \
      --time \
      --update-env

  pm2 save >/dev/null
}

apply_slot_file_caps() {
  local slot="$1"
  local setcap_bin=""
  local coolforkit_bin=""
  local coolmount_bin=""

  set_slot_context "$slot"

  coolforkit_bin="$SLOT_ROOT/coolforkit-caps"
  coolmount_bin="$SLOT_ROOT/coolmount"
  setcap_bin="$(command -v setcap || true)"
  if [[ -z "$setcap_bin" && -x /usr/sbin/setcap ]]; then
    setcap_bin="/usr/sbin/setcap"
  elif [[ -z "$setcap_bin" && -x /sbin/setcap ]]; then
    setcap_bin="/sbin/setcap"
  fi

  if [[ -z "$setcap_bin" ]]; then
    echo "[prod] setcap not found; cannot prepare slot $slot file capabilities" >&2
    exit 1
  fi

  if [[ ! -x "$coolforkit_bin" || ! -x "$coolmount_bin" ]]; then
    echo "[prod] Missing runtime binaries for slot $slot; expected $coolforkit_bin and $coolmount_bin" >&2
    exit 1
  fi

  echo "[prod] Applying runtime capabilities for slot $slot"
  sudo -n "$setcap_bin" cap_chown,cap_fowner,cap_sys_chroot=ep "$coolforkit_bin"
  sudo -n "$setcap_bin" cap_sys_admin=ep "$coolmount_bin"
}

check_slot_health() {
  local slot="$1"
  local url="$2"
  local consecutive_success=0

  echo "[prod] Health check for $slot: $url"
  for _ in $(seq 1 "$HEALTH_MAX_ATTEMPTS"); do
    if curl -fsS "$url" >/dev/null 2>&1; then
      consecutive_success=$((consecutive_success + 1))
      if [[ "$consecutive_success" -ge "$HEALTH_CONSECUTIVE_SUCCESS" ]]; then
        echo "[prod] Slot $slot is healthy"
        return 0
      fi
    else
      consecutive_success=0
    fi
    sleep "$HEALTH_SLEEP_SECONDS"
  done
  return 1
}

deploy_slot() {
  local slot="$1"
  local requested_ref="$2"
  local deployed_ref=""

  deployed_ref="$(resolve_deploy_ref "$requested_ref")"
  ensure_slot_checkout "$slot" "$deployed_ref"
  sync_slot_runtime_files "$slot"
  configure_slot "$slot"
  build_slot "$slot"
  sync_slot_fonts "$slot"
  apply_slot_file_caps "$slot"
  start_slot "$slot"

  set_slot_context "$slot"
  if ! check_slot_health "$slot" "$SLOT_HEALTH_URL"; then
    echo "[prod] Slot $slot failed health check after deployment" >&2
    pm2 logs "$SLOT_PM2_NAME" --lines 200 --nostream || true
    exit 1
  fi

  record_slot_checkpoint "$slot" "$deployed_ref"
  echo "$slot" > "$LAST_DEPLOYED_SLOT_PATH"
  echo "[prod] Updated checkpoints for $slot (current=$deployed_ref)"
}

switch_slot() {
  local slot="$1"
  local switch_cmd=""
  local switch_script=""

  case "$slot" in
    blue) switch_cmd="${PRODUCTION_SWITCH_TO_BLUE_CMD:-}" ;;
    green) switch_cmd="${PRODUCTION_SWITCH_TO_GREEN_CMD:-}" ;;
    *)
      echo "[prod] Unknown slot: $slot" >&2
      exit 1
      ;;
  esac

  set_slot_context "$slot"
  switch_script="$SLOT_ROOT/scripts/deploy/switch_nginx_slot.sh"

  echo "$slot" > "$PENDING_SLOT_PATH"
  if [[ -n "$switch_cmd" ]]; then
    echo "[prod] Switching traffic to $slot using configured command"
    bash -lc "$switch_cmd"
  elif [[ -x "$switch_script" ]]; then
    echo "[prod] Switching traffic to $slot using slot script"
    "$switch_script" "$slot"
  else
    echo "[prod] Switching traffic to $slot using repository script"
    "$REPO_ROOT/scripts/deploy/switch_nginx_slot.sh" "$slot"
  fi

  echo "$slot" > "$ACTIVE_SLOT_PATH"
  rm -f "$PENDING_SLOT_PATH"
}

stop_slot() {
  local slot="$1"
  set_slot_context "$slot"
  stop_slot_processes "$slot"
  if command -v pm2 >/dev/null 2>&1; then
    pm2 save >/dev/null
  fi
}

retire_inactive_slot() {
  local slot="$1"
  local expected_active_slot="$2"
  local reason="${3:-inactive after traffic switch}"
  local force_retire="${PRODUCTION_FORCE_RETIRE_AFTER_DRAIN_TIMEOUT:-true}"
  local drain_status=0
  local routed_slot=""

  set_slot_context "$slot"
  echo "[prod] Retiring inactive slot $slot ($reason)."
  wait_for_slot_connections_to_drain "$slot" "$SLOT_PORT" || drain_status=$?
  if (( drain_status != 0 )); then
    if (( drain_status != 1 )); then
      echo "[prod] Slot $slot remains online because drain inspection failed." >&2
      return 1
    fi
    case "${force_retire,,}" in
      1|true|yes|on) ;;
      *)
        echo "[prod] Slot $slot remains online because forced retirement is disabled." >&2
        return 1
        ;;
    esac
    if [[ -z "$expected_active_slot" || "$(current_active_slot)" != "$expected_active_slot" ]]; then
      echo "[prod] Refusing forced retirement because $expected_active_slot is not the recorded active slot." >&2
      return 1
    fi
    if ! routed_slot="$(confirm_slot_is_not_nginx_routed "$slot" "$UPSTREAM_FILE" "$BLUE_PORT" "$GREEN_PORT")"; then
      echo "[prod] Refusing forced retirement because inactive Nginx routing could not be proven." >&2
      return 1
    fi
    if [[ "$routed_slot" != "$expected_active_slot" ]]; then
      echo "[prod] Refusing forced retirement because Nginx routes to $routed_slot, not $expected_active_slot." >&2
      return 1
    fi
    echo "[prod] Drain grace period expired; force-retiring inactive slot $slot and disconnecting its remaining sessions." >&2
  else
    echo "[prod] Stopping drained inactive slot: $slot"
  fi

  stop_slot "$slot"
  if (( drain_status == 1 )); then
    echo "[prod] Inactive slot $slot was force-retired after its drain grace period."
  fi
}

prepare_target_slot() {
  local slot="$1"
  local active_slot="$2"

  if [[ -n "$active_slot" && "$slot" == "$active_slot" ]]; then
    echo "[prod] Refusing to rebuild active slot $slot while nginx is sending traffic to it." >&2
    echo "[prod] Use deploy_only, or explicitly deploy the inactive color." >&2
    return 1
  fi

  echo "[prod] Preparing inactive target slot $slot before build."
  retire_inactive_slot "$slot" "$active_slot" "stale inactive target before rebuild"
}

retire_old_slot_if_enabled() {
  local old_slot="$1"
  local new_slot="$2"
  local shutdown_old="${PRODUCTION_SHUTDOWN_OLD_SLOT:-true}"

  if [[ -z "$old_slot" || "$old_slot" == "$new_slot" ]]; then
    return
  fi

  case "${shutdown_old,,}" in
    1|true|yes|on)
      retire_inactive_slot "$old_slot" "$new_slot" "replaced by $new_slot"
      ;;
  esac
}

deploy_and_switch() {
  local slot="$1"
  local requested_ref="$2"
  local previous_active=""

  previous_active="$(current_active_slot)"
  prepare_target_slot "$slot" "$previous_active"
  deploy_slot "$slot" "$requested_ref"
  switch_slot "$slot"
  retire_old_slot_if_enabled "$previous_active" "$slot"
}

deploy_inactive_checkpoint() {
  local slot="$1"
  local requested_ref="$2"
  local active_slot=""

  active_slot="$(current_active_slot)"
  prepare_target_slot "$slot" "$active_slot"
  deploy_slot "$slot" "$requested_ref"
  echo "[prod] Stopping refreshed inactive slot $slot after its health check."
  stop_slot "$slot"
}

echo "[prod] Mode: $MODE"
if [[ -f "$ACTIVE_SLOT_PATH" ]]; then
  echo "[prod] Active slot: $(read_slot_file "$ACTIVE_SLOT_PATH")"
fi

case "$MODE" in
  deploy_only)
    TARGET_SLOT="$(next_slot)"
    echo "[prod] Rolling target slot: $TARGET_SLOT"
    deploy_and_switch "$TARGET_SLOT" "$DEFAULT_DEPLOY_REF"
    ;;
  deploy_blue)
    deploy_and_switch "blue" "$DEFAULT_DEPLOY_REF"
    ;;
  deploy_green)
    deploy_and_switch "green" "$DEFAULT_DEPLOY_REF"
    ;;
  deploy_both)
    TARGET_SLOT="$(next_slot)"
    FOLLOWUP_SLOT="$(other_slot "$TARGET_SLOT")"
    echo "[prod] Deploying both slots (switch to $TARGET_SLOT, then refresh $FOLLOWUP_SLOT)"
    deploy_and_switch "$TARGET_SLOT" "$DEFAULT_DEPLOY_REF"
    deploy_inactive_checkpoint "$FOLLOWUP_SLOT" "$DEFAULT_DEPLOY_REF"
    ;;
  rollback_blue)
    ROLLBACK_REF="$(checkpoint_ref_for_slot "blue")"
    echo "[prod] Rolling back blue to checkpoint: $ROLLBACK_REF"
    deploy_and_switch "blue" "$ROLLBACK_REF"
    ;;
  rollback_green)
    ROLLBACK_REF="$(checkpoint_ref_for_slot "green")"
    echo "[prod] Rolling back green to checkpoint: $ROLLBACK_REF"
    deploy_and_switch "green" "$ROLLBACK_REF"
    ;;
  rollback_both)
    TARGET_SLOT="$(next_slot)"
    FOLLOWUP_SLOT="$(other_slot "$TARGET_SLOT")"
    TARGET_REF="$(checkpoint_ref_for_slot "$TARGET_SLOT")"
    FOLLOWUP_REF="$(checkpoint_ref_for_slot "$FOLLOWUP_SLOT")"
    echo "[prod] Rolling back both slots (switch to $TARGET_SLOT, then refresh $FOLLOWUP_SLOT)"
    deploy_and_switch "$TARGET_SLOT" "$TARGET_REF"
    deploy_inactive_checkpoint "$FOLLOWUP_SLOT" "$FOLLOWUP_REF"
    ;;
esac

echo "[prod] Collabora blue-green deployment complete."
