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
# - PRODUCTION_SHUTDOWN_OLD_SLOT=true|false
# - PRODUCTION_BLUE_HEALTH_URL / PRODUCTION_GREEN_HEALTH_URL
# - PRODUCTION_HEALTH_MAX_ATTEMPTS / PRODUCTION_HEALTH_SLEEP_SECONDS / PRODUCTION_HEALTH_CONSECUTIVE_SUCCESS

REPO_ROOT="${1:-$(pwd)}"
cd "$REPO_ROOT"

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "[prod] Not a git repository: $REPO_ROOT" >&2
  exit 1
fi

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

HEALTH_MAX_ATTEMPTS="${PRODUCTION_HEALTH_MAX_ATTEMPTS:-90}"
HEALTH_SLEEP_SECONDS="${PRODUCTION_HEALTH_SLEEP_SECONDS:-2}"
HEALTH_CONSECUTIVE_SUCCESS="${PRODUCTION_HEALTH_CONSECUTIVE_SUCCESS:-2}"
MAKE_JOBS="${PRODUCTION_MAKE_JOBS:-$(nproc)}"
DEFAULT_DEPLOY_REF="${PRODUCTION_DEPLOY_REF:-prod}"
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

configure_slot() {
  local slot="$1"
  set_slot_context "$slot"

  if [[ -x "$SLOT_ROOT/config.status" ]]; then
    echo "[prod] Rechecking configure in $SLOT_ROOT"
    (
      cd "$SLOT_ROOT"
      ./config.status --recheck
      ./config.status config_version.h
    )
    return
  fi

  local configure_args
  configure_args="$(extract_configure_args)"
  echo "[prod] Bootstrapping configure in $SLOT_ROOT from blue slot settings"
  (
    cd "$SLOT_ROOT"
    if [[ ! -x "./configure" ]]; then
      echo "[prod] Generating configure script in $SLOT_ROOT"
      eval "./autogen.sh ${configure_args}"
    else
      eval "./configure ${configure_args}"
    fi
    ./config.status config_version.h
  )
}

sync_slot_runtime_files() {
  local slot="$1"
  set_slot_context "$slot"

  # Fresh worktrees do not include our local deployment-only runtime files
  # until they are committed upstream, so copy the active control copies in.
  if [[ "$SLOT_ROOT" == "$REPO_ROOT" ]]; then
    return
  fi

  cp -f "$REPO_ROOT/start-coolwsd.sh" "$SLOT_ROOT/start-coolwsd.sh"
  cp -f "$REPO_ROOT/coolwsd_prod.xml" "$SLOT_ROOT/coolwsd_prod.xml"
}

build_slot() {
  local slot="$1"
  set_slot_context "$slot"

  echo "[prod] Building slot $slot in $SLOT_ROOT with make -j${MAKE_JOBS}"
  (
    cd "$SLOT_ROOT"
    make -j"${MAKE_JOBS}"
  )
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

  case "$slot" in
    blue) switch_cmd="${PRODUCTION_SWITCH_TO_BLUE_CMD:-}" ;;
    green) switch_cmd="${PRODUCTION_SWITCH_TO_GREEN_CMD:-}" ;;
    *)
      echo "[prod] Unknown slot: $slot" >&2
      exit 1
      ;;
  esac

  echo "$slot" > "$PENDING_SLOT_PATH"
  if [[ -n "$switch_cmd" ]]; then
    echo "[prod] Switching traffic to $slot using configured command"
    bash -lc "$switch_cmd"
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
}

maybe_stop_old_slot() {
  local old_slot="$1"
  local new_slot="$2"
  local shutdown_old="${PRODUCTION_SHUTDOWN_OLD_SLOT:-false}"

  if [[ -z "$old_slot" || "$old_slot" == "$new_slot" ]]; then
    return
  fi

  case "${shutdown_old,,}" in
    1|true|yes|on)
      echo "[prod] Stopping old slot: $old_slot"
      stop_slot "$old_slot"
      ;;
  esac
}

deploy_and_switch() {
  local slot="$1"
  local requested_ref="$2"
  local previous_active=""

  previous_active="$(current_active_slot)"
  deploy_slot "$slot" "$requested_ref"
  switch_slot "$slot"
  maybe_stop_old_slot "$previous_active" "$slot"
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
    deploy_slot "$FOLLOWUP_SLOT" "$DEFAULT_DEPLOY_REF"
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
    deploy_slot "$FOLLOWUP_SLOT" "$FOLLOWUP_REF"
    ;;
esac

echo "[prod] Collabora blue-green deployment complete."
