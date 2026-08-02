#!/usr/bin/env bash

# Sourceable helpers for waiting until nginx has released every established
# connection to an inactive Collabora slot. The caller decides whether a
# timeout is fatal; this file never stops a PM2 process itself.

is_non_negative_number() {
  [[ "${1:-}" =~ ^[0-9]+([.][0-9]+)?$ ]]
}

count_established_slot_connections() {
  local port="$1"
  local ss_bin="${PRODUCTION_SS_BIN:-ss}"
  local output=""

  if ! command -v "$ss_bin" >/dev/null 2>&1; then
    echo "[drain] Cannot inspect connections because '$ss_bin' is unavailable." >&2
    return 2
  fi
  if [[ ! "$port" =~ ^[0-9]+$ ]] || (( port < 1 || port > 65535 )); then
    echo "[drain] Invalid Collabora slot port: $port" >&2
    return 2
  fi
  if ! output="$("$ss_bin" -Htn state established "( sport = :$port )")"; then
    echo "[drain] Failed to inspect established connections on port $port." >&2
    return 2
  fi

  if [[ -z "$output" ]]; then
    echo 0
  else
    awk 'END { print NR }' <<< "$output"
  fi
}

wait_for_slot_connections_to_drain() {
  local slot="$1"
  local port="$2"
  # Collabora editors can keep WebSockets open much longer than an API
  # request. Thirty minutes fits inside the one-hour deployment timeout while
  # still refusing to force-disconnect a long-running editing session.
  local max_attempts="${PRODUCTION_DRAIN_MAX_ATTEMPTS:-900}"
  local sleep_seconds="${PRODUCTION_DRAIN_SLEEP_SECONDS:-2}"
  local consecutive_required="${PRODUCTION_DRAIN_CONSECUTIVE_ZERO_CHECKS:-3}"
  local attempt=0
  local connection_count=0
  local consecutive_zero=0

  if [[ ! "$max_attempts" =~ ^[1-9][0-9]*$ ]]; then
    echo "[drain] PRODUCTION_DRAIN_MAX_ATTEMPTS must be a positive integer." >&2
    return 2
  fi
  if ! is_non_negative_number "$sleep_seconds"; then
    echo "[drain] PRODUCTION_DRAIN_SLEEP_SECONDS must be a non-negative number." >&2
    return 2
  fi
  if [[ ! "$consecutive_required" =~ ^[1-9][0-9]*$ ]]; then
    echo "[drain] PRODUCTION_DRAIN_CONSECUTIVE_ZERO_CHECKS must be a positive integer." >&2
    return 2
  fi

  echo "[drain] Waiting for $slot connections on port $port to reach zero."
  for ((attempt = 1; attempt <= max_attempts; attempt++)); do
    if ! connection_count="$(count_established_slot_connections "$port")"; then
      return 2
    fi

    if (( connection_count == 0 )); then
      consecutive_zero=$((consecutive_zero + 1))
      echo "[drain] $slot has zero established connections (${consecutive_zero}/${consecutive_required})."
      if (( consecutive_zero >= consecutive_required )); then
        echo "[drain] $slot is drained."
        return 0
      fi
    else
      consecutive_zero=0
      echo "[drain] $slot still has $connection_count established connection(s)."
    fi

    if (( attempt < max_attempts )); then
      sleep "$sleep_seconds"
    fi
  done

  echo "[drain] Timed out without draining $slot; its PM2 process must remain running." >&2
  return 1
}
