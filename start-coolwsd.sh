#!/bin/bash
# start-coolwd.sh - Optimized Collabora Online startup script
#
# This script starts coolwsd with performance-optimized settings for remote server deployment.
# Copy this to /srv/apps/collabora-code/ on your server and run: ./start-coolwd.sh
#
# Performance optimizations included:
#   - logging.level=warning: reduces log I/O overhead
#   - browser_logging=false: eliminates browser→server log round-trips
#   - trace_event=false: disables tracing instrumentation
#   - num_prespawn_children=4: pre-warms Kit processes to avoid cold start delays
#
# Expected improvement: 21s → 5-10s (remaining is network latency).


set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONFIG_FILE="${COOLWSD_CONFIG_FILE:-${SCRIPT_DIR}/coolwsd.xml}"
CLIENT_PORT="${COOLWSD_PORT:-9980}"
SYS_TEMPLATE_PATH="${COOLWSD_SYS_TEMPLATE_PATH:-${SCRIPT_DIR}/systemplate}"
CHILD_ROOT_PATH="${COOLWSD_CHILD_ROOT_PATH:-${SCRIPT_DIR}/jails}"
CACHE_PATH="${COOLWSD_CACHE_PATH:-${SCRIPT_DIR}/cache}"
ADMIN_USERNAME="${COOLWSD_ADMIN_USERNAME:-admin}"
ADMIN_PASSWORD="${COOLWSD_ADMIN_PASSWORD:-admin}"
CONTENT_SECURITY_POLICY="${COOLWSD_CONTENT_SECURITY_POLICY:-frame-ancestors *;}"
LOG_LEVEL="${COOLWSD_LOG_LEVEL:-warning}"
LOG_LEVEL_STARTUP="${COOLWSD_LOG_LEVEL_STARTUP:-warning}"
NUM_PRESPAWN_CHILDREN="${COOLWSD_NUM_PRESPAWN_CHILDREN:-4}"

mkdir -p "${CHILD_ROOT_PATH}" "${CACHE_PATH}"

exec "${SCRIPT_DIR}/coolwsd" \
  --config-file="${CONFIG_FILE}" \
  --port="${CLIENT_PORT}" \
  --o:sys_template_path="${SYS_TEMPLATE_PATH}" \
  --o:child_root_path="${CHILD_ROOT_PATH}" \
  --o:cache_files.path="${CACHE_PATH}" \
  --o:storage.filesystem[@allow]=true \
  --o:ssl.enable=false \
  --o:ssl.termination=true \
  --o:admin_console.username="${ADMIN_USERNAME}" \
  --o:admin_console.password="${ADMIN_PASSWORD}" \
  --o:net.content_security_policy="${CONTENT_SECURITY_POLICY}" \
  --o:logging.level="${LOG_LEVEL}" \
  --o:logging.level_startup="${LOG_LEVEL_STARTUP}" \
  --o:logging.file[@enable]=false \
  --o:logging_ui_cmd.file[@enable]=false \
  --o:logging.protocol=false \
  --o:browser_logging=false \
  --o:trace_event[@enable]=false \
  --o:num_prespawn_children="${NUM_PRESPAWN_CHILDREN}"
