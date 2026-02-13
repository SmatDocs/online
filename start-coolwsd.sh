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

exec "${SCRIPT_DIR}/coolwsd" \
  --config-file="${CONFIG_FILE}" \
  --o:sys_template_path="${SCRIPT_DIR}/systemplate" \
  --o:child_root_path="${SCRIPT_DIR}/jails" \
  --o:cache_files.path="${SCRIPT_DIR}/cache" \
  --o:storage.filesystem[@allow]=true \
  --o:ssl.enable=false \
  --o:ssl.termination=true \
  --o:admin_console.username=admin \
  --o:admin_console.password=admin \
  --o:net.content_security_policy="frame-ancestors *;" \
  --o:logging.level=warning \
  --o:logging.level_startup=warning \
  --o:logging.file[@enable]=false \
  --o:logging_ui_cmd.file[@enable]=false \
  --o:logging.protocol=false \
  --o:browser_logging=false \
  --o:trace_event[@enable]=false \
  --o:num_prespawn_children=4
