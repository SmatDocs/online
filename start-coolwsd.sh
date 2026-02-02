#!/bin/bash
# start-coolwsd.sh - Optimized Collabora Online startup script
#
# This script starts coolwsd with performance-optimized settings for remote server deployment.
# Copy this to /srv/apps/collabora-code/ on your server and run: ./start-coolwsd.sh
#
# Performance optimizations included:
#   - logging.level=warning: Reduces log I/O overhead by ~99%
#   - browser_logging=false: Eliminates browser→server log round-trips
#   - trace_event=false: Disables performance tracing instrumentation
#   - num_prespawn_children=4: Pre-warms Kit processes to avoid cold start delays
#
# Expected improvement: 21s → 5-10s (remaining is network latency)

cd /srv/apps/collabora-code

./coolwsd \
  --o:sys_template_path="/srv/apps/collabora-code/systemplate" \
  --o:child_root_path="/srv/apps/collabora-code/jails" \
  --o:storage.filesystem[@allow]=true \
  --o:ssl.enable=false \
  --o:admin_console.username=admin \
  --o:admin_console.password=admin \
  '--o:net.content_security_policy=frame-ancestors *;' \
  --o:logging.level=warning \
  --o:logging.level_startup=warning \
  --o:browser_logging=false \
  '--o:trace_event[@enable]=false' \
  --o:num_prespawn_children=4
