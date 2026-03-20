# Overten Collabora Production Blue-Green

This repository contains the deployment files for Overten's Collabora host.

## Slot Layout

- Blue root defaults to `/home/humphry/online`
- Green root defaults to `/home/humphry/online-green`
- Blue port defaults to `9980`
- Green port defaults to `9981`
- PM2 app names default to `coolwsd-blue` and `coolwsd-green`

## Runtime Config

Each slot must run with its own `coolwsd_prod.xml`.

Do not copy `coolwsd_prod.xml` over `coolwsd.xml`.
The startup path is explicit:

```bash
COOLWSD_CONFIG_FILE=/home/humphry/online/coolwsd_prod.xml /home/humphry/online/start-coolwsd.sh
```

The deploy script starts PM2 explicitly with the slot's `COOLWSD_*` variables.
The ecosystem file can still be kept as a reference, but deployment does not rely on PM2 auto-detecting it as an ecosystem config.

## Deploy Flow

`scripts/deploy/deploy_production_blue_green.sh` does the following:

1. Resolves the target git ref.
2. Deploys the inactive slot by default.
3. Bootstraps a fresh green slot using the current blue slot's stored configure arguments when needed.
   - If the fresh worktree does not yet contain a generated `configure`, the script runs `./autogen.sh` first.
   - It also copies `start-coolwsd.sh` and `coolwsd_prod.xml` from the control repo into the slot before startup.
4. Rebuilds the slot with:
   - `./config.status --recheck`
   - `./config.status config_version.h`
   - `make -j${PRODUCTION_MAKE_JOBS:-$(nproc)}`
5. Starts the slot through PM2 with `coolwsd_prod.xml`.
6. Health-checks `/hosting/discovery`.
7. Switches `docs.overten.ai` to the selected slot through Nginx.

## First Server Cleanup

The first blue deployment is also the normalization step:

- replace the legacy PM2 app name `coolwsd` with `coolwsd-blue`
- stop stray `coolwsd` processes running from `/home/humphry/online`
- keep the active slot tracked in `.deploy/active_slot`

## Required Automation Setup

For GitHub Actions to switch Nginx non-interactively, the Collabora host user needs passwordless sudo for the Nginx switch commands.

If you do not want to grant that, provide `PRODUCTION_SWITCH_TO_BLUE_CMD` and `PRODUCTION_SWITCH_TO_GREEN_CMD` externally and handle the switch outside the repository script.

## Low-RAM Hosts

If the host is memory-constrained, set `PRODUCTION_MAKE_JOBS` lower than `$(nproc)`.
On a small Droplet, `PRODUCTION_MAKE_JOBS=2` is a safer starting point than building with all cores.
