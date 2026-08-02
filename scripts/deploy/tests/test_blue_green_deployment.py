"""Contract tests for safe Collabora blue/green process retirement."""

from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[3]
DRAIN_HELPER = REPO_ROOT / "scripts/deploy/deployment_drain.sh"


def _read(relative_path: str) -> str:
    return (REPO_ROOT / relative_path).read_text(encoding="utf-8")


def _make_fake_ss(directory: Path, body: str) -> Path:
    fake_ss = directory / "fake-ss"
    fake_ss.write_text(
        "#!/usr/bin/env bash\nset -euo pipefail\n" + body,
        encoding="utf-8",
    )
    fake_ss.chmod(0o755)
    return fake_ss


def _run_drain(
    fake_ss: Path,
    *,
    attempts: int,
    zero_checks: int,
) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env.update(
        {
            "PRODUCTION_SS_BIN": str(fake_ss),
            "PRODUCTION_DRAIN_MAX_ATTEMPTS": str(attempts),
            "PRODUCTION_DRAIN_SLEEP_SECONDS": "0",
            "PRODUCTION_DRAIN_CONSECUTIVE_ZERO_CHECKS": str(zero_checks),
        }
    )
    return subprocess.run(
        [
            "bash",
            "-c",
            'source "$1"; wait_for_slot_connections_to_drain blue 9980',
            "drain-test",
            str(DRAIN_HELPER),
        ],
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )


class DrainHelperTests(unittest.TestCase):
    def test_requires_consecutive_zero_connection_checks(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            fake_ss = _make_fake_ss(Path(temp_dir), "exit 0\n")
            result = _run_drain(fake_ss, attempts=3, zero_checks=3)

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("zero established connections (3/3)", result.stdout)
        self.assertIn("blue is drained", result.stdout)

    def test_timeout_leaves_the_slot_running(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            fake_ss = _make_fake_ss(
                Path(temp_dir),
                'echo "ESTAB connection-to-collabora"\n',
            )
            result = _run_drain(fake_ss, attempts=2, zero_checks=2)

        self.assertEqual(result.returncode, 1)
        self.assertIn("PM2 process must remain running", result.stderr)
        self.assertNotIn("blue is drained", result.stdout)


class DeploymentContractTests(unittest.TestCase):
    def test_rolling_deploy_drains_and_retires_old_slot_by_default(self) -> None:
        script = _read("scripts/deploy/deploy_production_blue_green.sh")

        self.assertIn('source "$DRAIN_HELPER"', script)
        self.assertIn("PRODUCTION_SHUTDOWN_OLD_SLOT:-true", script)
        self.assertIn('prepare_target_slot "$slot" "$previous_active"', script)
        self.assertIn(
            'wait_for_slot_connections_to_drain "$old_slot" "$SLOT_PORT"',
            script,
        )
        self.assertIn('retire_old_slot_if_enabled "$previous_active" "$slot"', script)
        self.assertIn("pm2 save", script)

        prepare_position = script.index(
            'prepare_target_slot "$slot" "$previous_active"'
        )
        deploy_position = script.index(
            'deploy_slot "$slot" "$requested_ref"', prepare_position
        )
        switch_position = script.index('switch_slot "$slot"', deploy_position)
        retire_position = script.index(
            'retire_old_slot_if_enabled "$previous_active" "$slot"',
            switch_position,
        )
        self.assertLess(prepare_position, deploy_position)
        self.assertLess(deploy_position, switch_position)
        self.assertLess(switch_position, retire_position)

    def test_deploy_both_still_finishes_with_inactive_slot_stopped(self) -> None:
        script = _read("scripts/deploy/deploy_production_blue_green.sh")

        self.assertIn("deploy_inactive_checkpoint", script)
        self.assertEqual(
            script.count(
                'deploy_inactive_checkpoint "$FOLLOWUP_SLOT" "$DEFAULT_DEPLOY_REF"'
            ),
            1,
        )
        self.assertEqual(
            script.count(
                'deploy_inactive_checkpoint "$FOLLOWUP_SLOT" "$FOLLOWUP_REF"'
            ),
            1,
        )
        self.assertIn(
            'echo "[prod] Stopping refreshed inactive slot $slot after its health check."',
            script,
        )

    def test_github_workflows_ship_helper_and_enable_retirement(self) -> None:
        for relative_path in (
            ".github/workflows/production-deploy.yml",
            ".github/workflows/production-deploy-sa.yml",
        ):
            workflow = _read(relative_path)
            self.assertIn(
                'git show "${SCRIPT_REF}:scripts/deploy/deployment_drain.sh"',
                workflow,
            )
            self.assertIn("export PRODUCTION_SHUTDOWN_OLD_SLOT=true", workflow)

if __name__ == "__main__":
    unittest.main()
