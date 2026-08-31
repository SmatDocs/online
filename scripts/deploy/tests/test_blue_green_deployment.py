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


def _run_routing_check(
    upstream_file: Path,
    retiring_slot: str,
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            "bash",
            "-c",
            'source "$1"; confirm_slot_is_not_nginx_routed "$2" "$3" 9980 9981',
            "routing-test",
            str(DRAIN_HELPER),
            retiring_slot,
            str(upstream_file),
        ],
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

    def test_timeout_is_reported_to_the_deployment_caller(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            fake_ss = _make_fake_ss(
                Path(temp_dir),
                'echo "ESTAB connection-to-collabora"\n',
            )
            result = _run_drain(fake_ss, attempts=2, zero_checks=2)

        self.assertEqual(result.returncode, 1)
        self.assertIn("Timed out without draining blue", result.stderr)
        self.assertNotIn("blue is drained", result.stdout)

    def test_forced_retirement_guard_accepts_only_the_inactive_slot(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            upstream_file = Path(temp_dir) / "collabora-upstream.conf"
            upstream_file.write_text(
                "upstream collabora_backend { server 127.0.0.1:9981; }\n",
                encoding="utf-8",
            )
            inactive_result = _run_routing_check(upstream_file, "blue")
            active_result = _run_routing_check(upstream_file, "green")

        self.assertEqual(inactive_result.returncode, 0, inactive_result.stderr)
        self.assertEqual(inactive_result.stdout.strip(), "green")
        self.assertEqual(active_result.returncode, 1)
        self.assertIn("Nginx still routes to green", active_result.stderr)


class DeploymentContractTests(unittest.TestCase):
    def test_rolling_deploy_drains_and_retires_old_slot_by_default(self) -> None:
        script = _read("scripts/deploy/deploy_production_blue_green.sh")
        drain_helper = _read("scripts/deploy/deployment_drain.sh")

        self.assertIn('source "$DRAIN_HELPER"', script)
        self.assertIn("PRODUCTION_DRAIN_MAX_ATTEMPTS:-450", drain_helper)
        self.assertIn("PRODUCTION_SHUTDOWN_OLD_SLOT:-true", script)
        self.assertIn("PRODUCTION_FORCE_RETIRE_AFTER_DRAIN_TIMEOUT:-true", script)
        self.assertIn('prepare_target_slot "$slot" "$previous_active"', script)
        self.assertIn("retire_inactive_slot", script)
        self.assertIn(
            'retire_inactive_slot "$old_slot" "$new_slot" "replaced by $new_slot"',
            script,
        )
        self.assertIn('retire_old_slot_if_enabled "$previous_active" "$slot"', script)
        self.assertIn("pm2 save", script)
        self.assertIn(
            'confirm_slot_is_not_nginx_routed "$slot" "$UPSTREAM_FILE"',
            script,
        )
        self.assertIn("force-retired after its drain grace period", script)

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

    def test_stale_inactive_target_uses_force_retirement_path(self) -> None:
        script = _read("scripts/deploy/deploy_production_blue_green.sh")

        self.assertIn(
            'retire_inactive_slot "$slot" "$active_slot" '
            '"stale inactive target before rebuild"',
            script,
        )
        self.assertNotIn(
            "Inactive target $slot was not stopped because it still has active sessions",
            script,
        )

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
            self.assertIn(
                "export PRODUCTION_FORCE_RETIRE_AFTER_DRAIN_TIMEOUT=true",
                workflow,
            )

if __name__ == "__main__":
    unittest.main()
