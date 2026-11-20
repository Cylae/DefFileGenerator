"""
Smoke and Integration Tests for deffilegen CLI Subcommands.
Verifies that all primary and versioning CLI subcommands execute cleanly:
- deffilegen --help
- deffilegen history (text & --json)
- deffilegen duplicates --path production (text & --json)
- deffilegen diff (default HEAD vs parent, explicit commit, --json)
- deffilegen commit (default snapshot message & custom -m)
- deffilegen restore --dry-run (default HEAD & explicit commit)
"""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator._version import __version__

REPO_ROOT = Path(__file__).resolve().parent.parent.parent


def _run_cli_cmd(args: list[str], cwd: Path | str | None = None) -> subprocess.CompletedProcess[str]:
    """Helper to run deffilegen via module invocation."""
    cmd = [sys.executable, "-m", "DefFileGenerator.main"] + args
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=str(cwd or REPO_ROOT),
    )


class TestCliSmoke(unittest.TestCase):
    """End-to-end smoke verification for deffilegen CLI commands."""

    def test_cli_help(self) -> None:
        """Verify deffilegen --help returns 0 and lists all subcommands."""
        res = _run_cli_cmd(["--help"])
        self.assertEqual(res.returncode, 0, f"deffilegen --help failed:\n{res.stderr}")
        self.assertIn("usage: deffilegen", res.stdout)
        self.assertIn("history", res.stdout)
        self.assertIn("duplicates", res.stdout)
        self.assertIn("diff", res.stdout)
        self.assertIn("commit", res.stdout)
        self.assertIn("restore", res.stdout)

    def test_cli_version(self) -> None:
        """Verify deffilegen --version returns version string."""
        res = _run_cli_cmd(["--version"])
        self.assertEqual(res.returncode, 0)
        self.assertIn(f"deffilegen {__version__}", res.stdout)

    def test_cli_history_production(self) -> None:
        """Verify deffilegen history runs cleanly on existing store."""
        res = _run_cli_cmd(["history"])
        self.assertEqual(res.returncode, 0, f"deffilegen history failed:\n{res.stderr}")
        self.assertIn("DEFINITION VERSION HISTORY", res.stdout)

    def test_cli_history_json(self) -> None:
        """Verify deffilegen history --json produces valid JSON."""
        res = _run_cli_cmd(["history", "--json", "-n", "2"])
        self.assertEqual(res.returncode, 0, f"deffilegen history --json failed:\n{res.stderr}")
        parsed = json.loads(res.stdout)
        self.assertIsInstance(parsed, list)
        if parsed:
            self.assertIn("commit_id", parsed[0])
            self.assertIn("author", parsed[0])

    def test_cli_duplicates_production(self) -> None:
        """Verify deffilegen duplicates --path production runs cleanly."""
        res = _run_cli_cmd(["duplicates", "--path", "production"])
        self.assertEqual(res.returncode, 0, f"deffilegen duplicates failed:\n{res.stderr}")
        self.assertIn("DUPLICATE DEFINITION AUDIT", res.stdout)
        self.assertIn("Total Files Scanned", res.stdout)

    def test_cli_duplicates_json(self) -> None:
        """Verify deffilegen duplicates --json produces valid JSON."""
        res = _run_cli_cmd(["duplicates", "--path", "production", "--json"])
        self.assertEqual(res.returncode, 0, f"deffilegen duplicates --json failed:\n{res.stderr}")
        parsed = json.loads(res.stdout)
        self.assertIn("total_files_scanned", parsed)
        self.assertIn("exact_hash_clusters", parsed)

    def test_cli_diff_default_head(self) -> None:
        """Verify deffilegen diff with no positional args diffs HEAD against parent."""
        res = _run_cli_cmd(["diff"])
        self.assertEqual(res.returncode, 0, f"deffilegen diff failed:\n{res.stderr}")
        self.assertIn("DEFINITION DIFF", res.stdout)
        self.assertIn("Files Added", res.stdout)

    def test_cli_diff_json(self) -> None:
        """Verify deffilegen diff --json produces structured diff output."""
        res = _run_cli_cmd(["diff", "--json"])
        self.assertEqual(res.returncode, 0, f"deffilegen diff --json failed:\n{res.stderr}")
        parsed = json.loads(res.stdout)
        self.assertIn("commit_a", parsed)
        self.assertIn("added_files", parsed)
        self.assertIn("modified_files", parsed)

    def test_cli_restore_dry_run_default_head(self) -> None:
        """Verify deffilegen restore --dry-run simulates restore cleanly without modifying files."""
        res = _run_cli_cmd(["restore", "--dry-run"])
        self.assertEqual(res.returncode, 0, f"deffilegen restore --dry-run failed:\n{res.stderr}")
        self.assertIn("(Dry-run)", res.stdout)
        self.assertIn("Restored", res.stdout)

    def test_cli_isolated_commit_diff_restore_cycle(self) -> None:
        """Verify isolated lifecycle of commit, diff, and restore in a clean temp store."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            store_dir = tmp_path / "test_store"
            prod_dir = tmp_path / "test_prod"
            restore_dir = tmp_path / "test_restore"
            prod_dir.mkdir()
            restore_dir.mkdir()

            # Create sample definition file
            sample_def = prod_dir / "INV_Sample_Model.csv"
            sample_def.write_text(
                "Modbus;Inverter;Sample;Model;;;;;;;\n"
                "1;3;30001;U16;;ActivePower;P_AC;1.000000;0.000000;W;4\n"
                "2;3;30002;U16;;Frequency;FREQ;0.010000;0.000000;Hz;4\n",
                encoding="utf-8",
            )

            # 1. Commit initial snapshot
            res_c1 = _run_cli_cmd([
                "commit",
                "--path", str(prod_dir),
                "--store-dir", str(store_dir),
                "-m", "Initial test commit",
                "--author", "Test Author",
                "-t", "v1.0.0",
            ])
            self.assertEqual(res_c1.returncode, 0, f"Commit 1 failed:\n{res_c1.stderr}")
            self.assertIn("[OK] Commit", res_c1.stdout)

            # 2. History check
            res_h = _run_cli_cmd([
                "history",
                "--store-dir", str(store_dir),
                "--json",
            ])
            self.assertEqual(res_h.returncode, 0)
            history = json.loads(res_h.stdout)
            self.assertEqual(len(history), 1)
            cid1 = history[0]["commit_id"]

            # 3. Modify definition and make a second commit with default message
            sample_def.write_text(
                "Modbus;Inverter;Sample;Model;;;;;;;\n"
                "1;3;30001;U16;;ActivePower;P_AC;1.000000;0.000000;W;4\n"
                "2;3;30002;U16;;Frequency;FREQ;0.010000;0.000000;Hz;4\n"
                "3;3;30003;U16;;ReactivePower;Q_AC;1.000000;0.000000;VAr;4\n",
                encoding="utf-8",
            )
            res_c2 = _run_cli_cmd([
                "commit",
                "--path", str(prod_dir),
                "--store-dir", str(store_dir),
            ])
            self.assertEqual(res_c2.returncode, 0, f"Commit 2 failed:\n{res_c2.stderr}")

            # 4. Diff between commits
            res_diff = _run_cli_cmd([
                "diff",
                "--store-dir", str(store_dir),
                "--json",
            ])
            self.assertEqual(res_diff.returncode, 0, f"Diff failed:\n{res_diff.stderr}")
            diff_obj = json.loads(res_diff.stdout)
            self.assertIn("INV_Sample_Model.csv", diff_obj.get("modified_files", []))

            # 5. Restore commit 1 in dry-run mode
            res_rst_dry = _run_cli_cmd([
                "restore",
                cid1,
                "--dest", str(restore_dir),
                "--store-dir", str(store_dir),
                "--dry-run",
            ])
            self.assertEqual(res_rst_dry.returncode, 0)
            self.assertIn("(Dry-run)", res_rst_dry.stdout)
            self.assertFalse((restore_dir / "INV_Sample_Model.csv").exists())

            # 6. Actual restore commit 1
            res_rst = _run_cli_cmd([
                "restore",
                cid1,
                "--dest", str(restore_dir),
                "--store-dir", str(store_dir),
            ])
            self.assertEqual(res_rst.returncode, 0)
            restored_file = restore_dir / "INV_Sample_Model.csv"
            self.assertTrue(restored_file.exists())
            self.assertNotIn("ReactivePower", restored_file.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
