"""
Adversarial Unit Test Suite for Definition Version Control Engine.

Tests:
1. Content-Addressable Storage (CAS) blob deduplication and integrity.
2. Commit creation, parent linkage, authoring, and timestamps.
3. Commit log retrieval and single-file historical lineage filtering.
4. Semantic Modbus diff computation across commits (headers, added/removed/modified registers).
5. File and directory restoration with atomic replacement.
6. Named tagging and checkout by tag.
7. CLI subcommand execution (commit, history, log, diff, restore).
"""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.main import main as cli_main
from DefFileGenerator.version_manager import DefinitionVersionControl


class TestDefinitionVersionControl(unittest.TestCase):
    """Verifies all invariants of DefinitionVersionControl."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp_dir.name)
        self.store_dir = self.root / ".def_vc"
        self.work_dir = self.root / "work"
        self.work_dir.mkdir(parents=True, exist_ok=True)
        self.vc = DefinitionVersionControl(repo_root=self.root, store_dir=self.store_dir)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def _create_sample_file(
        self,
        name: str,
        manufacturer: str = "Acme",
        model: str = "Solar1000",
        category: str = "Inverter",
        registers: list[tuple[str, str, str, str]] | None = None,
    ) -> Path:
        p = self.work_dir / name
        rows = [
            ["Modbus", category, manufacturer, model, "1", "", "", "", "", "", ""],
        ]
        if registers:
            for idx, (fc, addr, dtype, tag) in enumerate(registers, start=1):
                rows.append(
                    [
                        str(idx),
                        fc,
                        addr,
                        dtype,
                        "0",
                        f"Reg_{addr}",
                        tag,
                        "1.000000",
                        "0.000000",
                        "W",
                        "4",
                    ]
                )
        else:
            rows.append(
                ["1", "3", "100", "U16", "0", "AC Power", "P_AC", "1.000000", "0.000000", "W", "4"]
            )

        with open(p, "w", encoding="utf-8", newline="") as fp:
            writer = csv.writer(fp, delimiter=";")
            writer.writerows(rows)
        return p

    def test_repository_initialization(self) -> None:
        """Verify storage folders and index.json are created."""
        self.assertTrue(self.store_dir.is_dir())
        self.assertTrue((self.store_dir / "commits").is_dir())
        self.assertTrue((self.store_dir / "blobs").is_dir())
        self.assertTrue((self.store_dir / "tags").is_dir())
        self.assertTrue((self.store_dir / "index.json").is_file())
        self.assertIsNone(self.vc.get_head())

    def test_first_commit_and_cas_deduplication(self) -> None:
        """Verify commit creation, manifest integrity, and blob storage."""
        self._create_sample_file("INV_Acme_Solar1000.csv", "Acme", "Solar1000")
        self._create_sample_file("INV_Acme_Solar2000.csv", "Acme", "Solar2000")

        c1 = self.vc.commit(
            source_dir=self.work_dir,
            message="Initial commit",
            author="DevEngineer",
            tags=["v1.0"],
        )

        self.assertIsNotNone(c1.commit_id)
        self.assertIsNone(c1.parent_id)
        self.assertEqual(c1.author, "DevEngineer")
        self.assertEqual(c1.files_count, 2)
        self.assertEqual(c1.changes_summary["added"], 2)
        self.assertEqual(self.vc.get_head(), c1.commit_id)

        # Blobs should be stored in blobs/
        blob_files = list(self.vc.blobs_dir.rglob("*"))
        # Only files, not dirs
        stored_blobs = [b for b in blob_files if b.is_file()]
        # Two distinct files (different headers)
        self.assertEqual(len(stored_blobs), 2)

    def test_commit_lineage_and_diff(self) -> None:
        """Verify second commit diff, parent linkage, and register changes."""
        # 1. Initial commit
        self._create_sample_file(
            "INV_Acme_Solar1000.csv",
            registers=[("3", "100", "U16", "P_AC")],
        )
        c1 = self.vc.commit(self.work_dir, message="v1")

        # 2. Modify file (add registers and modify tag)
        self._create_sample_file(
            "INV_Acme_Solar1000.csv",
            registers=[
                ("3", "100", "I32", "P_AC"),  # changed type to I32
                ("3", "102", "U32", "E_TOT"),  # added register
            ],
        )
        # Add a new file
        self._create_sample_file("MET_Acme_Meter.csv", category="Meter")

        c2 = self.vc.commit(self.work_dir, message="v2: upgrade registers")

        self.assertEqual(c2.parent_id, c1.commit_id)
        self.assertEqual(c2.changes_summary["modified"], 1)
        self.assertEqual(c2.changes_summary["added"], 1)

        # Test diff between c1 and c2
        diff_res = self.vc.diff(c2.commit_id, c1.commit_id)
        self.assertIn("MET_Acme_Meter.csv", diff_res["added_files"])
        self.assertIn("INV_Acme_Solar1000.csv", diff_res["modified_files"])

        f_diff = diff_res["file_diffs"]["INV_Acme_Solar1000.csv"]
        self.assertEqual(f_diff["added_registers_count"], 1)
        self.assertIn("102", f_diff["added_addresses"])
        self.assertEqual(f_diff["modified_registers_count"], 1)
        self.assertEqual(f_diff["modified_registers"][0]["address"], "100")
        self.assertEqual(
            f_diff["modified_registers"][0]["changes"]["type"],
            ("U16", "I32"),
        )

    def test_log_and_file_filtering(self) -> None:
        """Verify log listing and per-file history tracking."""
        self._create_sample_file("INV_A.csv")
        self._create_sample_file("INV_B.csv")
        c1 = self.vc.commit(self.work_dir, message="c1")

        # Only modify INV_A
        self._create_sample_file("INV_A.csv", model="Model_A2")
        c2 = self.vc.commit(self.work_dir, message="c2")

        # Global log has 2 commits
        full_log = self.vc.log()
        self.assertEqual(len(full_log), 2)
        self.assertEqual(full_log[0]["commit_id"], c2.commit_id)
        self.assertEqual(full_log[1]["commit_id"], c1.commit_id)

        # Filtered log for INV_B should only have c1 (since c2 didn't modify it)
        log_b = self.vc.log(file_name="INV_B.csv")
        self.assertEqual(len(log_b), 1)
        self.assertEqual(log_b[0]["commit_id"], c1.commit_id)

    def test_restore_single_file_and_directory(self) -> None:
        """Verify atomic restoration of previous states."""
        self._create_sample_file("INV_A.csv", model="V1_Original")
        c1 = self.vc.commit(self.work_dir, message="c1")

        # Overwrite with V2
        self._create_sample_file("INV_A.csv", model="V2_Corrupted")
        self.vc.commit(self.work_dir, message="c2")

        # Restore c1 back into work_dir
        restored = self.vc.restore(commit_id=c1.commit_id, target_dir=self.work_dir)
        self.assertEqual(len(restored), 1)

        restored_content = (self.work_dir / "INV_A.csv").read_text(encoding="utf-8")
        self.assertIn("V1_Original", restored_content)
        self.assertNotIn("V2_Corrupted", restored_content)

    def test_tagging_system(self) -> None:
        """Verify creating and retrieving named release tags."""
        self._create_sample_file("INV_A.csv")
        c1 = self.vc.commit(self.work_dir, message="c1")

        self.vc.create_tag("release-2026", c1.commit_id, description="First production release")
        tags = self.vc.list_tags()
        self.assertIn("release-2026", tags)
        self.assertEqual(tags["release-2026"], c1.commit_id)

    def test_cli_subcommands(self) -> None:
        """Verify CLI invocations for commit, history, diff, restore."""
        self._create_sample_file("INV_CLI.csv")

        # CLI commit
        cli_main(
            [
                "commit",
                "-m",
                "CLI initial commit",
                "--path",
                str(self.work_dir),
                "--store-dir",
                str(self.store_dir),
                "-t",
                "cli-tag",
            ]
        )

        head = self.vc.get_head()
        self.assertIsNotNone(head)
        assert head is not None

        # CLI history
        cli_main(
            [
                "history",
                "--store-dir",
                str(self.store_dir),
            ]
        )

        # CLI diff
        cli_main(
            [
                "diff",
                head,
                "--store-dir",
                str(self.store_dir),
            ]
        )

        # CLI restore dry-run
        cli_main(
            [
                "restore",
                head,
                "--dest",
                str(self.work_dir),
                "--dry-run",
                "--store-dir",
                str(self.store_dir),
            ]
        )


if __name__ == "__main__":
    unittest.main()
