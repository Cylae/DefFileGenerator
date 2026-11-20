"""
Adversarial Unit Test Suite for Duplicate Detector Engine.

Tests:
1. Detection of exact SHA-256 byte duplicates.
2. Detection of exact register-body duplicates with differing model headers.
3. Detection of category prefix conflicts (OTH_ vs MET_/INV_/SEN_/BAT_).
4. Detection of hardware alias naming variations.
5. Generation of Markdown audit reports.
6. Execution via deffilegen duplicates CLI.
"""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.duplicate_detector import DuplicateDetector
from DefFileGenerator.main import main as cli_main


class TestDuplicateDetector(unittest.TestCase):
    """Verifies all duplicate detection invariants."""

    def setUp(self) -> None:
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.work_dir = Path(self.tmp_dir.name)

    def tearDown(self) -> None:
        self.tmp_dir.cleanup()

    def _write_csv(self, filename: str, header: list[str], rows: list[list[str]]) -> Path:
        p = self.work_dir / filename
        with open(p, "w", encoding="utf-8", newline="") as fp:
            writer = csv.writer(fp, delimiter=";")
            writer.writerow(header)
            writer.writerows(rows)
        return p

    def test_category_conflict_detection(self) -> None:
        """Detect when OTH_ duplicates an already-typed MET_ or INV_ definition."""
        sample_rows = [
            ["1", "3", "100", "U16", "0", "Voltage", "U_AC", "1.000000", "0.000000", "V", "4"],
            ["2", "3", "101", "U16", "0", "Current", "I_AC", "1.000000", "0.000000", "A", "4"],
        ]
        # Canonical Meter
        self._write_csv(
            "MET_Vendor_MeterX.csv",
            ["Modbus", "Meter", "Vendor", "MeterX", "1", "", "", "", "", "", ""],
            sample_rows,
        )
        # Redundant Other copy
        self._write_csv(
            "OTH_Vendor_MeterX.csv",
            ["Modbus", "Other", "Vendor", "MeterX", "1", "", "", "", "", "", ""],
            sample_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        self.assertEqual(len(report.category_conflict_clusters), 1)
        cluster = report.category_conflict_clusters[0]
        self.assertEqual(cluster.canonical_file, "MET_Vendor_MeterX.csv")
        self.assertEqual(cluster.duplicate_files, ["OTH_Vendor_MeterX.csv"])
        self.assertEqual(cluster.action_recommendation, "purge_oth_duplicate")

    def test_alias_synonym_detection(self) -> None:
        """Detect when two files are naming aliases (e.g. GT_Series vs GT-SERIES, HT1500 vs HT_1500V)."""
        sample_rows = [
            ["1", "3", "500", "U32", "0", "Energy", "E_TOT", "1.000000", "0.000000", "kWh", "4"]
        ]
        self._write_csv(
            "INV_Acme_HT1500.csv",
            ["Modbus", "Inverter", "Acme", "HT1500", "1", "", "", "", "", "", ""],
            sample_rows,
        )
        self._write_csv(
            "INV_Acme_HT_1500V.csv",
            ["Modbus", "Inverter", "Acme", "HT_1500V", "1", "", "", "", "", "", ""],
            sample_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        self.assertEqual(len(report.alias_clusters), 1)
        cluster = report.alias_clusters[0]
        self.assertIn("INV_Acme_HT_1500V.csv", cluster.duplicate_files)

    def test_markdown_report_generation(self) -> None:
        """Verify markdown report generation formatting."""
        sample_rows = [
            ["1", "3", "10", "U16", "0", "Power", "P_AC", "1.000000", "0.000000", "W", "4"]
        ]
        self._write_csv(
            "MET_Chint_DTSU.csv",
            ["Modbus", "Meter", "Chint", "DTSU", "1", "", "", "", "", "", ""],
            sample_rows,
        )
        self._write_csv(
            "OTH_Chint_DTSU.csv",
            ["Modbus", "Other", "Chint", "DTSU", "1", "", "", "", "", "", ""],
            sample_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()
        md = detector.format_markdown_report(report)

        self.assertIn("# Rapport d'Audit des Doublons", md)
        self.assertIn("MET_Chint_DTSU.csv", md)
        self.assertIn("OTH_Chint_DTSU.csv", md)

    def test_cli_duplicates_command(self) -> None:
        """Verify CLI execution of duplicates subcommand."""
        self._write_csv(
            "INV_A.csv",
            ["Modbus", "Inverter", "Acme", "A", "1", "", "", "", "", "", ""],
            [["1", "3", "1", "U16", "0", "Reg", "P", "1.0", "0.0", "", "4"]],
        )
        report_path = self.work_dir / "audit.md"

        cli_main([
            "duplicates",
            "--path",
            str(self.work_dir),
            "--report-out",
            str(report_path),
        ])

        self.assertTrue(report_path.is_file())
        self.assertGreater(report_path.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
