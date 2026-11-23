"""
Adversarial Unit Test Suite for Duplicate Detector Engine.

Tests:
1. Detection of exact SHA-256 byte duplicates.
2. Detection of exact register-body duplicates with differing model headers.
3. Detection of category prefix conflicts (OTH_ vs MET_/INV_/SEN_/BAT_).
4. Detection of hardware alias naming variations.
5. Detection and stripping of copy suffixes (-_copie, _copie, -_copy, _copy).
6. Normalization of vendor tokens (fox_ess -> foxess, sl_energy -> slenergy).
7. Filtering of exact hash duplicates from hardware family body-matching.
8. Generation of Markdown audit reports with exact SHA-256 section.
9. Execution via deffilegen duplicates CLI.
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
        # Redundant Other copy with different header model to avoid exact byte duplicate
        self._write_csv(
            "OTH_Vendor_MeterX.csv",
            ["Modbus", "Other", "Vendor", "MeterX_Legacy", "1", "", "", "", "", "", ""],
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

    def test_copy_suffix_and_vendor_alias_handling(self) -> None:
        """Verify _is_alias_pair strips copy suffixes and normalizes vendor tokens."""
        detector = DuplicateDetector(self.work_dir)

        # Copy suffix variations
        self.assertTrue(
            detector._is_alias_pair("INV_Vendor_Device.csv", "INV_Vendor_Device-_copie.csv")
        )
        self.assertTrue(
            detector._is_alias_pair("INV_Vendor_Device.csv", "INV_Vendor_Device_copie.csv")
        )
        self.assertTrue(
            detector._is_alias_pair("INV_Vendor_Device.csv", "INV_Vendor_Device-_copy.csv")
        )
        self.assertTrue(
            detector._is_alias_pair("INV_Vendor_Device.csv", "INV_Vendor_Device_copy.csv")
        )
        self.assertTrue(
            detector._is_alias_pair(
                "BAT_Robestec_battery_Robestec_BESS.csv",
                "BAT_Robestec_battery_Robestec_BESS_-_Copie.csv",
            )
        )

        # Vendor token normalization
        self.assertTrue(
            detector._is_alias_pair("BAT_FoxESS_P3_H3_PRO.csv", "BAT_Fox_ESS_FoxESS_P3_H3_PRO.csv")
        )
        self.assertTrue(
            detector._is_alias_pair(
                "BAT_SLEnergy_Generator.csv", "BAT_SL_Energy_SLEnergy_Generator.csv"
            )
        )

        # Distinct equipment models must NOT be considered aliases
        self.assertFalse(
            detector._is_alias_pair("MET_Schneider_iEM3150.csv", "MET_Schneider_iEM3155.csv")
        )

    def test_exact_sha256_duplicate_detection(self) -> None:
        """Exact SHA-256 byte duplicates must be classified into exact_hash_clusters and NOT bleed into exact_body_clusters."""
        sample_rows = [
            ["1", "3", "100", "U16", "0", "Voltage", "U_AC", "1.000000", "0.000000", "V", "4"],
            ["2", "3", "101", "U16", "0", "Current", "I_AC", "1.000000", "0.000000", "A", "4"],
        ]
        self._write_csv(
            "INV_ABB_PVS800-57B.csv",
            ["Modbus", "Inverter", "ABB", "PVS800-57B", "1", "", "", "", "", "", ""],
            sample_rows,
        )
        self._write_csv(
            "INV_ABB_PVS800-57B_2.csv",
            ["Modbus", "Inverter", "ABB", "PVS800-57B", "1", "", "", "", "", "", ""],
            sample_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        self.assertEqual(len(report.exact_hash_clusters), 1)
        cluster = report.exact_hash_clusters[0]
        self.assertEqual(cluster.cluster_type, "exact_sha256")
        self.assertEqual(cluster.canonical_file, "INV_ABB_PVS800-57B.csv")
        self.assertEqual(cluster.duplicate_files, ["INV_ABB_PVS800-57B_2.csv"])
        self.assertEqual(cluster.register_count, 2)

        # Invariant: Must NOT bleed into hardware families
        self.assertEqual(len(report.exact_body_clusters), 0)
        self.assertEqual(len(report.category_conflict_clusters), 0)
        self.assertEqual(len(report.alias_clusters), 0)
        self.assertEqual(report.total_duplicate_files, 1)

    def test_exact_sha256_duplicates_do_not_bleed_into_hardware_families_with_distinct_models(
        self,
    ) -> None:
        """When distinct models share registers with byte duplicates present, byte duplicates must be excluded from hardware families."""
        shared_rows = [["1", "3", "100", "U16", "0", "Power", "P", "1.0", "0.0", "W", "4"]]
        # Model A - canonical
        self._write_csv(
            "INV_ModelA_1.csv",
            ["Modbus", "Inverter", "Brand", "ModelA", "1", "", "", "", "", "", ""],
            shared_rows,
        )
        # Model A - exact byte duplicate
        self._write_csv(
            "INV_ModelA_2.csv",
            ["Modbus", "Inverter", "Brand", "ModelA", "1", "", "", "", "", "", ""],
            shared_rows,
        )
        # Model B - distinct hardware model sharing same register body
        self._write_csv(
            "INV_ModelB.csv",
            ["Modbus", "Inverter", "Brand", "ModelB", "1", "", "", "", "", "", ""],
            shared_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        # Exact duplicate detected
        self.assertEqual(len(report.exact_hash_clusters), 1)
        self.assertEqual(report.exact_hash_clusters[0].canonical_file, "INV_ModelA_1.csv")
        self.assertEqual(report.exact_hash_clusters[0].duplicate_files, ["INV_ModelA_2.csv"])

        # Hardware family only contains distinct models
        self.assertEqual(len(report.exact_body_clusters), 1)
        body_cluster = report.exact_body_clusters[0]
        self.assertEqual(body_cluster.canonical_file, "INV_ModelB.csv")
        self.assertEqual(body_cluster.duplicate_files, ["INV_ModelA_1.csv"])
        self.assertNotIn("INV_ModelA_2.csv", body_cluster.duplicate_files)
        self.assertNotEqual("INV_ModelA_2.csv", body_cluster.canonical_file)

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
            ["Modbus", "Other", "Chint", "DTSU_Other", "1", "", "", "", "", "", ""],
            sample_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()
        md = detector.format_markdown_report(report)

        self.assertIn("# Rapport d'Audit des Doublons", md)
        self.assertIn("MET_Chint_DTSU.csv", md)
        self.assertIn("OTH_Chint_DTSU.csv", md)

    def test_markdown_report_includes_exact_sha256_section(self) -> None:
        """Markdown report must include dedicated Section 1 for Exact SHA-256 duplicates with required columns and recommendation."""
        sample_rows = [["1", "3", "50", "U16", "0", "Frequency", "FREQ", "1.0", "0.0", "Hz", "4"]]
        self._write_csv(
            "INV_ABB_PVS800-57B.csv",
            ["Modbus", "Inverter", "ABB", "PVS800", "1", "", "", "", "", "", ""],
            sample_rows,
        )
        self._write_csv(
            "INV_ABB_PVS800-57B_2.csv",
            ["Modbus", "Inverter", "ABB", "PVS800", "1", "", "", "", "", "", ""],
            sample_rows,
        )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()
        md = detector.format_markdown_report(report)

        self.assertIn("## 1. Doublons Stricts (Empreinte SHA-256 Identique)", md)
        self.assertIn(
            "| Fichier Canonique | Doublons Stricts | Modèle Matériel | Registres | Recommandation |",
            md,
        )
        self.assertIn("`INV_ABB_PVS800-57B.csv`", md)
        self.assertIn("`INV_ABB_PVS800-57B_2.csv`", md)
        self.assertIn("PVS800", md)
        self.assertIn("Supprimer le doublon strict ou créer un alias", md)
        self.assertIn("## 2. Conflits de Catégorie", md)
        self.assertIn("## 3. Alias Matériels et Variantes Synonymes de Nommage", md)
        self.assertIn("## 4. Familles Matérielles Partageant la Même Carte Modbus", md)

    def test_markdown_report_renders_all_hardware_families_above_25(self) -> None:
        """Markdown report must cleanly render all hardware family clusters when > 25 exist, without truncation."""
        # Create 30 distinct pairs of hardware families
        for i in range(30):
            reg_row = [
                [
                    "1",
                    "3",
                    str(100 + i),
                    "U16",
                    "0",
                    f"Param_{i}",
                    f"TAG_{i}",
                    "1.0",
                    "0.0",
                    "",
                    "4",
                ]
            ]
            self._write_csv(
                f"INV_Brand_M{i:02d}_A.csv",
                ["Modbus", "Inverter", "Brand", f"M{i:02d}_A", "1", "", "", "", "", "", ""],
                reg_row,
            )
            self._write_csv(
                f"INV_Brand_M{i:02d}_B.csv",
                ["Modbus", "Inverter", "Brand", f"M{i:02d}_B", "1", "", "", "", "", "", ""],
                reg_row,
            )

        detector = DuplicateDetector(self.work_dir)
        report = detector.scan()

        self.assertEqual(len(report.exact_body_clusters), 30)
        md = detector.format_markdown_report(report)

        self.assertIn("## 4. Familles Matérielles Partageant la Même Carte Modbus", md)
        self.assertIn("Total : **30** familles répertoriées.", md)
        for i in range(30):
            self.assertIn(f"INV_Brand_M{i:02d}_A.csv", md)

    def test_cli_duplicates_command(self) -> None:
        """Verify CLI execution of duplicates subcommand."""
        self._write_csv(
            "INV_A.csv",
            ["Modbus", "Inverter", "Acme", "A", "1", "", "", "", "", "", ""],
            [["1", "3", "1", "U16", "0", "Reg", "P", "1.0", "0.0", "", "4"]],
        )
        report_path = self.work_dir / "audit.md"

        cli_main(
            [
                "duplicates",
                "--path",
                str(self.work_dir),
                "--report-out",
                str(report_path),
            ]
        )

        self.assertTrue(report_path.is_file())
        self.assertGreater(report_path.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
