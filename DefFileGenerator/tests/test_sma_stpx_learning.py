#!/usr/bin/env python3
"""
Test suite validating SMA STP-X Modbus telemetry inference, accent normalization,
definition file compliance, and dataset knowledge retraining.
"""

from __future__ import annotations

import csv
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.knowledge_miner import (
    clean_variable_key,
    infer_sunpm_tag,
    is_canonical_sunpm_tag,
    load_dataset_knowledge,
)


class TestSmaStpxLearning(unittest.TestCase):
    """Hermetic test battery for SMA STP-X telemetry tagging and learning."""

    def test_infer_sunpm_tag_sma_stpx_french_telemetry(self) -> None:
        """Verifies that SMA STP-X French variable names infer exact SunPM canonical tags."""
        test_cases = [
            ("Puissance", "P_AC"),
            ("Puissance L1", "P_AC_1"),
            ("Puissance L2", "P_AC_2"),
            ("Puissance L3", "P_AC_3"),
            ("Courant du reseau", "I_AC"),
            ("Courant du réseau", "I_AC"),
            ("Courant du reseau phase L1_0", "I_AC_1"),
            ("Courant du reseau phase L2_0", "I_AC_2"),
            ("Courant du reseau phase L3_0", "I_AC_3"),
            ("Courant du reseau phase L1_1", "I_AC_1"),
            ("Courant du réseau phase L1_1", "I_AC_1"),
            ("Tension du reseau phase L1", "U_AC_1"),
            ("Tension du reseau phase L2", "U_AC_2"),
            ("Tension du reseau phase L3", "U_AC_3"),
            ("Tension du réseau phase L1", "U_AC_1"),
            ("Tension de reseau entre phase L1 et L2", "U_AC_L1L2"),
            ("Tension de reseau phase L2 contre L3", "U_AC_L2L3"),
            ("Tension de reseau phase L3 contre L1", "U_AC_L3L1"),
            ("Tension de réseau entre phase L1 et L2", "U_AC_L1L2"),
            ("Temperature interieure", "TEMP_INT"),
            ("Température intérieure", "TEMP_INT"),
            ("Tension du circuit intermediaire", "U_DC"),
            ("Tension du circuit intermédiaire", "U_DC"),
            ("Entree courant DC_1", "I_DC_MPPT_1"),
            ("Entree tension DC_1", "U_DC_MPPT_1"),
            ("Entree puissance DC_1", "P_DC_MPPT_1"),
            ("Entree courant DC_2", "I_DC_MPPT_2"),
            ("Entree tension DC_2", "U_DC_MPPT_2"),
            ("Entree puissance DC_2", "P_DC_MPPT_2"),
            ("Frequence du reseau", "FREQ"),
            ("Fréquence du réseau", "FREQ"),
            ("Puissance reactive", "Q_AC"),
            ("Puissance réactive", "Q_AC"),
            ("Puissance reactive L1", "Q_AC_1"),
            ("Puissance reactive L2", "Q_AC_2"),
            ("Puissance reactive L3", "Q_AC_3"),
            ("Puissance apparente", "S_AC"),
            ("Puissance apparente L1", "S_AC_1"),
            ("Puissance apparente L2", "S_AC_2"),
            ("Puissance apparente L3", "S_AC_3"),
            ("Facteur de decalage", "PF"),
            ("Facteur de décalage", "PF"),
            ("Etat de charge actuel de la batterie", "SOC"),
            ("État de charge actuel de la batterie", "SOC"),
            ("Etat", "STATE"),
            ("Resistance d isolement", "R_ISO"),
            ("Résistance d'isolement", "R_ISO"),
            ("Rendement total_0", "E_TOT"),
            ("Temps de service_0", "TOT"),
        ]

        for name, expected_tag in test_cases:
            with self.subTest(name=name):
                tag = infer_sunpm_tag(name)
                self.assertEqual(
                    tag,
                    expected_tag,
                    f"Failed inference for '{name}': expected '{expected_tag}', got '{tag}'",
                )
                self.assertTrue(
                    is_canonical_sunpm_tag(tag),
                    f"Inferred tag '{tag}' for '{name}' is not in SUNPM_CANONICAL_TAGS",
                )

        # Disambiguation safety: Excitation type must NEVER infer Power Factor (PF)
        for non_pf in (
            "Type d excitation du cos phi",
            "Type d'excitation du cos phi",
            "Consigne cos phi externe_ type d_excitation en cas de puissance active absorbee",
        ):
            with self.subTest(non_pf=non_pf):
                self.assertNotEqual(
                    infer_sunpm_tag(non_pf),
                    "PF",
                    f"Excitation type '{non_pf}' must not infer 'PF'",
                )

    def test_sma_stpx_accent_normalization(self) -> None:
        """Verifies that clean_variable_key normalizes accents consistently."""
        self.assertEqual(
            clean_variable_key("Tension du réseau phase L1"),
            clean_variable_key("Tension du reseau phase L1"),
        )
        self.assertEqual(
            clean_variable_key("Température intérieure"),
            clean_variable_key("Temperature interieure"),
        )
        self.assertEqual(
            clean_variable_key("Puissance réactive"),
            clean_variable_key("Puissance reactive"),
        )
        self.assertEqual(
            clean_variable_key("Résistance d'isolement"),
            clean_variable_key("Resistance d isolement"),
        )

    def test_sma_stpx_definition_files_compliance(self) -> None:
        """Verifies 100% compliance of both SMA STP-X definition files with the 5 Pillars."""
        gen = Generator()
        files = [
            Path("production/INV_SMA_STP_X.csv"),
            Path("Definitions/INV_SMA_STP_X.csv"),
        ]

        for p in files:
            with self.subTest(file=p.name):
                self.assertTrue(p.is_file(), f"File {p} does not exist.")
                report = gen.validate_csv_detailed(str(p))
                self.assertTrue(
                    report.is_valid,
                    f"Validation failed for {p}: {report.issues}",
                )
                self.assertEqual(
                    len(report.issues),
                    0,
                    f"Expected 0 issues for {p}, got: {report.issues}",
                )

                # Verify tag uniqueness and count
                with open(p, encoding="utf-8-sig") as f:
                    reader = list(csv.reader(f, delimiter=";"))
                tags = [r[6].strip() for r in reader[1:] if len(r) > 6 and r[6].strip()]
                self.assertEqual(
                    len(tags),
                    len(set(tags)),
                    f"Duplicate tags detected in {p}: {[t for t in tags if tags.count(t) > 1]}",
                )
                self.assertGreaterEqual(
                    len(tags),
                    45,
                    f"Expected at least 45 tagged variables in {p}, found {len(tags)}",
                )

    def test_dataset_knowledge_contains_sma_stpx(self) -> None:
        """Verifies that dataset_knowledge.json includes retrained SMA STP-X mappings."""
        kb = load_dataset_knowledge()
        self.assertIsNotNone(kb, "Failed to load dataset_knowledge.json")
        assert kb is not None

        expected_mappings = {
            "puissance": "P_AC",
            "puissance l1": "P_AC_1",
            "puissance l2": "P_AC_2",
            "puissance l3": "P_AC_3",
            "courant du reseau": "I_AC",
            "tension du reseau phase l1": "U_AC_1",
            "tension du reseau phase l2": "U_AC_2",
            "tension du reseau phase l3": "U_AC_3",
            "tension de reseau entre phase l1 et l2": "U_AC_L1L2",
            "tension de reseau phase l2 contre l3": "U_AC_L2L3",
            "tension de reseau phase l3 contre l1": "U_AC_L3L1",
            "temperature interieure": "TEMP_INT",
            "frequence du reseau": "FREQ",
            "puissance reactive": "Q_AC",
            "puissance reactive l1": "Q_AC_1",
            "puissance apparente": "S_AC",
            "puissance apparente l1": "S_AC_1",
            "facteur de decalage": "PF",
            "etat de charge actuel de la batterie": "SOC",
            "etat": "STATE",
            "tension du circuit intermediaire": "U_DC",
            "entree courant dc 1": "I_DC_MPPT_1",
            "entree tension dc 1": "U_DC_MPPT_1",
            "entree puissance dc 1": "P_DC_MPPT_1",
            "resistance d isolement": "R_ISO",
            "rendement total 0": "E_TOT",
            "temps de service 0": "TOT",
        }

        for key, expected_tag in expected_mappings.items():
            with self.subTest(key=key):
                learned_tag = kb.get_learned_tag(key)
                self.assertEqual(
                    learned_tag,
                    expected_tag,
                    f"Key '{key}' learned tag is '{learned_tag}', expected '{expected_tag}'",
                )


if __name__ == "__main__":
    unittest.main()
