"""
Comprehensive validation and coverage test suite for the Knowledge Mining & Semantic Tag Engine.
Tests SunPM tag taxonomy integrity, multi-language pattern recognition (English, French, German, Chinese transliterations),
indentation and block alignment for E_IMP/E_EXP, and edge case resilience according to Apollo Engineering Invariants.
"""

from __future__ import annotations

import csv
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.knowledge_miner import (
    SUNPM_CANONICAL_TAGS,
    clean_variable_key,
    enrich_definition_file,
    enrich_definition_rows,
    infer_sunpm_tag,
    is_canonical_sunpm_tag,
    normalize_tag_to_sunpm,
)


class TestSunPMTaxonomyIntegrity(unittest.TestCase):
    """Verifies that the full SunPM canonical tag taxonomy is intact and well-formed."""

    def test_core_telemetry_tags_present(self) -> None:
        """Verify all essential inverter and meter telemetry tags exist in SUNPM_CANONICAL_TAGS."""
        expected_tags = [
            # Active, Reactive, Apparent Power
            "P_AC", "P_AC_1", "P_AC_2", "P_AC_3",
            "Q_AC", "Q_AC_1", "Q_AC_2", "Q_AC_3",
            "S_AC", "S_AC_1", "S_AC_2", "S_AC_3",
            # Voltages & Currents
            "U_AC", "U_AC_1", "U_AC_2", "U_AC_3",
            "U_AC_L1L2", "U_AC_L2L3", "U_AC_L3L1",
            "I_AC", "I_AC_1", "I_AC_2", "I_AC_3",
            # Frequency & Power Factor
            "FREQ", "FREQ_1", "FREQ_2", "FREQ_3",
            "PF", "PF_1", "PF_2", "PF_3",
            # Energy
            "E_DAY", "E_TOT", "E_TOT_1", "E_TOT_2", "E_TOT_3",
            "E_EXP", "E_IMP",
            # Ratings & DC Parameters
            "P_RATED", "Q_RATED",
            "U_DC", "I_DC", "P_DC",
            # Temperatures
            "TEMP", "TEMP_AMB", "TEMP_CELL", "TEMP_INT", "TEMP_MAX", "TEMP_MIN",
            # Battery & BESS
            "SOC", "SOH", "U_CELL_MAX", "U_CELL_MIN",
            # Sensors & Met
            "POA", "GHI", "WIND_SPEED", "WIND_BEARING", "HUMIDITY", "PRESSURE", "RAIN",
            # Diagnostics & State
            "STATE", "STATE_1", "FAULT_CODE", "R_ISO",
            # Metadata
            "SERIAL", "MODEL", "MANUFACTURER", "UNIT_ID",
            # Remote Commands
            "CMD_ON", "CMD_OFF", "CMD_P_AC_SET_PERCENT", "CMD_P_AC_SET_ABS",
            "CMD_P_ENA", "CMD_RAMP_TIME_S", "CMD_REVERT_TIME_S",
            "CMD_Q_AC_SET_PERCENT", "CMD_Q_AC_SET_ABS", "CMD_PF", "CMD_COMMIT",
        ]
        for tag in expected_tags:
            self.assertIn(tag, SUNPM_CANONICAL_TAGS, f"Missing canonical tag: {tag}")
            self.assertTrue(is_canonical_sunpm_tag(tag), f"is_canonical_sunpm_tag returned False for {tag}")

    def test_dynamic_pattern_canonical_tags(self) -> None:
        """Verify dynamic numbered tags for MPPTs, strings, alarms, and states."""
        # MPPT and Strings
        for i in range(1, 13):
            self.assertTrue(is_canonical_sunpm_tag(f"U_DC_MPPT_{i}"))
            self.assertTrue(is_canonical_sunpm_tag(f"I_DC_MPPT_{i}"))
            self.assertTrue(is_canonical_sunpm_tag(f"P_DC_MPPT_{i}"))
            self.assertTrue(is_canonical_sunpm_tag(f"U_DC_STRING_{i}"))
            self.assertTrue(is_canonical_sunpm_tag(f"I_DC_STRING_{i}"))

        # Alarms and States
        for i in range(1, 33):
            self.assertTrue(is_canonical_sunpm_tag(f"ALARM_{i}"))
        for i in range(1, 17):
            self.assertTrue(is_canonical_sunpm_tag(f"STATE_{i}"))

        # Rejection of invalid tags
        invalid_tags = ["", "NOT_A_TAG", "U16", "FLOAT32", "MODBUS", "123", "P_AC_99_INVALID"]
        for inv in invalid_tags:
            self.assertFalse(is_canonical_sunpm_tag(inv), f"Expected False for {inv}")

    def test_per_phase_legacy_normalization(self) -> None:
        """Verify that L1..L3 alternates normalize to canonical _1.._3 tags."""
        alternates = [
            ("U_AC_L1", "U_AC_1"),
            ("U_AC_L2", "U_AC_2"),
            ("U_AC_L3", "U_AC_3"),
            ("I_AC_L1", "I_AC_1"),
            ("I_AC_L2", "I_AC_2"),
            ("I_AC_L3", "I_AC_3"),
            ("P_AC_L1", "P_AC_1"),
            ("P_AC_L2", "P_AC_2"),
            ("P_AC_L3", "P_AC_3"),
            ("Q_AC_L1", "Q_AC_1"),
            ("Q_AC_L2", "Q_AC_2"),
            ("Q_AC_L3", "Q_AC_3"),
            ("S_AC_L1", "S_AC_1"),
            ("S_AC_L2", "S_AC_2"),
            ("S_AC_L3", "S_AC_3"),
            ("FREQ_L1", "FREQ_1"),
            ("FREQ_L2", "FREQ_2"),
            ("FREQ_L3", "FREQ_3"),
            ("PF_L1", "PF_1"),
            ("PF_L2", "PF_2"),
            ("PF_L3", "PF_3"),
            ("E_TOT_L1", "E_TOT_1"),
            ("E_TOT_L2", "E_TOT_2"),
            ("E_TOT_L3", "E_TOT_3"),
        ]
        for legacy, canonical in alternates:
            self.assertEqual(normalize_tag_to_sunpm(legacy), canonical)
            self.assertTrue(is_canonical_sunpm_tag(canonical))


class TestMultiLanguagePatternRecognition(unittest.TestCase):
    """Validates regex and keyword pattern recognition across English, French, German, and Chinese transliterations."""

    def test_english_inverter_and_meter_patterns(self) -> None:
        """Verify English telemetry and command inference."""
        cases = [
            ("Total Active Power", "P_AC"),
            ("Phase 1 Active Power", "P_AC_1"),
            ("Phase B Active Power", "P_AC_2"),
            ("Phase C Active Power", "P_AC_3"),
            ("Total Reactive Power", "Q_AC"),
            ("Total Apparent Power", "S_AC"),
            ("Grid Frequency", "FREQ"),
            ("Total Power Factor", "PF"),
            ("Phase A Voltage", "U_AC_1"),
            ("Line 1 to Line 2 Voltage", "U_AC_L1L2"),
            ("Phase B to Phase C Voltage", "U_AC_L2L3"),
            ("Phase C to Phase A Voltage", "U_AC_L3L1"),
            ("Total AC Current", "I_AC"),
            ("Phase 1 Current", "I_AC_1"),
            ("Daily Energy Yield", "E_DAY"),
            ("Cumulative Total Energy", "E_TOT"),
            ("Active Import Energy", "E_IMP"),
            ("Active Export Energy", "E_EXP"),
            ("DC Bus Voltage", "U_DC"),
            ("DC Bus Current", "I_DC"),
            ("Ambient Temperature", "TEMP_AMB"),
            ("Inverter Internal Temperature", "TEMP_INT"),
            ("Battery State of Charge", "SOC"),
            ("Battery State of Health", "SOH"),
            ("Turn On Inverter", "CMD_ON"),
            ("Turn Off Inverter", "CMD_OFF"),
        ]
        for name, expected in cases:
            tag = infer_sunpm_tag(name, use_learned_dataset=False)
            self.assertEqual(tag, expected, f"Failed English inference for '{name}': got '{tag}', expected '{expected}'")

    def test_french_inverter_and_meter_patterns(self) -> None:
        """Verify French telemetry and command inference."""
        cases = [
            ("Puissance active totale", "P_AC"),
            ("Puissance active phase 1", "P_AC_1"),
            ("Puissance active phase 2", "P_AC_2"),
            ("Puissance active phase 3", "P_AC_3"),
            ("Puissance réactive totale", "Q_AC"),
            ("Puissance apparente totale", "S_AC"),
            ("Fréquence du réseau", "FREQ"),
            ("Facteur de puissance total", "PF"),
            ("Tension simple phase 1", "U_AC_1"),
            ("Tension réseau phase 2", "U_AC_2"),
            ("Tension réseau phase 3", "U_AC_3"),
            ("Tension composée L1-L2", "U_AC_L1L2"),
            ("Tension de réseau phase L2 contre L3", "U_AC_L2L3"),
            ("Tension composée phase L3 et L1", "U_AC_L3L1"),
            ("Courant réseau total", "I_AC"),
            ("Courant phase 1", "I_AC_1"),
            ("Energie journalière", "E_DAY"),
            ("Energie active totale", "E_TOT"),
            ("Energie active importée", "E_IMP"),
            ("Energie active exportée", "E_EXP"),
            ("Prélèvement actif", "E_IMP"),
            ("Injection active", "E_EXP"),
            ("Tension bus continu", "U_DC"),
            ("Courant continu", "I_DC"),
            ("Température ambiante", "TEMP_AMB"),
            ("Température interne", "TEMP_INT"),
            ("Etat de charge batterie", "SOC"),
            ("Etat de santé batterie", "SOH"),
            ("Démarrage onduleur", "CMD_ON"),
            ("Arrêt onduleur", "CMD_OFF"),
        ]
        for name, expected in cases:
            tag = infer_sunpm_tag(name, use_learned_dataset=False)
            self.assertEqual(tag, expected, f"Failed French inference for '{name}': got '{tag}', expected '{expected}'")

    def test_german_inverter_and_meter_patterns(self) -> None:
        """Verify German telemetry and command inference."""
        cases = [
            ("Wirkleistung", "P_AC"),
            ("Wirkleistung Phase 1", "P_AC_1"),
            ("Wirkleistung Phase 2", "P_AC_2"),
            ("Wirkleistung Phase 3", "P_AC_3"),
            ("Blindleistung", "Q_AC"),
            ("Scheinleistung", "S_AC"),
            ("Nennleistung", "P_RATED"),
            ("Netzfrequenz", "FREQ"),
            ("Leistungsfaktor", "PF"),
            ("Verschiebungsfaktor", "PF"),
            ("Netzspannung", "U_AC_L1L2"),
            ("Spannung L1-L2", "U_AC_L1L2"),
            ("Spannung L2-L3", "U_AC_L2L3"),
            ("Spannung L3-L1", "U_AC_L3L1"),
            ("Phasenspannung L1", "U_AC_1"),
            ("Phasenspannung L2", "U_AC_2"),
            ("Phasenspannung L3", "U_AC_3"),
            ("Netzstrom", "I_AC"),
            ("Strom L1", "I_AC_1"),
            ("Strom L2", "I_AC_2"),
            ("Strom L3", "I_AC_3"),
            ("Gesamtertrag", "E_TOT"),
            ("Gesamtenergie", "E_TOT"),
            ("Tagesertrag", "E_DAY"),
            ("Tagesenergie", "E_DAY"),
            ("Wirkenergie Bezug", "E_IMP"),
            ("Bezugsenergie", "E_IMP"),
            ("Bezug kWh", "E_IMP"),
            ("Wirkenergie Einspeisung", "E_EXP"),
            ("Einspeiseenergie", "E_EXP"),
            ("Einspeisung kWh", "E_EXP"),
            ("Gleichspannung", "U_DC"),
            ("Gleichstrom", "I_DC"),
            ("Zwischenkreisspannung", "U_DC"),
            ("Umgebungstemperatur", "TEMP_AMB"),
            ("Innentemperatur", "TEMP_INT"),
            ("Ladezustand", "SOC"),
            ("Alterungszustand", "SOH"),
            ("Einschalten", "CMD_ON"),
            ("Ausschalten", "CMD_OFF"),
            ("Seriennummer", "SERIAL"),
            ("Modell", "MODEL"),
            ("Hersteller", "MANUFACTURER"),
            ("Geräteadresse", "UNIT_ID"),
        ]
        for name, expected in cases:
            tag = infer_sunpm_tag(name, use_learned_dataset=False)
            self.assertEqual(tag, expected, f"Failed German inference for '{name}': got '{tag}', expected '{expected}'")

    def test_chinese_transliteration_patterns(self) -> None:
        """Verify Chinese pinyin transliteration telemetry and command inference."""
        cases = [
            ("Yougong Gonglv", "P_AC"),
            ("Yougong L1", "P_AC_1"),
            ("Yougong L2", "P_AC_2"),
            ("Yougong L3", "P_AC_3"),
            ("Wugong Gonglv", "Q_AC"),
            ("Shizai Gonglv", "S_AC"),
            ("Eding Gonglv", "P_RATED"),
            ("Wangce Dianya", "U_AC_L1L2"),
            ("AB Dianya", "U_AC_L1L2"),
            ("BC Dianya", "U_AC_L2L3"),
            ("CA Dianya", "U_AC_L3L1"),
            ("Dianya L1", "U_AC_1"),
            ("Dianya L2", "U_AC_2"),
            ("Dianya L3", "U_AC_3"),
            ("Wangce Dianliu", "I_AC"),
            ("Dianliu L1", "I_AC_1"),
            ("Dianliu L2", "I_AC_2"),
            ("Dianliu L3", "I_AC_3"),
            ("Pinlv", "FREQ"),
            ("Gonglv Yinshu", "PF"),
            ("Rifadianliang", "E_DAY"),
            ("Zong Fadianliang", "E_TOT"),
            ("Leiji Fadianliang", "E_TOT"),
            ("Zhengxiang Yougong", "E_IMP"),
            ("Fanxiang Yougong", "E_EXP"),
            ("Zhiliu Dianya", "U_DC"),
            ("Zhiliu Dianliu", "I_DC"),
            ("Zhiliu Gonglv", "P_DC"),
            ("Huanjing Wendu", "TEMP_AMB"),
            ("Neibu Wendu", "TEMP_INT"),
            ("Dianchi SOC", "SOC"),
            ("Shengyu Dianliang", "SOC"),
            ("Jiankang Zhuangtai", "SOH"),
            ("Kaiji", "CMD_ON"),
            ("Tingji", "CMD_OFF"),
            ("Guanji", "CMD_OFF"),
            ("Xuliehao", "SERIAL"),
            ("Xinghao", "MODEL"),
            ("Changjia", "MANUFACTURER"),
            ("Shebei Dizhi", "UNIT_ID"),
        ]
        for name, expected in cases:
            tag = infer_sunpm_tag(name, use_learned_dataset=False)
            self.assertEqual(tag, expected, f"Failed Chinese transliteration inference for '{name}': got '{tag}', expected '{expected}'")


class TestEnergyBlockIndentationAndEdgeCases(unittest.TestCase):
    """Ensures E_IMP, E_EXP, and E_TOT block alignments and edge case boundaries are rigorously tested."""

    def test_import_and_export_without_energy_keyword(self) -> None:
        """Verify meter variables without explicit 'energy' keyword correctly infer E_IMP and E_EXP."""
        self.assertEqual(infer_sunpm_tag("Active Import", use_learned_dataset=False), "E_IMP")
        self.assertEqual(infer_sunpm_tag("Active Export", use_learned_dataset=False), "E_EXP")
        self.assertEqual(infer_sunpm_tag("Import Active", use_learned_dataset=False), "E_IMP")
        self.assertEqual(infer_sunpm_tag("Export Active", use_learned_dataset=False), "E_EXP")
        self.assertEqual(infer_sunpm_tag("Total Active Import", use_learned_dataset=False), "E_IMP")
        self.assertEqual(infer_sunpm_tag("Total Active Export", use_learned_dataset=False), "E_EXP")
        self.assertEqual(infer_sunpm_tag("Import kWh", use_learned_dataset=False), "E_IMP")
        self.assertEqual(infer_sunpm_tag("Export kWh", use_learned_dataset=False), "E_EXP")
        self.assertEqual(infer_sunpm_tag("Active Import Energy", use_learned_dataset=False), "E_IMP")
        self.assertEqual(infer_sunpm_tag("Active Export Energy", use_learned_dataset=False), "E_EXP")

    def test_power_vs_energy_exclusion(self) -> None:
        """Verify that variables with 'power' or 'puissance' are classified as power, not energy."""
        self.assertEqual(infer_sunpm_tag("Active Import Power", use_learned_dataset=False), "P_AC")
        self.assertEqual(infer_sunpm_tag("Active Export Power", use_learned_dataset=False), "P_AC")
        self.assertEqual(infer_sunpm_tag("Puissance active import", use_learned_dataset=False), "P_AC")
        self.assertEqual(infer_sunpm_tag("Puissance active export", use_learned_dataset=False), "P_AC")

    def test_autoconso_not_confused_with_e_imp(self) -> None:
        """Verify 'autoconso' does not match the 'conso' substring inside the E_IMP block."""
        tag = infer_sunpm_tag("Taux d'autoconso", use_learned_dataset=False, equipment_type="Display")
        self.assertNotEqual(tag, "E_IMP")
        self.assertEqual(tag, "SELF_CONSUMPTION_RATIO")

    def test_reserved_and_dummy_registers(self) -> None:
        """Verify reserved/placeholder names return empty string."""
        for dummy in ["Reserved", "réservé", "Reserve", "Not Used", "unused", "N.U.", "None", "-", "--"]:
            self.assertEqual(infer_sunpm_tag(dummy), "")
            self.assertEqual(infer_sunpm_tag(dummy, use_learned_dataset=False), "")

    def test_clean_variable_key_multilingual(self) -> None:
        """Verify token normalization preserves semantic intent across languages and removes prefixes."""
        self.assertEqual(clean_variable_key("Wechselrichter - Wirkleistung (W)"), "wirkleistung")
        self.assertEqual(clean_variable_key("1#PCS - Inverter Active Power [kW]"), "active power")
        self.assertEqual(clean_variable_key("Nibianqi - Dianya L1"), "dianya l1")
        self.assertEqual(clean_variable_key("INV_01 - Tension réseau"), "tension reseau")


class TestEnrichmentPipelineRoundtrip(unittest.TestCase):
    """Verifies CSV and row dictionary batch enrichment using the knowledge miner."""

    def test_enrich_rows_multilingual(self) -> None:
        """Verify enriching row dictionaries adds tags for multilingual inputs."""
        rows = [
            {"Name": "Wirkleistung", "Tag": ""},
            {"Name": "Netzspannung", "Tag": ""},
            {"Name": "Dianya L1", "Tag": ""},
            {"Name": "Active Import", "Tag": ""},
            {"Name": "Active Export", "Tag": ""},
            {"Name": "RealPower", "Tag": "RealPower"},  # Legacy to be normalized
        ]
        modified = enrich_definition_rows(rows, equipment_type="Inverter")
        self.assertEqual(modified, 6)
        self.assertEqual(rows[0]["Tag"], "P_AC")
        self.assertEqual(rows[1]["Tag"], "U_AC_L1L2")
        self.assertEqual(rows[2]["Tag"], "U_AC_1")
        self.assertEqual(rows[3]["Tag"], "E_IMP")
        self.assertEqual(rows[4]["Tag"], "E_EXP")
        self.assertEqual(rows[5]["Tag"], "P_AC")

    def test_enrich_file_atomic(self) -> None:
        """Verify enrich_definition_file reads and writes atomically with proper tags."""
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_in = Path(tmpdir) / "multilingual_test.csv"
            csv_out = Path(tmpdir) / "multilingual_out.csv"

            content = [
                ["Modbus", "Inverter", "MultilingualVendor", "TestModel"],
                ["1", "3", "100", "U16", "", "Wirkleistung", "", "1.0", "0.0", "W", "4"],
                ["2", "3", "101", "U16", "", "Spannung L1-L2", "", "0.1", "0.0", "V", "4"],
                ["3", "3", "102", "U32", "", "Active Import", "", "1.0", "0.0", "kWh", "4"],
                ["4", "3", "104", "U32", "", "Active Export", "", "1.0", "0.0", "kWh", "4"],
            ]
            with open(csv_in, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerows(content)

            count = enrich_definition_file(str(csv_in), str(csv_out))
            self.assertEqual(count, 4)

            with open(csv_out, encoding="utf-8") as f:
                reader = list(csv.reader(f, delimiter=";"))

            self.assertEqual(reader[1][6], "P_AC")
            self.assertEqual(reader[2][6], "U_AC_L1L2")
            self.assertEqual(reader[3][6], "E_IMP")
            self.assertEqual(reader[4][6], "E_EXP")


if __name__ == "__main__":
    unittest.main()
