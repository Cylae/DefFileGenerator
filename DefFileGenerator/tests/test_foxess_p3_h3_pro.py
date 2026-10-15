#!/usr/bin/env python3
"""
Unit and integration tests for Fox ESS P3 Pro / H3 Pro regenerated definitions:
- Inverter (FOXESS_P3_H3_PRO_INVERTER.csv)
- Meter (FOXESS_P3_H3_PRO_METER.csv)
- Battery (FOXESS_P3_H3_PRO_BATTERY.csv)
"""

import csv
import json
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.report_generator import analyze_definition_file

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
TAGGED_DIR = REPO_ROOT / "tagged_definitions"
GEN_DIR = REPO_ROOT / "generated_definitions"
DOCS_DIR = REPO_ROOT / "docs"


class TestFoxESSP3H3Pro(unittest.TestCase):
    """Test suite verifying 100% compliance of Fox ESS P3 Pro / H3 Pro definitions."""

    generator: Generator
    inv_file: Path
    met_file: Path
    bat_file: Path

    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = Generator()
        cls.inv_file = TAGGED_DIR / "INV_Fox_ESS_FOXESS_P3_H3_PRO_INVERTER.csv"
        cls.met_file = TAGGED_DIR / "MET_Fox_ESS_FOXESS_P3_H3_PRO_METER.csv"
        cls.bat_file = TAGGED_DIR / "BAT_Fox_ESS_FOXESS_P3_H3_PRO_BATTERY.csv"

    def test_inverter_definition_compliance(self) -> None:
        """Verify Inverter definition schema, mandatory tags, and commands."""
        self.assertTrue(self.inv_file.exists(), f"Missing {self.inv_file}")
        report = self.generator.validate_csv_detailed(str(self.inv_file), strict=False, strict_overlap=False)
        self.assertTrue(report.is_valid, f"Inverter validation errors: {[i.message for i in report.issues if i.severity == 'ERROR']}")

        with open(self.inv_file, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            self.assertEqual(header[1], "Inverter")
            self.assertEqual(header[2], "FoxESS")
            rows = list(reader)

        self.assertGreaterEqual(len(rows), 280)
        tags = {r[6].strip() for r in rows if r[6].strip()}
        mandatory_tags = {"MODEL", "SERIAL", "MANUFACTURER", "P_RATED", "P_AC", "Q_AC", "PF", "FREQ", "TEMP_INT", "E_TOT", "E_DAY", "CMD_ON", "CMD_OFF"}
        for mt in mandatory_tags:
            self.assertIn(mt, tags, f"Inverter definition must contain mandatory SunPM tag '{mt}'")

        # Verify command actions
        for r in rows:
            if r[6].startswith("CMD_"):
                self.assertIn(r[10], ("1", "0", "RW", "W"))

    def test_meter_definition_compliance(self) -> None:
        """Verify Meter definition schema and 3-phase grid telemetry tags."""
        self.assertTrue(self.met_file.exists(), f"Missing {self.met_file}")
        report = self.generator.validate_csv_detailed(str(self.met_file), strict=False, strict_overlap=False)
        self.assertTrue(report.is_valid, f"Meter validation errors: {[i.message for i in report.issues if i.severity == 'ERROR']}")

        with open(self.met_file, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            self.assertEqual(header[1], "Meter")
            self.assertEqual(header[2], "FoxESS")
            rows = list(reader)

        self.assertGreaterEqual(len(rows), 60)
        tags = {r[6].strip() for r in rows if r[6].strip()}
        mandatory_tags = {
            "SERIAL", "MANUFACTURER", "MODEL",
            "U_AC_1", "U_AC_2", "U_AC_3",
            "I_AC_1", "I_AC_2", "I_AC_3",
            "P_AC", "P_AC_1", "P_AC_2", "P_AC_3",
            "Q_AC", "Q_AC_1", "Q_AC_2", "Q_AC_3",
            "S_AC", "S_AC_1", "S_AC_2", "S_AC_3",
            "PF", "FREQ",
        }
        for mt in mandatory_tags:
            self.assertIn(mt, tags, f"Meter definition must contain mandatory SunPM tag '{mt}'")

    def test_battery_definition_compliance(self) -> None:
        """Verify Battery definition schema and BMS telemetry/alarm tags."""
        self.assertTrue(self.bat_file.exists(), f"Missing {self.bat_file}")
        report = self.generator.validate_csv_detailed(str(self.bat_file), strict=False, strict_overlap=False)
        self.assertTrue(report.is_valid, f"Battery validation errors: {[i.message for i in report.issues if i.severity == 'ERROR']}")

        with open(self.bat_file, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            self.assertEqual(header[1], "Battery")
            self.assertEqual(header[2], "FoxESS")
            rows = list(reader)

        self.assertGreaterEqual(len(rows), 890)
        tags = {r[6].strip() for r in rows if r[6].strip()}
        mandatory_tags = {"MODEL", "SERIAL", "MANUFACTURER", "U_DC", "I_DC", "TEMP", "SOC", "SOH", "TEMP_MAX", "TEMP_MIN", "U_CELL_MAX", "U_CELL_MIN", "ALARM_1"}
        for mt in mandatory_tags:
            self.assertIn(mt, tags, f"Battery definition must contain mandatory SunPM tag '{mt}'")

        # Verify no duplicate unique tags
        tag_counts: dict[str, int] = {}
        for r in rows:
            t = r[6].strip()
            if t:
                tag_counts[t] = tag_counts.get(t, 0) + 1
        for unique_tag in ("MODEL", "SERIAL", "MANUFACTURER", "SOC", "SOH", "TEMP", "U_DC", "I_DC"):
            if unique_tag in tag_counts:
                self.assertEqual(tag_counts[unique_tag], 1, f"Tag '{unique_tag}' must be unique, found {tag_counts[unique_tag]} occurrences")

    def test_technical_report_generated(self) -> None:
        """Verify technical report exists in docs and tagged_definitions."""
        doc_report = DOCS_DIR / "RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md"
        tagged_report = TAGGED_DIR / "RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md"
        json_report = TAGGED_DIR / "foxess_p3_h3_pro_report.json"

        self.assertTrue(doc_report.exists(), "Missing docs technical report")
        self.assertTrue(tagged_report.exists(), "Missing tagged_definitions technical report")
        self.assertTrue(json_report.exists(), "Missing json metrics report")

        content = doc_report.read_text(encoding="utf-8")
        self.assertIn("P3 Pro / H3 Pro Series", content)
        self.assertIn("FOXESS_P3_H3_PRO_INVERTER.csv", content)
        self.assertIn("FOXESS_P3_H3_PRO_METER.csv", content)
        self.assertIn("FOXESS_P3_H3_PRO_BATTERY.csv", content)

        data = json.loads(json_report.read_text(encoding="utf-8"))
        self.assertTrue(data["summary"]["all_valid"])
        self.assertEqual(data["manufacturer"], "Fox ESS")


if __name__ == "__main__":
    unittest.main()
