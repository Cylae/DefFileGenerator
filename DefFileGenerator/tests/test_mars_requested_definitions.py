#!/usr/bin/env python3
"""
Unit and integration tests for Mars Renewable regenerated definition files:
- MarsRenewable_LocalControler.csv (ID 247: LC Overview FC4 + EMS Writes FC3)
- MarsRenewable_PCS-Bank-Rack.csv (Device ID 1: FC2 Alarms + DIO1 + PCS1 + Bank1)
- MarsRenewable_Rack-Cell.csv (Device ID 2: Rack 1 + Cell 1)
- MarsRenewable_LC-AC.csv (Device ID 1: Liquid Cooling 1 + Air Conditioning 1)
- MarsRenewable_Meter.csv (Device ID 1: Bidirectional Meter 1)
- MarsRenewable_PV.csv (Device ID 1: PV Inverter 1)
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.mars_generator import generate_mars_requested_definitions

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
XLSX_PATH = REPO_ROOT / "stress_test_data" / "MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
GEN_DIR = REPO_ROOT / "generated_definitions"


class TestMarsRequestedDefinitions(unittest.TestCase):
    """Test suite verifying 100% compliance of Mars Renewable definitions."""

    generator: Generator
    lc_file: Path
    pcs_bank_file: Path
    rack_cell_file: Path
    lc_ac_file: Path
    meter_file: Path
    pv_file: Path

    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = Generator()
        cls.lc_file = GEN_DIR / "MarsRenewable_LocalControler.csv"
        cls.pcs_bank_file = GEN_DIR / "MarsRenewable_PCS-Bank-Rack.csv"
        cls.rack_cell_file = GEN_DIR / "MarsRenewable_Rack-Cell.csv"
        cls.lc_ac_file = GEN_DIR / "MarsRenewable_LC-AC.csv"
        cls.meter_file = GEN_DIR / "MarsRenewable_Meter.csv"
        cls.pv_file = GEN_DIR / "MarsRenewable_PV.csv"

        # Ensure files are generated fresh before running tests
        if XLSX_PATH.exists():
            generate_mars_requested_definitions(XLSX_PATH, GEN_DIR)

    def _read_definition(self, path: Path) -> tuple[list[str], list[list[str]]]:
        self.assertTrue(path.exists(), f"Definition file does not exist: {path}")
        with open(path, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            rows = list(reader)
        return header, rows

    def test_mars_localcontroller_compliance(self) -> None:
        """Verify MarsRenewable_LocalControler.csv schema, FC4 Overview, and FC3 EMS setpoints."""
        report = self.generator.validate_csv_detailed(str(self.lc_file), strict=True, strict_overlap=True)
        self.assertTrue(report.is_valid, f"Validation errors: {[i.message for i in report.issues]}")

        header, rows = self._read_definition(self.lc_file)
        self.assertEqual(len(header), 11)
        self.assertEqual(header[0], "modbusTCP")
        self.assertEqual(header[1], "Inverter")
        self.assertEqual(header[2], "MarsRenewable")
        self.assertEqual(header[3], "LocalControler")
        self.assertEqual(len(rows), 12)

        # Overview FC4
        fc4_rows = [r for r in rows if r[1] == "4"]
        self.assertEqual(len(fc4_rows), 5)
        self.assertEqual(fc4_rows[0][2], "0")
        self.assertIn("Measurement point version number", fc4_rows[0][5])

        # EMS Writes FC3
        fc3_rows = [r for r in rows if r[1] == "3"]
        self.assertEqual(len(fc3_rows), 7)
        cmd_p_row = [r for r in fc3_rows if r[6] == "CMD_P_AC_SET_ABS"]
        self.assertEqual(len(cmd_p_row), 1)
        self.assertEqual(cmd_p_row[0][2], "4")
        self.assertEqual(cmd_p_row[0][9], "W")
        self.assertEqual(cmd_p_row[0][10], "1")

        # AC Frequency with 0.01 scaling
        freq_row = [r for r in fc3_rows if r[2] == "9"][0]
        self.assertEqual(freq_row[7], "0.010000")
        self.assertEqual(freq_row[9], "Hz")

    def test_mars_pcs_bank_rack_compliance(self) -> None:
        """Verify MarsRenewable_PCS-Bank-Rack.csv (Device ID 1: FC2 Alarms + DIO1 + PCS1 + Bank1)."""
        report = self.generator.validate_csv_detailed(str(self.pcs_bank_file), strict=True, strict_overlap=True)
        self.assertTrue(report.is_valid, f"Validation errors: {[i.message for i in report.issues]}")

        header, rows = self._read_definition(self.pcs_bank_file)
        self.assertEqual(len(header), 11)
        self.assertEqual(header[0], "modbusTCP")
        self.assertEqual(header[1], "Battery")
        self.assertEqual(header[2], "MarsRenewable")
        self.assertEqual(len(rows), 80)

        # Check FC2 discrete alarms (4 alarms)
        fc2_rows = [r for r in rows if r[1] == "2"]
        self.assertEqual(len(fc2_rows), 4)
        for r in fc2_rows:
            self.assertEqual(r[10], "9")

        # Check DIO1 channels (14 registers: 20006..20021)
        dio_rows = [r for r in rows if r[5].startswith("DIO1 ")]
        self.assertEqual(len(dio_rows), 14)

        # Check PCS1 telemetry and branches (49 registers)
        pcs_rows = [r for r in rows if r[5].startswith("PCS1 ")]
        self.assertEqual(len(pcs_rows), 49)

        # Check Bank1 telemetry (13 registers)
        bank_rows = [r for r in rows if r[5].startswith("Bank1 ")]
        self.assertEqual(len(bank_rows), 13)

        # Check canonical tags
        tags = {r[6] for r in rows if r[6]}
        expected_tags = {
            "U_AC_L1L2", "U_AC_L2L3", "U_AC_L3L1",
            "I_AC_1", "I_AC_2", "I_AC_3",
            "FREQ", "P_AC", "Q_AC", "S_AC", "PF",
            "TEMP_INT", "TEMP_AMB", "E_TOT", "E_DAY",
            "U_DC", "I_DC", "SOC", "SOH", "R_ISO",
        }
        for et in expected_tags:
            self.assertIn(et, tags, f"Missing expected SunPM tag '{et}' in PCS-Bank-Rack")

    def test_mars_rack_cell_compliance(self) -> None:
        """Verify MarsRenewable_Rack-Cell.csv (Device ID 2: Rack 1 + Cell 1)."""
        report = self.generator.validate_csv_detailed(str(self.rack_cell_file), strict=True, strict_overlap=True)
        self.assertTrue(report.is_valid, f"Validation errors: {[i.message for i in report.issues]}")

        header, rows = self._read_definition(self.rack_cell_file)
        self.assertEqual(len(header), 11)
        self.assertEqual(header[0], "modbusTCP")
        self.assertEqual(header[1], "Battery")
        self.assertEqual(header[2], "MarsRenewable")
        self.assertEqual(len(rows), 35)

        # 30 Rack registers + 5 Cell registers
        rack_rows = [r for r in rows if r[5].startswith("Rack ")]
        cell_rows = [r for r in rows if r[5].startswith("Cell 1 ")]
        self.assertEqual(len(rack_rows), 30)
        self.assertEqual(len(cell_rows), 5)

        # Canonical tags
        tags = {r[6] for r in rows if r[6]}
        self.assertIn("U_DC", tags)
        self.assertIn("I_DC", tags)
        self.assertIn("SOC", tags)
        self.assertIn("SOH", tags)
        self.assertIn("R_ISO", tags)
        self.assertIn("TEMP_CELL", tags)

    def test_mars_lc_ac_compliance(self) -> None:
        """Verify MarsRenewable_LC-AC.csv (Device ID 1: Liquid Cooling 1 + Air Conditioning 1)."""
        report = self.generator.validate_csv_detailed(str(self.lc_ac_file), strict=True, strict_overlap=True)
        self.assertTrue(report.is_valid, f"Validation errors: {[i.message for i in report.issues]}")

        header, rows = self._read_definition(self.lc_ac_file)
        self.assertEqual(len(header), 11)
        self.assertEqual(header[0], "modbusTCP")
        self.assertEqual(header[1], "Inverter")
        self.assertEqual(header[2], "MarsRenewable")
        self.assertEqual(len(rows), 34)

        lc_rows = [r for r in rows if r[5].startswith("LiquidCooling1 ")]
        ac_rows = [r for r in rows if r[5].startswith("AC1 ")]
        self.assertEqual(len(lc_rows), 18)
        self.assertEqual(len(ac_rows), 16)

        tags = {r[6] for r in rows if r[6]}
        self.assertIn("TEMP_AMB", tags)
        self.assertIn("TEMP", tags)
        self.assertIn("HUMIDITY", tags)

    def test_mars_standalone_lc_and_ac_compliance(self) -> None:
        """Verify standalone MarsRenewable_LC.csv (18 regs) and MarsRenewable_AC.csv (16 regs)."""
        lc_path = GEN_DIR / "MarsRenewable_LC.csv"
        ac_path = GEN_DIR / "MarsRenewable_AC.csv"

        self.assertTrue(lc_path.exists(), f"Missing {lc_path}")
        self.assertTrue(ac_path.exists(), f"Missing {ac_path}")

        rep_lc = self.generator.validate_csv_detailed(str(lc_path), strict=True, strict_overlap=True)
        self.assertTrue(rep_lc.is_valid, f"LC validation errors: {[i.message for i in rep_lc.issues]}")

        rep_ac = self.generator.validate_csv_detailed(str(ac_path), strict=True, strict_overlap=True)
        self.assertTrue(rep_ac.is_valid, f"AC validation errors: {[i.message for i in rep_ac.issues]}")

        _, lc_rows = self._read_definition(lc_path)
        self.assertEqual(len(lc_rows), 18)
        lc_tags = {r[6] for r in lc_rows if r[6]}
        self.assertIn("TEMP_AMB", lc_tags)

        _, ac_rows = self._read_definition(ac_path)
        self.assertEqual(len(ac_rows), 16)
        ac_tags = {r[6] for r in ac_rows if r[6]}
        self.assertIn("TEMP", ac_tags)
        self.assertIn("HUMIDITY", ac_tags)

    def test_mars_meter_compliance(self) -> None:
        """Verify MarsRenewable_Meter.csv (Device ID 1: Bidirectional Meter 1)."""
        report = self.generator.validate_csv_detailed(str(self.meter_file), strict=True, strict_overlap=True)
        self.assertTrue(report.is_valid, f"Validation errors: {[i.message for i in report.issues]}")

        header, rows = self._read_definition(self.meter_file)
        self.assertEqual(len(header), 11)
        self.assertEqual(header[0], "modbusTCP")
        self.assertEqual(header[1], "Meter")
        self.assertEqual(header[2], "MarsRenewable")
        self.assertEqual(len(rows), 11)

        tags = {r[6] for r in rows if r[6]}
        self.assertIn("E_IMP", tags)
        self.assertIn("E_EXP", tags)
        self.assertIn("P_AC", tags)

    def test_mars_pv_compliance(self) -> None:
        """Verify MarsRenewable_PV.csv (Device ID 1: PV Inverter 1 with 6 MPPT branches)."""
        report = self.generator.validate_csv_detailed(str(self.pv_file), strict=True, strict_overlap=True)
        self.assertTrue(report.is_valid, f"Validation errors: {[i.message for i in report.issues]}")

        header, rows = self._read_definition(self.pv_file)
        self.assertEqual(len(header), 11)
        self.assertEqual(header[0], "modbusTCP")
        self.assertEqual(header[1], "Inverter")
        self.assertEqual(header[2], "MarsRenewable")
        self.assertEqual(len(rows), 49)

        tags = {r[6] for r in rows if r[6]}
        expected_ac_tags = {
            "U_AC_1", "U_AC_2", "U_AC_3",
            "I_AC_1", "I_AC_2", "I_AC_3",
            "FREQ", "P_AC", "Q_AC", "S_AC", "PF",
            "TEMP_INT", "TEMP_AMB", "E_TOT", "E_DAY",
        }
        for eat in expected_ac_tags:
            self.assertIn(eat, tags, f"Missing expected AC tag '{eat}' in PV")

        # Verify 6 MPPT DC branch tags
        for b in range(1, 7):
            self.assertIn(f"P_DC_MPPT_{b}", tags)
            self.assertIn(f"U_DC_MPPT_{b}", tags)
            self.assertIn(f"I_DC_MPPT_{b}", tags)

    def test_zero_duplicate_tags_all_files(self) -> None:
        """Verify that within every single definition file, all assigned tags are 100% unique."""
        files = [self.lc_file, self.pcs_bank_file, self.rack_cell_file, self.lc_ac_file, self.meter_file, self.pv_file]
        for f in files:
            _, rows = self._read_definition(f)
            assigned_tags = [r[6] for r in rows if r[6]]
            self.assertEqual(
                len(assigned_tags),
                len(set(assigned_tags)),
                f"Duplicate tags found in {f.name}: {[t for t in assigned_tags if assigned_tags.count(t) > 1]}"
            )

    def test_strict_webdynsunpm_formatting(self) -> None:
        """Verify strict WebdynSunPM formatting: 11 columns, 6 decimals, valid actions."""
        files = [self.lc_file, self.pcs_bank_file, self.rack_cell_file, self.lc_ac_file, self.meter_file, self.pv_file]
        for f in files:
            header, rows = self._read_definition(f)
            self.assertEqual(len(header), 11, f"Header in {f.name} does not have 11 columns")
            for idx, r in enumerate(rows, start=1):
                self.assertEqual(len(r), 11, f"Row {idx} in {f.name} does not have 11 columns: {r}")
                self.assertEqual(r[0], str(idx), f"Index mismatch in {f.name} line {idx}")
                self.assertIn(r[1], ("1", "2", "3", "4"), f"Invalid FC {r[1]} in {f.name} line {idx}")
                self.assertIn(r[3], ("U16", "U32", "I16", "I32", "F32"), f"Invalid type {r[3]} in {f.name} line {idx}")
                self.assertIn(r[10], ("0", "1", "4", "9"), f"Invalid action {r[10]} in {f.name} line {idx}")
                # Verify 6-decimal float format
                self.assertRegex(r[7], r"^-?\d+\.\d{6}$", f"CoefA not 6 decimals in {f.name} line {idx}: {r[7]}")
                self.assertRegex(r[8], r"^-?\d+\.\d{6}$", f"CoefB not 6 decimals in {f.name} line {idx}: {r[8]}")


if __name__ == "__main__":
    unittest.main()
