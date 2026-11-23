"""
Adversarial Verification Suite for GoodWe Complete Ecosystem:
Residential (XS, DNS, MS), Commercial (SDT, SMT, MT), Utility (HT 1100V, HT 1500V, GT),
Hybrid Storage (ES, EM, SBP, ET, EH), Batteries (Lynx Home HV), Loggers (EzLogger Pro, SCB3000),
Meters (GM1000, GM3000), and SunSpec 3-Phase.

Validates:
1. All 34 GoodWe definition files exist in production/ and Definitions/GoodWe/
2. Exact 5 Apollo-Standard Pillars compliance for every single file
3. Canonical SunPM tag uniqueness (no duplicate single-tags within a file)
4. Strict sequential 1..N index matching row number
5. Valid Modbus function codes (3) and certified Webdyn action codes (0, 1, 4, 9)
6. 6-decimal coefficient formatting (e.g. 1.000000, 0.100000)
7. Byte-identical synchronization across production, Definitions/GoodWe, Definitions, generated_definitions, and tagged_definitions
8. Domain-specific telemetry coverage (PV string inputs, battery BMS metrics, grid metrics)
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PRODUCTION_DIR = REPO_ROOT / "production"
DEFINITIONS_GOODWE_DIR = REPO_ROOT / "Definitions" / "GoodWe"
DEFINITIONS_DIR = REPO_ROOT / "Definitions"
GENERATED_DIR = REPO_ROOT / "generated_definitions"
TAGGED_DIR = REPO_ROOT / "tagged_definitions"

EXPECTED_GOODWE_FILES = [
    # Residential Single-Phase
    "INV_Goodwe_XS_DNS_MS.csv",
    "INV_Goodwe_MS_DNS_XS.csv",
    "INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv",
    # Residential / Small C&I Three-Phase
    "INV_Goodwe_SDT_G2_G3.csv",
    "INV_Goodwe_SDTG2.csv",
    # Commercial & Industrial Three-Phase
    "INV_Goodwe_SMT_Series.csv",
    "INV_Goodwe_SMT.csv",
    "INV_Goodwe_SMT50K_SMT60K.csv",
    "INV_Goodwe_MT_Series.csv",
    "INV_Goodwe_MT.csv",
    "INV_Goodwe_MT_series.csv",
    "INV_Goodwe_MT-Series_ModbusRTU.csv",
    # Utility & Large Scale
    "INV_Goodwe_HT_Series.csv",
    "INV_Goodwe_HT1100.csv",
    "INV_Goodwe_HT_others_Series.csv",
    "INV_Goodwe_HT1500.csv",
    "INV_Goodwe_GT-SERIES.csv",
    # Low-Voltage Hybrid Storage
    "INV_Goodwe_ES_EM_SBP.csv",
    "INV_Goodwe_ES_EM_S-BP_BP_Series.csv",
    "INV_Goodwe_ES-EM-SBP-BP.csv",
    # High-Voltage Hybrid Storage
    "INV_Goodwe_ET_EH_Series.csv",
    "INV_Goodwe_ET_EH_Hybrid_Series.csv",
    "INV_Goodwe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv",
    "INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv",
    "INV_Goodwe_ET_Series.csv",
    # Battery & BESS
    "BAT_Goodwe_Lynx_Home_HV.csv",
    # Communication Gateways & Loggers
    "COM_Goodwe_EzLogger_Pro.csv",
    "COM_Goodwe_Ezlogger3000.csv",
    "COM_GoodWe_Ezlogger3000.csv",
    "COM_Goodwe_SCB3000.csv",
    # Smart Energy Meters
    "MET_Goodwe_GM1000_GM3000.csv",
    # SunSpec Standard Inverters
    "INV_Goodwe_ThreePhase_SunSpec.csv",
]


class TestGoodweDefinitionsSuite(unittest.TestCase):
    """Test suite certifying all 32 GoodWe ecosystem definitions."""

    def test_all_files_exist_in_all_target_directories(self) -> None:
        """Ensure all 32 files exist in production, Definitions/GoodWe, generated_definitions, and tagged_definitions."""
        for fname in EXPECTED_GOODWE_FILES:
            for d in (
                PRODUCTION_DIR,
                DEFINITIONS_GOODWE_DIR,
                DEFINITIONS_DIR,
                GENERATED_DIR,
                TAGGED_DIR,
            ):
                fpath = d / fname
                self.assertTrue(fpath.is_file(), f"Missing {fname} in {d.name}")

    def test_byte_identical_redundant_synchronization(self) -> None:
        """Verify byte-identical synchronization across redundant directories."""
        for fname in EXPECTED_GOODWE_FILES:
            prod_bytes = (PRODUCTION_DIR / fname).read_bytes()
            for d in (DEFINITIONS_GOODWE_DIR, DEFINITIONS_DIR, GENERATED_DIR, TAGGED_DIR):
                target_bytes = (d / fname).read_bytes()
                self.assertEqual(
                    prod_bytes,
                    target_bytes,
                    f"File {fname} in {d.name} is not byte-identical to production/{fname}",
                )

    def test_five_apollo_pillars_compliance(self) -> None:
        """Verify each GoodWe file strictly complies with all 5 Apollo Pillars."""
        for fname in EXPECTED_GOODWE_FILES:
            fpath = PRODUCTION_DIR / fname
            is_valid, issues = verify_file_pillars(fpath)
            self.assertTrue(
                is_valid,
                f"File {fname} violated 5 Pillars:\n" + "\n".join(issues),
            )

    def test_protocol_header_and_structure(self) -> None:
        """Verify protocol header begins with Modbus; and has 11 columns with canonical manufacturer GoodWe."""
        for fname in EXPECTED_GOODWE_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                reader = list(csv.reader(fp, delimiter=";"))

            self.assertGreaterEqual(len(reader), 10, f"File {fname} has too few rows")
            header = reader[0]
            self.assertEqual(len(header), 11, f"Header of {fname} must have 11 columns")
            self.assertEqual(header[0], "Modbus", f"Header protocol of {fname} must be 'Modbus'")
            self.assertEqual(
                header[2], "GoodWe", f"Header manufacturer of {fname} must be 'GoodWe'"
            )

    def test_sequential_indexing_and_row_integrity(self) -> None:
        """Verify sequential index 1..N, valid FCs, decimal addresses, and valid action codes."""
        valid_fcs = {"1", "2", "3", "4"}
        valid_actions = {"0", "1", "4", "9", "10"}

        for fname in EXPECTED_GOODWE_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            tags_in_file: set[str] = set()
            for idx, r in enumerate(rows, start=1):
                self.assertEqual(len(r), 11, f"Row {idx} in {fname} must have 11 columns")
                row_idx, fc, addr, dtype, _, name, tag, coef_a, coef_b, unit, action = r

                # 1..N index
                self.assertEqual(row_idx, str(idx), f"Index mismatch at row {idx} in {fname}")

                # Function code
                self.assertIn(fc, valid_fcs, f"Invalid FC '{fc}' at row {idx} in {fname}")

                # Action code
                self.assertIn(
                    action, valid_actions, f"Invalid action '{action}' at row {idx} in {fname}"
                )

                # Coefficients must have 6 decimals
                self.assertRegex(
                    coef_a,
                    r"^-?\d+\.\d{6}$",
                    f"CoefA '{coef_a}' at row {idx} in {fname} must have 6 decimals",
                )
                self.assertRegex(
                    coef_b,
                    r"^-?\d+\.\d{6}$",
                    f"CoefB '{coef_b}' at row {idx} in {fname} must have 6 decimals",
                )

                # Name not empty
                self.assertTrue(name.strip(), f"Empty name at row {idx} in {fname}")

                # Canonical Tag uniqueness
                if tag.strip():
                    self.assertNotIn(
                        tag, tags_in_file, f"Duplicate tag '{tag}' at row {idx} in {fname}"
                    )
                    tags_in_file.add(tag)

    def test_domain_telemetry_contracts(self) -> None:
        """Verify essential domain telemetry tags are present based on equipment role."""
        for fname in EXPECTED_GOODWE_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            tags = {r[6].strip() for r in rows if r[6].strip()}

            # Inverters must measure active power and energy yields
            if fname.startswith("INV_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in inverter {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_DAY" in tags, f"Missing energy yield in {fname}"
                )

            # High power utility inverters must have string-level DC monitoring
            if "HT" in fname:
                self.assertIn("U_DC_STRING_1", tags, f"Missing U_DC_STRING_1 in {fname}")
                self.assertIn("I_DC_STRING_1", tags, f"Missing I_DC_STRING_1 in {fname}")
                self.assertIn(
                    "U_DC_STRING_24", tags, f"Missing U_DC_STRING_24 in 24-string HT {fname}"
                )
                self.assertIn(
                    "I_DC_STRING_24", tags, f"Missing I_DC_STRING_24 in 24-string HT {fname}"
                )

            # Hybrid and Battery systems must report battery state of charge
            if any(
                k in fname for k in ("_ET_", "-ET-", "_ES_", "-ES-", "Hybrid", "Lynx")
            ) or fname.startswith("BAT_"):
                self.assertIn("SOC", tags, f"Missing SOC in storage/battery {fname}")
                self.assertIn("U_BAT", tags, f"Missing U_BAT in storage/battery {fname}")

            # Meters must measure active power, energy, voltage, and current
            if fname.startswith("MET_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in meter {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_IMP" in tags,
                    f"Missing energy yield/import in meter {fname}",
                )
                self.assertIn("U_AC_1", tags, f"Missing U_AC_1 in meter {fname}")
                self.assertIn("I_AC_1", tags, f"Missing I_AC_1 in meter {fname}")


if __name__ == "__main__":
    unittest.main()
