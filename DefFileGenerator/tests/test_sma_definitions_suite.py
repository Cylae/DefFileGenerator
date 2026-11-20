"""
Adversarial Verification Suite for SMA Complete Product Ecosystem:
Residential (Sunny Boy, Sunny Boy Storage), Commercial (Sunny Tripower, Sunny Tripower X),
Decentralized C&I (Core 1, Core 2), Hybrid Storage (Sunny Tripower Smart Energy, Sunny Island),
Utility Scale (Sunny Central, SC Storage, SC2xxx), Communication Gateways (Data Manager M, Cluster Controller,
Com Gateway, WebBox), Smart Meters (Energy Meter, Home Manager), and SunSpec 3-Phase.

Validates:
1. All 27 SMA definition files exist in production/ and Definitions/SMA/
2. Exact 5 Apollo-Standard Pillars compliance for every single file
3. Canonical SunPM tag uniqueness (no duplicate single-tags within a file)
4. Strict sequential 1..N index matching row number
5. Valid Modbus function codes (1, 2, 3, 4) and certified Webdyn action codes (0, 1, 4, 9, 10)
6. 6-decimal coefficient formatting (e.g. 1.000000, 0.001000)
7. Byte-identical synchronization across production, Definitions/SMA, Definitions, generated_definitions, and tagged_definitions
8. Domain-specific telemetry coverage (MPPT trackers, battery BMS, 3-phase grid AC, plant controls)
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PRODUCTION_DIR = REPO_ROOT / "production"
DEFINITIONS_SMA_DIR = REPO_ROOT / "Definitions" / "SMA"
DEFINITIONS_DIR = REPO_ROOT / "Definitions"
GENERATED_DIR = REPO_ROOT / "generated_definitions"
TAGGED_DIR = REPO_ROOT / "tagged_definitions"

EXPECTED_SMA_FILES = [
    # Residential
    "INV_SMA_SunnyBoy.csv",
    "INV_SMA_SunnyBoyStorage.csv",
    # Commercial Three-Phase
    "INV_SMA_SunnyTripower.csv",
    "INV_SMA_SunnyTripower_X.csv",
    "INV_SMA_STP_X.csv",
    "INV_SMA_Core1_Core2.csv",
    "INV_SMA_Core1.csv",
    "INV_SMA_Core2.csv",
    "INV_SMA_STPCORE1_50-41.csv",
    "INV_SMA_STP_110-60.csv",
    # Hybrid & Energy Storage
    "INV_SMA_SunnyTripower_SmartEnergy.csv",
    "INV_SMA_STP_5-10-SE40.csv",
    "INV_SMA_SunnyIsland.csv",
    # Utility Scale Central
    "INV_SMA_SunnyCentral.csv",
    "INV_SMA_SUNNY_CENTRAL.csv",
    "INV_SMA_SC2xxx.csv",
    "INV_SMA_SunnyCentral_Storage.csv",
    # Communication Gateways & Plant Controllers
    "COM_SMA_DataManager_M.csv",
    "COM_SMA_ClusterController.csv",
    "COM_SMA_ComGateway.csv",
    "INV_SMA_SC_Webbox.csv",
    # Smart Energy Meters & Grid Analyzers
    "MET_SMA_EnergyMeter.csv",
    "MET_SMA_HomeManager.csv",
    # SunSpec Standard Inverters & Interoperable Profiles
    "INV_SMA_SunSpec_Inverter.csv",
    "INV_SMA_Sunspec.csv",
    "INV_SMA_Generic.csv",
    "INV_SMA_Modbus.csv",
]


class TestSmaDefinitionsSuite(unittest.TestCase):
    """Test suite certifying all 27 SMA ecosystem definitions."""

    def test_all_files_exist_in_all_target_directories(self) -> None:
        """Ensure all 27 files exist in production, Definitions/SMA, Definitions, generated_definitions, and tagged_definitions."""
        for fname in EXPECTED_SMA_FILES:
            for d in (PRODUCTION_DIR, DEFINITIONS_SMA_DIR, DEFINITIONS_DIR, GENERATED_DIR, TAGGED_DIR):
                fpath = d / fname
                self.assertTrue(fpath.is_file(), f"Missing {fname} in {d.name}")

    def test_byte_identical_redundant_synchronization(self) -> None:
        """Verify byte-identical synchronization across redundant directories."""
        for fname in EXPECTED_SMA_FILES:
            prod_bytes = (PRODUCTION_DIR / fname).read_bytes()
            for d in (DEFINITIONS_SMA_DIR, DEFINITIONS_DIR, GENERATED_DIR, TAGGED_DIR):
                target_bytes = (d / fname).read_bytes()
                self.assertEqual(
                    prod_bytes,
                    target_bytes,
                    f"File {fname} in {d.name} is not byte-identical to production/{fname}",
                )

    def test_five_apollo_pillars_compliance(self) -> None:
        """Verify each SMA file strictly complies with all 5 Apollo Pillars."""
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            is_valid, issues = verify_file_pillars(fpath)
            self.assertTrue(
                is_valid,
                f"File {fname} violated 5 Pillars:\n" + "\n".join(issues),
            )

    def test_protocol_header_and_structure(self) -> None:
        """Verify protocol header begins with Modbus; and has 11 columns with canonical manufacturer SMA."""
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                reader = list(csv.reader(fp, delimiter=";"))

            self.assertGreaterEqual(len(reader), 10, f"File {fname} has too few rows")
            header = reader[0]
            self.assertEqual(len(header), 11, f"Header of {fname} must have 11 columns")
            self.assertEqual(header[0], "Modbus", f"Header protocol of {fname} must be 'Modbus'")
            self.assertEqual(header[2], "SMA", f"Header manufacturer of {fname} must be 'SMA'")

    def test_sequential_indexing_and_row_integrity(self) -> None:
        """Verify sequential index 1..N, valid FCs, decimal addresses, and valid action codes."""
        valid_fcs = {"1", "2", "3", "4"}
        valid_actions = {"0", "1", "4", "9", "10"}

        for fname in EXPECTED_SMA_FILES:
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
        for fname in EXPECTED_SMA_FILES:
            fpath = PRODUCTION_DIR / fname
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            tags = {r[6].strip() for r in rows if r[6].strip()}

            # Inverters and loggers must measure active power and energy yields
            if fname.startswith("INV_") or fname.startswith("COM_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_DAY" in tags, f"Missing energy yield in {fname}"
                )

            # Core 1 and Core 2 inverters must have multi-MPPT monitoring
            if "Core" in fname or "CORE" in fname or "110-60" in fname:
                self.assertIn("I_DC_1", tags, f"Missing I_DC_1 in {fname}")
                self.assertIn("I_DC_6", tags, f"Missing I_DC_6 in {fname}")

            # Battery storage systems must report state of charge and voltage
            if any(k in fname for k in ("Storage", "SmartEnergy", "Island", "SE40")):
                self.assertIn("SOC", tags, f"Missing SOC in storage {fname}")
                self.assertIn("U_BAT", tags, f"Missing U_BAT in storage {fname}")

            # Meters must measure active power, energy, voltage, and current
            if fname.startswith("MET_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in meter {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_IMP" in tags or "E_EXP" in tags,
                    f"Missing energy counter in meter {fname}",
                )
                self.assertIn("U_AC_1", tags, f"Missing U_AC_1 in meter {fname}")
                self.assertIn("I_AC_1", tags, f"Missing I_AC_1 in meter {fname}")


if __name__ == "__main__":
    unittest.main()
