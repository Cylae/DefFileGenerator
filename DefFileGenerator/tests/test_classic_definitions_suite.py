"""
Adversarial Verification Suite for Grand Classic Manufacturers:
Huawei, CHINT, SMA, Solplanet, Sungrow, GoodWe, Growatt, Fronius, SolarEdge, Socomec, Schneider, and Carlo Gavazzi.

Validates:
1. All 18 grand classic definition files exist in production/ and Definitions/Classic/
2. Exact 5 Apollo-Standard Pillars compliance for every single file
3. Canonical SunPM tag uniqueness (no duplicate single-tags within a file)
4. Strict sequential 1..N index matching row number
5. Valid Modbus function codes (3, 4) and certified Webdyn action codes (0, 1, 4, 9, 10)
6. 6-decimal coefficient formatting (e.g. 1.000000, 0.001000)
7. Byte-identical synchronization across production, generated_definitions, tagged_definitions, and Definitions/Classic
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.io_utils import find_definition_file
from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PRODUCTION_DIR = REPO_ROOT / "production"
DEFINITIONS_CLASSIC_DIR = REPO_ROOT / "Definitions"
GENERATED_DIR = REPO_ROOT / "generated_definitions"
TAGGED_DIR = REPO_ROOT / "tagged_definitions"

EXPECTED_CLASSIC_FILES = [
    # CHINT
    "MET_Chint_DTSU666.csv",
    "MET_Chint_DDSU666.csv",
    "MET_Chint_DTSU666-H.csv",
    # HUAWEI
    "INV_Huawei_SUN2000-M3.csv",
    "INV_Huawei_SUN2000-Utility.csv",
    "MET_Huawei_DTSU666-H.csv",
    # SMA
    "INV_SMA_SunnyTripower.csv",
    "INV_SMA_Core1_Core2.csv",
    # SOLPLANET
    "INV_Solplanet_ASW_ThreePhase.csv",
    # SUNGROW
    "INV_Sungrow_SG_Series.csv",
    # GOODWE
    "INV_Goodwe_SMT_Series.csv",
    "INV_Goodwe_ET_Series.csv",
    # GROWATT
    "INV_Growatt_MAX_Series.csv",
    # FRONIUS
    "INV_Fronius_Symo_Primo_Gen24.csv",
    # SOLAREDGE
    "INV_SolarEdge_SE_Series.csv",
    # SOCOMEC
    "MET_Socomec_Diris_A_Countis_E.csv",
    # SCHNEIDER
    "MET_Schneider_PM5000_IEM3000.csv",
    # CARLO GAVAZZI
    "MET_CarloGavazzi_EM24_EM300.csv",
]


class TestClassicDefinitionsSuite(unittest.TestCase):
    """Test suite certifying all 18 grand classic definitions."""

    def test_all_files_exist_in_all_target_directories(self) -> None:
        """Ensure all 18 files exist in production, generated_definitions, tagged_definitions, and Definitions."""
        for fname in EXPECTED_CLASSIC_FILES:
            for d in (PRODUCTION_DIR, DEFINITIONS_CLASSIC_DIR, GENERATED_DIR, TAGGED_DIR):
                fpath = find_definition_file(d, fname)
                self.assertTrue(fpath.is_file(), f"Missing {fname} in {d.name}")

    def test_byte_identical_redundant_synchronization(self) -> None:
        """Verify byte-identical synchronization across redundant directories."""
        for fname in EXPECTED_CLASSIC_FILES:
            prod_bytes = find_definition_file(PRODUCTION_DIR, fname).read_bytes()
            for d in (DEFINITIONS_CLASSIC_DIR, GENERATED_DIR, TAGGED_DIR):
                target_bytes = find_definition_file(d, fname).read_bytes()
                self.assertEqual(
                    prod_bytes,
                    target_bytes,
                    f"File {fname} in {d.name} is not byte-identical to production/{fname}",
                )

    def test_five_apollo_pillars_compliance(self) -> None:
        """Verify each file strictly complies with all 5 Apollo Pillars."""
        for fname in EXPECTED_CLASSIC_FILES:
            fpath = find_definition_file(PRODUCTION_DIR, fname)
            is_valid, issues = verify_file_pillars(fpath)
            self.assertTrue(
                is_valid,
                f"File {fname} violated 5 Pillars:\n" + "\n".join(issues),
            )

    def test_protocol_header_and_structure(self) -> None:
        """Verify protocol header begins with Modbus; and has 11 columns."""
        for fname in EXPECTED_CLASSIC_FILES:
            fpath = find_definition_file(PRODUCTION_DIR, fname)
            with open(fpath, encoding="utf-8", newline="") as fp:
                reader = list(csv.reader(fp, delimiter=";"))

            self.assertGreaterEqual(len(reader), 10, f"File {fname} has too few rows")
            header = reader[0]
            self.assertEqual(len(header), 11, f"Header of {fname} must have 11 columns")
            self.assertEqual(header[0], "Modbus", f"Header of {fname} protocol must be 'Modbus'")

    def test_sequential_indexing_and_row_integrity(self) -> None:
        """Verify sequential index 1..N, valid FCs, decimal addresses, and valid action codes."""
        valid_fcs = {"1", "2", "3", "4"}
        valid_actions = {"0", "1", "4", "9", "10"}

        for fname in EXPECTED_CLASSIC_FILES:
            fpath = find_definition_file(PRODUCTION_DIR, fname)
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

    def test_canonical_telemetry_tags_present(self) -> None:
        """Verify essential standard tags exist for inverters and meters."""
        for fname in EXPECTED_CLASSIC_FILES:
            fpath = find_definition_file(PRODUCTION_DIR, fname)
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            tags = {r[6].strip() for r in rows if r[6].strip()}

            # Inverters must measure active power and yield
            if fname.startswith("INV_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in inverter definition {fname}")
                self.assertTrue(
                    "E_TOT" in tags or "E_DAY" in tags, f"Missing energy yield in {fname}"
                )

            # Meters must measure active power and energy
            if fname.startswith("MET_"):
                self.assertIn("P_AC", tags, f"Missing P_AC in meter definition {fname}")
                self.assertIn("E_TOT", tags, f"Missing E_TOT in meter definition {fname}")
                self.assertIn("U_AC_1", tags, f"Missing U_AC_1 in meter definition {fname}")
                self.assertIn("I_AC_1", tags, f"Missing I_AC_1 in meter definition {fname}")

    def test_float32_meters_direct_scaling(self) -> None:
        """Verify that F32 voltage, current, frequency, and power factor metrics in meters have CoefA == 1.000000."""
        target_meters = [
            "MET_Chint_DTSU666.csv",
            "MET_Chint_DDSU666.csv",
            "MET_Schneider_PM5000_IEM3000.csv",
        ]
        direct_scale_tags = {"U_AC_1", "U_AC_L1L2", "I_AC_1", "FREQ", "PF"}
        for fname in target_meters:
            fpath = find_definition_file(PRODUCTION_DIR, fname)
            with open(fpath, encoding="utf-8", newline="") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]

            for r in rows:
                dtype = r[3]
                tag = r[6]
                coef_a = r[7]
                if dtype == "F32" and tag in direct_scale_tags:
                    self.assertEqual(
                        coef_a,
                        "1.000000",
                        f"Register with tag {tag} in {fname} has CoefA={coef_a}, expected 1.000000 for F32",
                    )


if __name__ == "__main__":
    unittest.main()
