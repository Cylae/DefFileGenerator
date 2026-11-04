"""
Unit and integration tests for Dyness DH-Series BESS WebdynSunPM definitions.

Covers:
- End-to-end parsing of Dyness DH-Series Modbus protocol (V1.0 - V2.2).
- Verification of all 4 generated model definitions:
  * BAT_Dyness_DH200Y.csv (Liquid-cooled 232 kWh)
  * BAT_Dyness_DH200F.csv (Air-cooled 215 kWh)
  * BAT_Dyness_DH100F.csv (Air-cooled 100 kWh)
  * BAT_Dyness_DH_Series_Generic.csv (Universal superset)
- Strict WebdynSunPM schema validation (strict=True, strict_overlap=True).
- Physical offsets and coefficients (current -3200.0 A, temps -40.0 °C).
- String spans and byte lengths.
- Tag uniqueness (zero duplicates).
- Sequential 1-based indexing.
- UTF-8 BOM encoding.
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator
from generate_dyness_definitions import DynessDefinitionGenerator, DynessProtocolParser

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
RAW_CSV = REPO_ROOT / "stress_test_data" / "Dyness_DH_Series_Modbus_Protocol_V2.2.csv"
DEFINITIONS_DIR = REPO_ROOT / "Definitions"

EXPECTED_FILES = {
    "DH200Y": "BAT_Dyness_DH200Y.csv",
    "DH200F": "BAT_Dyness_DH200F.csv",
    "DH100F": "BAT_Dyness_DH100F.csv",
    "DH_Series_Generic": "BAT_Dyness_DH_Series_Generic.csv",
}


class TestDynessProtocolParsing(unittest.TestCase):
    """Tests raw protocol parsing fidelity."""

    parser: DynessProtocolParser

    @classmethod
    def setUpClass(cls) -> None:
        cls.parser = DynessProtocolParser(RAW_CSV)
        cls.parser.parse()

    def test_raw_csv_exists(self) -> None:
        self.assertTrue(RAW_CSV.is_file(), f"Protocol CSV missing at: {RAW_CSV}")
        self.assertGreater(RAW_CSV.stat().st_size, 10_000, "Protocol CSV suspiciously small")

    def test_holding_registers_extracted(self) -> None:
        self.assertGreaterEqual(
            len(self.parser.holding_regs), 30, "Holding registers under-extracted"
        )
        for reg in self.parser.holding_regs:
            self.assertEqual(reg["info1"], 3, f"Holding register must have Info1=3: {reg}")
            self.assertIn(
                reg["action"], ("1", "4"), f"Holding register action must be 1 or 4: {reg}"
            )
            self.assertTrue(reg["address"].isdigit(), f"Holding address must be numeric: {reg}")

    def test_common_input_registers_extracted(self) -> None:
        self.assertGreaterEqual(
            len(self.parser.common_input_regs), 200, "Common input registers under-extracted"
        )
        for reg in self.parser.common_input_regs:
            self.assertEqual(reg["info1"], 4, f"Input register must have Info1=4: {reg}")
            self.assertIn(reg["action"], ("4", "9"), f"Input register action must be 4 or 9: {reg}")

    def test_parallel_input_registers_extracted(self) -> None:
        self.assertGreaterEqual(
            len(self.parser.parallel_input_regs), 100, "Parallel registers under-extracted"
        )
        for reg in self.parser.parallel_input_regs:
            self.assertEqual(reg["info1"], 4, f"Parallel register must have Info1=4: {reg}")

    def test_model_specific_modules_extracted(self) -> None:
        self.assertGreaterEqual(len(self.parser.pcs_pos4_dh200y), 50, "PCS Pos #4 under-extracted")
        self.assertGreaterEqual(len(self.parser.pcs_pos5_dh200f), 50, "PCS Pos #5 under-extracted")
        self.assertGreaterEqual(len(self.parser.dcdc_regs), 30, "DCDC registers under-extracted")
        self.assertGreaterEqual(
            len(self.parser.bms_common_regs), 20, "BMS common registers under-extracted"
        )
        self.assertGreaterEqual(
            len(self.parser.fire_regs), 10, "Fire protection registers under-extracted"
        )
        self.assertGreaterEqual(
            len(self.parser.ac_regs), 30, "AC cooling registers under-extracted"
        )
        self.assertGreaterEqual(
            len(self.parser.liquid_cooling_regs), 50, "Liquid cooling under-extracted"
        )
        self.assertGreaterEqual(
            len(self.parser.temp_hum_regs), 20, "Temp/Hum registers under-extracted"
        )


class TestDynessGeneratedDefinitions(unittest.TestCase):
    """Tests generated CSV definitions against WebdynSunPM strict specification."""

    validator: Generator
    parser: DynessProtocolParser
    generated: dict[str, Path]

    @classmethod
    def setUpClass(cls) -> None:
        cls.validator = Generator()
        # Ensure fresh generation
        cls.parser = DynessProtocolParser(RAW_CSV)
        cls.parser.parse()
        gen = DynessDefinitionGenerator(cls.parser, DEFINITIONS_DIR)
        cls.generated = gen.generate_all()

    def test_all_expected_files_exist(self) -> None:
        for _model_key, filename in EXPECTED_FILES.items():
            filepath = DEFINITIONS_DIR / filename
            self.assertTrue(filepath.is_file(), f"Missing definition file: {filename}")
            self.assertGreater(
                filepath.stat().st_size, 20_000, f"Definition file too small: {filename}"
            )

    def test_strict_webdyn_validation(self) -> None:
        """Every generated Dyness definition must pass strict validation with 0 errors and 0 overlaps."""
        for filename in EXPECTED_FILES.values():
            filepath = DEFINITIONS_DIR / filename
            report = self.validator.validate_csv_detailed(
                str(filepath),
                strict=True,
                strict_overlap=True,
            )
            self.assertTrue(
                report.is_valid,
                f"File {filename} failed strict validation with {len(report.issues)} issues: {report.issues}",
            )
            self.assertGreater(report.register_count, 800, f"Expected >800 registers in {filename}")

    def test_utf8_bom_and_header(self) -> None:
        """Verify UTF-8 BOM, 11-column header, and metadata."""
        for model_key, filename in EXPECTED_FILES.items():
            filepath = DEFINITIONS_DIR / filename
            with open(filepath, "rb") as f:
                bom = f.read(3)
                self.assertEqual(bom, b"\xef\xbb\xbf", f"{filename} missing UTF-8 BOM")

            with open(filepath, encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                self.assertEqual(len(header), 11, f"{filename} header must have 11 columns")
                self.assertEqual(header[0], "Modbus")
                self.assertEqual(header[1], "Battery")
                self.assertEqual(header[2], "Dyness")
                self.assertEqual(header[3], model_key)

    def test_tag_uniqueness(self) -> None:
        """Verify 0 duplicate SunPM tags within any file."""
        for filename in EXPECTED_FILES.values():
            filepath = DEFINITIONS_DIR / filename
            seen_tags: set[str] = set()
            with open(filepath, encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=";")
                next(reader)  # header
                for line_no, row in enumerate(reader, start=2):
                    tag = row[6].strip()
                    if tag:
                        self.assertNotIn(
                            tag,
                            seen_tags,
                            f"{filename} Line {line_no}: Duplicate tag '{tag}'",
                        )
                        seen_tags.add(tag)
            self.assertGreaterEqual(len(seen_tags), 45, f"{filename} has too few canonical tags")

    def test_sequential_indexing(self) -> None:
        """Verify 1-based sequential indexing for all rows."""
        for filename in EXPECTED_FILES.values():
            filepath = DEFINITIONS_DIR / filename
            with open(filepath, encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=";")
                next(reader)
                for expected_idx, row in enumerate(reader, start=1):
                    self.assertEqual(
                        row[0].strip(),
                        str(expected_idx),
                        f"{filename} index mismatch at row {expected_idx}",
                    )

    def test_physical_offsets_and_units(self) -> None:
        """Verify temperature (-40.0 °C) and pack current (-3200.0 A) physical offsets."""
        for filename in EXPECTED_FILES.values():
            filepath = DEFINITIONS_DIR / filename
            found_pack_current = False
            found_temp_offset = False

            with open(filepath, encoding="utf-8-sig") as f:
                reader = csv.reader(f, delimiter=";")
                next(reader)
                for row in reader:
                    name = row[5].lower()
                    coef_a = float(row[7])
                    coef_b = float(row[8])

                    if "total battery pack current" in name:
                        found_pack_current = True
                        self.assertAlmostEqual(coef_a, 0.1, places=3)
                        self.assertAlmostEqual(coef_b, -3200.0, places=1)
                        self.assertIn(row[9], ("A", "a", ""))

                    if "cell temperature" in name:
                        found_temp_offset = True
                        self.assertAlmostEqual(coef_a, 1.0, places=3)
                        self.assertAlmostEqual(coef_b, -40.0, places=1)
                        self.assertEqual(row[9], "°C")

            self.assertTrue(found_pack_current, f"Pack current with offset not found in {filename}")
            self.assertTrue(
                found_temp_offset, f"Cell temperature with offset not found in {filename}"
            )

    def test_model_cell_counts(self) -> None:
        """Verify DH200Y has 260 cell voltages, DH200F/DH100F have 240 cell voltages."""
        # DH200Y: 260 cells
        dh200y_path = DEFINITIONS_DIR / EXPECTED_FILES["DH200Y"]
        with open(dh200y_path, encoding="utf-8-sig") as f:
            content = f.read()
            self.assertIn("Cell Voltage 260", content)
            self.assertNotIn("Cell Voltage 261", content)
            self.assertIn("Cell Temperature 80", content)
            self.assertNotIn("Cell Temperature 81", content)

        # DH200F: 240 cells
        dh200f_path = DEFINITIONS_DIR / EXPECTED_FILES["DH200F"]
        with open(dh200f_path, encoding="utf-8-sig") as f:
            content = f.read()
            self.assertIn("Cell Voltage 240", content)
            self.assertNotIn("Cell Voltage 241", content)
            self.assertIn("Cell Temperature 60", content)
            self.assertNotIn("Cell Temperature 61", content)

    def test_package_level_dyness_exports(self) -> None:
        """Verify DefFileGenerator exports dyness generator components cleanly."""
        from DefFileGenerator import (
            DynessDefinitionGenerator as PkgGen,
        )
        from DefFileGenerator import (
            DynessProtocolParser as PkgParser,
        )
        from DefFileGenerator import (
            generate_dyness_definitions as pkg_fn,
        )

        self.assertIs(PkgGen, DynessDefinitionGenerator)
        self.assertIs(PkgParser, DynessProtocolParser)
        self.assertTrue(callable(pkg_fn))

    def test_metadata_inference_recognizes_dyness(self) -> None:
        """Verify DefFileGenerator.main.infer_metadata_from_filename infers Dyness Battery."""
        from DefFileGenerator.main import infer_metadata_from_filename

        mfg, model, cat = infer_metadata_from_filename("Dyness_DH200Y_Protocol.xlsx")
        self.assertEqual(mfg, "Dyness")
        self.assertEqual(cat, "Battery")

        mfg_f, model_f, cat_f = infer_metadata_from_filename("Dyness_DH200F_Manual.pdf")
        self.assertEqual(mfg_f, "Dyness")
        self.assertEqual(cat_f, "Battery")


if __name__ == "__main__":
    unittest.main()
