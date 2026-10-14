"""
Unit and integration tests for the organized Definitions library.

Verifies:
1. All definition files adhere to the canonical naming scheme (<PREFIX>_<Manufacturer>_<Model>.csv).
2. Filename prefixes match the equipment types in the CSV headers.
3. Every CSV file strictly satisfies the Webdyn definition schema and passes validation.
4. Definitions catalog (Definitions/README.md) is present and accurately documents the library.
"""

import csv
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import Generator

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
DEFINITIONS_DIR = REPO_ROOT / "Definitions"

EXPECTED_PREFIX_MAP = {
    "Inverter": "INV",
    "Meter": "MET",
    "Sensor": "SEN",
    "Logger": "COM",
    "Io": "OTH",
    "Display": "OTH",
}

VALID_PREFIXES = {"INV", "MET", "SEN", "COM", "OTH"}


class TestDefinitionsLibrary(unittest.TestCase):
    """Validation test suite for the Definitions/ library."""

    generator: Generator
    csv_files: list[Path]

    @classmethod
    def setUpClass(cls) -> None:
        cls.generator = Generator()
        cls.csv_files = sorted(DEFINITIONS_DIR.glob("*.csv"))

    def test_definitions_directory_populated(self) -> None:
        """Verify Definitions directory exists and contains expected files."""
        self.assertTrue(DEFINITIONS_DIR.is_dir(), "Definitions directory must exist")
        self.assertGreaterEqual(
            len(self.csv_files),
            120,
            f"Expected at least 120 definitions, found {len(self.csv_files)}",
        )

    def test_naming_convention_and_prefixes(self) -> None:
        """Verify each file has a valid 3-letter prefix and matches its header equipment type."""
        for csv_path in self.csv_files:
            parts = csv_path.stem.split("_")
            prefix = parts[0]

            self.assertIn(
                prefix,
                VALID_PREFIXES,
                f"File {csv_path.name} has invalid prefix '{prefix}'. Expected one of {VALID_PREFIXES}",
            )
            self.assertGreaterEqual(
                len(parts),
                3,
                f"File {csv_path.name} must follow <PREFIX>_<Manufacturer>_<Model>.csv pattern",
            )

            # Check header
            with open(csv_path, encoding="utf-8", errors="ignore") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader, [])

            self.assertGreaterEqual(
                len(header),
                2,
                f"File {csv_path.name} header must contain at least [Protocol, EquipmentType]",
            )
            eq_type = header[1]
            expected_prefix = EXPECTED_PREFIX_MAP.get(eq_type)
            self.assertIsNotNone(
                expected_prefix,
                f"Unrecognized equipment type '{eq_type}' in {csv_path.name}",
            )
            self.assertEqual(
                prefix,
                expected_prefix,
                f"Prefix mismatch for {csv_path.name}: header is '{eq_type}' (expects '{expected_prefix}'), got '{prefix}'",
            )

    def test_all_definitions_validate_successfully(self) -> None:
        """Verify every definition file in Definitions/ passes Webdyn schema validation."""
        failures = []
        for csv_path in self.csv_files:
            report = self.generator.validate_csv_detailed(
                str(csv_path),
                strict=True,
                strict_overlap=False,
            )
            if not report.is_valid:
                failures.append(f"{csv_path.name}: {[str(i) for i in report.issues]}")

        self.assertEqual(
            failures,
            [],
            f"{len(failures)} definition files failed validation:\n" + "\n".join(failures[:10]),
        )

    def test_catalog_readme_exists_and_matches_count(self) -> None:
        """Verify Definitions/README.md exists and documents the full set of files."""
        readme_path = DEFINITIONS_DIR / "README.md"
        self.assertTrue(readme_path.is_file(), "Definitions/README.md must exist")
        content = readme_path.read_text(encoding="utf-8")

        self.assertIn(f"{len(self.csv_files)} fichiers de définition", content)
        for pfx in VALID_PREFIXES:
            self.assertIn(f"`{pfx}_`", content)


if __name__ == "__main__":
    unittest.main()
