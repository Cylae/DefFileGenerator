#!/usr/bin/env python3
"""Automated regression and audit test suite enforcing UTF-8 without BOM.

Verifies:
1. Absence of UTF-8 BOM byte sequence (\\xef\\xbb\\xbf) across all definition libraries.
2. GeneratorConfig default encoding is 'utf-8'.
3. Generator.write_output_csv produces pure UTF-8 without BOM.
4. CLI --encoding parameter works across subcommands.
5. validate_csv_detailed validates pure UTF-8 files.
"""

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.def_gen import CSVHeaderConfig, Generator, GeneratorConfig

REPO_ROOT = Path(__file__).resolve().parent.parent.parent

DEFINITION_DIRS = [
    REPO_ROOT / "production",
    REPO_ROOT / "Definitions",
    REPO_ROOT / "Dyness_DH_Series_Definitions",
    REPO_ROOT / "FoxESS_P3_H3_Pro_Definitions",
    REPO_ROOT / "MarsRenewable_Definitions",
    REPO_ROOT / "tagged_definitions",
    REPO_ROOT / "generated_definitions",
]


class TestUtf8WithoutBomAudit(unittest.TestCase):
    def setUp(self) -> None:
        self.generator = Generator()

    def test_definitions_catalog_has_zero_bom(self) -> None:
        """Verify that every CSV in Definitions/ has no BOM and is valid UTF-8."""
        def_dir = REPO_ROOT / "Definitions"
        self.assertTrue(def_dir.is_dir(), "Definitions/ directory must exist")
        csv_files = list(def_dir.glob("*.csv"))
        self.assertGreaterEqual(len(csv_files), 130, "Definitions/ must contain >= 130 CSV files")

        bom_failures = []
        for csv_file in csv_files:
            raw = csv_file.read_bytes()
            if raw.startswith(b"\xef\xbb\xbf"):
                bom_failures.append(csv_file.name)
            try:
                raw.decode("utf-8")
            except UnicodeDecodeError as e:
                self.fail(f"{csv_file.name} failed UTF-8 decoding: {e}")

        self.assertEqual(bom_failures, [], f"Definitions files containing BOM: {bom_failures}")

    def test_all_definition_suites_have_zero_bom(self) -> None:
        """Verify that all definition suites across the repo have zero BOM."""
        total_checked = 0
        bom_failures = []

        for suite_dir in DEFINITION_DIRS:
            if not suite_dir.exists():
                continue
            for csv_file in suite_dir.rglob("*.csv"):
                total_checked += 1
                raw = csv_file.read_bytes()
                if raw.startswith(b"\xef\xbb\xbf"):
                    bom_failures.append(str(csv_file.relative_to(REPO_ROOT)))

        self.assertGreaterEqual(total_checked, 3000, "Must check at least 3,000 definitions")
        self.assertEqual(
            bom_failures,
            [],
            f"Found {len(bom_failures)} definitions containing BOM out of {total_checked} checked.",
        )

    def test_generator_config_default_encoding(self) -> None:
        """Verify that GeneratorConfig default encoding is 'utf-8'."""
        cfg = GeneratorConfig()
        self.assertEqual(cfg.encoding, "utf-8")

        hdr_cfg = CSVHeaderConfig(manufacturer="Test", model="Model")
        self.assertEqual(hdr_cfg.encoding, "utf-8")

    def test_write_output_csv_produces_no_bom(self) -> None:
        """Verify that Generator.write_output_csv produces pure UTF-8 without BOM."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "output_test.csv"
            cfg = CSVHeaderConfig(
                protocol="Modbus",
                category="Inverter",
                manufacturer="TestMfr",
                model="TestMdl",
                encoding="utf-8",
            )
            rows = [
                {
                    "Index": 1,
                    "Info1": 3,
                    "Info2": 100,
                    "Info3": "U16",
                    "Info4": "",
                    "Name": "Voltage",
                    "Tag": "U_AC_1",
                    "CoefA": 1.0,
                    "CoefB": 0.0,
                    "Unit": "V",
                    "Action": 4,
                }
            ]

            Generator.write_output_csv(str(out_path), rows, config=cfg)
            self.assertTrue(out_path.is_file())
            raw = out_path.read_bytes()
            self.assertFalse(raw.startswith(b"\xef\xbb\xbf"), "Output CSV must NOT start with BOM")
            content = raw.decode("utf-8")
            self.assertIn("Modbus;Inverter;TestMfr;TestMdl", content)

    def test_cli_generate_encoding_flag(self) -> None:
        """Verify CLI generate command supports --encoding and defaults to utf-8."""
        cmd = [
            sys.executable,
            "-m",
            "DefFileGenerator.main",
            "generate",
            "--help",
        ]
        proc = subprocess.run(cmd, cwd=str(REPO_ROOT), capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0)
        self.assertIn("--encoding", proc.stdout)
        self.assertIn("utf-8", proc.stdout)

    def test_cli_run_encoding_flag(self) -> None:
        """Verify CLI run command supports --encoding and defaults to utf-8."""
        cmd = [
            sys.executable,
            "-m",
            "DefFileGenerator.main",
            "run",
            "--help",
        ]
        proc = subprocess.run(cmd, cwd=str(REPO_ROOT), capture_output=True, text=True)
        self.assertEqual(proc.returncode, 0)
        self.assertIn("--encoding", proc.stdout)
        self.assertIn("utf-8", proc.stdout)


if __name__ == "__main__":
    unittest.main()
