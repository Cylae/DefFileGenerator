"""
Apollo-Standard Adversarial Test Suite for Sungrow Master Definitions.

Tests full compliance and domain fidelity for:
  - PowerStack ST255CS Liquid-Cooled BESS & LC3000
  - FullMap v2 & standard battery
  - SG Series Generic & Three-phase
  - Logger 1000/3000, EMS300CP, LC300
  - PV Combiner Box
"""

from __future__ import annotations

import csv
from pathlib import Path

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"

SUNGROW_MODELS = [
    ("BAT", "PowerStack_ST255CS"),
    ("BAT", "LC3000"),
    ("BAT", "FullMap_v2"),
    ("BAT", "Battery"),
    ("INV", "SG_Series_Generic"),
    ("INV", "SG_Series"),
    ("INV", "Generic"),
    ("COM", "Logger1000_3000"),
    ("COM", "EMS300CP"),
    ("COM", "LC300_Modbus"),
    ("OTH", "PVCombinerBox"),
]


class TestSungrowDefinitionsSuite:
    """Rigorous verification suite for Sungrow definition files."""

    def test_all_sungrow_definitions_exist(self) -> None:
        """Verify that all target Sungrow definition files exist in production/."""
        for pfx, model in SUNGROW_MODELS:
            file_path = PROD_DIR / f"{pfx}_Sungrow_{model}.csv"
            assert file_path.is_file(), f"Missing definition file for {model} at {file_path}"
            assert file_path.stat().st_size > 500, f"File suspiciously small: {file_path.name}"

    def test_all_sungrow_5pillars_strict_certification(self) -> None:
        """Verify that all Sungrow definitions satisfy the 5 Webdyn Pillars with strict CRLF."""
        for pfx, model in SUNGROW_MODELS:
            file_path = PROD_DIR / f"{pfx}_Sungrow_{model}.csv"
            valid, issues = verify_file_pillars(file_path, strict_crlf=True)
            assert valid, f"{file_path.name} failed 5-pillars certification: {issues}"

    def test_powerstack_and_lc3000_registers(self) -> None:
        """Verify high-capacity ESS batteries have authentic registers."""
        for model in ["PowerStack_ST255CS", "LC3000"]:
            file_path = PROD_DIR / f"BAT_Sungrow_{model}.csv"
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[2] == "Sungrow"
                rows = list(reader)
                assert len(rows) > 900

    def test_multi_directory_synchronization(self) -> None:
        """Verify all Sungrow definitions are synchronized across target directories."""
        for pfx, model in SUNGROW_MODELS:
            filename = f"{pfx}_Sungrow_{model}.csv"
            prod_bytes = (PROD_DIR / filename).read_bytes()
            assert (REPO_ROOT / "Definitions" / "Sungrow" / filename).read_bytes() == prod_bytes
            assert (REPO_ROOT / "Definitions" / filename).read_bytes() == prod_bytes
            assert (REPO_ROOT / "generated_definitions" / filename).read_bytes() == prod_bytes
            assert (REPO_ROOT / "tagged_definitions" / filename).read_bytes() == prod_bytes
