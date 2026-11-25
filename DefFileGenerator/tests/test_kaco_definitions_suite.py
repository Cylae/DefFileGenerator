"""
Apollo-Standard Adversarial Test Suite for KACO new energy Master Definitions.

Tests full compliance and domain fidelity for:
  - blueplanet 125 NX3 M10, 100 NX3 M8, 125 TL3, 50.0 TL3, 50.0 NX3, 20.0 TL3, 15.0 TL3
  - blueplanet NX1 & NX3 series
  - Powador 60.0 TL3, 360TL, 600TL
"""

from __future__ import annotations

import csv
from pathlib import Path

from DefFileGenerator.io_utils import find_definition_file

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"

KACO_MODELS = [
    ("INV", "blueplanet_125_NX3_M10"),
    ("INV", "blueplanet_100_NX3_M8"),
    ("INV", "blueplanet_125_TL3"),
    ("INV", "blueplanet_50.0_TL3"),
    ("INV", "blueplanet_50.0_NX3_M5"),
    ("INV", "blueplanet_20.0_TL3"),
    ("INV", "blueplanet_15.0_TL3"),
    ("INV", "blueplanet_NX1"),
    ("INV", "blueplanet_NX3"),
    ("INV", "Powador_60.0_TL3"),
]


class TestKacoDefinitionsSuite:
    """Rigorous verification suite for KACO new energy definition files."""

    def test_all_kaco_definitions_exist(self) -> None:
        """Verify that all target Kaco definition files exist in production/."""
        for pfx, model in KACO_MODELS:
            file_path = find_definition_file(PROD_DIR, f"{pfx}_Kaco_{model}.csv")
            assert file_path.is_file(), f"Missing definition file for {model} at {file_path}"
            assert file_path.stat().st_size > 500, f"File suspiciously small: {file_path.name}"

    def test_all_kaco_5pillars_strict_certification(self) -> None:
        """Verify that all Kaco definitions satisfy the 5 Webdyn Pillars with strict CRLF."""
        for pfx, model in KACO_MODELS:
            file_path = find_definition_file(PROD_DIR, f"{pfx}_Kaco_{model}.csv")
            valid, issues = verify_file_pillars(file_path, strict_crlf=True)
            assert valid, f"{file_path.name} failed 5-pillars certification: {issues}"

    def test_blueplanet_nx3_registers_and_tags(self) -> None:
        """Verify high-power blueplanet NX3 models have SunSpec tags."""
        for model in ["blueplanet_125_NX3_M10", "blueplanet_100_NX3_M8"]:
            file_path = find_definition_file(PROD_DIR, f"INV_Kaco_{model}.csv")
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[2] == "Kaco"
                rows = list(reader)
                assert len(rows) > 900
                tags = [r[6] for r in rows if r[6]]
                assert "MODEL" in tags or "SERIAL" in tags

    def test_powador_series_modbus(self) -> None:
        """Verify Powador series inverters have valid addresses and action codes."""
        for model in ["Powador_60.0_TL3"]:
            file_path = find_definition_file(PROD_DIR, f"INV_Kaco_{model}.csv")
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[2] == "Kaco"
                rows = list(reader)
                assert len(rows) >= 40
                actions = {r[10] for r in rows}
                assert actions.issubset({"0", "1", "2", "4", "9"})

    def test_multi_directory_synchronization(self) -> None:
        """Verify all Kaco definitions are synchronized across target directories."""
        for pfx, model in KACO_MODELS:
            filename = f"{pfx}_Kaco_{model}.csv"
            prod_bytes = find_definition_file(PROD_DIR, filename).read_bytes()
            assert find_definition_file(REPO_ROOT / "Definitions", filename).read_bytes() == prod_bytes
            assert find_definition_file(REPO_ROOT / "Definitions", filename).read_bytes() == prod_bytes
            assert find_definition_file(REPO_ROOT / "generated_definitions", filename).read_bytes() == prod_bytes
            assert find_definition_file(REPO_ROOT / "tagged_definitions", filename).read_bytes() == prod_bytes
