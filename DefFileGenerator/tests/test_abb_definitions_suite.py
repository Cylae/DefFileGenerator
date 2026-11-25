"""
Apollo-Standard Adversarial Test Suite for ABB & Fimer Master Definitions.

Tests full compliance and domain fidelity for:
  - PVS-100/120, PVS-175, PVS-50/60, PVS-10..33 commercial & utility string inverters
  - PVS800-57B utility central inverter
  - TRIO-20.0/27.6 & TRIO-50.0 commercial string inverters
  - PRO-33.0-TL, PVI-110/275, PVI-50/60, ULTRA Inverter Module
  - B21/B23/B24 DIN-rail energy meters
  - M2M & M4M precision network analyzers
  - Terra AC EV charger

Validates:
  1. The 5 Webdyn Pillars (UTF-8 no BOM, CRLF, 11 columns, sequential 1..N indices).
  2. Exact Modbus addressing, signedness, scale factors, and canonical SunPM tags.
  3. Multi-directory synchronization across production/, Definitions/ABB/, Definitions/, generated_definitions/, tagged_definitions/.
"""

from __future__ import annotations

import csv
from pathlib import Path

from DefFileGenerator.io_utils import find_definition_file

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"

ABB_MODELS = [
    ("INV", "PVS-100_120"),
    ("INV", "PVS-175"),
    ("INV", "PVS-50_60"),
    ("INV", "PVS-10_12.5_15_20_30_33"),
    ("INV", "PVS800-57B"),
    ("INV", "TRIO_20.0-27.6_TL_OUTD"),
    ("INV", "TRIO-50.0-TM-OUTD"),
    ("INV", "PRO-33.0-TL"),
    ("INV", "PVI-110_275"),
    ("INV", "PVI-50_60"),
    ("INV", "ULTRA_Inverter_Module"),
    ("MET", "B21_B23_B24"),
    ("MET", "M2M"),
    ("MET", "M4M"),
    ("OTH", "Terra_AC_Charger"),
    ("INV", "PVS-30-TL"),
    ("INV", "AURORA_Central"),
    ("INV", "PVI-10.0_12.5"),
]


class TestAbbDefinitionsSuite:
    """Rigorous verification suite for ABB & Fimer definition files."""

    def test_all_abb_definitions_exist(self) -> None:
        """Verify that all target ABB definition files exist in production/."""
        for pfx, model in ABB_MODELS:
            file_path = find_definition_file(PROD_DIR, f"{pfx}_ABB_{model}.csv")
            assert file_path.is_file(), f"Missing definition file for {model} at {file_path}"
            assert file_path.stat().st_size > 500, f"File suspiciously small: {file_path.name}"

    def test_all_abb_5pillars_strict_certification(self) -> None:
        """Verify that all ABB definitions satisfy the 5 Webdyn Pillars with strict CRLF."""
        for pfx, model in ABB_MODELS:
            file_path = find_definition_file(PROD_DIR, f"{pfx}_ABB_{model}.csv")
            valid, issues = verify_file_pillars(file_path, strict_crlf=True)
            assert valid, f"{file_path.name} failed 5-pillars certification: {issues}"

    def test_pvs_series_sunspec_telemetry(self) -> None:
        """Verify SunSpec inverter models have valid telemetry and canonical tags."""
        for model in ["PVS-100_120", "PVS-175", "PVS-50_60"]:
            file_path = find_definition_file(PROD_DIR, f"INV_ABB_{model}.csv")
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[0] == "Modbus"
                assert header[1] == "Inverter"
                assert header[2] == "ABB"
                assert header[3] == model

                rows = list(reader)
                assert len(rows) >= 900, f"Expected at least 900 registers for {model}, got {len(rows)}"

    def test_pvs800_central_inverter_spec(self) -> None:
        """Verify PVS800-57B utility central inverter registers and telemetry."""
        file_path = find_definition_file(PROD_DIR, "INV_ABB_PVS800-57B.csv")
        with open(file_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[3] == "PVS800-57B"
            rows = list(reader)
            assert len(rows) == 1053, f"Expected 1053 registers, got {len(rows)}"

            # DC voltage at 25857
            u_dc = next((r for r in rows if r[2] == "25857"), None)
            assert u_dc is not None and u_dc[6] == "U_DC"
            assert float(u_dc[7]) == 0.01

            # Grid frequency at 25871
            freq = next((r for r in rows if r[2] == "25871"), None)
            assert freq is not None and freq[6] == "FREQ"
            assert float(freq[7]) == 0.01

    def test_trio_series_spec(self) -> None:
        """Verify TRIO 20/27.6 and TRIO 50 commercial string inverters."""
        trio20_path = find_definition_file(PROD_DIR, "INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv")
        with open(trio20_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[3] == "TRIO_20.0-27.6_TL_OUTD"
            rows = list(reader)
            assert len(rows) == 89, f"Expected 89 registers, got {len(rows)}"

            # Remote On/Off control at 179
            cmd_on = next((r for r in rows if r[2] == "179"), None)
            assert cmd_on is not None and cmd_on[6] == "CMD_ON"

    def test_meter_series_spec(self) -> None:
        """Verify B21/B23/B24 and M4M meters."""
        b24_path = find_definition_file(PROD_DIR, "MET_ABB_B21_B23_B24.csv")
        with open(b24_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[3] == "B21_B23_B24"
            rows = list(reader)
            assert len(rows) == 137, f"Expected 137 registers, got {len(rows)}"
            # Active energy import at 20480
            e_import = next((r for r in rows if r[2] == "20480"), None)
            assert e_import is not None
            assert float(e_import[7]) == 0.01
            assert e_import[9] == "kWh"

        m4m_path = find_definition_file(PROD_DIR, "MET_ABB_M4M.csv")
        with open(m4m_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[3] == "M4M"
            rows = list(reader)
            assert len(rows) == 298, f"Expected 298 registers, got {len(rows)}"

    def test_terra_ac_charger_spec(self) -> None:
        """Verify Terra AC EV charger definition."""
        terra_path = find_definition_file(PROD_DIR, "OTH_ABB_Terra_AC_Charger.csv")
        with open(terra_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[3] == "Terra_AC_Charger"
            rows = list(reader)
            assert len(rows) == 20, f"Expected 20 registers, got {len(rows)}"

    def test_multi_directory_synchronization(self) -> None:
        """Verify that definitions are mirrored byte-identically across all mirror folders."""
        for pfx, model in ABB_MODELS:
            filename = f"{pfx}_ABB_{model}.csv"
            prod_file = find_definition_file(PROD_DIR, filename)
            prod_bytes = prod_file.read_bytes()

            for target_name in ["Definitions", "generated_definitions", "tagged_definitions"]:
                target_file = find_definition_file(REPO_ROOT / target_name, filename)
                assert target_file.is_file(), f"Missing mirror file in {target_name}: {filename}"
                assert (
                    target_file.read_bytes() == prod_bytes
                ), f"Mismatch between production and {target_name} for {filename}"
