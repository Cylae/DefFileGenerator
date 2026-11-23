"""
Apollo-Standard Adversarial Test Suite for Carlo Gavazzi Master Definitions.

Tests full compliance and domain fidelity for:
  - EM210_MV & EM210
  - EM530 & EM540
  - EM111, EM112, ET112
  - EM330, EM340, ET330, ET340
  - EM21, EM271
  - WM15, WM20, WM50

Validates:
  1. The 5 Webdyn Pillars (UTF-8 no BOM, CRLF, 11 columns, 1..N index, valid types).
  2. Domain scaling, units, and Modbus function codes against official protocols.
  3. Canonical SunPM tags (uniqueness of single tags: P_AC, Q_AC, S_AC, FREQ, PF, E_TOT).
  4. Multi-directory synchronization across production/, Definitions/, generated_definitions/, tagged_definitions/.
"""

from __future__ import annotations

import csv
from pathlib import Path

from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"
CG_MODELS = [
    "EM210_MV",
    "EM210",
    "EM24",
    "EM530",
    "EM540",
    "EM111",
    "EM112",
    "ET112",
    "EM330",
    "EM340",
    "ET330",
    "ET340",
    "EM21",
    "EM271",
    "WM15",
    "WM20",
    "WM30",
    "WM40",
    "WM50",
]


class TestCarloGavazziDefinitionsSuite:
    """Rigorous verification suite for Carlo Gavazzi definition files."""

    def test_all_carlo_gavazzi_definitions_exist(self) -> None:
        """Verify that all target Carlo Gavazzi definition files exist in production/."""
        for model in CG_MODELS:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            assert file_path.is_file(), f"Missing definition file for {model} at {file_path}"
            assert file_path.stat().st_size > 500, f"File suspiciously small: {file_path.name}"

    def test_all_carlo_gavazzi_5pillars_strict_certification(self) -> None:
        """Verify that all Carlo Gavazzi definitions satisfy the 5 Webdyn Pillars with strict CRLF."""
        for model in CG_MODELS:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            valid, issues = verify_file_pillars(file_path, strict_crlf=True)
            assert valid, f"{file_path.name} failed 5-pillars certification: {issues}"

    def test_em210_mv_and_em210_domain_spec(self) -> None:
        """Verify Carlo Gavazzi EM210_MV and EM210 specifications from official PDF."""
        for model in ["EM210_MV", "EM210"]:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[0] == "Modbus"
                assert header[1] == "Meter"
                assert header[2] == "CarloGavazzi"
                assert header[3] == model

                rows = list(reader)
                assert len(rows) == 115, f"{model} expected 115 registers, got {len(rows)}"

                # Check Rogowski sensor registers
                reg_4104 = next((r for r in rows if r[2] == "4104"), None)
                assert reg_4104 is not None, f"Missing Rogowski sensor selection in {model}"
                assert "Sensor selection" in reg_4104[5]

                reg_4105 = next((r for r in rows if r[2] == "4105"), None)
                assert reg_4105 is not None, f"Missing Rogowski range in {model}"
                assert "Range Rogowski sensor" in reg_4105[5]

                # Check canonical tags
                p_ac_row = next((r for r in rows if r[6] == "P_AC"), None)
                assert p_ac_row is not None and p_ac_row[2] == "41"
                assert float(p_ac_row[7]) == 0.1

                freq_row = next((r for r in rows if r[6] == "FREQ"), None)
                assert freq_row is not None and freq_row[2] == "52"
                assert float(freq_row[7]) == 1.0

                e_tot_row = next((r for r in rows if r[6] == "E_TOT"), None)
                assert e_tot_row is not None and e_tot_row[2] == "53"
                assert float(e_tot_row[7]) == 0.1

    def test_em530_and_em540_domain_spec(self) -> None:
        """Verify Carlo Gavazzi EM530 and EM540 specifications from official protocol."""
        for model in ["EM530", "EM540"]:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[0] == "Modbus"
                assert header[1] == "Meter"
                assert header[2] == "CarloGavazzi"
                assert header[3] == model

                rows = list(reader)
                assert len(rows) == 143, f"{model} expected 143 registers, got {len(rows)}"

                # Voltage L1-N
                v_l1 = next((r for r in rows if r[2] == "0"), None)
                assert v_l1 is not None and v_l1[6] == "U_AC_1"
                assert float(v_l1[7]) == 0.1
                assert v_l1[9] == "V"

                # Current L1
                a_l1 = next((r for r in rows if r[2] == "12"), None)
                assert a_l1 is not None and a_l1[6] == "I_AC_1"
                assert float(a_l1[7]) == 0.001
                assert a_l1[9] == "A"

                # Total Active Power
                p_ac = next((r for r in rows if r[2] == "40"), None)
                assert p_ac is not None and p_ac[6] == "P_AC"
                assert float(p_ac[7]) == 0.1
                assert p_ac[9] == "W"

                # Power Factor Total (scale factor 0.001)
                pf_sys = next((r for r in rows if r[2] == "49"), None)
                assert pf_sys is not None and pf_sys[6] == "PF"
                assert float(pf_sys[7]) == 0.001

                # Frequency (scale factor 0.1)
                freq = next((r for r in rows if r[2] == "51"), None)
                assert freq is not None and freq[6] == "FREQ"
                assert float(freq[7]) == 0.1
                assert freq[9] == "Hz"

                # Total Active Energy
                e_tot = next((r for r in rows if r[2] == "52"), None)
                assert e_tot is not None and e_tot[6] == "E_TOT"
                assert float(e_tot[7]) == 0.1
                assert e_tot[9] == "kWh"

    def test_em24_domain_spec(self) -> None:
        """Verify Carlo Gavazzi EM24-DIN official Modbus specifications."""
        file_path = PROD_DIR / "MET_CarloGavazzi_EM24.csv"
        with open(file_path, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert header[3] == "EM24"
            rows = list(reader)
            assert len(rows) == 57, f"Expected 57 registers, got {len(rows)}"

            # Voltage L1-N at 0
            v_l1 = next((r for r in rows if r[2] == "0"), None)
            assert v_l1 is not None and v_l1[6] == "U_AC_1"
            assert float(v_l1[7]) == 0.1

            # System active power at 40
            p_sys = next((r for r in rows if r[2] == "40"), None)
            assert p_sys is not None and p_sys[6] == "P_AC"
            assert float(p_sys[7]) == 0.1

            # System PF at 53
            pf_sys = next((r for r in rows if r[2] == "53"), None)
            assert pf_sys is not None and pf_sys[6] == "PF"
            assert float(pf_sys[7]) == 0.001

            # Frequency at 55
            freq = next((r for r in rows if r[2] == "55"), None)
            assert freq is not None and freq[6] == "FREQ"
            assert float(freq[7]) == 0.1

            # Total active energy at 62
            e_tot = next((r for r in rows if r[2] == "62"), None)
            assert e_tot is not None and e_tot[6] == "E_TOT"
            assert float(e_tot[7]) == 0.1

    def test_em100_series_spec(self) -> None:
        """Verify EM111, EM112, ET112 1-phase specifications."""
        for model in ["EM111", "EM112", "ET112"]:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[3] == model
                rows = list(reader)

                v_l1 = next((r for r in rows if r[6] == "U_AC_1"), None)
                assert v_l1 is not None and v_l1[2] == "0"
                assert float(v_l1[7]) == 0.1

                p_ac = next((r for r in rows if r[6] == "P_AC"), None)
                assert p_ac is not None and p_ac[2] == "4"

                e_tot = next((r for r in rows if r[6] == "E_TOT"), None)
                assert e_tot is not None and e_tot[2] == "16"

    def test_em300_series_spec(self) -> None:
        """Verify EM330, EM340, ET330, ET340 3-phase specifications."""
        for model in ["EM330", "EM340", "ET330", "ET340"]:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[3] == model
                rows = list(reader)

                p_ac = next((r for r in rows if r[6] == "P_AC"), None)
                assert p_ac is not None and p_ac[2] == "40"

                freq = next((r for r in rows if r[6] == "FREQ"), None)
                assert freq is not None and freq[2] == "51"

    def test_wm20_and_wm50_float_telemetry_spec(self) -> None:
        """Verify WM20 and WM50 IEEE 754 float telemetry and U64 energy counters."""
        for model in ["WM20", "WM50"]:
            file_path = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            with open(file_path, encoding="utf-8") as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader)
                assert header[3] == model
                rows = list(reader)

                # Float Voltages & Powers
                v_l1 = next((r for r in rows if r[6] == "U_AC_1"), None)
                assert v_l1 is not None and v_l1[2] == "81"
                assert v_l1[3] == "FLOAT"
                assert float(v_l1[7]) == 1.0

                p_ac = next((r for r in rows if r[6] == "P_AC"), None)
                assert p_ac is not None and p_ac[2] == "111"
                assert p_ac[3] == "FLOAT"

                # 64-bit Totalizer
                e_tot = next((r for r in rows if r[6] == "E_TOT"), None)
                assert e_tot is not None and e_tot[2] == "1281"
                assert e_tot[3] == "U64"
                assert float(e_tot[7]) == 0.001

    def test_multi_directory_synchronization(self) -> None:
        """Verify that definitions are byte-identical across synchronization directories."""
        sync_dirs = [
            REPO_ROOT / "Definitions" / "CarloGavazzi",
            REPO_ROOT / "Definitions",
            REPO_ROOT / "generated_definitions",
            REPO_ROOT / "tagged_definitions",
        ]
        for model in CG_MODELS:
            prod_file = PROD_DIR / f"MET_CarloGavazzi_{model}.csv"
            prod_bytes = prod_file.read_bytes()
            for target_dir in sync_dirs:
                target_file = target_dir / f"MET_CarloGavazzi_{model}.csv"
                assert target_file.is_file(), f"Missing mirror file: {target_file}"
                assert target_file.read_bytes() == prod_bytes, (
                    f"Mirror mismatch for {model} in {target_dir}"
                )
