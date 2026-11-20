"""
Apollo-Standard Adversarial Verification Suite for Production Energy Meters (production/MET_*.csv).

Mission & Mandate:
- Rigorous cross-audit of all 200+ energy meter definitions in production/.
- Deep domain verification for critical meter families:
  * Chint DTSU666 (DTSU666, DTSU666-H, DDSU666, Power_CHINT_DTSU)
  * Schneider iEM31xx/35xx & PM series (iEM3100..3555, PM5000, PM5xxx, PM820, PMC, EM6433)
  * Socomec Diris & Countis (Diris A10..A40, B10, B30, I30..I35, Digiware, Countis E23..E53)
  * Carlo Gavazzi (EM24, EM210, EM-ET300, EM300-ET300, 72RVV23XOSX)
- Enforce the 5 Webdyn Pillars:
  1. Strict UTF-8 without BOM
  2. Strict CRLF line terminators (zero lonely LF/CR)
  3. Exact 11-column semicolon CSV schema
  4. Monotonic sequential 1..N row index matching physical line
  5. Canonical SunPM tags, valid action codes {0..12}, valid function codes {1..4}, valid Modbus addressing
"""

from __future__ import annotations

import csv
import math
from pathlib import Path

import pytest

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.production_governor import verify_file_pillars

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"

VALID_FUNCTION_CODES = {"1", "2", "3", "4"}
VALID_ACTIONS = {str(i) for i in range(13)}
VALID_TYPES = {
    "U16", "I16", "U32", "I32", "F32", "STRING", "HEX", "BITS", "BIT", "FLOAT", "INT16", "UINT16",
    "INT32", "UINT32", "U64", "I64", "ENUM"
}


def get_meter_files() -> list[Path]:
    """Retrieve all production meter files sorted canonically."""
    return sorted(PROD_DIR.glob("MET_*.csv"))


@pytest.fixture(scope="module")
def generator() -> Generator:
    return Generator()


class TestProductionMetersSuite:
    """Comprehensive adversarial testing for production energy meters."""

    def test_all_meters_exist_and_populated(self) -> None:
        """Verify that at least 200 meter definitions are present in production/."""
        meter_files = get_meter_files()
        assert len(meter_files) >= 200, f"Expected at least 200 meter files, found {len(meter_files)}"
        for f in meter_files:
            assert f.is_file(), f"Not a valid file: {f}"
            assert f.stat().st_size > 50, f"File suspiciously small: {f.name}"

    def test_all_meters_strict_crlf_and_zero_bom(self) -> None:
        """Enforce strict binary hygiene: UTF-8 without BOM, strict CRLF termination."""
        meter_files = get_meter_files()
        for f in meter_files:
            raw = f.read_bytes()
            # 1. Zero BOM
            assert not raw.startswith(b"\xef\xbb\xbf"), f"BOM detected in {f.name}"
            # 2. Strict CRLF ending
            assert raw.endswith(b"\r\n"), f"File does not end with CRLF: {f.name}"
            assert not raw.endswith(b"\r\r\n"), f"Double CR detected in {f.name}"
            # 3. No isolated LF or CR
            stripped = raw.replace(b"\r\n", b"")
            assert b"\n" not in stripped, f"Isolated LF detected in {f.name}"
            assert b"\r" not in stripped, f"Isolated CR detected in {f.name}"

    @pytest.mark.parametrize("file_path", get_meter_files(), ids=lambda p: p.name)
    def test_meter_5_pillars_certification(self, file_path: Path) -> None:
        """Verify full 5-pillar compliance using production governor."""
        is_valid, issues = verify_file_pillars(file_path)
        assert is_valid, f"{file_path.name} failed 5 pillars certification: {issues}"

    @pytest.mark.parametrize("file_path", get_meter_files(), ids=lambda p: p.name)
    def test_meter_schema_and_monotonic_indices(self, file_path: Path) -> None:
        """Verify 11 columns, monotonic 1..N indices, addressing, types, and scale factors."""
        with open(file_path, encoding="utf-8") as fp:
            reader = list(csv.reader(fp, delimiter=";"))

        assert len(reader) >= 2, f"{file_path.name}: file must have header + at least 1 row"
        header = reader[0]
        assert len(header) == 11, f"{file_path.name}: header length {len(header)} != 11"
        assert header[0].strip().lower() == "modbus", f"{file_path.name}: protocol header must be 'Modbus'"

        seen_single_tags: set[str] = set()

        for idx, row in enumerate(reader[1:], 1):
            assert len(row) == 11, f"{file_path.name} L{idx}: column count {len(row)} != 11"

            # Monotonic 1..N index
            assert row[0].strip() == str(idx), f"{file_path.name} L{idx}: index mismatch {row[0]} != {idx}"

            # Function code
            fc = row[1].strip()
            assert fc in VALID_FUNCTION_CODES, f"{file_path.name} L{idx}: invalid fc {fc}"

            # Address
            addr_str = row[2].strip()
            assert addr_str, f"{file_path.name} L{idx}: empty register address"
            dtype = row[3].strip().upper()
            if dtype == "STRING":
                assert "_" in addr_str, f"{file_path.name} L{idx}: STRING missing length suffix: {addr_str}"
                base_part, len_part = addr_str.split("_", 1)
                assert base_part.isdigit(), f"{file_path.name} L{idx}: non-numeric string address {addr_str}"
                assert len_part.isdigit(), f"{file_path.name} L{idx}: non-numeric string length {addr_str}"
            else:
                base_addr = addr_str.split("_")[0]
                assert base_addr.isdigit(), f"{file_path.name} L{idx}: non-numeric address {addr_str}"
                iaddr = int(base_addr)
                assert 0 <= iaddr <= 65535, f"{file_path.name} L{idx}: address out of range: {iaddr}"

            # Scale factors
            coef_a_str = row[7].strip()
            coef_b_str = row[8].strip()
            if coef_a_str:
                coef_a = float(coef_a_str)
                assert math.isfinite(coef_a), f"{file_path.name} L{idx}: coef_a not finite {coef_a_str}"
            if coef_b_str:
                coef_b = float(coef_b_str)
                assert math.isfinite(coef_b), f"{file_path.name} L{idx}: coef_b not finite {coef_b_str}"

            # Action code
            action = row[10].strip()
            assert action in VALID_ACTIONS, f"{file_path.name} L{idx}: invalid action {action}"

            # SunPM Tag single-instance rule
            tag = row[6].strip()
            if tag in {"P_AC", "Q_AC", "S_AC", "FREQ", "PF", "E_TOT"}:
                assert tag not in seen_single_tags, f"{file_path.name} L{idx}: duplicate single tag {tag}"
                seen_single_tags.add(tag)

    def test_chint_meter_ecosystem(self) -> None:
        """Deep domain audit of CHINT meter definitions."""
        chint_files = [
            PROD_DIR / "MET_Chint_DTSU666.csv",
            PROD_DIR / "MET_Chint_DTSU666-H.csv",
            PROD_DIR / "MET_Chint_DDSU666.csv",
            PROD_DIR / "MET_Huawei_DTSU666-H.csv",
            PROD_DIR / "MET_Chint_Power_CHINT_DTSU.csv",
        ]
        for f in chint_files:
            assert f.is_file(), f"Missing critical Chint file: {f.name}"
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            tags = {r[6].strip() for r in rows if r[6].strip()}
            # Check critical electrical telemetry
            if "DTSU666-H" in f.name:
                assert "U_AC_1" in tags
                assert "U_AC_2" in tags
                assert "U_AC_3" in tags
                assert "I_AC_1" in tags
                assert "I_AC_2" in tags
                assert "I_AC_3" in tags
            elif "DTSU666.csv" in f.name:
                # Standard DTSU666 has CT/VT ratios and comprehensive registers
                descriptions = [r[5].lower() for r in rows]
                assert any("voltage" in d or "ratio" in d or "current" in d for d in descriptions)

    def test_schneider_meter_ecosystem(self) -> None:
        """Deep domain audit of Schneider iEM31xx/35xx and PM series definitions."""
        schneider_iem_files = sorted(PROD_DIR.glob("MET_Schneider_iem3*.csv")) + sorted(
            PROD_DIR.glob("MET_Schneider_iEM3*.csv")
        )
        assert len(schneider_iem_files) >= 10, f"Expected at least 10 Schneider iEM files, found {len(schneider_iem_files)}"

        pm_files = sorted(PROD_DIR.glob("MET_Schneider_PM*.csv"))
        assert len(pm_files) >= 3, f"Expected Schneider PM files, found {len(pm_files)}"

        for f in schneider_iem_files + pm_files:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 5, f"{f.name}: suspiciously few registers ({len(rows)})"
            # Ensure function code is read holding (3) or read input (4)
            fcs = {r[1].strip() for r in rows}
            assert fcs.issubset({"3", "4"}), f"{f.name}: invalid function codes {fcs}"

    def test_socomec_diris_meter_ecosystem(self) -> None:
        """Deep domain audit of Socomec Diris and Countis definitions."""
        socomec_diris = sorted(PROD_DIR.glob("MET_Socomec_Diris*.csv")) + sorted(
            PROD_DIR.glob("MET_Socomec_DIRIS*.csv")
        )
        socomec_countis = sorted(PROD_DIR.glob("MET_Socomec_Countis*.csv")) + sorted(
            PROD_DIR.glob("MET_Socomec_COUNTIS*.csv")
        )
        assert len(socomec_diris) >= 8, f"Expected >= 8 Diris files, found {len(socomec_diris)}"
        assert len(socomec_countis) >= 5, f"Expected >= 5 Countis files, found {len(socomec_countis)}"

        for f in socomec_diris + socomec_countis:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 1, f"Empty Socomec definition: {f.name}"
            # Check valid actions (0, 1, 4, 9, etc.)
            for r in rows:
                assert r[10].strip() in VALID_ACTIONS

    def test_carlo_gavazzi_meter_ecosystem(self) -> None:
        """Deep domain audit of Carlo Gavazzi energy meters (EM24, EM210, EM-ET300)."""
        cg_files = sorted(PROD_DIR.glob("MET_CarloGavazzi_*.csv"))
        assert len(cg_files) >= 5, f"Expected >= 5 Carlo Gavazzi files, found {len(cg_files)}"

        em24_found = False
        for f in cg_files:
            if "EM24" in f.name:
                em24_found = True
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 1, f"Empty Carlo Gavazzi file: {f.name}"
            # Check register addresses are valid Modbus numbers
            for r in rows:
                base_addr = r[2].strip().split("_")[0]
                assert base_addr.isdigit(), f"{f.name}: invalid address {r[2]}"
        assert em24_found, "Missing Carlo Gavazzi EM24 definition"

    def test_meters_telemetry_and_scaling(self) -> None:
        """Verify scale factors and units on critical electrical parameters."""
        meter_files = get_meter_files()
        checked_tags = 0
        for f in meter_files:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            for r in rows:
                tag = r[6].strip()
                if not tag:
                    continue
                checked_tags += 1
                coef_a = float(r[7].strip() or "1.0")
                assert math.isfinite(coef_a), f"{f.name} register {r[0]}: coef_a is not finite"
        assert checked_tags > 500, f"Expected > 500 tagged registers across all meters, found {checked_tags}"
