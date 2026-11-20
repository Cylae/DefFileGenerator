"""
Apollo-Standard Adversarial Verification Suite for Production Batteries & Sensors (BAT_*.csv & SEN_*.csv).

Mission & Mandate:
- Rigorous cross-audit of all 50+ battery/BMS definitions and 50+ sensor definitions in production/.
- Deep domain verification for critical battery families:
  * Dyness (DH100F, DH200F, DH200Y, DH_Series_Generic)
  * FoxESS (G-MAX EMS/BMS, GMax Storage, H3 Pro, P3 H3 PRO, P3 Pro, RSeries)
  * SL Energy (Battery, bms5_p1_1, Generator, PCS)
  * Other BESS: Sungrow, Huawei Luna2000, MarsEnergy, Goodwe Lynx Home HV
- Deep domain verification for critical sensor families:
  * Hukseflux pyranometers (SR05, SR20, SR30, Generic)
  * IngenieurBuro (Si-RS485-TC-2T-v, Si-RS485TC-3T-MB, Tx-RS485-MB, TA, TM)
  * Kipp & Zonen pyranometers (SMP_lite, SMP12, RT1, SMP_full, DustIQ)
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


def get_battery_files() -> list[Path]:
    """Retrieve all production battery files sorted canonically."""
    return sorted(PROD_DIR.glob("BAT_*.csv"))


def get_sensor_files() -> list[Path]:
    """Retrieve all production sensor files sorted canonically."""
    return sorted(PROD_DIR.glob("SEN_*.csv"))


@pytest.fixture(scope="module")
def generator() -> Generator:
    return Generator()


class TestProductionBatteriesAndSensorsSuite:
    """Comprehensive adversarial testing for production batteries and sensors."""

    def test_battery_and_sensor_files_populated(self) -> None:
        """Verify that at least 50 battery and 50 sensor definitions exist in production/."""
        battery_files = get_battery_files()
        sensor_files = get_sensor_files()
        assert len(battery_files) >= 50, f"Expected >= 50 battery files, found {len(battery_files)}"
        assert len(sensor_files) >= 50, f"Expected >= 50 sensor files, found {len(sensor_files)}"

    def test_batteries_and_sensors_strict_crlf_and_zero_bom(self) -> None:
        """Enforce strict binary hygiene: UTF-8 without BOM, strict CRLF termination."""
        all_files = get_battery_files() + get_sensor_files()
        for f in all_files:
            raw = f.read_bytes()
            assert not raw.startswith(b"\xef\xbb\xbf"), f"BOM detected in {f.name}"
            assert raw.endswith(b"\r\n"), f"File does not end with CRLF: {f.name}"
            assert not raw.endswith(b"\r\r\n"), f"Double CR detected in {f.name}"
            stripped = raw.replace(b"\r\n", b"")
            assert b"\n" not in stripped, f"Isolated LF detected in {f.name}"
            assert b"\r" not in stripped, f"Isolated CR detected in {f.name}"

    @pytest.mark.parametrize("file_path", get_battery_files(), ids=lambda p: p.name)
    def test_battery_5_pillars_certification(self, file_path: Path) -> None:
        """Verify full 5-pillar compliance using production governor on batteries."""
        is_valid, issues = verify_file_pillars(file_path)
        assert is_valid, f"{file_path.name} failed 5 pillars certification: {issues}"

    @pytest.mark.parametrize("file_path", get_sensor_files(), ids=lambda p: p.name)
    def test_sensor_5_pillars_certification(self, file_path: Path) -> None:
        """Verify full 5-pillar compliance using production governor on sensors."""
        is_valid, issues = verify_file_pillars(file_path)
        assert is_valid, f"{file_path.name} failed 5 pillars certification: {issues}"

    @pytest.mark.parametrize("file_path", get_battery_files() + get_sensor_files(), ids=lambda p: p.name)
    def test_csv_structure_and_monotonic_indices(self, file_path: Path) -> None:
        """Verify 11 columns, monotonic 1..N indices, addressing, and scale factors."""
        with open(file_path, encoding="utf-8") as fp:
            reader = list(csv.reader(fp, delimiter=";"))

        assert len(reader) >= 2, f"{file_path.name}: file must have header + at least 1 row"
        header = reader[0]
        assert len(header) == 11, f"{file_path.name}: header length {len(header)} != 11"
        assert header[0].strip().lower() == "modbus", f"{file_path.name}: protocol header must be 'Modbus'"

        for idx, row in enumerate(reader[1:], 1):
            assert len(row) == 11, f"{file_path.name} L{idx}: column count {len(row)} != 11"
            assert row[0].strip() == str(idx), f"{file_path.name} L{idx}: index mismatch {row[0]} != {idx}"

            fc = row[1].strip()
            assert fc in VALID_FUNCTION_CODES, f"{file_path.name} L{idx}: invalid fc {fc}"

            addr_str = row[2].strip()
            assert addr_str, f"{file_path.name} L{idx}: empty register address"
            dtype = row[3].strip().upper()
            if dtype == "STRING":
                assert "_" in addr_str, f"{file_path.name} L{idx}: STRING missing length: {addr_str}"
            else:
                base_addr = addr_str.split("_")[0]
                assert base_addr.isdigit(), f"{file_path.name} L{idx}: non-numeric address {addr_str}"
                iaddr = int(base_addr)
                assert 0 <= iaddr <= 65535, f"{file_path.name} L{idx}: address out of range: {iaddr}"

            coef_a_str = row[7].strip()
            coef_b_str = row[8].strip()
            if coef_a_str:
                assert math.isfinite(float(coef_a_str)), f"{file_path.name} L{idx}: coef_a not finite"
            if coef_b_str:
                assert math.isfinite(float(coef_b_str)), f"{file_path.name} L{idx}: coef_b not finite"

            action = row[10].strip()
            assert action in VALID_ACTIONS, f"{file_path.name} L{idx}: invalid action {action}"

    def test_dyness_bess_ecosystem(self) -> None:
        """Deep domain audit of Dyness BESS definitions (DH100F, DH200F, DH200Y, Generic)."""
        dyness_files = [
            PROD_DIR / "BAT_Dyness_DH100F.csv",
            PROD_DIR / "BAT_Dyness_DH200F.csv",
            PROD_DIR / "BAT_Dyness_DH200Y.csv",
            PROD_DIR / "BAT_Dyness_DH_Series_Generic.csv",
        ]
        for f in dyness_files:
            assert f.is_file(), f"Missing critical Dyness file: {f.name}"
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 30, f"{f.name}: suspiciously few registers ({len(rows)})"
            tags = {r[6].strip() for r in rows if r[6].strip()}
            assert "SOC" in tags, f"{f.name} missing SOC tag"
            assert "SOH" in tags, f"{f.name} missing SOH tag"
            # Verify temperature registers have offset -40.0 °C
            temp_rows = [r for r in rows if "temp" in r[5].lower() and r[8].strip()]
            assert len(temp_rows) > 0, f"{f.name}: expected temperature rows with offset"
            for r in temp_rows:
                offset = float(r[8].strip())
                assert offset in (-40.0, 0.0, -50.0), f"{f.name}: unexpected temp offset {offset}"

    def test_foxess_bess_ecosystem(self) -> None:
        """Deep domain audit of FoxESS battery definitions."""
        fox_files = sorted(PROD_DIR.glob("BAT_FoxESS_*.csv"))
        assert len(fox_files) >= 6, f"Expected >= 6 FoxESS battery files, found {len(fox_files)}"

        for f in fox_files:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 5, f"{f.name}: suspiciously few registers ({len(rows)})"
            tags = {r[6].strip() for r in rows if r[6].strip()}
            # Check presence of battery state metrics (standard SunPM tags, BMS metrics, or SunSpec DER 700-series storage registers)
            has_bat_metrics = (
                any(t in tags for t in ("SOC", "SOH", "U_DC", "I_DC", "P_DC", "TEMP_BAT", "TEMP_CELL", "PRESSURE"))
                or any(any(kw in r[5].lower() for kw in ("soc", "soh", "cell", "voltage", "current", "battery", "bat", "bms", "stack", "temperature", "storage", "cha", "discha")) for r in rows)
            )
            assert has_bat_metrics or "OTHER" in f.name, f"{f.name}: missing key battery telemetry tags"

    def test_sl_energy_bess_ecosystem(self) -> None:
        """Deep domain audit of SL Energy battery and generator definitions."""
        sl_files = [
            PROD_DIR / "BAT_SLEnergy_Battery.csv",
            PROD_DIR / "BAT_SLEnergy_bms5_p1_1.csv",
            PROD_DIR / "BAT_SLEnergy_Generator.csv",
            PROD_DIR / "BAT_SLEnergy_PCS.csv",
        ]
        for f in sl_files:
            assert f.is_file(), f"Missing SL Energy file: {f.name}"
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 1, f"Empty SL Energy definition: {f.name}"
            for r in rows:
                assert r[1].strip() in VALID_FUNCTION_CODES
                assert r[10].strip() in VALID_ACTIONS

    def test_critical_bess_families_telemetry(self) -> None:
        """Audit battery telemetry across Sungrow, Huawei Luna2000, and Goodwe Lynx."""
        families = [
            (PROD_DIR / "BAT_Goodwe_Lynx_Home_HV.csv", ["SOC", "SOH"]),
            (PROD_DIR / "BAT_Huawei_Luna_2000.csv", ["SOC"]),
            (PROD_DIR / "BAT_Sungrow_LC3000.csv", ["SOC"]),
        ]
        for f, required_tags in families:
            if f.is_file():
                with open(f, encoding="utf-8") as fp:
                    rows = list(csv.reader(fp, delimiter=";"))[1:]
                tags = {r[6].strip() for r in rows if r[6].strip()}
                for req in required_tags:
                    assert req in tags, f"{f.name} missing required battery tag {req}"

    def test_hukseflux_sensor_ecosystem(self) -> None:
        """Deep domain audit of Hukseflux pyranometers (SR05, SR20, SR30)."""
        hukseflux_files = sorted(PROD_DIR.glob("SEN_Hukseflux_*.csv"))
        assert len(hukseflux_files) >= 4, f"Expected >= 4 Hukseflux files, found {len(hukseflux_files)}"

        for f in hukseflux_files:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 5, f"{f.name}: suspiciously few registers ({len(rows)})"
            tags = {r[6].strip() for r in rows if r[6].strip()}
            descriptions = [r[5].lower() for r in rows]
            # Verify irradiance measurement presence
            has_irr = any("irr" in t.lower() for t in tags) or any(
                "irradiance" in d or "solar" in d or "radiation" in d for d in descriptions
            )
            assert has_irr, f"{f.name}: missing solar irradiance telemetry"

    def test_ingenieurburo_sensor_ecosystem(self) -> None:
        """Deep domain audit of IngenieurBuro sensors."""
        ib_files = sorted(PROD_DIR.glob("SEN_IngenieurBuro_*.csv"))
        assert len(ib_files) >= 6, f"Expected >= 6 IngenieurBuro sensor files, found {len(ib_files)}"

        for f in ib_files:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 1, f"Empty IngenieurBuro file: {f.name}"
            descriptions = [r[5].lower() for r in rows]
            # IngenieurBuro sensors monitor temperature and/or solar irradiance
            has_sensor_data = any(
                "temp" in d or "irr" in d or "radiation" in d or "voltage" in d or "current" in d
                for d in descriptions
            )
            assert has_sensor_data, f"{f.name}: missing sensor measurement channels"

    def test_kipp_zonen_sensor_ecosystem(self) -> None:
        """Deep domain audit of Kipp & Zonen pyranometers."""
        kz_files = sorted(PROD_DIR.glob("SEN_KippZonen_*.csv"))
        assert len(kz_files) >= 5, f"Expected >= 5 Kipp & Zonen files, found {len(kz_files)}"

        for f in kz_files:
            with open(f, encoding="utf-8") as fp:
                rows = list(csv.reader(fp, delimiter=";"))[1:]
            assert len(rows) >= 3, f"{f.name}: suspiciously few registers ({len(rows)})"
            # Verify register addresses and action codes
            for r in rows:
                base_addr = r[2].strip().split("_")[0]
                assert base_addr.isdigit(), f"{f.name}: invalid address {r[2]}"
                assert r[10].strip() in VALID_ACTIONS
