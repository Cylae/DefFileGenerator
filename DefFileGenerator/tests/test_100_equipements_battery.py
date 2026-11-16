#!/usr/bin/env python3
"""
Industrial 100+ Equipment Multi-Manufacturer Test Battery.

Executes a comprehensive, falsifiable validation battery across 100+ distinct industrial
equipment models, covering 50+ world-leading manufacturers and 6 functional categories:
  - Batteries & BESS (BAT)
  - Loggers & Gateways (COM)
  - Displays, IO & Auxiliary Devices (OTH)
  - Pyranometers, Weather & Environmental Sensors (SEN)
  - Multi-function Energy Meters & Network Analyzers (MET)
  - Grid-tie, Hybrid & Central Solar Inverters (INV)

Also verifies live document extractions across multiple vendor media formats:
  - Multi-sheet Excel workbooks (e.g. Mars BESS Local Controller)
  - Modbus XML register dictionaries (5,000+ registers)
  - Multi-page Modbus PDF engineering manuals (Answer Drives Solargate, FoxESS)
"""

from __future__ import annotations

import csv
import os
import sys
import tempfile
import time
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest

# Ensure project root is available on sys.path
REPO_ROOT = Path(__file__).resolve().parent.parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator  # noqa: E402
from DefFileGenerator.extractor import Extractor  # noqa: E402

DEFINITIONS_DIR = REPO_ROOT / "Definitions"
STRESS_DATA_DIR = REPO_ROOT / "stress_test_data"


def count_csv_registers(filepath: str | Path) -> int:
    """Counts data register rows in a WebdynSunPM definition CSV."""
    with open(filepath, encoding="utf-8-sig", errors="replace") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader, None)  # Skip header row
        return sum(
            1 for row in reader if row and len(row) >= 4 and any(cell.strip() for cell in row)
        )


@dataclass(frozen=True)
class EquipmentItem:
    """Metadata specification for an equipment definition test item."""

    file_name: str
    manufacturer: str
    model: str
    category: str
    min_registers: int = 1


# 101 Curated Industrial Equipment Models across 50 Manufacturers
EQUIPMENT_CATALOG: list[EquipmentItem] = [
    # =========================================================================
    # 1. BATTERIES & BESS (BAT) - 4 models
    # =========================================================================
    EquipmentItem("BAT_Dyness_DH100F.csv", "Dyness", "DH100F", "Battery", 900),
    EquipmentItem("BAT_Dyness_DH200F.csv", "Dyness", "DH200F", "Battery", 900),
    EquipmentItem("BAT_Dyness_DH200Y.csv", "Dyness", "DH200Y", "Battery", 900),
    EquipmentItem(
        "BAT_Dyness_DH_Series_Generic.csv", "Dyness", "DH_Series_Generic", "Battery", 900
    ),
    # =========================================================================
    # 2. LOGGERS & GATEWAYS (COM) - 2 models
    # =========================================================================
    EquipmentItem("COM_GoodWe_Ezlogger3000.csv", "GoodWe", "Ezlogger3000", "Logger", 15),
    EquipmentItem("COM_Huawei_SL3000.csv", "Huawei", "SL3000", "Logger", 30),
    # =========================================================================
    # 3. DISPLAYS, IO & PERIPHERALS (OTH) - 5 models
    # =========================================================================
    EquipmentItem("OTH_ICPDAS_M7028.csv", "ICPDAS", "M7028", "Io", 10),
    EquipmentItem("OTH_SIEBERT_XC4XX.csv", "SIEBERT", "XC4XX", "Display", 3),
    EquipmentItem("OTH_Webdyn_Sens_IO_16DI.csv", "Webdyn", "Sens_IO_16DI", "Io", 20),
    EquipmentItem("OTH_Webdyn_Sens_IO_8AI.csv", "Webdyn", "Sens_IO_8AI", "Io", 15),
    EquipmentItem("OTH_Webdyn_Sens_IO_8DI-6DO.csv", "Webdyn", "Sens_IO_8DI-6DO", "Io", 20),
    # =========================================================================
    # 4. SENSORS & PYRANOMETERS (SEN) - 15 models across 13 manufacturers
    # =========================================================================
    EquipmentItem("SEN_ATERSA_DATASOLMET.csv", "ATERSA", "DATASOLMET", "Sensor", 5),
    EquipmentItem("SEN_BD_SENSOR_FUEL_SENSOR.csv", "BD_SENSOR", "FUEL_SENSOR", "Sensor", 4),
    EquipmentItem("SEN_DELTAOHM_HD523D.csv", "DELTAOHM", "HD523D", "Sensor", 10),
    EquipmentItem("SEN_DELTA_OHM_LP_PYRA_03_.csv", "DELTA_OHM", "LP_PYRA_03_", "Sensor", 5),
    EquipmentItem("SEN_EKO_MS40.csv", "EKO", "MS40", "Sensor", 5),
    EquipmentItem("SEN_Hukseflux_SR05-D1A3.csv", "Hukseflux", "SR05-D1A3", "Sensor", 10),
    EquipmentItem("SEN_Hukseflux_SR30.csv", "Hukseflux", "SR30", "Sensor", 15),
    EquipmentItem(
        "SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv",
        "IngenieurBuro",
        "Si-RS485TC-2T-v-MB",
        "Sensor",
        5,
    ),
    EquipmentItem("SEN_KIPPZONEN_RT1.csv", "KIPPZONEN", "RT1", "Sensor", 10),
    EquipmentItem("SEN_KippZonen_DustIQ.csv", "KippZonen", "DustIQ", "Sensor", 15),
    EquipmentItem("SEN_KippZonen_SMP12.csv", "KippZonen", "SMP12", "Sensor", 15),
    EquipmentItem("SEN_LSI_LASTEM_DPA870_873.csv", "LSI_LASTEM", "DPA870_873", "Sensor", 10),
    EquipmentItem("SEN_Lufft_WS10.csv", "Lufft", "WS10", "Sensor", 20),
    EquipmentItem("SEN_MEIER-NT_ADL-SR.csv", "MEIER-NT", "ADL-SR", "Sensor", 5),
    EquipmentItem("SEN_Seven_Sensor_3S-IS.csv", "Seven_Sensor", "3S-IS", "Sensor", 5),
    # =========================================================================
    # 5. METERS & NETWORK ANALYZERS (MET) - 25 models across 15 manufacturers
    # =========================================================================
    EquipmentItem("MET_ABB_B21-B23-B24.csv", "ABB", "B21-B23-B24", "Meter", 30),
    EquipmentItem("MET_ABB_M2M_Meter.csv", "ABB", "M2M_Meter", "Meter", 40),
    EquipmentItem("MET_Algodue_UPM209-309.csv", "Algodue", "UPM209-309", "Meter", 50),
    EquipmentItem("MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv", "CHAUVIN-ARNOUX", "Ulys_TT-M", "Meter", 15),
    EquipmentItem("MET_CHINT_DTSU666.csv", "CHINT", "DTSU666", "Meter", 20),
    EquipmentItem("MET_CIRCUTOR_CVME3MINI.csv", "CIRCUTOR", "CVME3MINI", "Meter", 25),
    EquipmentItem("MET_Carlo_Gavazzi_EM210.csv", "Carlo_Gavazzi", "EM210", "Meter", 30),
    EquipmentItem("MET_ECO_ADAPT_POWERELEC_6.csv", "ECO_ADAPT", "POWERELEC_6", "Meter", 15),
    EquipmentItem("MET_EDMI_Mk6.csv", "EDMI", "Mk6", "Meter", 20),
    EquipmentItem("MET_HAGER_Generic.csv", "HAGER", "Generic", "Meter", 15),
    EquipmentItem("MET_HUAWEI_Power_Meter.csv", "HUAWEI", "Power_Meter", "Meter", 20),
    EquipmentItem("MET_Inepro_PRO380-MOD-CT.csv", "Inepro", "PRO380-MOD-CT", "Meter", 20),
    EquipmentItem("MET_Janitza_Generic.csv", "Janitza", "Generic", "Meter", 50),
    EquipmentItem("MET_Legrand_LEG412052.csv", "Legrand", "LEG412052", "Meter", 25),
    EquipmentItem("MET_Lettel_MAP4-34RJ.csv", "Lettel", "MAP4-34RJ", "Meter", 30),
    EquipmentItem("MET_Lettel_MCT4-45.csv", "Lettel", "MCT4-45", "Meter", 25),
    EquipmentItem("MET_Lettel_MCX4-34.csv", "Lettel", "MCX4-34", "Meter", 25),
    EquipmentItem("MET_Lettel_MCX4-34V.csv", "Lettel", "MCX4-34V", "Meter", 25),
    EquipmentItem("MET_POLIER_MTR5LMOD.csv", "POLIER", "MTR5LMOD", "Meter", 20),
    EquipmentItem("MET_SOCOMEC_COUNTIS_E23.csv", "SOCOMEC", "COUNTIS_E23", "Meter", 15),
    EquipmentItem("MET_SOCOMEC_DIRIS_B10.csv", "SOCOMEC", "DIRIS_B10", "Meter", 20),
    EquipmentItem("MET_SOCOMEC_Diris_A40.csv", "SOCOMEC", "Diris_A40", "Meter", 30),
    EquipmentItem("MET_Schneider_EM1200-EM1220.csv", "Schneider", "EM1200-EM1220", "Meter", 20),
    EquipmentItem("MET_Schneider_PM5xxx.csv", "Schneider", "PM5xxx", "Meter", 30),
    EquipmentItem("MET_Weidmuller_610.csv", "Weidmuller", "610", "Meter", 25),
    # =========================================================================
    # 6. INVERTERS (INV) - 50 models across 20 manufacturers
    # =========================================================================
    EquipmentItem("INV_ABB_10.5.csv", "ABB", "10.5", "Inverter", 15),
    EquipmentItem("INV_ABB_PVS800-57B.csv", "ABB", "PVS800-57B", "Inverter", 1000),
    EquipmentItem(
        "INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv", "ABB", "TRIO_20.0-27.6_TL_OUTD", "Inverter", 50
    ),
    EquipmentItem(
        "INV_ABB_Ultra_Inverter_Module.csv", "ABB", "Ultra_Inverter_Module", "Inverter", 50
    ),
    EquipmentItem("INV_AnswerDrives_Solargate.csv", "AnswerDrives", "Solargate", "Inverter", 250),
    EquipmentItem("INV_CyberPower_1-Phase.csv", "CyberPower", "1-Phase", "Inverter", 20),
    EquipmentItem("INV_CyberPower_3-Phases.csv", "CyberPower", "3-Phases", "Inverter", 30),
    EquipmentItem(
        "INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv",
        "CyberPower",
        "CPSHB6000ETL48-CPSHB6000ETLFR",
        "Inverter",
        30,
    ),
    EquipmentItem("INV_CyberPower_GENERIC.csv", "CyberPower", "GENERIC", "Inverter", 20),
    EquipmentItem("INV_EFACEC_Efasolar.csv", "EFACEC", "Efasolar", "Inverter", 40),
    EquipmentItem(
        "INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv",
        "GoodWe",
        "ET_EH_BT_BH_EHB_AES_ABP_BTC_Series",
        "Inverter",
        80,
    ),
    EquipmentItem(
        "INV_Goodwe_ES_EM_S-BP_BP_Series.csv", "Goodwe", "ES_EM_S-BP_BP_Series", "Inverter", 40
    ),
    EquipmentItem("INV_Goodwe_GT_Series.csv", "Goodwe", "GT_Series", "Inverter", 40),
    EquipmentItem("INV_Goodwe_HT1100.csv", "Goodwe", "HT1100", "Inverter", 60),
    EquipmentItem("INV_Goodwe_HT1500.csv", "Goodwe", "HT1500", "Inverter", 60),
    EquipmentItem("INV_Goodwe_HT_Series.csv", "Goodwe", "HT_Series", "Inverter", 70),
    EquipmentItem("INV_Goodwe_MS_DNS_XS.csv", "Goodwe", "MS_DNS_XS", "Inverter", 40),
    EquipmentItem("INV_Goodwe_MT.csv", "Goodwe", "MT", "Inverter", 40),
    EquipmentItem("INV_Goodwe_SDTG2.csv", "Goodwe", "SDTG2", "Inverter", 40),
    EquipmentItem("INV_Goodwe_SMT.csv", "Goodwe", "SMT", "Inverter", 40),
    EquipmentItem("INV_Goodwe_SMT50K_SMT60K.csv", "Goodwe", "SMT50K_SMT60K", "Inverter", 50),
    EquipmentItem("INV_Growatt_Generic.csv", "Growatt", "Generic", "Inverter", 30),
    EquipmentItem(
        "INV_Growatt_MAX1500V_MAX-XLV_Series.csv",
        "Growatt",
        "MAX1500V_MAX-XLV_Series",
        "Inverter",
        60,
    ),
    EquipmentItem(
        "INV_Growatt_MOD_TL3-XH_Series.csv", "Growatt", "MOD_TL3-XH_Series", "Inverter", 50
    ),
    EquipmentItem(
        "INV_Growatt_TL-X_TL-XH_Series.csv", "Growatt", "TL-X_TL-XH_Series", "Inverter", 40
    ),
    EquipmentItem("INV_Growatt_TL3-X_Series.csv", "Growatt", "TL3-X_Series", "Inverter", 40),
    EquipmentItem("INV_Huawei_V2.csv", "Huawei", "V2", "Inverter", 50),
    EquipmentItem("INV_Huawei_V3.csv", "Huawei", "V3", "Inverter", 60),
    EquipmentItem("INV_Huawei_V4.csv", "Huawei", "V4", "Inverter", 70),
    EquipmentItem("INV_JEMA_Generic.csv", "JEMA", "Generic", "Inverter", 25),
    EquipmentItem("INV_KEHUA_SPI_100K-250K-B.csv", "KEHUA", "SPI_100K-250K-B", "Inverter", 50),
    EquipmentItem("INV_KOSTAL_PIKO_CI.csv", "KOSTAL", "PIKO_CI", "Inverter", 40),
    EquipmentItem("INV_KSTAR_Generic.csv", "KSTAR", "Generic", "Inverter", 30),
    EquipmentItem("INV_KSTAR_KIG1-250K.csv", "KSTAR", "KIG1-250K", "Inverter", 40),
    EquipmentItem("INV_Kaco_NX1-NX3_Series.csv", "Kaco", "NX1-NX3_Series", "Inverter", 40),
    EquipmentItem("INV_Kehua_SinglePhase.csv", "Kehua", "SinglePhase", "Inverter", 30),
    EquipmentItem("INV_Kehua_ThreePhase.csv", "Kehua", "ThreePhase", "Inverter", 40),
    EquipmentItem("INV_Kostal_Modbus.csv", "Kostal", "Modbus", "Inverter", 35),
    EquipmentItem("INV_SAJ_R6.csv", "SAJ", "R6", "Inverter", 30),
    EquipmentItem("INV_SAJ_Sununo_Plus_Suntrio.csv", "SAJ", "Sununo_Plus_Suntrio", "Inverter", 30),
    EquipmentItem("INV_SATCON_PowerGate.csv", "SATCON", "PowerGate", "Inverter", 30),
    EquipmentItem(
        "INV_SIEL_10_TL_SOLAR_INVERTER.csv", "SIEL", "10_TL_SOLAR_INVERTER", "Inverter", 15
    ),
    EquipmentItem("INV_SIEL_DSPX.csv", "SIEL", "DSPX", "Inverter", 15),
    EquipmentItem("INV_SIEL_Solar_Converters.csv", "SIEL", "Solar_Converters", "Inverter", 15),
    EquipmentItem("INV_SMA_Generic.csv", "SMA", "Generic", "Inverter", 35),
    EquipmentItem("INV_SMA_SC2xxx.csv", "SMA", "SC2xxx", "Inverter", 50),
    EquipmentItem("INV_SMA_STP_X.csv", "SMA", "STP_X", "Inverter", 60),
    EquipmentItem(
        "INV_Schneider_CONEXT_CORE_XC.csv", "Schneider", "CONEXT_CORE_XC", "Inverter", 45
    ),
    EquipmentItem("INV_SofarSolar_1-40KTL.csv", "SofarSolar", "1-40KTL", "Inverter", 40),
    EquipmentItem("INV_Solax_Generic.csv", "Solax", "Generic", "Inverter", 30),
]


# =============================================================================
# PARAMETRIZED 101 INDIVIDUAL EQUIPMENT TESTS
# =============================================================================


@pytest.mark.parametrize(
    "item",
    EQUIPMENT_CATALOG,
    ids=[f"{idx + 1:03d}_{item.file_name[:-4]}" for idx, item in enumerate(EQUIPMENT_CATALOG)],
)
def test_equipment_profile(item: EquipmentItem) -> None:
    """
    Validates an individual equipment definition file against Apollo-Standard contracts.

    Verifies:
      1. File existence and non-empty size in Definitions/
      2. First-line header metadata (Protocol, Category, Manufacturer, Model)
      3. Modbus register row invariants (address bounds, length >= 1, valid types, finite mul)
      4. Webdyn schema validation via Generator.validate_csv_detailed
      5. Minimum register count threshold for the given hardware model
    """
    file_path = DEFINITIONS_DIR / item.file_name
    assert file_path.is_file(), f"Definition file missing: {file_path}"
    assert file_path.stat().st_size > 0, f"Definition file is empty: {file_path}"

    generator = Generator()

    # 1. Inspect header metadata
    with open(file_path, encoding="utf-8-sig", errors="replace") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader, None)
        assert header is not None, f"Missing header row in {item.file_name}"
        assert len(header) >= 4, f"Header row too short ({len(header)}) in {item.file_name}"

        protocol, category, mfg, model = (
            header[0].strip(),
            header[1].strip(),
            header[2].strip(),
            header[3].strip(),
        )
        assert protocol.lower() in {"modbus", "modbusrtu", "modbustcp"}, (
            f"Invalid protocol '{protocol}' in {item.file_name}"
        )
        assert category == item.category, (
            f"Category mismatch: expected '{item.category}', got '{category}'"
        )
        assert mfg, f"Missing manufacturer in header of {item.file_name}"
        assert model, f"Missing model in header of {item.file_name}"

        # 2. Inspect individual register rows
        rows: list[list[str]] = [r for r in reader if r and any(cell.strip() for cell in r)]

    assert len(rows) >= item.min_registers, (
        f"Register count {len(rows)} below expected minimum {item.min_registers} for {item.file_name}"
    )

    valid_types = {
        "U16",
        "S16",
        "I16",
        "UINT16",
        "INT16",
        "U16_B",
        "S16_B",
        "I16_B",
        "U8",
        "S8",
        "I8",
        "UINT8",
        "INT8",
        "U32",
        "S32",
        "I32",
        "UINT32",
        "INT32",
        "U32_W",
        "S32_W",
        "I32_W",
        "UINT32_W",
        "INT32_W",
        "U32_WB",
        "S32_WB",
        "I32_WB",
        "FLOAT32",
        "FLOAT",
        "FLOAT32_W",
        "F32",
        "F32_W",
        "F32_WB",
        "FLOAT_W",
        "FLOAT64",
        "DOUBLE",
        "FLOAT64_W",
        "U64",
        "S64",
        "I64",
        "UINT64",
        "INT64",
        "U64_W",
        "I64_W",
        "STR",
        "STRING",
        "RAW",
        "HEX",
        "BITS",
        "BIT",
        "IP",
        "IPV6",
        "MAC",
        "DATE",
        "TIME",
        "DATETIME",
        "",
    }

    for line_idx, row in enumerate(rows, start=2):
        assert len(row) >= 4, (
            f"Row {line_idx} in {item.file_name} has {len(row)} columns, expected at least 4"
        )
        info1 = row[1].strip()
        reg_str = row[2].strip()
        signtype = row[3].strip().upper()
        name = row[5].strip() if len(row) > 5 else ""
        coefa_str = row[7].strip() if len(row) > 7 else ""

        assert name, f"Line {line_idx}: empty register name in {item.file_name}"
        assert info1 in {"1", "2", "3", "4"}, (
            f"Line {line_idx}: invalid Info1 (Function Code) '{info1}' in {item.file_name}"
        )

        # Parse register address (allow standard integer or bit-addressed format e.g. 100_0_1)
        base_addr_str = reg_str.split("_")[0]
        try:
            base_addr = int(base_addr_str)
            assert 0 <= base_addr <= 1000000, (
                f"Line {line_idx}: address {base_addr} out of range in {item.file_name}"
            )
        except ValueError:
            pytest.fail(
                f"Line {line_idx}: non-numeric register address '{reg_str}' in {item.file_name}"
            )

        # SignType validation
        assert signtype in valid_types, (
            f"Line {line_idx}: invalid SignType '{signtype}' in {item.file_name}"
        )

        # Multiplier finity
        if coefa_str:
            try:
                mul = float(coefa_str)
                assert not (mul != mul or mul == float("inf") or mul == float("-inf")), (
                    f"Line {line_idx}: non-finite multiplier '{coefa_str}' in {item.file_name}"
                )
            except ValueError:
                pytest.fail(
                    f"Line {line_idx}: invalid float multiplier '{coefa_str}' in {item.file_name}"
                )

    # 3. Webdyn detailed validation report
    report = generator.validate_csv_detailed(
        str(file_path),
        strict=True,
        strict_overlap=False,
    )
    assert report.is_valid, (
        f"Webdyn schema validation failed for {item.file_name}: {[str(i) for i in report.issues]}"
    )


# =============================================================================
# LIVE DOCUMENT EXTRACTION & COMPILATION TESTS
# =============================================================================


def test_live_extraction_mars_bess_multisheet_excel() -> None:
    """
    Validates live extraction and compilation from a multi-sheet industrial BESS workbook:
    'MarsLocalController_IO List_ModbusTCP_V7.0.xlsx'.
    """
    excel_path = STRESS_DATA_DIR / "MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
    if not excel_path.is_file():
        pytest.skip(f"Mars Excel file not found: {excel_path}")

    extractor = Extractor()
    generator = Generator()

    raw_tables = extractor.extract_from_excel(str(excel_path))
    mapped_stream = extractor.map_and_clean(raw_tables, address_offset=0)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_csv = os.path.join(tmp_dir, "BAT_Mars_MultiSheet_BESS.csv")
        cfg = GeneratorConfig(
            input_file=str(excel_path),
            output=out_csv,
            protocol="ModbusTCP",
            category="Battery",
            manufacturer="Mars",
            model="LocalControllerBESS",
            forced_write="True",
        )

        success = run_generator(cfg, input_data=mapped_stream)
        assert success, "run_generator failed for Mars multi-sheet Excel extraction"
        assert os.path.isfile(out_csv), f"Output file was not created: {out_csv}"
        assert os.path.getsize(out_csv) > 1000, "Output definition file unexpectedly small"

        report = generator.validate_csv_detailed(out_csv, strict=True, strict_overlap=False)
        assert report.is_valid, f"Validation issues: {[str(i) for i in report.issues]}"


def test_live_extraction_mars_ems_excel() -> None:
    """
    Validates targeted single-sheet extraction from Mars Local Controller:
    sheet 'LC Information'.
    """
    excel_path = STRESS_DATA_DIR / "MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
    if not excel_path.is_file():
        pytest.skip(f"Mars Excel file not found: {excel_path}")

    extractor = Extractor()
    generator = Generator()

    raw_tables = extractor.extract_from_excel(str(excel_path), sheet_name="LC Information")
    mapped_stream = extractor.map_and_clean(raw_tables, address_offset=0)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_csv = os.path.join(tmp_dir, "BAT_Mars_LC_Information.csv")
        cfg = GeneratorConfig(
            input_file=str(excel_path),
            output=out_csv,
            protocol="ModbusTCP",
            category="Battery",
            manufacturer="Mars",
            model="LC_Info",
            forced_write="True",
        )

        success = run_generator(cfg, input_data=mapped_stream)
        assert success, "run_generator failed for Mars LC Information extraction"
        assert os.path.isfile(out_csv), f"Output file missing: {out_csv}"

        report = generator.validate_csv_detailed(out_csv, strict=True, strict_overlap=False)
        assert report.is_valid, f"Validation issues: {[str(i) for i in report.issues]}"


def test_live_extraction_solargate_industrial_pdf() -> None:
    """
    Validates live extraction from real industrial inverter documentation:
    Answer Drives Solargate Manual (Pages 160-173).
    """
    candidate_paths = [
        Path(r"c:\Users\alexandre.calais\Downloads\ADSolargateManENHWSW 02-11-09 (1).pdf"),
        REPO_ROOT / "ADSolargateManENHWSW 02-11-09 (1).pdf",
        STRESS_DATA_DIR / "ADSolargateManENHWSW 02-11-09 (1).pdf",
    ]
    pdf_path = next((p for p in candidate_paths if p.is_file()), None)
    if pdf_path is None:
        pytest.skip("Solargate PDF manual not found in expected paths")

    extractor = Extractor()
    generator = Generator()

    raw_tables = extractor.extract_from_pdf(str(pdf_path), pages="160-173")
    mapped_stream = extractor.map_and_clean(raw_tables, address_offset=0)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_csv = os.path.join(tmp_dir, "INV_AnswerDrives_Solargate_Live.csv")
        cfg = GeneratorConfig(
            input_file=str(pdf_path),
            output=out_csv,
            protocol="ModbusRTU",
            category="Inverter",
            manufacturer="AnswerDrives",
            model="Solargate",
            forced_write="True",
        )

        success = run_generator(cfg, input_data=mapped_stream)
        assert success, "run_generator failed for Solargate PDF extraction"
        assert os.path.isfile(out_csv), f"Output file missing: {out_csv}"

        report = generator.validate_csv_detailed(out_csv, strict=True, strict_overlap=False)
        assert report.is_valid, f"Validation issues: {[str(i) for i in report.issues]}"
        assert count_csv_registers(out_csv) >= 200, (
            "Expected at least 200 registers extracted from PDF"
        )


def test_live_extraction_foxess_industrial_pdf() -> None:
    """
    Validates live extraction from FoxESS multi-page Modbus protocol manual.
    """
    pdf_candidates = list(STRESS_DATA_DIR.glob("*FoxESS*Modbus*Protocol*.pdf"))
    if not pdf_candidates:
        pytest.skip("FoxESS Modbus PDF protocol file not found in stress_test_data/")

    fox_path = pdf_candidates[0]
    extractor = Extractor()
    generator = Generator()

    raw_tables = extractor.extract_from_pdf(str(fox_path), pages="1-10")
    mapped_stream = extractor.map_and_clean(raw_tables, address_offset=0)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_csv = os.path.join(tmp_dir, "INV_FoxESS_Live.csv")
        cfg = GeneratorConfig(
            input_file=str(fox_path),
            output=out_csv,
            protocol="ModbusRTU",
            category="Inverter",
            manufacturer="FoxESS",
            model="P3_H3_Pro",
            forced_write="True",
        )

        success = run_generator(cfg, input_data=mapped_stream)
        assert success, "run_generator failed for FoxESS PDF extraction"
        assert os.path.isfile(out_csv), f"Output file missing: {out_csv}"

        report = generator.validate_csv_detailed(out_csv, strict=True, strict_overlap=False)
        assert report.is_valid, f"Validation issues: {[str(i) for i in report.issues]}"


def test_live_extraction_stress_controller_xml() -> None:
    """
    Validates live extraction and Webdyn definition generation from a massive 5,000-register XML map.
    """
    xml_path = STRESS_DATA_DIR / "stress.xml"
    if not xml_path.is_file():
        pytest.skip(f"Stress XML file not found: {xml_path}")

    extractor = Extractor()
    generator = Generator()

    raw_tables = extractor.extract_from_xml(str(xml_path))
    mapped_stream = extractor.map_and_clean(raw_tables, address_offset=0)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_csv = os.path.join(tmp_dir, "OTH_Industrial_Stress_XML.csv")
        cfg = GeneratorConfig(
            input_file=str(xml_path),
            output=out_csv,
            protocol="ModbusTCP",
            category="Io",
            manufacturer="Industrial",
            model="StressController",
            forced_write="True",
        )

        success = run_generator(cfg, input_data=mapped_stream)
        assert success, "run_generator failed for Stress XML extraction"
        assert os.path.isfile(out_csv), f"Output file missing: {out_csv}"

        reg_count = count_csv_registers(out_csv)
        assert reg_count >= 5000, f"Expected at least 5000 registers, got {reg_count}"

        report = generator.validate_csv_detailed(out_csv, strict=True, strict_overlap=False)
        assert report.is_valid, f"Validation issues: {[str(i) for i in report.issues]}"


# =============================================================================
# OVERALL BATTERY DIVERSITY & AGGREGATE CONTRACTS
# =============================================================================


def test_battery_diversity_and_totals() -> None:
    """
    Verifies that the 100+ equipment battery satisfies broad industrial diversity contracts:
      - At least 100 distinct equipment catalog items
      - At least 35 distinct industrial hardware manufacturers
      - All 6 functional equipment categories represented
      - Cumulative volume exceeding 18,000 Modbus registers verified
    """
    assert len(EQUIPMENT_CATALOG) >= 100, (
        f"Expected >= 100 equipment items, got {len(EQUIPMENT_CATALOG)}"
    )

    manufacturers = {item.manufacturer.lower() for item in EQUIPMENT_CATALOG}
    assert len(manufacturers) >= 35, (
        f"Expected >= 35 distinct manufacturers, got {len(manufacturers)}"
    )

    categories = {item.category for item in EQUIPMENT_CATALOG}
    expected_categories = {"Battery", "Logger", "Io", "Display", "Sensor", "Meter", "Inverter"}
    assert expected_categories.issubset(categories), (
        f"Missing categories: {expected_categories - categories}"
    )

    total_registers = 0
    for item in EQUIPMENT_CATALOG:
        csv_file = DEFINITIONS_DIR / item.file_name
        assert csv_file.is_file(), f"File missing: {csv_file}"
        total_registers += count_csv_registers(csv_file)

    assert total_registers >= 18000, (
        f"Expected >= 18,000 registers across catalog, verified {total_registers}"
    )


# =============================================================================
# STANDALONE BENCHMARK & REPORTING RUNNER
# =============================================================================


def run_benchmark_and_generate_report() -> None:
    """Runs all 100+ equipment tests sequentially, measures latency, and outputs a Markdown audit."""
    print("=" * 80)
    print("APOLLO-STANDARD 100+ INDUSTRIAL EQUIPMENT VALIDATION BATTERY")
    print("=" * 80)

    start_total = time.perf_counter()
    results: list[dict[str, Any]] = []

    print(
        f"{'Idx':<4} | {'Category':<10} | {'Manufacturer':<16} | {'Model':<24} | {'Regs':<6} | {'Status':<6} | {'Time (ms)'}"
    )
    print("-" * 80)

    for idx, item in enumerate(EQUIPMENT_CATALOG, start=1):
        t0 = time.perf_counter()
        file_path = DEFINITIONS_DIR / item.file_name
        status = "PASS"
        regs = 0
        try:
            regs = count_csv_registers(file_path)
            test_equipment_profile(item)
        except Exception as e:
            status = f"FAIL ({e})"

        elapsed_ms = (time.perf_counter() - t0) * 1000
        results.append(
            {
                "idx": idx,
                "category": item.category,
                "manufacturer": item.manufacturer,
                "model": item.model,
                "file": item.file_name,
                "registers": regs,
                "status": status,
                "time_ms": elapsed_ms,
            }
        )
        print(
            f"{idx:<4} | {item.category:<10} | {item.manufacturer:<16} | {item.model[:24]:<24} | {regs:<6} | {status:<6} | {elapsed_ms:6.2f}"
        )

    total_time = time.perf_counter() - start_total
    passed = sum(1 for r in results if r["status"] == "PASS")
    total_regs = sum(r["registers"] for r in results)

    print("=" * 80)
    print(
        f"BENCHMARK COMPLETE: {passed}/{len(results)} PASSED in {total_time:.2f}s ({total_regs:,} registers verified)"
    )
    print("=" * 80)

    # Generate Markdown Report
    report_path = REPO_ROOT / "docs" / "audit" / "BATTERY_100_EQUIPMENTS_REPORT.md"
    report_path.parent.mkdir(parents=True, exist_ok=True)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# Apollo-Standard 100+ Industrial Equipment Battery Report\n\n")
        f.write(f"- **Execution Date**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"- **Total Equipments Tested**: {len(results)}\n")
        f.write(
            f"- **Total Distinct Manufacturers**: {len({r['manufacturer'] for r in results})}\n"
        )
        f.write(f"- **Total Registers Verified**: {total_regs:,}\n")
        f.write(f"- **Pass Rate**: {passed}/{len(results)} ({passed / len(results) * 100:.1f}%)\n")
        f.write(f"- **Cumulative Execution Time**: {total_time:.2f} seconds\n")
        f.write(
            f"- **Average Verification Latency**: {total_time / len(results) * 1000:.2f} ms / equipment\n\n"
        )

        f.write("## 1. Equipment Category Breakdown\n\n")
        f.write("| Category | Models Tested | Total Registers |\n")
        f.write("| :--- | :---: | :---: |\n")
        cat_models: dict[str, int] = defaultdict(int)
        cat_regs: dict[str, int] = defaultdict(int)
        for r in results:
            cat_models[r["category"]] += 1
            cat_regs[r["category"]] += r["registers"]
        for cat in sorted(cat_models):
            f.write(f"| **{cat}** | {cat_models[cat]} | {cat_regs[cat]:,} |\n")

        f.write("\n## 2. Detailed Equipment Test Ledger\n\n")
        f.write(
            "| # | Category | Manufacturer | Model | Registers | Verification Latency | Status |\n"
        )
        f.write("| :-: | :--- | :--- | :--- | -: | -: | :-: |\n")
        for r in results:
            f.write(
                f"| {r['idx']} | {r['category']} | {r['manufacturer']} | {r['model']} | "
                f"{r['registers']} | {r['time_ms']:.2f} ms | `{r['status']}` |\n"
            )

    print(f"\nReport generated successfully: {report_path}")


if __name__ == "__main__":
    run_benchmark_and_generate_report()
