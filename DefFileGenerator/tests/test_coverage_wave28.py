"""
Targeted unit tests for DefFileGenerator/knowledge_miner.py covering edge branches,
tag inference, action code, function code text parsing, and definition file/row enrichment.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

from DefFileGenerator.knowledge_miner import (
    enrich_definition_file,
    enrich_definition_rows,
    infer_action_code,
    infer_sunpm_tag,
    infer_webdyn_tag,
    normalize_tag_to_sunpm,
    parse_function_code_from_text,
)


def test_normalize_tag_to_sunpm_edge_cases():
    assert normalize_tag_to_sunpm("") == ""
    assert normalize_tag_to_sunpm("   ") == ""
    # Exact case match
    assert normalize_tag_to_sunpm("RealPower") == "P_AC"
    # Case-insensitive match
    assert normalize_tag_to_sunpm("realpower") == "P_AC"
    # Unknown tag returns cleaned original string
    assert normalize_tag_to_sunpm("  UNKNOWN_TAG_XYZ  ") == "UNKNOWN_TAG_XYZ"


def test_infer_webdyn_tag_prefix_stripping():
    # Tag with existing tag returns existing
    assert infer_webdyn_tag("PCS1 - Total Active Power", existing_tag="P_AC") == "P_AC"
    # Equipment prefix stripping
    assert infer_webdyn_tag("1#PCS - Total Active Power") == "RealPower"
    # Unmatched variable name returns empty
    assert infer_webdyn_tag("Random Variable 123") == ""


def test_infer_sunpm_tag_existing_tag_branching():
    # Existing tag with normalize_existing=True (default)
    assert infer_sunpm_tag("Active Power", existing_tag="RealPower", normalize_existing=True) == "P_AC"
    # Existing tag with normalize_existing=False
    assert infer_sunpm_tag("Active Power", existing_tag="RealPower", normalize_existing=False) == "RealPower"
    # Existing tag without normalization match
    assert infer_sunpm_tag("Custom Name", existing_tag="CUSTOM_TAG", normalize_existing=True) == "CUSTOM_TAG"


def test_infer_sunpm_tag_control_commands():
    assert infer_sunpm_tag("Control Power Enable (wmaxlim)") == "CMD_P_ENA"
    assert infer_sunpm_tag("Reactive Power Setpoint", equipment_type="Inverter") == "CMD_Q_AC_SET_PERCENT"
    assert infer_sunpm_tag("Reactive Power Setpoint (kVAR)") == "CMD_Q_AC_SET_ABS"
    assert infer_sunpm_tag("Reactive Power Setpoint (VAR)") == "CMD_Q_AC_SET_ABS"
    assert infer_sunpm_tag("Reactive Power Setpoint (abs)") == "CMD_Q_AC_SET_ABS"
    assert infer_sunpm_tag("Active Power Limit (kW)") == "CMD_P_AC_SET_ABS"
    assert infer_sunpm_tag("Active Power Limit (W)") == "CMD_P_AC_SET_ABS"
    assert infer_sunpm_tag("Active Power Limit (abs)") == "CMD_P_AC_SET_ABS"
    assert infer_sunpm_tag("Active Power Limitation Ratio (%)") == "CMD_P_AC_SET_PERCENT"
    assert infer_sunpm_tag("Commit Parameters Command") == "CMD_COMMIT"


def test_infer_sunpm_tag_ac_measurements():
    # AC Voltage per phase neutral / line
    assert infer_sunpm_tag("Voltage UA") == "U_AC_1"
    assert infer_sunpm_tag("Voltage UB") == "U_AC_2"
    assert infer_sunpm_tag("Voltage UC") == "U_AC_3"

    # AC Current phase branches
    assert infer_sunpm_tag("Reactive Power Phase A (qac1)") == "Q_AC_1"
    assert infer_sunpm_tag("Reactive Power Phase B (qac2)") == "Q_AC_2"
    assert infer_sunpm_tag("Reactive Power Phase C (qac3)") == "Q_AC_3"

    assert infer_sunpm_tag("Active Power Phase A (pac1)") == "P_AC_1"
    assert infer_sunpm_tag("Active Power Phase B (pac2)") == "P_AC_2"
    assert infer_sunpm_tag("Active Power Phase C (pac3)") == "P_AC_3"

    assert infer_sunpm_tag("Apparent Power (sac1)") == "S_AC_1"
    assert infer_sunpm_tag("Apparent Power (sac2)") == "S_AC_2"
    assert infer_sunpm_tag("Apparent Power (sac3)") == "S_AC_3"
    assert infer_sunpm_tag("Apparent Power") == "S_AC"

    # Frequency per phase
    assert infer_sunpm_tag("Grid Frequency Phase A (l1)") == "FREQ_1"
    assert infer_sunpm_tag("Grid Frequency Phase B (l2)") == "FREQ_2"
    assert infer_sunpm_tag("Grid Frequency Phase C (l3)") == "FREQ_3"

    # Power Factor per phase
    assert infer_sunpm_tag("Power Factor Phase A (l1)") == "PF_1"
    assert infer_sunpm_tag("Power Factor Phase B (l2)") == "PF_2"
    assert infer_sunpm_tag("Power Factor Phase C (l3)") == "PF_3"

    # Energy
    assert infer_sunpm_tag("Energy Exported (kWh)") == "E_EXP"
    assert infer_sunpm_tag("Total Energy L1") == "E_TOT_1"
    assert infer_sunpm_tag("Total Energy L2") == "E_TOT_2"
    assert infer_sunpm_tag("Total Energy L3") == "E_TOT_3"


def test_infer_sunpm_tag_mppt_strings_dc_and_sensors():
    # MPPT
    assert infer_sunpm_tag("MPPT 1 Current") == "I_DC_MPPT_1"
    assert infer_sunpm_tag("MPPT 2 Voltage") == "U_DC_MPPT_2"
    assert infer_sunpm_tag("MPPT 3 Power") == "P_DC_MPPT_3"

    # PV Strings
    assert infer_sunpm_tag("PV String 1 Current") == "I_DC_STRING_1"
    assert infer_sunpm_tag("PV String 2 Voltage") == "U_DC_STRING_2"

    # DC Bus
    assert infer_sunpm_tag("DC Bus Voltage") == "U_DC"
    assert infer_sunpm_tag("DC Bus Current") == "I_DC"
    assert infer_sunpm_tag("DC Bus Power") == "P_DC"

    # Temperature & Weather
    assert infer_sunpm_tag("Module Temperature") == "TEMP_CELL"
    assert infer_sunpm_tag("Heatsink Temperature") == "TEMP_INT"
    assert infer_sunpm_tag("Global Horizontal Irradiance") == "GHI"
    assert infer_sunpm_tag("Plane of Array Irradiance") == "POA"
    assert infer_sunpm_tag("Wind Bearing") == "WIND_BEARING"
    assert infer_sunpm_tag("Ambient Humidity") == "HUMIDITY"
    assert infer_sunpm_tag("Atmospheric Pressure") == "PRESSURE"
    assert infer_sunpm_tag("Rainfall") == "RAIN"

    # BESS & Displays
    assert infer_sunpm_tag("State of Health", equipment_type="BESS") == "SOH"
    assert infer_sunpm_tag("Insulation Resistance (R_ISO)") == "R_ISO"
    assert infer_sunpm_tag("Manufacturer Name") == "MANUFACTURER"
    assert infer_sunpm_tag("Modbus Slave Address") == "UNIT_ID"

    # Display / Public counters
    assert infer_sunpm_tag("Daily Avoided CO2 Emissions", equipment_type="Display") == "AVOIDED_CO2_DAY"
    assert infer_sunpm_tag("Avoided CO2 Emissions", equipment_type="Display") == "AVOIDED_CO2"
    assert infer_sunpm_tag("Self-Consumption Rate", equipment_type="Display") == "SELF_CONSUMPTION_RATIO"
    assert infer_sunpm_tag("Load Power Consumption", equipment_type="Display") == "LOAD_P_AC"


def test_enrich_definition_rows():
    rows = [
        {"Name": "Grid Frequency", "Tag": "GridFrequency"},
        {"Name": "Active Power", "Tag": "P_AC"},
        {"Name": "Custom Unmatched", "Tag": ""},
    ]
    # enrich without overwrite_existing (normalizes legacy)
    modified = enrich_definition_rows(rows, overwrite_existing=False, normalize_legacy=True)
    assert modified == 1
    assert rows[0]["Tag"] == "FREQ"
    assert rows[1]["Tag"] == "P_AC"
    assert rows[2]["Tag"] == ""

    # enrich with overwrite_existing=True
    rows_overwrite = [{"Name": "Grid Frequency", "Tag": "CUSTOM_FREQ"}]
    modified_ov = enrich_definition_rows(rows_overwrite, overwrite_existing=True)
    assert modified_ov == 1
    assert rows_overwrite[0]["Tag"] == "FREQ"


def test_enrich_definition_file_edge_cases():
    with tempfile.TemporaryDirectory() as tmp_dir:
        # 1. Empty file
        empty_csv = Path(tmp_dir) / "empty.csv"
        empty_csv.write_text("", encoding="utf-8")
        assert enrich_definition_file(str(empty_csv)) == 0

        # 2. File with header and short rows + full rows
        content = (
            "Header1;Inverter;Header3\n"
            "Row;Too;Short\n"
            "1;30000;4;1;1;Grid Frequency;GridFrequency\n"
            "1;30001;4;1;1;Custom Parameter;\n"
        )
        sample_csv = Path(tmp_dir) / "sample.csv"
        sample_csv.write_text(content, encoding="utf-8")

        out_csv = Path(tmp_dir) / "enriched.csv"
        modified = enrich_definition_file(str(sample_csv), output_path=str(out_csv))
        assert modified == 1
        written = out_csv.read_text(encoding="utf-8")
        assert "FREQ" in written
        assert "Row;Too;Short" in written


def test_infer_action_code():
    assert infer_action_code("Reserved register") == "0"
    assert infer_action_code("RES") == "0"
    assert infer_action_code("Write Active Power Limit", fc=6) == "1"
    assert infer_action_code("High Temperature Alarm") == "9"
    assert infer_action_code("Overvoltage Fault") == "9"
    assert infer_action_code("Grid Voltage L1") == "4"


def test_parse_function_code_from_text():
    assert parse_function_code_from_text("Function Code: 0x01 - Read Coils") == "1"
    assert parse_function_code_from_text("Function Code: 0x02 - Read Discrete Inputs") == "2"
    assert parse_function_code_from_text("Function Code: 0x04 - Read Input Registers") == "4"
    assert parse_function_code_from_text("Function Code: FC03 - Holding Registers") == "3"
    assert parse_function_code_from_text("Code Fonction: 05 - Write Single Coil") == "1"
    assert parse_function_code_from_text("FC 02 Discrete") == "2"
    assert parse_function_code_from_text("FC 16 Write Multiple") == "3"
    assert parse_function_code_from_text("Code Fonction = 4") == "4"
    assert parse_function_code_from_text("No function code in this text") == "4"
