"""
Targeted Unit & Integration Tests for Wave 28 Coverage & Real-World Dataset Intelligence.

Tests covering:
- DatasetKnowledge JSON preloading, missing file handling, and corrupt JSON resilience.
- get_active_knowledge caching and singleton behavior.
- train_knowledge_from_definitions recursive directory traversal and duplicate file deduplication.
- DatasetKnowledge query methods (units, types, action codes) and entry recording edge cases.
- infer_sunpm_tag accuracy across real-world industrial telemetry and command metrics.
- Multilingual and prefixed variable normalization (clean_variable_key).
- Real-world manufacturer definition auditing from the network share (Acrel, SMA, Carlo Gavazzi, Legrand, Solax).
- Auto-fixing legacy Webdyn tags to canonical SunPM tags.
- Edge branches, tag inference, action codes, function code text parsing, and definition file/row enrichment.
"""

from __future__ import annotations

import csv
import os
import tempfile
from pathlib import Path

import pytest

from DefFileGenerator.knowledge_miner import (
    DEFAULT_DATASET_PATH,
    DEFAULT_EQUIPEMENTIERS_DIR,
    DatasetKnowledge,
    audit_definition_file,
    auto_fix_definition_file,
    clean_variable_key,
    enrich_definition_file,
    enrich_definition_rows,
    get_active_knowledge,
    infer_action_code,
    infer_sunpm_tag,
    infer_webdyn_tag,
    load_dataset_knowledge,
    normalize_tag_to_sunpm,
    parse_function_code_from_text,
    train_knowledge_from_definitions,
)


def test_load_dataset_knowledge_from_default_path():
    """Verify that load_dataset_knowledge loads the pre-trained dataset_knowledge.json correctly."""
    assert DEFAULT_DATASET_PATH.is_file(), f"{DEFAULT_DATASET_PATH} must exist"
    knowledge = load_dataset_knowledge()
    assert knowledge is not None
    assert isinstance(knowledge, DatasetKnowledge)

    # Validate scale of knowledge mined from 220,000+ registers
    assert knowledge.total_definitions_mined >= 800
    assert knowledge.total_registers_mined >= 200_000
    assert len(knowledge.name_to_tag) >= 800

    # Validate essential canonical tags
    assert "P_AC" in knowledge.tag_counts
    assert "CMD_P_AC_SET_PERCENT" in knowledge.tag_counts
    assert "SERIAL" in knowledge.tag_counts
    assert "MODEL" in knowledge.tag_counts
    assert "E_TOT" in knowledge.tag_counts
    assert "E_DAY" in knowledge.tag_counts


def test_load_dataset_knowledge_missing_and_corrupt(tmp_path: Path):
    """Verify that load_dataset_knowledge handles non-existent or corrupted files gracefully."""
    missing_file = tmp_path / "non_existent_knowledge.json"
    assert load_dataset_knowledge(missing_file) is None

    corrupt_file = tmp_path / "corrupt_knowledge.json"
    corrupt_file.write_text("{This is not valid JSON!}", encoding="utf-8")
    assert load_dataset_knowledge(corrupt_file) is None


def test_get_active_knowledge_singleton():
    """Verify that get_active_knowledge returns a valid instance and reuses it across calls."""
    k1 = get_active_knowledge()
    assert isinstance(k1, DatasetKnowledge)
    assert k1.total_definitions_mined > 0

    k2 = get_active_knowledge()
    assert k1 is k2


def test_train_knowledge_from_definitions_recursive_vs_flat(tmp_path: Path):
    """Verify that recursive=True traverses nested subdirectories whereas recursive=False does not."""
    header = "modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n"
    row = "1;3;100;U16;;Active Power;P_AC;1;0;W;4\n"

    (tmp_path / "root_def.csv").write_text(header + row, encoding="utf-8")
    sub1 = tmp_path / "sub1"
    sub1.mkdir()
    (sub1 / "sub1_def.csv").write_text(header + row, encoding="utf-8")
    sub2 = sub1 / "sub2"
    sub2.mkdir()
    (sub2 / "sub2_def.csv").write_text(header + row, encoding="utf-8")

    # Flat scan
    flat_knowledge = train_knowledge_from_definitions(
        directories=tmp_path, auto_activate=False, force_retrain=True, recursive=False
    )
    assert flat_knowledge.total_definitions_mined == 1

    # Recursive scan
    rec_knowledge = train_knowledge_from_definitions(
        directories=tmp_path, auto_activate=False, force_retrain=True, recursive=True
    )
    assert rec_knowledge.total_definitions_mined == 3


def test_train_knowledge_from_definitions_seen_files_deduplication(tmp_path: Path):
    """Verify that identically named files across multiple search directories are deduplicated."""
    header = "modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n"
    row = "1;3;100;U16;;Active Power;P_AC;1;0;W;4\n"

    dir1 = tmp_path / "dir1"
    dir2 = tmp_path / "dir2"
    dir1.mkdir()
    dir2.mkdir()

    # Identical filename in both directories
    (dir1 / "same_model.csv").write_text(header + row, encoding="utf-8")
    (dir2 / "same_model.csv").write_text(header + row, encoding="utf-8")

    knowledge = train_knowledge_from_definitions(
        directories=[dir1, dir2], auto_activate=False, force_retrain=True, recursive=True
    )
    assert knowledge.total_definitions_mined == 1


def test_dataset_knowledge_metric_query_edge_cases():
    """Verify DatasetKnowledge metric queries and empty/invalid entry recording."""
    k = DatasetKnowledge()

    # Querying empty model
    assert k.get_metric_units("P_AC") == {}
    assert k.get_metric_types("P_AC") == {}
    assert k.get_metric_actions("P_AC") == {}

    # Recording valid entry
    k.record_variable(
        name="Active Power Total",
        tag="P_AC",
        dtype="U32",
        unit="kW",
        action="4",
    )
    assert k.get_metric_units("P_AC") == {"kW": 1}
    assert k.get_metric_types("P_AC") == {"U32": 1}
    assert k.get_metric_actions("P_AC") == {"4": 1}

    # Querying by raw variable name
    assert k.get_metric_units("Active Power Total") == {"kW": 1}

    # Invalid / empty recording should not crash
    k.record_variable(name="", tag="P_AC")
    k.record_variable(name="Something", tag="")
    assert k.total_definitions_mined == 0


def test_clean_variable_key_multilingual_and_prefixes():
    """Verify token normalization on prefixes, multilingual text, and special characters."""
    assert clean_variable_key("1#PCS - Inverter Active Power (kW)") == "active power"
    assert clean_variable_key("INV_01 - Grid Voltage [V]") == "grid voltage"
    assert clean_variable_key("Puissance active totale") == "puissance active totale"
    assert clean_variable_key("Wirkleistung Phase L1 (W)") == "wirkleistung phase l1"
    assert clean_variable_key("  MPPT2_Current {A}  ") == "mppt2 current"
    assert clean_variable_key("---") == ""
    assert clean_variable_key("") == ""


def test_sunpm_tag_inference_real_world_corpus():
    """Verify infer_sunpm_tag resolves real-world industrial telemetry and command variable names."""
    k = get_active_knowledge()

    # Power & Energy
    assert infer_sunpm_tag("Total Active Power", knowledge=k) == "P_AC"
    assert infer_sunpm_tag("Daily Power Generation", knowledge=k) == "E_DAY"
    assert infer_sunpm_tag("Total Forward Active Energy", knowledge=k) == "E_TOT"
    assert infer_sunpm_tag("CosPhi", knowledge=k) == "PF"

    # Commands & Setpoints
    assert infer_sunpm_tag("Power limitation percent", knowledge=k) == "CMD_P_AC_SET_PERCENT"
    assert infer_sunpm_tag("Turn On Command", knowledge=k) == "CMD_ON"
    assert infer_sunpm_tag("Turn Off Command", knowledge=k) == "CMD_OFF"


def test_auto_fix_definition_legacy_to_canonical(tmp_path: Path):
    """Verify that auto_fix_definition_file upgrades legacy Webdyn tags to official SunPM tags."""
    def_file = tmp_path / "legacy_tags.csv"
    lines = [
        "modbusRTU;Inverter;Vendor;Model;;;;;;;",
        "1;3;100;U32;;Total Active Power;RealPower;1.000000;0;kW;4",
        "2;3;102;U32;;Total Cumulative Energy;EnergyTotal;1.000000;0;kWh;4",
        "3;3;104;U16;;Daily Energy Yield;EnergyDay;1.000000;0;kWh;4",
    ]
    def_file.write_text("\n".join(lines), encoding="utf-8-sig")

    repairs = auto_fix_definition_file(def_file)
    assert repairs["tags_normalized"] >= 1
    assert def_file.is_file()

    with open(def_file, encoding="utf-8-sig") as f:
        reader = list(csv.reader(f, delimiter=";", lineterminator="\n"))
    assert len(reader) == 4
    # Check that tags were normalized
    assert reader[1][6] == "P_AC"
    assert reader[2][6] == "E_TOT"
    assert reader[3][6] == "E_DAY"


@pytest.mark.skipif(
    not os.path.exists(DEFAULT_EQUIPEMENTIERS_DIR),
    reason=f"Network share not accessible: {DEFAULT_EQUIPEMENTIERS_DIR}",
)
def test_real_world_definitions_audit_from_share():
    """Audits real definition files from diverse manufacturers on the network share."""
    share = Path(DEFAULT_EQUIPEMENTIERS_DIR)
    target_rel_files = [
        r"Acrel\METER_ACREL_DTSD1352.csv",
        r"SMA\SMA_Modbus.csv",
        r"CarloGavazzi\Meter\CarloGavazzi_EM210.csv",
        r"Legrand\Meter\METER_LEGRAND_LG412041.csv",
        r"Solax\SOLAX-X1-X3-GENERIC.csv",
    ]

    for rel in target_rel_files:
        full_path = share / rel
        if full_path.is_file():
            report = audit_definition_file(full_path, strict_overlap=False)
            assert report.register_count > 0, f"{rel} should yield registers"
            assert report.filepath == str(full_path)


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
    assert (
        infer_sunpm_tag("Active Power", existing_tag="RealPower", normalize_existing=True) == "P_AC"
    )
    # Existing tag with normalize_existing=False
    assert (
        infer_sunpm_tag("Active Power", existing_tag="RealPower", normalize_existing=False)
        == "RealPower"
    )
    # Existing tag without normalization match
    assert (
        infer_sunpm_tag("Custom Name", existing_tag="CUSTOM_TAG", normalize_existing=True)
        == "CUSTOM_TAG"
    )


def test_infer_sunpm_tag_control_commands():
    assert infer_sunpm_tag("Control Power Enable (wmaxlim)") == "CMD_P_ENA"
    assert (
        infer_sunpm_tag("Reactive Power Setpoint", equipment_type="Inverter")
        == "CMD_Q_AC_SET_PERCENT"
    )
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
    assert (
        infer_sunpm_tag("Daily Avoided CO2 Emissions", equipment_type="Display")
        == "AVOIDED_CO2_DAY"
    )
    assert infer_sunpm_tag("Avoided CO2 Emissions", equipment_type="Display") == "AVOIDED_CO2"
    assert (
        infer_sunpm_tag("Self-Consumption Rate", equipment_type="Display")
        == "SELF_CONSUMPTION_RATIO"
    )
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
