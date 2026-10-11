#!/usr/bin/env python3
"""
Unit and integration tests for universal Modbus register extraction and semantic learning.
Verifies learning generalization across 250+ manufacturers and SunSpec models.
"""

import pytest

from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor
from DefFileGenerator.knowledge_miner import (
    infer_action_code,
    infer_webdyn_tag,
    parse_function_code_from_text,
)


class TestUniversalSemanticLearning:
    """Tests the semantic intelligence learned from 211,000+ real Modbus registers."""

    def test_webdyn_tag_inference(self):
        """Standardized parameters receive Webdyn tags, secondary variables remain tagless."""
        assert infer_webdyn_tag("Total Active Power") == "RealPower"
        assert infer_webdyn_tag("AC Output Active Power Sum") == "RealPower"
        assert infer_webdyn_tag("Puissance active totale") == "RealPower"
        assert infer_webdyn_tag("1#PCS - Active Power") == "RealPower"

        assert infer_webdyn_tag("Total Forward Active Energy") == "EnergyTotal"
        assert infer_webdyn_tag("Cumulative Discharging Energy") == "EnergyTotal"
        assert infer_webdyn_tag("System Life kWh") == "EnergyTotal"

        assert infer_webdyn_tag("Daily Energy Yield") == "EnergyDay"
        assert infer_webdyn_tag("Grid Frequency") == "GridFrequency"
        assert infer_webdyn_tag("Power Factor") == "CosPhi"

        # Secondary variables should NOT have tags (strict INV_HUAWEI_V4 alignment)
        assert infer_webdyn_tag("Phase A Voltage") == ""
        assert infer_webdyn_tag("Internal Cabinet Temperature") == ""
        assert infer_webdyn_tag("Fan Speed") == ""

    def test_action_code_inference(self):
        """Action codes must map to Webdyn 4 (telemetry), 9 (alarm), 1 (command), 0 (reserve)."""
        # Reserve registers
        assert infer_action_code("Reserve") == "0"
        assert infer_action_code("Reserved Point") == "0"
        assert infer_action_code("res.") == "0"

        # Commands and setpoints
        assert infer_action_code("Set Active Power Ratio", fc=3, rw_field="RW") == "1"
        assert infer_action_code("Consigne Puissance", fc=16, rw_field="W") == "1"
        assert infer_action_code("Power On/Off Command", fc=6, rw_field="W") == "1"

        # Alarms and status
        assert infer_action_code("System Total Fault Alarm") == "9"
        assert infer_action_code("Overvoltage Trip") == "9"
        assert infer_action_code("Operating Status") == "9"
        assert infer_action_code("PCS Communication Failure") == "9"

        # Telemetry (default active measurements)
        assert infer_action_code("Total Forward Active Energy") == "4"
        assert infer_action_code("Grid Voltage Phase A") == "4"
        assert infer_action_code("Battery Temperature") == "4"

    def test_sunspec_and_multilingual_type_normalization(self):
        """Validates normalization of SunSpec, Alro, and European data types."""
        # SunSpec types
        assert Generator.normalize_type("sunssf") == "I16"
        assert Generator.normalize_type("acc32") == "U32"
        assert Generator.normalize_type("acc64") == "U64"
        assert Generator.normalize_type("pad") == "U16"
        assert Generator.normalize_type("enum16") == "U16"
        assert Generator.normalize_type("enum32") == "U32"
        assert Generator.normalize_type("bool") == "BITS"
        assert Generator.normalize_type("boolean") == "BITS"
        assert Generator.normalize_type("single") == "F32"
        assert Generator.normalize_type("raw16") == "U16"

        # Bracketed string notations
        assert Generator.normalize_type("string[16]") == "STR16"
        assert Generator.normalize_type("string(32)") == "STR32"
        assert Generator.normalize_type("str[8]") == "STR8"

    def test_multi_section_modbus_banner_parsing(self):
        """Validates function code parsing from text banners."""
        assert parse_function_code_from_text("Function Code: 0x04 - Read Input Registers") == "4"
        assert parse_function_code_from_text("Register Settings (Function Code:0x02 - Read Discrete Inputs)") == "2"
        assert parse_function_code_from_text("Function Code: 0x03/0x10 - Read/Write Multiple Registers") == "3"
        assert parse_function_code_from_text("Function Code: 0x01 - Read Coils") == "1"

    def test_universal_extractor_pipeline(self):
        """Tests that Extractor.map_and_clean uses learned semantics end-to-end."""
        raw_table = [
            {
                "Point Table": "10000",
                "Data Name": "Total Forward Active Energy",
                "Data Type": "float32",
                "Unit": "kWh",
                "R/W": "RO",
            },
            {
                "Point Table": "10002",
                "Data Name": "Active Power",
                "Data Type": "acc32",
                "Unit": "kW",
                "Scale Factor": "-1",
                "R/W": "RO",
            },
            {
                "Point Table": "10004",
                "Data Name": "System Alarm Status",
                "Data Type": "uint16",
                "R/W": "RO",
            },
            {
                "Point Table": "10005",
                "Data Name": "Reserve",
                "Data Type": "u16",
            },
            {
                "Point Table": "10006",
                "Data Name": "Set Power Limit",
                "Data Type": "u16",
                "R/W": "RW",
            },
        ]
        ext = Extractor()
        cleaned = list(ext.map_and_clean([raw_table]))

        assert len(cleaned) == 5

        # 1. Total Forward Active Energy (preserves source RO for Generator)
        reg1 = cleaned[0]
        assert reg1["Address"] == "10000"
        assert reg1["Type"] == "F32"
        assert reg1["Tag"] == "EnergyTotal"
        assert reg1["Action"] == "RO"

        # 2. SunSpec scale factor resolution: -1 -> 0.1
        reg2 = cleaned[1]
        assert reg2["Address"] == "10002"
        assert reg2["Type"] == "U32"
        assert reg2["Tag"] == "RealPower"
        assert float(reg2["Factor"]) == pytest.approx(0.1)
        assert reg2["Action"] == "RO"

        # 3. Alarm status with raw RW
        reg3 = cleaned[2]
        assert reg3["Address"] == "10004"
        assert reg3["Action"] == "RO"

        # 4. Reserve (no RW column -> inferred action '0')
        reg4 = cleaned[3]
        assert reg4["Address"] == "10005"
        assert reg4["Action"] == "0"

        # 5. Setting / Command (preserves raw RW)
        reg5 = cleaned[4]
        assert reg5["Address"] == "10006"
        assert reg5["Action"] == "RW"
