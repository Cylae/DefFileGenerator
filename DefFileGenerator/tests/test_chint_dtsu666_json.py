"""
Unit and integration tests for CHINT DTSU666 Modbus JSON specification extraction
and WebdynSunPM definition file generation.
"""

from __future__ import annotations

import csv
import json
import os
import tempfile
import unittest
from pathlib import Path
from typing import ClassVar

from DefFileGenerator.extractor import Extractor
from DefFileGenerator.knowledge_miner import audit_definition_file, infer_sunpm_tag
from generate_webdyn_def import generate_webdyn_definition


class TestChintDtsu666Json(unittest.TestCase):
    """Test suite for CHINT DTSU666 JSON extraction and definition generation."""

    json_download_path: ClassVar[Path] = Path(
        r"c:\Users\alexandre.calais\Downloads\chint-dtsu666.json"
    )
    sample_json_path: ClassVar[Path] = (
        Path(__file__).parent / "fixtures" / "chint_dtsu666_sample.json"
    )

    @classmethod
    def setUpClass(cls) -> None:
        cls.json_download_path = Path(r"c:\Users\alexandre.calais\Downloads\chint-dtsu666.json")
        cls.sample_json_path = Path(__file__).parent / "fixtures" / "chint_dtsu666_sample.json"
        # Create fixtures dir if needed
        cls.sample_json_path.parent.mkdir(parents=True, exist_ok=True)
        if cls.json_download_path.is_file():
            # Copy to fixture for durable CI execution
            content = cls.json_download_path.read_text(encoding="utf-8")
            cls.sample_json_path.write_text(content, encoding="utf-8")
        elif not cls.sample_json_path.is_file():
            # Fallback fixture if download file not present
            fallback_data = {
                "manufacturer": "CHINT",
                "model": "DTSU666",
                "holding_registers": [
                    {
                        "address": 0,
                        "name": "Software version",
                        "data_type": "Int16",
                        "units": "",
                        "scale_factor": 1,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 6,
                        "name": "Current transformer ratio (IrAt)",
                        "data_type": "Int16",
                        "units": "",
                        "scale_factor": 1,
                        "offset": 0,
                        "access": "RW",
                    },
                    {
                        "address": 7,
                        "name": "Voltage transformer ratio (UrAt)",
                        "data_type": "Int16",
                        "units": "",
                        "scale_factor": 1,
                        "offset": 0,
                        "access": "RW",
                    },
                    {
                        "address": 4126,
                        "name": "Forward active energy, total",
                        "data_type": "Float32",
                        "units": "kWh",
                        "scale_factor": 1,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 4136,
                        "name": "Reverse active energy, total",
                        "data_type": "Float32",
                        "units": "kWh",
                        "scale_factor": 1,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 8192,
                        "name": "Line voltage Uab",
                        "data_type": "Float32",
                        "units": "V",
                        "scale_factor": 0.1,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 8204,
                        "name": "Phase current Ia",
                        "data_type": "Float32",
                        "units": "A",
                        "scale_factor": 0.001,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 8210,
                        "name": "Combined active power",
                        "data_type": "Float32",
                        "units": "W",
                        "scale_factor": 0.1,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 8218,
                        "name": "Combined reactive power",
                        "data_type": "Float32",
                        "units": "var",
                        "scale_factor": 0.1,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 8234,
                        "name": "Combined power factor",
                        "data_type": "Float32",
                        "units": "",
                        "scale_factor": 0.001,
                        "offset": 0,
                        "access": "R",
                    },
                    {
                        "address": 8260,
                        "name": "Frequency",
                        "data_type": "Float32",
                        "units": "Hz",
                        "scale_factor": 0.01,
                        "offset": 0,
                        "access": "R",
                    },
                ],
            }
            cls.sample_json_path.write_text(json.dumps(fallback_data), encoding="utf-8")

    def test_json_extractor_parses_sections(self) -> None:
        """Verifies that Extractor.extract_from_json parses holding registers accurately."""
        extractor = Extractor()
        target_path = (
            str(self.json_download_path)
            if self.json_download_path.is_file()
            else str(self.sample_json_path)
        )
        tables = list(extractor.extract_from_json(target_path))
        self.assertEqual(len(tables), 1)

        rows = list(tables[0])
        self.assertGreaterEqual(len(rows), 11)

        # Check key register mappings
        addrs = [r.get("address") for r in rows]
        self.assertIn(8192, addrs)
        self.assertIn(4126, addrs)
        self.assertIn(8210, addrs)

    def test_dtsu666_tag_inference_accuracy(self) -> None:
        """Verifies zero false positives and exact canonical SunPM inferences for DTSU666."""
        cases = [
            ("Forward active energy, total", "E_TOT"),
            ("Reverse active energy, total", "E_EXP"),
            ("Line voltage Uab", "U_AC_L1L2"),
            ("Line voltage Ubc", "U_AC_L2L3"),
            ("Line voltage Uca", "U_AC_L3L1"),
            ("Phase voltage Ua", "U_AC_1"),
            ("Phase voltage Ub", "U_AC_2"),
            ("Phase voltage Uc", "U_AC_3"),
            ("Phase current Ia", "I_AC_1"),
            ("Phase current Ib", "I_AC_2"),
            ("Phase current Ic", "I_AC_3"),
            ("Combined active power", "P_AC"),
            ("Phase A active power", "P_AC_1"),
            ("Phase B active power", "P_AC_2"),
            ("Phase C active power", "P_AC_3"),
            ("Combined reactive power", "Q_AC"),
            ("Phase A reactive power", "Q_AC_1"),
            ("Phase B reactive power", "Q_AC_2"),
            ("Phase C reactive power", "Q_AC_3"),
            ("Combined power factor", "PF"),
            ("Phase A power factor", "PF_1"),
            ("Phase B power factor", "PF_2"),
            ("Phase C power factor", "PF_3"),
            ("Frequency", "FREQ"),
        ]
        for name, expected in cases:
            with self.subTest(name=name):
                self.assertEqual(
                    infer_sunpm_tag(name),
                    expected,
                    f"Tag for '{name}' was '{infer_sunpm_tag(name)}', expected '{expected}'",
                )

        # Disambiguation safety guards
        for non_tag in (
            "Current transformer ratio (IrAt)",
            "Voltage transformer ratio (UrAt)",
            "Display scroll time",
            "Backlight time",
            "Protocol / parity selection",
            "Baud rate",
            "Wiring mode",
            "Software version",
        ):
            with self.subTest(non_tag=non_tag):
                tag = infer_sunpm_tag(non_tag)
                self.assertNotIn(
                    tag,
                    {"MODEL", "SERIAL", "PF", "P_AC", "FREQ"},
                    f"Configuration setting '{non_tag}' must not be tagged with telemetry/device info '{tag}'",
                )

    def test_dtsu666_end_to_end_generation_and_validation(self) -> None:
        """Verifies end-to-end definition generation from DTSU666 JSON to valid CSV."""
        target_path = (
            str(self.json_download_path)
            if self.json_download_path.is_file()
            else str(self.sample_json_path)
        )
        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
            tmp_path = tmp.name

        try:
            success = generate_webdyn_definition(
                input_file=target_path,
                output_file=tmp_path,
                manufacturer="CHINT",
                model="DTSU666",
                protocol="modbusRTU",
                category="Meter",
                address_offset=0,
                strict_validation=True,
            )
            self.assertTrue(success, "Definition generation from JSON failed.")

            # Audit generated definition
            report = audit_definition_file(tmp_path, strict_overlap=True)
            self.assertTrue(report.is_valid, f"Audit failed with errors: {report.errors}")
            self.assertEqual(len(report.errors), 0)

            # Inspect CSV rows directly
            with open(tmp_path, encoding="utf-8") as f:
                reader = list(csv.reader(f, delimiter=";"))

            # Header
            self.assertEqual(reader[0][0], "modbusRTU")
            self.assertEqual(reader[0][1], "Meter")
            self.assertEqual(reader[0][2], "CHINT")
            self.assertEqual(reader[0][3], "DTSU666")

            # Check register count
            rows = reader[1:]
            self.assertGreaterEqual(len(rows), 11)

            # Find row 8192 (Uab)
            uab_row = next((r for r in rows if r[2] == "8192"), None)
            self.assertIsNotNone(uab_row)
            assert uab_row is not None
            self.assertEqual(uab_row[3], "F32")
            self.assertEqual(uab_row[7], "0.100000")
            self.assertEqual(uab_row[9], "V")
            self.assertEqual(uab_row[10], "4")

            # Find row 8204 (Ia)
            ia_row = next((r for r in rows if r[2] == "8204"), None)
            self.assertIsNotNone(ia_row)
            assert ia_row is not None
            self.assertEqual(ia_row[3], "F32")
            self.assertEqual(ia_row[7], "0.001000")
            self.assertEqual(ia_row[9], "A")

            # Find row 8210 (P_AC)
            p_row = next((r for r in rows if r[2] == "8210"), None)
            self.assertIsNotNone(p_row)
            assert p_row is not None
            self.assertEqual(p_row[3], "F32")
            self.assertEqual(p_row[7], "0.100000")
            self.assertEqual(p_row[9], "W")

        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def test_production_and_definitions_dtsu666_compliance(self) -> None:
        """Verifies that Definitions and production DTSU666 CSVs meet all 5 quality pillars."""
        for rel_path in (
            "Definitions/MET_CHINT_DTSU666.csv",
            "production/MET_Chint_DTSU666.csv",
        ):
            path = Path(__file__).parent.parent.parent / rel_path
            self.assertTrue(path.is_file(), f"File {path} does not exist.")

            report = audit_definition_file(str(path), strict_overlap=True)
            self.assertTrue(
                report.is_valid,
                f"File {path} is invalid: errors={report.errors}, warnings={report.warnings}",
            )
            self.assertEqual(len(report.errors), 0)
            self.assertEqual(len(report.warnings), 0)

            # Strict UTF-8 without BOM check
            raw_bytes = path.read_bytes()
            self.assertFalse(
                raw_bytes.startswith(b"\xef\xbb\xbf"),
                f"File {path} must be strictly UTF-8 without BOM.",
            )


if __name__ == "__main__":
    unittest.main()
