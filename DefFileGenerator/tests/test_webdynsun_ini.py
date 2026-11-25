#!/usr/bin/env python3
"""
Tests for WebdynSun Classic (.ini) conversion, validation, and FoxESS
<Manufacturer>/<Range>/ repository layout.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from DefFileGenerator.webdynsun_ini import (
    convert_sunpm_csv_to_webdynsun_ini,
    render_webdynsun_ini_from_rows,
    verify_webdynsun_ini,
)

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
PROD_DIR = REPO_ROOT / "production"


class TestWebdynSunIni(unittest.TestCase):
    """Validates WebdynSun .ini generation and FoxESS dual-target layout."""

    def test_types_signedness_and_varposition_mapping(self) -> None:
        """Verify U16/I32/F32_W/STRING mapping, signedness, and 1-based varPosition."""
        rows = [
            ["1", "3", "39000", "U16", "", "Model ID", "MODEL", "1.000000", "0.000000", "", "4"],
            ["2", "3", "39001", "I32", "", "Rated Power", "P_NOM", "0.001000", "0.000000", "kW", "4"],
            ["3", "3", "39003", "F32_W", "", "Freq", "FREQ", "1.000000", "0.000000", "Hz", "9"],
            ["4", "3", "46001", "U16", "", "Power Limit", "CMD_PLIM", "1.000000", "0.000000", "%", "1"],
            ["5", "3", "30000_32", "STRING", "", "Serial Number", "SERIAL", "1.000000", "0.000000", "", "4"],
        ]
        ini_text, skipped = render_webdynsun_ini_from_rows(rows)
        self.assertEqual(skipped, [])
        self.assertIn("Modbus_RequestsTables={", ini_text)
        self.assertIn("Modbus_VariablesTables={", ini_text)
        # U16 (addr 39000) pos 1, I32 (addr 39001) pos 2, F32_W (addr 39003) pos 4, alarm 9 -> 8
        self.assertIn("1;1;Model ID;3;2;1;;;1;0;;4", ini_text)
        self.assertIn("2;1;Rated Power;5;1;2;;;0.001;0;kW;4", ini_text)
        self.assertIn("3;1;Freq;8;1;4;;;1;0;Hz;8", ini_text)
        # Writable register gets write FC 6 and enableWriting=1
        self.assertIn(";3;6;46001;1;1;1;;", ini_text)
        self.assertIn(";Power Limit;3;2;1;;;1;0;%;1", ini_text)
        # String uses varType 9 and 16 registers
        self.assertIn(";3;0;30000;16;1;2;;", ini_text)
        self.assertIn(";Serial Number;9;2;1;;;1;0;;4", ini_text)

    def test_64bit_not_silently_truncated(self) -> None:
        """Verify U64/I64/F64 are explicitly documented as skipped rather than truncated."""
        rows = [
            ["1", "3", "100", "U64", "", "Total Energy 64", "E_TOT", "1.000000", "0.000000", "kWh", "4"],
        ]
        ini_text, skipped = render_webdynsun_ini_from_rows(rows)
        self.assertEqual(len(skipped), 1)
        self.assertIn("# NOTE: Skipped unsupported WebdynSun type/address: Total Energy 64", ini_text)

    def test_verify_rejects_bad_ini(self) -> None:
        """Verify validator catches malformed varType, out-of-range registers, and BOM."""
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "bad.ini"
            bad.write_bytes(
                b"\xef\xbb\xbfModbus_RequestsTables={\n1;Req1;3;0;70000;32;1;2;;\n}\n"
                b"Modbus_VariablesTables={\n1;99;Var;99;2;0;;;1;0;;4\n}\n"
            )
            ok, issues = verify_webdynsun_ini(bad)
            self.assertFalse(ok)
            self.assertGreaterEqual(len(issues), 4)

    def test_convert_csv_to_ini_roundtrip(self) -> None:
        """Verify atomic CSV -> INI export passes verify_webdynsun_ini with CRLF and no BOM."""
        with tempfile.TemporaryDirectory() as tmp:
            csv_p = Path(tmp) / "INV_Demo_Test.csv"
            ini_p = Path(tmp) / "INV_Demo_Test.ini"
            csv_p.write_text(
                "Modbus;Inverter;Demo;Test;;;;;;;\r\n"
                "1;3;100;U16;;Voltage;U_AC_1;0.100000;0.000000;V;4\r\n",
                encoding="utf-8",
                newline="",
            )
            ok_conv, skipped = convert_sunpm_csv_to_webdynsun_ini(csv_p, ini_p)
            self.assertTrue(ok_conv)
            self.assertEqual(skipped, [])
            ok_val, issues = verify_webdynsun_ini(ini_p)
            self.assertTrue(ok_val, issues)
            raw = ini_p.read_bytes()
            self.assertFalse(raw.startswith(b"\xef\xbb\xbf"))
            self.assertIn(b"\r\n", raw)

    def test_foxess_ranges_dual_format_and_sources(self) -> None:
        """Verify every FoxESS range has WebdynSunPM .csv, WebdynSun .ini, and sources/."""
        fox_root = PROD_DIR / "FoxESS"
        self.assertTrue(fox_root.is_dir())
        expected_ranges = {"P3_H3_Pro", "P3_Pro", "H3_Pro", "G_MAX", "R_Series", "R_V_H3plus"}
        actual_ranges = {d.name for d in fox_root.iterdir() if d.is_dir()}
        self.assertTrue(expected_ranges.issubset(actual_ranges), actual_ranges)
        for rng in expected_ranges:
            rng_dir = fox_root / rng
            csvs = list(rng_dir.glob("*.csv"))
            self.assertGreater(len(csvs), 0, f"No CSVs in FoxESS/{rng}")
            self.assertTrue((rng_dir / "sources" / "README.md").is_file(), f"Missing sources README in {rng}")
            for c in csvs:
                ini = c.with_suffix(".ini")
                self.assertTrue(ini.is_file(), f"Missing .ini for {c}")
                ok, issues = verify_webdynsun_ini(ini)
                self.assertTrue(ok, f"{ini}: {issues}")


if __name__ == "__main__":
    unittest.main()
