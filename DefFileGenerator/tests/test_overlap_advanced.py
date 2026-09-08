"""Batch 4: Advanced address overlap, validate_csv, and write_output_csv tests."""

import csv
import io
import logging
import os
import sys
import tempfile
import unittest

from DefFileGenerator.def_gen import Generator


class TestOverlapDetection(unittest.TestCase):
    """_check_address_overlap edge cases."""

    def setUp(self):
        self.g = Generator()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def _process(self, rows):
        return list(self.g.process_rows(rows))

    def test_no_overlap_adjacent(self):
        rows = [
            {"Name": "A", "Address": "100", "Type": "U16"},
            {"Name": "B", "Address": "101", "Type": "U16"},
        ]
        self.assertEqual(len(self._process(rows)), 2)

    def test_overlap_u32_and_u16(self):
        """U32 occupies 2 registers; placing a U16 on the 2nd one should warn."""
        logging.disable(logging.NOTSET)
        rows = [
            {"Name": "P", "Tag": "p", "Address": "100", "Type": "U32"},
            {"Name": "Q", "Tag": "q", "Address": "101", "Type": "U16"},
        ]
        with self.assertLogs(level="WARNING") as cm:
            result = self._process(rows)
        self.assertTrue(any("overlap" in m.lower() for m in cm.output))
        # Both rows are still emitted
        self.assertEqual(len(result), 2)

    def test_bits_same_base_address_no_overlap(self):
        """Multiple BITS registers at the same base address do NOT overlap."""
        rows = [
            {"Name": "Bit0", "Address": "100_0_1", "Type": "BITS"},
            {"Name": "Bit1", "Address": "100_1_1", "Type": "BITS"},
        ]
        result = self._process(rows)
        self.assertEqual(len(result), 2)

    def test_different_register_types_no_cross_type_overlap(self):
        """Holding (info1=3) and Input (info1=4) at same address do not overlap."""
        rows = [
            {"Name": "H", "Address": "100", "Type": "U16", "RegisterType": "Holding Register"},
            {"Name": "I", "Address": "100", "Type": "U16", "RegisterType": "Input Register"},
        ]
        result = self._process(rows)
        self.assertEqual(len(result), 2)

    def test_u64_occupies_four_registers(self):
        """A U64 at 100 blocks 100-103."""
        logging.disable(logging.NOTSET)
        rows = [
            {"Name": "BigVal", "Tag": "bv", "Address": "100", "Type": "U64"},
            {"Name": "SmallVal", "Tag": "sv", "Address": "102", "Type": "U16"},
        ]
        with self.assertLogs(level="WARNING") as cm:
            result = self._process(rows)
        self.assertTrue(any("overlap" in m.lower() for m in cm.output))

    def test_mac_type_occupies_three_registers(self):
        """MAC occupies 3 registers."""
        self.assertEqual(self.g.get_register_count("MAC", "100"), 3)

    def test_ipv6_type_occupies_eight_registers(self):
        """IPV6 occupies 8 registers."""
        self.assertEqual(self.g.get_register_count("IPV6", "100"), 8)


class TestValidateCsvAdvanced(unittest.TestCase):
    """validate_csv edge cases."""

    def setUp(self):
        self.g = Generator()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _write(self, rows, header=None):
        p = os.path.join(self.tmpdir.name, "v.csv")
        if header is None:
            header = ["modbusRTU", "Inverter", "M", "X", "", "", "", "", "", "", ""]
        with open(p, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header)
            for i, row in enumerate(rows, start=1):
                writer.writerow([str(i)] + list(row))
        return p

    def test_bad_header_fails(self):
        p = os.path.join(self.tmpdir.name, "badheader.csv")
        with open(p, "w", encoding="utf-8") as f:
            f.write("")
        self.assertFalse(self.g.validate_csv(p))

    def test_file_not_found(self):
        self.assertFalse(self.g.validate_csv("/no/such/file.csv"))

    def test_insufficient_columns_row_warns_not_fatal(self):
        """A row with < 11 columns emits a warning but does not set valid=False."""
        rows = [
            ["3", "100", "U16", ""],  # too short
            ["3", "101", "U16", "", "OK", "ok_tag", "1.0", "0.0", "V", "4"],
        ]
        p = self._write(rows)
        # Insufficient columns are warnings, not fatal; valid row exists -> True
        result = self.g.validate_csv(p, strict=False)
        self.assertTrue(result)

    def test_invalid_type_makes_invalid(self):
        rows = [["3", "100", "GARBAGE", "", "V", "v_tag", "1.0", "0.0", "V", "4"]]
        p = self._write(rows)
        self.assertFalse(self.g.validate_csv(p))

    def test_strict_overlap_returns_false(self):
        rows = [
            ["3", "100", "U32", "", "A", "a_tag", "1.0", "0.0", "W", "4"],
            ["3", "101", "U16", "", "B", "b_tag", "1.0", "0.0", "W", "4"],
        ]
        p = self._write(rows)
        self.assertFalse(self.g.validate_csv(p, strict=True))

    def test_lenient_overlap_returns_true(self):
        rows = [
            ["3", "100", "U32", "", "A", "a_tag", "1.0", "0.0", "W", "4"],
            ["3", "101", "U16", "", "B", "b_tag", "1.0", "0.0", "W", "4"],
        ]
        p = self._write(rows)
        self.assertTrue(self.g.validate_csv(p, strict=False))

    def test_nonexistent_input_file_returns_false(self):
        self.assertFalse(self.g.validate_csv("/totally/nonexistent.csv"))

    def test_duplicate_tag_returns_false(self):
        rows = [
            ["3", "100", "U16", "", "V1", "dup", "1.0", "0.0", "V", "4"],
            ["3", "101", "U16", "", "V2", "dup", "1.0", "0.0", "V", "4"],
        ]
        p = self._write(rows)
        self.assertFalse(self.g.validate_csv(p))

    def test_empty_tag_not_counted_as_duplicate(self):
        """Rows with empty tags should not trigger duplicate-tag errors."""
        rows = [
            ["3", "100", "U16", "", "V1", "", "1.0", "0.0", "V", "4"],
            ["3", "101", "U16", "", "V2", "", "1.0", "0.0", "V", "4"],
        ]
        p = self._write(rows)
        self.assertTrue(self.g.validate_csv(p, strict=False))

    def test_utf16_bom_file(self):
        """validate_csv must handle UTF-16 BOM encoded definition files."""
        p = os.path.join(self.tmpdir.name, "utf16.csv")
        header = ["modbusRTU", "Inverter", "M", "X", "", "", "", "", "", "", ""]
        rows_data = [["1", "3", "100", "U16", "", "Var", "v_tag", "1.0", "0.0", "V", "4"]]
        with open(p, "w", newline="", encoding="utf-16") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(header)
            for r in rows_data:
                writer.writerow(r)
        self.assertTrue(self.g.validate_csv(p, strict=False))


class TestWriteOutputCsv(unittest.TestCase):
    """write_output_csv correctness beyond the atomicity tests."""

    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def _rows(self, n=1):
        return [
            {
                "Info1": "3", "Info2": str(100 + i), "Info3": "U16",
                "Info4": "", "Name": f"Var{i}", "Tag": f"var{i}",
                "CoefA": "1.000000", "CoefB": "0.000000", "Unit": "V", "Action": "4",
            }
            for i in range(n)
        ]

    def test_writes_header_and_data(self):
        out = os.path.join(self.tmpdir.name, "out.csv")
        Generator.write_output_csv(out, iter(self._rows(3)), "M", "X")
        with open(out, encoding="utf-8-sig") as f:
            content = f.read()
        self.assertIn("modbusRTU", content)
        self.assertIn("Var0", content)
        self.assertIn("Var2", content)

    def test_writes_to_stdout_when_none(self):
        buf = io.StringIO()
        captured = []
        import sys
        old = sys.stdout
        sys.stdout = buf
        try:
            Generator.write_output_csv(None, iter(self._rows(2)), "A", "B")
        finally:
            sys.stdout = old
        content = buf.getvalue()
        self.assertIn("modbusRTU", content)

    def test_writes_to_file_object(self):
        buf = io.StringIO()
        Generator.write_output_csv(buf, iter(self._rows(1)), "P", "Q")
        content = buf.getvalue()
        self.assertIn("modbusRTU", content)
        self.assertIn("Var0", content)

    def test_type_counts_logged_per_info1(self):
        rows = [
            {"Info1": "3", "Info2": "1", "Info3": "U16", "Info4": "",
             "Name": "H", "Tag": "h", "CoefA": "1.0", "CoefB": "0.0", "Unit": "", "Action": "4"},
            {"Info1": "4", "Info2": "2", "Info3": "U16", "Info4": "",
             "Name": "I", "Tag": "i", "CoefA": "1.0", "CoefB": "0.0", "Unit": "", "Action": "4"},
        ]
        out = os.path.join(self.tmpdir.name, "count.csv")
        # Should not raise
        Generator.write_output_csv(out, iter(rows), "M", "X")

    def test_custom_protocol_and_category(self):
        out = os.path.join(self.tmpdir.name, "custom.csv")
        Generator.write_output_csv(out, iter(self._rows(1)), "M", "X",
                                   protocol="modbustcp", category="Battery")
        with open(out, encoding="utf-8-sig") as f:
            first_line = f.readline()
        self.assertIn("modbustcp", first_line)
        self.assertIn("Battery", first_line)

    def test_forced_write_field_written(self):
        out = os.path.join(self.tmpdir.name, "fw.csv")
        Generator.write_output_csv(out, iter(self._rows(1)), "M", "X", forced_write="1;2")
        with open(out, encoding="utf-8-sig") as f:
            first_line = f.readline()
        self.assertIn("1", first_line)

    def test_index_increments_correctly(self):
        out = os.path.join(self.tmpdir.name, "idx.csv")
        Generator.write_output_csv(out, iter(self._rows(3)), "M", "X")
        with open(out, encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            rows = list(reader)
        # Header on row[0], data starts at row[1]
        self.assertEqual(rows[1][0], "1")
        self.assertEqual(rows[2][0], "2")
        self.assertEqual(rows[3][0], "3")


if __name__ == "__main__":
    unittest.main()