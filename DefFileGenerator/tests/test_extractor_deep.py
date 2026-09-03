"""Batch 1: Deep tests for Extractor.
"""

import csv
import logging
import os
import tempfile
import unittest


class TestExtractorColumnMapping(unittest.TestCase):

    def setUp(self):
        from DefFileGenerator.extractor import Extractor
        self.ex = Extractor()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_map_and_clean_none_tables(self):
        result = list(self.ex.map_and_clean(None))
        self.assertEqual(result, [])

    def test_map_and_clean_empty_list(self):
        result = list(self.ex.map_and_clean([]))
        self.assertEqual(result, [])

    def test_map_and_clean_table_with_no_rows(self):
        result = list(self.ex.map_and_clean([iter([])]))
        self.assertEqual(result, [])

    def test_map_and_clean_rows_lacking_name_and_address(self):
        raw = [[{"Type": "U16", "Unit": "V"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result, [])

    def test_exact_header_match_address(self):
        raw = [[{"Address": "100", "Name": "Var", "Type": "U16"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Address"], "100")

    def test_synonym_header_description_maps_to_name(self):
        raw = [[{"description": "Output Voltage", "address": "300", "type": "U16"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Name"], "Output Voltage")

    def test_synonym_header_param_maps_to_name(self):
        raw = [[{"parameter": "Grid Freq", "addr": "50", "data type": "U16"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Name"], "Grid Freq")
        self.assertEqual(result[0]["Address"], "50")

    def test_synonym_datatype_maps_to_type(self):
        raw = [[{"Name": "Power", "Address": "500", "datatype": "U32"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Type"], "U32")

    def test_synonym_access_maps_to_action(self):
        raw = [[{"Name": "Param", "Address": "10", "Type": "U16", "access": "R"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Action"], "R")

    def test_synonym_scale_maps_to_factor(self):
        raw = [[{"Name": "Voltage", "Address": "10", "Type": "U16", "scale": "0.1"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Factor"], "0.1")

    def test_synonym_units_maps_to_unit(self):
        raw = [[{"Name": "Freq", "Address": "20", "Type": "U16", "units": "Hz"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Unit"], "Hz")

    def test_address_offset_applied(self):
        raw = [[{"Name": "Var", "Address": "100", "Type": "U16"}]]
        result = list(self.ex.map_and_clean(raw, address_offset=10))
        self.assertEqual(result[0]["Address"], "110")

    def test_address_offset_zero(self):
        raw = [[{"Name": "Var", "Address": "200", "Type": "U16"}]]
        result = list(self.ex.map_and_clean(raw, address_offset=0))
        self.assertEqual(result[0]["Address"], "200")

    def test_address_offset_negative_warns_but_still_works(self):
        logging.disable(logging.NOTSET)
        raw = [[{"Name": "Var", "Address": "5", "Type": "U16"}]]
        with self.assertLogs(level="WARNING"):
            result = list(self.ex.map_and_clean(raw, address_offset=-10))
        self.assertIn("-5", result[0]["Address"])

    def test_missing_register_type_defaults_to_holding(self):
        raw = [[{"Name": "Var", "Address": "100", "Type": "U16"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["RegisterType"], "Holding Register")

    def test_explicit_register_type_preserved(self):
        raw = [[{"Name": "Var", "Address": "100", "Type": "U16", "Register Type": "Input Register"}]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["RegisterType"], "Input Register")

    def test_bits_address_with_start_and_length(self):
        raw = [[{
            "Name": "BitFlag", "Address": "100", "Type": "BITS",
            "StartBit": "3", "Length": "2",
        }]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Address"], "100_3_2")

    def test_bits_address_default_length_is_one(self):
        raw = [[{
            "Name": "BitFlag", "Address": "100", "Type": "BITS", "StartBit": "7",
        }]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Address"], "100_7_1")

    def test_bits_address_already_has_underscore_not_altered(self):
        raw = [[{
            "Name": "BitFlag", "Address": "100_7_1", "Type": "BITS",
            "StartBit": "7", "Length": "2",
        }]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Address"], "100_7_1")

    def test_string_address_with_length_appended(self):
        raw = [[{
            "Name": "SerialNo", "Address": "200", "Type": "STRING", "Length": "20",
        }]]
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(result[0]["Address"], "200_20")

    def test_explicit_mapping_overrides_auto(self):
        from DefFileGenerator.extractor import Extractor
        ex = Extractor(mapping={"Name": "VariableName", "Address": "RegAddr"})
        raw = [[{"VariableName": "Grid Voltage", "RegAddr": "1234", "datatype": "U16"}]]
        result = list(ex.map_and_clean(raw))
        self.assertEqual(result[0]["Name"], "Grid Voltage")
        self.assertEqual(result[0]["Address"], "1234")

    def test_multiple_tables_all_processed(self):
        table1 = [{"Name": "V1", "Address": "1", "Type": "U16"}]
        table2 = [{"Name": "V2", "Address": "2", "Type": "U32"}]
        result = list(self.ex.map_and_clean([iter(table1), iter(table2)]))
        self.assertEqual(len(result), 2)
        names = {r["Name"] for r in result}
        self.assertIn("V1", names)
        self.assertIn("V2", names)

    def test_empty_table_followed_by_data_table(self):
        result = list(self.ex.map_and_clean([
            iter([]),
            iter([{"Name": "X", "Address": "5", "Type": "U16"}])
        ]))
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["Name"], "X")


class TestExtractorCsvExtraction(unittest.TestCase):

    def setUp(self):
        from DefFileGenerator.extractor import Extractor
        self.ex = Extractor()
        self.tmpdir = tempfile.TemporaryDirectory()
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def _write_csv(self, content, filename="test.csv", encoding="utf-8"):
        path = os.path.join(self.tmpdir.name, filename)
        with open(path, "w", encoding=encoding, newline="") as f:
            f.write(content)
        return path

    def test_extract_comma_delimited(self):
        path = self._write_csv("Address,Name,Type\n100,Voltage,U16\n")
        tables = list(self.ex.extract_from_csv(path))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Voltage")

    def test_extract_semicolon_delimited(self):
        path = self._write_csv("Address;Name;Type\n100;Power;U32\n")
        tables = list(self.ex.extract_from_csv(path))
        rows = list(tables[0])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Power")

    def test_extract_csv_utf8_bom(self):
        path = os.path.join(self.tmpdir.name, "bom.csv")
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write("Address,Name,Type\n100,BomTest,U16\n")
        tables = list(self.ex.extract_from_csv(path))
        rows = list(tables[0])
        self.assertEqual(rows[0]["Name"], "BomTest")

    def test_extract_csv_skips_blank_rows(self):
        path = self._write_csv("Address,Name,Type\n100,V,U16\n   ,   ,   \n200,I,U16\n")
        tables = list(self.ex.extract_from_csv(path))
        rows = list(tables[0])
        self.assertEqual(len(rows), 2)

    def test_extract_csv_missing_file_logs_error(self):
        logging.disable(logging.NOTSET)
        tables = list(self.ex.extract_from_csv("/nonexistent/path.csv"))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_extract_csv_empty_file(self):
        path = self._write_csv("")
        tables = list(self.ex.extract_from_csv(path))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_extract_csv_header_only(self):
        path = self._write_csv("Address,Name,Type\n")
        tables = list(self.ex.extract_from_csv(path))
        rows = list(tables[0])
        self.assertEqual(rows, [])

    def test_extract_csv_and_map_end_to_end(self):
        path = self._write_csv("Address,Name,Type\n100,FreqVar,U16\n101,PowerVar,U32\n")
        raw = self.ex.extract_from_csv(path)
        result = list(self.ex.map_and_clean(raw))
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["Name"], "FreqVar")
        self.assertEqual(result[1]["Name"], "PowerVar")


class TestExtractorNormalizeType(unittest.TestCase):

    def test_normalize_type_via_extractor(self):
        from DefFileGenerator.extractor import Extractor
        ex = Extractor()
        self.assertEqual(ex.normalize_type("float32"), "F32")
        self.assertEqual(ex.normalize_type("uint16"), "U16")
        self.assertEqual(ex.normalize_type(None), "U16")
        self.assertEqual(ex.normalize_type(""), "U16")
        self.assertEqual(ex.normalize_type("string 10"), "STR10")


if __name__ == "__main__":
    unittest.main()