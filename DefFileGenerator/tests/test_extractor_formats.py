import json
import os
import tempfile
import unittest

from DefFileGenerator.extractor import Extractor


class TestExtractorNewFormats(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.extractor = Extractor()

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_json_extraction_direct_array(self):
        data = [
            {"Address": 40001, "Name": "Active Power", "Type": "U16", "Unit": "W"},
            {"Address": 40002, "Name": "Reactive Power", "Type": "I16", "Unit": "var"},
        ]
        p = os.path.join(self.tmpdir.name, "registers.json")
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f)

        tables = self.extractor.extract_from_json(p)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Name"], "Active Power")
        self.assertEqual(rows[0]["Address"], "40001")

    def test_json_extraction_nested_object(self):
        data = {
            "device": "Inverter",
            "modbus_table": [
                {"Address": "0x0001", "Name": "DC Voltage", "Type": "U16", "Unit": "V"},
                {"Address": "0x0002", "Name": "DC Current", "Type": "U16", "Unit": "A"},
            ],
        }
        p = os.path.join(self.tmpdir.name, "nested.json")
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f)

        tables = self.extractor.extract_from_json(p)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Name"], "DC Voltage")
        self.assertEqual(rows[0]["Address"], "1")

    def test_html_table_extraction(self):
        html_content = """
        <html>
            <body>
                <h2>Modbus Register Map</h2>
                <table>
                    <thead>
                        <tr><th>Register</th><th>Signal Name</th><th>Data Type</th><th>Unit</th></tr>
                    </thead>
                    <tbody>
                        <tr><td>30001</td><td>Grid Frequency</td><td>U16</td><td>Hz</td></tr>
                        <tr><td>30002</td><td>Grid Voltage</td><td>U16</td><td>V</td></tr>
                    </tbody>
                </table>
            </body>
        </html>
        """
        p = os.path.join(self.tmpdir.name, "doc.html")
        with open(p, "w", encoding="utf-8") as f:
            f.write(html_content)

        tables = self.extractor.extract_from_html(p)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Name"], "Grid Frequency")
        self.assertEqual(rows[0]["Address"], "30001")

    def test_tsv_and_markdown_text_extraction(self):
        # TSV format test
        tsv_content = "Address\tName\tType\tUnit\n100\tStatus\tU16\t\n101\tError Code\tU16\t"
        p_tsv = os.path.join(self.tmpdir.name, "registers.tsv")
        with open(p_tsv, "w", encoding="utf-8") as f:
            f.write(tsv_content)

        tables = self.extractor.extract_from_csv(p_tsv)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Name"], "Status")

    def test_extract_auto_sniffing(self):
        # Test unknown extension containing JSON content
        data = [{"Address": 500, "Name": "Temperature", "Type": "I16"}]
        p_unk_json = os.path.join(self.tmpdir.name, "doc.unkdata")
        with open(p_unk_json, "w", encoding="utf-8") as f:
            json.dump(data, f)

        tables = self.extractor.extract_auto(p_unk_json)
        rows = list(self.extractor.map_and_clean(tables))
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Name"], "Temperature")
        self.assertEqual(rows[0]["Address"], "500")

        # Test unknown extension containing HTML table
        html = (
            "<table><tr><th>Addr</th><th>Name</th></tr><tr><td>600</td><td>Power</td></tr></table>"
        )
        p_unk_html = os.path.join(self.tmpdir.name, "doc.table")
        with open(p_unk_html, "w", encoding="utf-8") as f:
            f.write(html)

        tables_html = self.extractor.extract_auto(p_unk_html)
        rows_html = list(self.extractor.map_and_clean(tables_html))
        self.assertEqual(len(rows_html), 1)
        self.assertEqual(rows_html[0]["Name"], "Power")
        self.assertEqual(rows_html[0]["Address"], "600")


if __name__ == "__main__":
    unittest.main()
