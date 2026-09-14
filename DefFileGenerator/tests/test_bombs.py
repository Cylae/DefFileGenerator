import logging
import os
import tempfile
import unittest
import zipfile

from DefFileGenerator.extractor import Extractor


class TestBombs(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()
        logging.disable(logging.NOTSET)

    def test_excel_zip_bomb(self):
        bomb_path = os.path.join(self.tmpdir.name, "bomb.xlsx")
        with zipfile.ZipFile(bomb_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("docProps/app.xml", b"A" * 1024 * 1024 * 10)  # 10MB of 'A's
            zf.writestr("xl/workbook.xml", b"<workbook></workbook>")
            zf.writestr("[Content_Types].xml", b"<Types></Types>")
            zf.writestr("_rels/.rels", b"<Relationships></Relationships>")

        ex = Extractor()
        raw = list(ex.extract_from_excel(bomb_path))
        self.assertTrue(len(raw) > 0)
        tables = list(raw[0])
        self.assertEqual(len(tables), 0)

    def test_pdf_bomb(self):
        bomb_path = os.path.join(self.tmpdir.name, "bomb.pdf")
        with open(bomb_path, "wb") as f:
            f.write(b"%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
            f.write(b"2 0 obj\n<< /Type /Pages /Count 100000 /Kids [ 3 0 R ] >>\nendobj\n")
            f.write(b"3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n")
            f.write(b"trailer\n<< /Root 1 0 R /Size 4 >>\nstartxref\n0\n%%EOF")

        ex = Extractor()
        raw = list(ex.extract_from_pdf(bomb_path))
        self.assertEqual(len(raw), 0)

    def test_csv_field_size(self):
        bomb_path = os.path.join(self.tmpdir.name, "bomb.csv")
        with open(bomb_path, "wb") as f:
            f.write(b"Name,Address\n")
            f.write(b'"' + b"A" * 131072 + b'"' + b",1000\n")

        ex = Extractor()
        raw = list(ex.extract_from_csv(bomb_path))
        self.assertEqual(len(raw), 1)


if __name__ == "__main__":
    unittest.main()
