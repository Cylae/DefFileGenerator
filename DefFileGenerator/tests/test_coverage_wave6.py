import io
import logging
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

from DefFileGenerator.main import main
from web.app import app

client = TestClient(app)


class TestCoverageWave6(unittest.TestCase):
    def setUp(self):
        logging.disable(logging.CRITICAL)

    def tearDown(self):
        logging.disable(logging.NOTSET)

    def test_cli_unexpected_exception(self):
        with patch("DefFileGenerator.main._run_cli", side_effect=Exception("Root level exception")):
            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 1)

    def test_web_missing_filename_fastapi_injection(self):
        csv_content = "Register,Name,Data Type,Unit,Scale,Access\n40001,Active Power,uint16,W,1,R\n"
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        response = client.post(
            "/api/convert",
            files={"file": (None, file_obj, "text/csv")},
            data={"manufacturer": "TestMfg", "model": "TestModel"},
        )
        self.assertIn(response.status_code, [400, 422])

    def test_web_static_files(self):
        response = client.get("/static/app.js")
        self.assertEqual(response.status_code, 200)

    def test_web_health(self):
        response = client.get("/api/health")
        self.assertEqual(response.status_code, 200)

    def test_web_root(self):
        response = client.get("/")
        self.assertIn(response.status_code, [200, 404])

    def test_web_convert_exception_in_generate(self):
        csv_content = "Register,Name,Data Type,Unit,Scale,Access\n40001,Active Power,uint16,W,1,R\n"
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        with patch("web.app.run_generator", side_effect=Exception("Test Exception")):
            response = client.post(
                "/api/convert",
                files={"file": ("test.csv", file_obj, "text/csv")},
                data={"manufacturer": "TestMfg", "model": "TestModel"},
            )
            self.assertEqual(response.status_code, 400)
            self.assertIn("Core generator processing failed", response.json()["detail"])


if __name__ == "__main__":
    unittest.main()
