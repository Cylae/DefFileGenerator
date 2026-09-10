import io
import unittest

from fastapi.testclient import TestClient

from web.app import app


class TestWebBackend(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_health_check_endpoint(self):
        response = self.client.get("/api/health")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["status"], "ok")

    def test_root_index_endpoint(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn("WebdynSunPM", response.text)

    def test_convert_valid_csv_upload(self):
        csv_content = (
            "Register,Name,Data Type,Unit,Scale,Access\n"
            "40001,Active Power,uint16,W,1,R\n"
            "40002,DC Voltage,uint16,V,0.1,R\n"
        )
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        response = self.client.post(
            "/api/convert",
            files={"file": ("sample.csv", file_obj, "text/csv")},
            data={
                "manufacturer": "WebdynMfg",
                "model": "WebdynModel",
                "protocol": "modbusRTU",
                "category": "Inverter",
                "address_offset": "0",
            },
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertEqual(data["register_count"], 2)
        self.assertIn("WebdynMfg;WebdynModel", data["csv_content"])

    def test_convert_path_traversal_filename(self):
        csv_content = "Register,Name,Data Type,Unit,Scale,Access\n40001,Active Power,uint16,W,1,R\n"
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        # Upload with malicious path traversal in filename
        response = self.client.post(
            "/api/convert",
            files={"file": ("../../../../etc/passwd.csv", file_obj, "text/csv")},
            data={"manufacturer": "TestMfg", "model": "TestModel"},
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()["success"])

    def test_validate_path_traversal_filename(self):
        def_content = (
            "modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n"
            "1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4\n"
        )
        file_obj = io.BytesIO(def_content.encode("utf-8"))
        response = self.client.post(
            "/api/validate",
            files={"file": ("../../../tmp/evil.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 200)
        # Should sanitize filename and return basename
        self.assertEqual(response.json()["filename"], "evil.csv")
        self.assertTrue(response.json()["valid"])

    def test_convert_unsupported_file_extension(self):
        file_obj = io.BytesIO(b"dummy binary data")
        response = self.client.post(
            "/api/convert",
            files={"file": ("unsupported.exe", file_obj, "application/octet-stream")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Unsupported file format", response.json()["detail"])

    def test_convert_corrupt_file_raises_400(self):
        # Corrupt file content that cannot be parsed as registers
        file_obj = io.BytesIO(b"corrupt non-csv data without registers")
        response = self.client.post(
            "/api/convert",
            files={"file": ("corrupt.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 400)

    def test_validate_definition_endpoint(self):
        def_content = (
            "modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n"
            "1;3;40001;U16;;Active Power;active_power;1.000000;0.000000;W;4\n"
        )
        file_obj = io.BytesIO(def_content.encode("utf-8"))
        response = self.client.post(
            "/api/validate",
            files={"file": ("definition.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()["valid"])

    def test_validate_invalid_csv_definition(self):
        def_content = "Header line\n1;3;-50;U16;;Name;var_name;1.0;0.0;V;4\n"
        file_obj = io.BytesIO(def_content.encode("utf-8"))
        response = self.client.post(
            "/api/validate",
            files={"file": ("invalid_def.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 200)
        self.assertFalse(response.json()["valid"])


if __name__ == "__main__":
    unittest.main()
