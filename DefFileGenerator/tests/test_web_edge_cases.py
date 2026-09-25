import io
import unittest
from unittest.mock import patch

from fastapi.testclient import TestClient

import web.app as web_app
from web.app import app


class TestWebEdgeCases(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_upload_exceeds_max_upload_bytes(self):
        # Create bytes larger than MAX_UPLOAD_BYTES (10 MiB)
        oversized_data = b"0" * (web_app.MAX_UPLOAD_BYTES + 1024)
        file_obj = io.BytesIO(oversized_data)
        response = self.client.post(
            "/api/convert",
            files={"file": ("oversized.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 413)
        self.assertIn("exceeds the 10 MiB limit", response.json()["detail"])

    def test_convert_empty_filename(self):
        file_obj = io.BytesIO(b"data")
        response = self.client.post(
            "/api/convert",
            files={"file": ("", file_obj, "text/csv")},
        )
        self.assertIn(response.status_code, (400, 422))

    def test_convert_invalid_dot_filename(self):
        file_obj = io.BytesIO(b"data")
        response = self.client.post(
            "/api/convert",
            files={"file": (".", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Invalid filename provided", response.json()["detail"])

    def test_validate_empty_filename(self):
        file_obj = io.BytesIO(b"data")
        response = self.client.post(
            "/api/validate",
            files={"file": ("", file_obj, "text/csv")},
        )
        self.assertIn(response.status_code, (400, 422))

    def test_validate_invalid_dot_filename(self):
        file_obj = io.BytesIO(b"data")
        response = self.client.post(
            "/api/validate",
            files={"file": ("..", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Invalid filename provided", response.json()["detail"])

    def test_validate_unsupported_extension(self):
        file_obj = io.BytesIO(b"data")
        response = self.client.post(
            "/api/validate",
            files={"file": ("doc.pdf", file_obj, "application/pdf")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Unsupported file format", response.json()["detail"])

    def test_convert_exceeds_max_web_registers(self):
        csv_content = "Register,Name,Data Type\n40001,Power,U16\n"
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        fake_mapped = [{"Address": f"{i}", "Name": f"V_{i}", "Type": "U16"} for i in range(65537)]
        with patch(
            "DefFileGenerator.extractor.Extractor.map_and_clean", return_value=iter(fake_mapped)
        ):
            response = self.client.post(
                "/api/convert",
                files={"file": ("large.csv", file_obj, "text/csv")},
            )
            self.assertEqual(response.status_code, 400)
            self.assertIn("exceeding the maximum allowable limit", response.json()["detail"])

    def test_convert_missing_output_file_returns_500(self):
        csv_content = "Register,Name,Data Type\n40001,Power,U16\n"
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        with patch("web.app.run_generator"):  # No-op so output file is not written
            response = self.client.post(
                "/api/convert",
                files={"file": ("sample.csv", file_obj, "text/csv")},
            )
            self.assertEqual(response.status_code, 500)
            self.assertIn("Failed to generate output CSV definition", response.json()["detail"])

    def test_convert_extraction_exception_returns_400(self):
        file_obj = io.BytesIO(b"dummy data")
        with patch(
            "DefFileGenerator.extractor.Extractor.extract_from_csv",
            side_effect=RuntimeError("Extraction explosion"),
        ):
            response = self.client.post(
                "/api/convert",
                files={"file": ("sample.csv", file_obj, "text/csv")},
            )
            self.assertEqual(response.status_code, 400)
            self.assertIn(
                "Failed to extract registers from uploaded file", response.json()["detail"]
            )

    def test_validate_exception_returns_400(self):
        file_obj = io.BytesIO(b"Header\n1;3;40001;U16;;N;t;1;0;V;4\n")
        with patch(
            "DefFileGenerator.def_gen.Generator.validate_csv_detailed",
            side_effect=RuntimeError("Validation crash"),
        ):
            response = self.client.post(
                "/api/validate",
                files={"file": ("sample.csv", file_obj, "text/csv")},
            )
            self.assertEqual(response.status_code, 400)
            self.assertIn("Failed to validate uploaded definition CSV", response.json()["detail"])

    def test_read_root_when_index_html_missing(self):
        orig_exists = web_app.os.path.exists

        def fake_exists(path):
            if str(path).endswith("index.html"):
                return False
            return orig_exists(path)

        with patch("web.app.os.path.exists", side_effect=fake_exists):
            response = self.client.get("/")
            self.assertEqual(response.status_code, 200)
            self.assertEqual(
                response.json(),
                {"message": "WebdynSunPM API server running. Frontend assets not found."},
            )


if __name__ == "__main__":
    unittest.main()
