"""
Wave 18 Edge Branch Coverage and Hardening Unit Tests.

Targeted test coverage for edge error paths in web/app.py and DefFileGenerator/main.py.
"""

import argparse
import asyncio
import os
import tempfile
import unittest
from unittest.mock import AsyncMock, MagicMock, patch

from fastapi import HTTPException, Request, UploadFile
from fastapi.routing import APIRoute
from fastapi.testclient import TestClient
from starlette.exceptions import HTTPException as StarletteHTTPException
from starlette.formparsers import MultiPartException

from DefFileGenerator.main import (
    _perform_extraction,
    _run_cli,
)
from web.app import (
    MAX_UPLOAD_BYTES,
    UploadBodyLimitMiddleware,
    UploadRoute,
    _save_upload,
    app,
    static_dir,
)


class TestWave18WebEdgeCases(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app, raise_server_exceptions=False)

    def test_upload_body_limit_invalid_content_length(self):
        """Verify middleware gracefully handles non-integer Content-Length header."""
        response = self.client.post(
            "/api/convert",
            headers={"Content-Length": "not_an_integer"},
            files={"file": ("test.csv", b"Name,Address\nVoltage,30001\n", "text/csv")},
            data={"manufacturer": "Test", "model": "Model"},
        )
        self.assertIn(response.status_code, (200, 400, 413, 422))

    def test_upload_body_limit_multipart_exception_propagate(self):
        """Verify MultiPartException re-raises when total bytes did not exceed maximum."""

        async def dummy_app(scope, receive, send):
            raise MultiPartException("Stream truncated unexpectedly")

        middleware = UploadBodyLimitMiddleware(dummy_app)
        scope = {
            "type": "http",
            "method": "POST",
            "path": "/api/convert",
            "headers": [(b"content-length", b"100")],
        }

        async def receive():
            return {"type": "http.request", "body": b"data"}

        async def send(message):
            pass

        with self.assertRaises(MultiPartException):
            asyncio.run(middleware(scope, receive, send))

    def test_upload_route_other_starlette_http_exception(self):
        """Verify UploadRoute propagates StarletteHTTPException when detail is not 'Too many files'."""
        route = UploadRoute("/test", MagicMock())

        async def custom_handler(request):
            raise StarletteHTTPException(status_code=400, detail="Custom bad request")

        async def dummy_receive():
            return {"type": "http.request", "body": b""}

        with patch.object(APIRoute, "get_route_handler", return_value=custom_handler):
            handler = route.get_route_handler()

            request = Request(
                {
                    "type": "http",
                    "method": "POST",
                    "path": "/api/convert",
                    "headers": [
                        (b"content-type", b"multipart/form-data; boundary=----WebKitFormBoundary")
                    ],
                },
                receive=dummy_receive,
            )
            with self.assertRaises(StarletteHTTPException) as ctx:
                asyncio.run(handler(request))
            self.assertEqual(ctx.exception.detail, "Custom bad request")

    def test_save_upload_exceeds_max_bytes(self):
        """Verify _save_upload raises 413 HTTPException when file chunking exceeds limit."""

        async def oversized_read(size):
            return b"A" * (MAX_UPLOAD_BYTES + 1024)

        mock_file = MagicMock(spec=UploadFile)
        mock_file.read = oversized_read
        mock_file.close = AsyncMock()

        with tempfile.TemporaryDirectory() as tmp_dir:
            dest = os.path.join(tmp_dir, "out.bin")
            with self.assertRaises(HTTPException) as ctx:
                asyncio.run(_save_upload(mock_file, dest))
            self.assertEqual(ctx.exception.status_code, 413)

    def test_convert_document_csv_validation_unexpected_exception(self):
        """Verify convert_file handles unexpected exception during CSV validation."""
        csv_payload = "Name,Address,Type,RegisterType\nVoltage,30001,U16,Holding\n"

        call_count = 0

        def mock_validate(output_path, *args, **kwargs):
            nonlocal call_count
            call_count += 1
            if call_count > 1:
                raise RuntimeError("Unexpected CSV validation error")
            from DefFileGenerator.def_gen import ValidationReport

            return ValidationReport(is_valid=True, register_count=1, issues=[])

        with patch("web.app.Generator.validate_csv_detailed", side_effect=mock_validate):
            response = self.client.post(
                "/api/convert",
                files={"file": ("test.csv", csv_payload.encode("utf-8"), "text/csv")},
                data={"manufacturer": "Test", "model": "M1"},
            )
            self.assertEqual(response.status_code, 400)
            self.assertIn("Failed to validate generated definition CSV.", response.json()["detail"])

    def test_validate_document_unexpected_exception(self):
        """Verify validate_file handles unexpected exception during CSV validation."""
        csv_payload = "modbus;1;1;1;1;1\n"
        with patch(
            "web.app.Generator.validate_csv_detailed",
            side_effect=RuntimeError("Unexpected CSV validation error"),
        ):
            response = self.client.post(
                "/api/validate",
                files={"file": ("test.csv", csv_payload.encode("utf-8"), "text/csv")},
            )
            self.assertEqual(response.status_code, 400)
            self.assertIn("Failed to validate uploaded definition CSV.", response.json()["detail"])

    def test_root_endpoint_index_html_missing(self):
        """Verify / root endpoint returns fallback JSON when index.html is missing."""
        index_path = os.path.join(static_dir, "index.html")
        with patch(
            "os.path.exists",
            side_effect=lambda path: False if path == index_path else os.path.exists(path),
        ):
            response = self.client.get("/")
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertIn("WebdynSunPM API server running", data.get("message", ""))


class TestWave18CLIEdgeCases(unittest.TestCase):
    def test_perform_extraction_corrupted_mapping_json(self):
        """Verify _perform_extraction exits on corrupted or unreadable mapping JSON file."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_csv = os.path.join(tmp_dir, "input.csv")
            with open(input_csv, "w", encoding="utf-8") as f:
                f.write("Name,Address\nVoltage,30001\n")

            bad_mapping = os.path.join(tmp_dir, "bad_mapping.json")
            with open(bad_mapping, "w", encoding="utf-8") as f:
                f.write("{invalid json syntax")

            args = argparse.Namespace(
                input_file=input_csv,
                mapping=bad_mapping,
                address_offset=0,
                pages=None,
                sheet=None,
            )
            with self.assertRaises(SystemExit) as ctx:
                _perform_extraction(args)
            self.assertEqual(ctx.exception.code, 1)

    def test_cli_pages_and_sheet_warnings_for_mismatched_filetypes(self):
        """Verify _run_cli logs warnings when --pages is used for CSV or --sheet is used for PDF."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            input_csv = os.path.join(tmp_dir, "test.csv")
            with open(input_csv, "w", encoding="utf-8") as f:
                f.write("Name,Address\nVoltage,30001\n")

            out_csv = os.path.join(tmp_dir, "out.csv")

            with patch("logging.warning") as mock_warn:
                try:
                    _run_cli(
                        [
                            "extract",
                            input_csv,
                            "-o",
                            out_csv,
                            "--pages",
                            "1-2",
                            "--sheet",
                            "Sheet1",
                            "--force",
                        ]
                    )
                except SystemExit:
                    pass

                warn_messages = [call.args[0] for call in mock_warn.call_args_list]
                self.assertTrue(any("--pages is only applicable" in msg for msg in warn_messages))
                self.assertTrue(any("--sheet is only applicable" in msg for msg in warn_messages))


if __name__ == "__main__":
    unittest.main()
