#!/usr/bin/env python3
"""
Extended test suite for Web API endpoints:
- POST /api/tag
- POST /api/audit
- GET /api/library
- GET /api/library/{filename}
- Edge case branches in web/app.py
"""

from __future__ import annotations

import io
import os

import pytest
from fastapi.testclient import TestClient

from web.app import app


@pytest.fixture
def client() -> TestClient:
    return TestClient(app)


class TestWebExtendedEndpoints:
    """Tests the new SunPM tagging, audit, and library API endpoints."""

    def test_get_library_catalog(self, client: TestClient) -> None:
        response = client.get("/api/library")
        assert response.status_code == 200
        catalog = response.json()
        assert isinstance(catalog, list)
        assert len(catalog) >= 100  # We have 125 definitions

        item = catalog[0]
        assert "filename" in item
        assert "prefix" in item
        assert "category" in item
        assert "manufacturer" in item
        assert "model" in item
        assert "size_kb" in item

    def test_download_library_file_success(self, client: TestClient) -> None:
        cat_resp = client.get("/api/library")
        catalog = cat_resp.json()
        assert len(catalog) > 0
        first_file = catalog[0]["filename"]

        response = client.get(f"/api/library/{first_file}")
        assert response.status_code == 200
        assert "text/csv" in response.headers.get("content-type", "")
        content = response.text
        assert ";" in content

    def test_download_library_file_not_found(self, client: TestClient) -> None:
        response = client.get("/api/library/non_existent_model_12345.csv")
        assert response.status_code == 404

    def test_download_library_file_invalid_extension(self, client: TestClient) -> None:
        response = client.get("/api/library/malicious.exe")
        assert response.status_code == 400

    def test_tag_file_endpoint_success(self, client: TestClient) -> None:
        csv_data = (
            "modbusRTU;Inverter;SMA;STP5000;;;;;;;\n"
            "1;3;30001;U16;;Active Power;active_power;1.000000;0.000000;W;4\n"
            "2;3;30002;U32;;Total Energy;total_energy;1.000000;0.000000;kWh;4\n"
        )
        files = {"file": ("sma_def.csv", io.BytesIO(csv_data.encode("utf-8")), "text/csv")}
        data = {
            "category": "Inverter",
            "overwrite": "true",
            "normalize_legacy": "true",
        }
        response = client.post("/api/tag", files=files, data=data)
        assert response.status_code == 200
        result = response.json()
        assert result["success"] is True
        assert result["tags_updated"] >= 1
        assert result["updated_count"] == result["tags_updated"]
        assert result["total_registers"] == result["register_count"]
        assert "P_AC" in result["csv_content"] or "E_TOT" in result["csv_content"]
        assert len(result["preview"]) >= 2

    def test_tag_file_invalid_extension(self, client: TestClient) -> None:
        files = {"file": ("manual.pdf", io.BytesIO(b"fake pdf"), "application/pdf")}
        response = client.post("/api/tag", files=files)
        assert response.status_code == 400
        assert "Unsupported file format" in response.json()["detail"]

    def test_audit_file_endpoint_success(self, client: TestClient) -> None:
        csv_data = (
            "modbusRTU;Inverter;SMA;STP5000;;;;;;;\n"
            "1;3;30001;U16;;Active Power;P_AC;1.000000;0.000000;W;4\n"
            "2;3;30002;U32;;Total Energy;E_TOT;1.000000;0.000000;kWh;4\n"
        )
        files = {"file": ("test_def.csv", io.BytesIO(csv_data.encode("utf-8")), "text/csv")}
        data = {
            "strict_overlap": "true",
            "auto_fix": "false",
        }
        response = client.post("/api/audit", files=files, data=data)
        assert response.status_code == 200
        result = response.json()
        assert result["success"] is True
        assert result["valid"] is True
        assert result["is_valid"] is True
        assert result["register_count"] == 2
        assert result["error_count"] == len(result["errors"])
        assert result["warning_count"] == len(result["warnings"])
        assert result["info_count"] == len(result["infos"])
        assert isinstance(result["issues"], list)
        assert isinstance(result["errors"], list)
        assert isinstance(result["warnings"], list)

    def test_audit_file_with_auto_fix(self, client: TestClient) -> None:
        csv_data = (
            "modbusRTU;Inverter;SMA;STP5000;;;;;;;\n"
            "1;3;30001;U16;;Active Power;RealPower;1.000000;0.000000;W;4\n"
        )
        files = {"file": ("legacy_def.csv", io.BytesIO(csv_data.encode("utf-8")), "text/csv")}
        data = {
            "auto_fix": "true",
        }
        response = client.post("/api/audit", files=files, data=data)
        assert response.status_code == 200
        result = response.json()
        assert "fix_applied" in result
        assert "csv_content" in result
        assert result["fixed_csv_content"] == result["csv_content"]
        assert "P_AC" in result["csv_content"]

    def test_audit_file_invalid_extension(self, client: TestClient) -> None:
        files = {"file": ("data.xlsx", io.BytesIO(b"fake xlsx"), "application/vnd.ms-excel")}
        response = client.post("/api/audit", files=files)
        assert response.status_code == 400
        assert "Unsupported file format" in response.json()["detail"]


class TestWebEdgeCases:
    """Tests edge cases in web/app.py error handling."""

    def test_convert_unsupported_file_extension(self, client: TestClient) -> None:
        files = {"file": ("bad_file.bin", io.BytesIO(b"raw bytes"), "application/octet-stream")}
        response = client.post("/api/convert", files=files)
        assert response.status_code == 400
        assert "Unsupported file format" in response.json()["detail"]

    def test_validate_unsupported_file_extension(self, client: TestClient) -> None:
        files = {"file": ("document.pdf", io.BytesIO(b"%PDF-1.4"), "application/pdf")}
        response = client.post("/api/validate", files=files)
        assert response.status_code == 400
        assert "Unsupported file format" in response.json()["detail"]

    def test_root_route_serves_html(self, client: TestClient) -> None:
        response = client.get("/")
        assert response.status_code == 200
        assert "<!DOCTYPE html>" in response.text

    def test_convert_csv_file(self, client: TestClient) -> None:
        csv_data = "address;name;type\n40001;Active Power;uint16\n"
        files = {"file": ("source.csv", io.BytesIO(csv_data.encode("utf-8")), "text/csv")}
        data = {
            "manufacturer": "TestBrand",
            "model": "ModelCSV",
        }
        response = client.post("/api/convert", files=files, data=data)
        assert response.status_code == 200
        assert response.json()["success"] is True

    def test_convert_xml_file(self, client: TestClient) -> None:
        xml_data = (
            "<root><registers>"
            "<register><address>40001</address><name>Active Power</name><type>uint16</type></register>"
            "</registers></root>"
        )
        files = {"file": ("source.xml", io.BytesIO(xml_data.encode("utf-8")), "application/xml")}
        data = {
            "manufacturer": "TestBrand",
            "model": "ModelXML",
        }
        response = client.post("/api/convert", files=files, data=data)
        assert response.status_code == 200
        assert response.json()["success"] is True

    def test_tag_file_empty_filename(self, client: TestClient) -> None:
        files = {"file": ("", io.BytesIO(b"data"), "text/csv")}
        response = client.post("/api/tag", files=files)
        assert response.status_code in (400, 422)

    def test_audit_file_empty_filename(self, client: TestClient) -> None:
        files = {"file": ("", io.BytesIO(b"data"), "text/csv")}
        response = client.post("/api/audit", files=files)
        assert response.status_code in (400, 422)

    def test_tag_file_invalid_filename_dots(self, client: TestClient) -> None:
        files = {"file": ("..", io.BytesIO(b"data"), "text/csv")}
        response = client.post("/api/tag", files=files)
        assert response.status_code in (400, 422)

    def test_audit_file_invalid_filename_dots(self, client: TestClient) -> None:
        files = {"file": ("..", io.BytesIO(b"data"), "text/csv")}
        response = client.post("/api/audit", files=files)
        assert response.status_code in (400, 422)

    def test_library_when_directory_missing(self, client: TestClient) -> None:
        with pytest.MonkeyPatch.context() as mp:
            original_exists = os.path.exists
            mp.setattr(
                os.path,
                "exists",
                lambda p: False if "Definitions" in str(p) else original_exists(p),
            )
            response = client.get("/api/library")
            assert response.status_code == 200
            assert response.json() == []

    def test_root_when_index_missing(self, client: TestClient) -> None:
        with pytest.MonkeyPatch.context() as mp:
            original_exists = os.path.exists
            mp.setattr(
                os.path,
                "exists",
                lambda p: False if "index.html" in str(p) else original_exists(p),
            )
            response = client.get("/")
            assert response.status_code == 200
            assert "Frontend assets not found" in response.json()["message"]

    def test_tag_document_exception_handling(self, client: TestClient) -> None:
        from unittest.mock import patch

        csv_data = "modbusRTU;Inverter;SMA;STP5000;;;;;;;\n1;3;10;U16;;P;p;1;0;W;4\n"
        files = {"file": ("sma.csv", io.BytesIO(csv_data.encode("utf-8")), "text/csv")}
        with patch(
            "DefFileGenerator.knowledge_miner.enrich_definition_file",
            side_effect=RuntimeError("Boom"),
        ):
            response = client.post("/api/tag", files=files)
            assert response.status_code == 400
            assert "Failed to enrich definition file" in response.json()["detail"]

    def test_audit_document_exception_handling(self, client: TestClient) -> None:
        from unittest.mock import patch

        csv_data = "modbusRTU;Inverter;SMA;STP5000;;;;;;;\n1;3;10;U16;;P;p;1;0;W;4\n"
        files = {"file": ("sma.csv", io.BytesIO(csv_data.encode("utf-8")), "text/csv")}
        with patch(
            "DefFileGenerator.knowledge_miner.audit_definition_file",
            side_effect=RuntimeError("Boom"),
        ):
            response = client.post("/api/audit", files=files)
            assert response.status_code == 400
            assert "Failed to audit definition file" in response.json()["detail"]
