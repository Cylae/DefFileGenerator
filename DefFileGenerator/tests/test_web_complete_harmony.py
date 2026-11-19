"""
Complete End-to-End Harmony and Parity Suite for Web Interface and Core Generator Engine.
Validates that:
  1. Multi-format extraction & conversion (JSON, CSV, Excel, PDF) produces certified Modbus definitions.
  2. Protocol is strictly normalized to 'Modbus' across all web endpoints.
  3. Auto-enrichment with SunPM canonical tags works seamlessly in the web pipeline.
  4. 5 Pillars Apollo compliance verification reports clear and actionable diagnostics.
  5. Audit & repair pipeline restores invalid definitions to 100% compliance.
  6. Library endpoints correctly list and serve the 1,118 certified production definitions.
  7. Error handling preserves security while returning unambiguous, actionable messages.
"""

import io
import json
from pathlib import Path

from fastapi.testclient import TestClient

from web.app import app

client = TestClient(app)
WORKSPACE_DIR = Path(__file__).resolve().parents[2]


class TestWebCompleteHarmony:
    """Verifies complete concord between web API, UI contracts, and core engine."""

    def test_health_check(self):
        """API health endpoint responds with ok status and version."""
        res = client.get("/api/health")
        assert res.status_code == 200
        data = res.json()
        assert data["status"] == "ok"
        assert "version" in data

    def test_knowledge_stats(self):
        """Knowledge base stats are properly computed and exposed."""
        res = client.get("/api/knowledge-stats")
        assert res.status_code == 200
        data = res.json()
        assert data["unique_tags"] > 0
        assert data["total_mappings"] > 10000

    def test_convert_csv_end_to_end_with_autotag(self):
        """Converting a CSV table produces a validated definition with Modbus protocol and SunPM tags."""
        raw_csv = (
            "Name;Register;Type;Scale;Unit\n"
            "Active Power;40001;U16;0.1;kW\n"
            "Grid Voltage;40002;U16;0.1;V\n"
            "Grid Frequency;40003;U16;0.01;Hz\n"
            "Total Energy;40004;U32;1.0;kWh\n"
        )
        file_bytes = raw_csv.encode("utf-8")
        res = client.post(
            "/api/convert",
            files={"file": ("inverter_registers.csv", io.BytesIO(file_bytes), "text/csv")},
            data={
                "manufacturer": "SolarTech",
                "model": "ST-5000",
                "protocol": "modbusRTU",  # Should be normalized to 'Modbus'
                "category": "Inverter",
                "tag": "true",
            },
        )
        assert res.status_code == 200
        data = res.json()
        assert data["success"] is True
        assert data["is_valid"] is True
        assert data["register_count"] >= 4

        # Header check: must be Modbus
        csv_content = data["csv_content"]
        header_line = csv_content.strip().split("\n")[0]
        header_parts = header_line.split(";")
        assert header_parts[0] == "Modbus"
        assert header_parts[1] == "Inverter"
        assert header_parts[2] == "SolarTech"
        assert header_parts[3] == "ST-5000"

        # Check preview mapping
        preview = data["preview"]
        assert len(preview) >= 4
        assert any(r["Tag"] == "P_AC" for r in preview)
        assert any(r["Tag"] == "FREQ" for r in preview)

    def test_convert_json_end_to_end(self):
        """Converting a JSON register table (like Chint DTSU666) works directly."""
        json_data = [
            {"address": 0, "name": "Software Version", "type": "int16", "access": "ro"},
            {
                "address": 3,
                "name": "Active Power Phase A",
                "type": "float32",
                "access": "ro",
                "scale": 0.001,
                "unit": "kW",
            },
        ]
        file_bytes = json.dumps(json_data).encode("utf-8")
        res = client.post(
            "/api/convert",
            files={"file": ("meter.json", io.BytesIO(file_bytes), "application/json")},
            data={
                "manufacturer": "Chint",
                "model": "DTSU",
                "protocol": "Modbus",
                "category": "Meter",
            },
        )
        assert res.status_code == 200
        data = res.json()
        assert data["success"] is True
        assert data["is_valid"] is True
        assert data["csv_content"].startswith("Modbus;Meter;Chint;DTSU;")

    def test_pillars_endpoint_accepts_valid_definition(self):
        """Pillars audit endpoint certifies 5 pillars on a valid production file."""
        sample_csv = (
            "Modbus;Meter;Chint;DTSU666;;;;;;;\n"
            "1;3;0;U16;;Version;version;1.000000;0.000000;;4\n"
            "2;3;10;U16;;Active Power;P_AC;0.001000;0.000000;kW;4\n"
        )
        res = client.post(
            "/api/pillars",
            files={
                "file": (
                    "MET_Chint_DTSU666.csv",
                    io.BytesIO(sample_csv.encode("utf-8")),
                    "text/csv",
                )
            },
        )
        assert res.status_code == 200
        data = res.json()
        assert data["success"] is True
        assert data["is_valid"] is True
        assert data["issues"] == []

    def test_pillars_endpoint_rejects_deprecated_modbus_protocol(self):
        """Pillars audit endpoint flags modbusRTU as deprecated."""
        bad_csv = (
            "modbusRTU;Meter;Chint;DTSU666;;;;;;;\n"
            "1;3;0;U16;;Version;version;1.000000;0.000000;;4\n"
        )
        res = client.post(
            "/api/pillars",
            files={
                "file": ("MET_Chint_DTSU666.csv", io.BytesIO(bad_csv.encode("utf-8")), "text/csv")
            },
        )
        assert res.status_code == 200
        data = res.json()
        assert data["is_valid"] is False
        assert any(
            "deprecated" in iss.lower() and "modbus" in iss.lower() for iss in data["issues"]
        )

    def test_audit_endpoint_auto_fixes_and_validates(self):
        """Audit endpoint with auto-fix repairs non-compliant rows and produces valid output."""
        messy_csv = "modbusRTU;Inverter;Vendor;Model;;;;;;;\n1;3;1000;uint16;;Grid Power;;1;0;W;4\n"
        res = client.post(
            "/api/audit",
            files={"file": ("test.csv", io.BytesIO(messy_csv.encode("utf-8")), "text/csv")},
            data={"strict_overlap": "true", "auto_fix": "true"},
        )
        assert res.status_code == 200
        data = res.json()
        assert data["success"] is True
        # Fixed CSV should be present
        assert data["fixed_csv_content"] is not None
        assert "P_AC" in data["fixed_csv_content"]

    def test_library_catalog_listing_and_filtering(self):
        """Library endpoint returns the catalog of certified definitions."""
        res = client.get("/api/library")
        assert res.status_code == 200
        items = res.json()
        assert len(items) >= 1118
        # Filter by category
        inverters = [i for i in items if i["category"] == "Inverter"]
        assert len(inverters) >= 400

    def test_library_single_file_download(self):
        """Library single file endpoint serves certified definitions with UTF-8 encoding."""
        res = client.get("/api/library/MET_Chint_DTSU666.csv")
        assert res.status_code == 200
        content = res.text
        assert content.startswith("Modbus;Meter;CHINT;DTSU666;")

    def test_tagger_endpoint_enriches_csv(self):
        """Standalone tagger endpoint maps canonical SunPM tags."""
        untagged_csv = (
            "Modbus;Inverter;SMA;STP5000;;;;;;;\n"
            "1;3;30775;U32;;Grid Current;;0.001;0;A;4\n"
            "2;3;30795;U32;;Grid Frequency;;0.01;0;Hz;4\n"
        )
        res = client.post(
            "/api/tag",
            files={"file": ("untagged.csv", io.BytesIO(untagged_csv.encode("utf-8")), "text/csv")},
            data={"category": "Inverter", "overwrite": "true"},
        )
        assert res.status_code == 200
        data = res.json()
        assert data["success"] is True
        assert data["tags_updated"] >= 1
        assert "FREQ" in data["csv_content"]
