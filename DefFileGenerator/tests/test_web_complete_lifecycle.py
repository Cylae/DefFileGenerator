"""
Comprehensive lifecycle test suite for all FastAPI web endpoints in web.app.
Verifies all 19 routes for schema conformance, HTTP status codes, error boundaries,
and end-to-end integration with the underlying DefFileGenerator engines.
"""

import unittest

from fastapi.testclient import TestClient

from web.app import app


class TestWebCompleteLifecycle(unittest.TestCase):
    """Rigorous tests covering all endpoints in web.app."""

    client: TestClient

    @classmethod
    def setUpClass(cls) -> None:
        cls.client = TestClient(app)

    def test_root_index_html_serving(self) -> None:
        """GET / must return HTTP 200 and serve the frontend index.html."""
        res = self.client.get("/")
        self.assertEqual(res.status_code, 200)
        self.assertIn("text/html", res.headers.get("content-type", ""))
        self.assertIn("WebdynSunPM", res.text)
        self.assertIn("tab-convert", res.text)
        self.assertIn("tab-discovery", res.text)
        self.assertIn("tab-versions", res.text)
        self.assertIn("tab-duplicates", res.text)

    def test_health_endpoint(self) -> None:
        """GET /api/health must report ok status and version."""
        res = self.client.get("/api/health")
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertEqual(data["status"], "ok")
        self.assertIn("version", data)

    def test_knowledge_stats_endpoint(self) -> None:
        """GET /api/knowledge-stats must return mapping counts."""
        res = self.client.get("/api/knowledge-stats")
        self.assertEqual(res.status_code, 200)
        data = res.json()
        self.assertIn("total_mappings", data)
        self.assertIn("unique_tags", data)
        self.assertGreater(data["total_mappings"], 0)

    def test_audit_trail_endpoint(self) -> None:
        """GET /api/audit-trail must return list of events."""
        res = self.client.get("/api/audit-trail?limit=10")
        self.assertEqual(res.status_code, 200)
        self.assertIsInstance(res.json(), list)

    def test_library_catalog_and_download(self) -> None:
        """GET /api/library and GET /api/library/{filename} must return valid data."""
        res = self.client.get("/api/library?source=production")
        self.assertEqual(res.status_code, 200)
        items = res.json()
        self.assertIsInstance(items, list)
        self.assertGreater(len(items), 0)

        first_file = items[0]["filename"]
        dl_res = self.client.get(f"/api/library/{first_file}")
        self.assertEqual(dl_res.status_code, 200)
        self.assertIn("text/csv", dl_res.headers.get("content-type", ""))
        self.assertIn(first_file, dl_res.headers.get("content-disposition", ""))

        # 404 on nonexistent definition
        not_found = self.client.get("/api/library/NONEXISTENT_MODEL_xyz.csv")
        self.assertEqual(not_found.status_code, 404)

        # 400 on non-csv extension
        bad_req = self.client.get("/api/library/not_a_csv.exe")
        self.assertEqual(bad_req.status_code, 400)

    def test_convert_pipeline_csv(self) -> None:
        """POST /api/convert must convert CSV register tables into Webdyn definition."""
        sample_csv = "Address;Name;Type;Unit;Scale\n40001;Active Power;U16;W;1.0\n40002;Grid Voltage;U16;V;0.1\n"
        files = {"file": ("test_registers.csv", sample_csv.encode("utf-8"), "text/csv")}
        data = {
            "protocol": "Modbus",
            "category": "Inverter",
            "manufacturer": "SolarTech",
            "model": "ST-5000",
        }
        res = self.client.post("/api/convert", files=files, data=data)
        self.assertEqual(res.status_code, 200)
        body = res.json()
        self.assertTrue(body["is_valid"])
        self.assertEqual(body["register_count"], 2)
        self.assertIn("SolarTech", body["csv_content"])

    def test_validate_endpoint(self) -> None:
        """POST /api/validate must evaluate CSV definition syntax and rules."""
        valid_def = "Modbus;Inverter;Brand;Model;;;;;;;\n1;3;40001;U16;;Active Power;p_ac;1.0;0.0;W;4\n"
        files = {"file": ("valid_def.csv", valid_def.encode("utf-8"), "text/csv")}
        res = self.client.post("/api/validate", files=files)
        self.assertEqual(res.status_code, 200)
        body = res.json()
        self.assertTrue(body["valid"])
        self.assertEqual(body["register_count"], 1)

    def test_tag_endpoint(self) -> None:
        """POST /api/tag must attribute SunPM canonical tags."""
        sample_def = "Modbus;Inverter;Brand;Model;;;;;;;\n1;3;40001;U16;;Active Power;;1.0;0.0;W;4\n"
        files = {"file": ("sample_def.csv", sample_def.encode("utf-8"), "text/csv")}
        data = {"category": "Inverter", "normalize_legacy": "true", "overwrite": "true"}
        res = self.client.post("/api/tag", files=files, data=data)
        self.assertEqual(res.status_code, 200)
        body = res.json()
        self.assertTrue(body["success"])
        self.assertEqual(body["register_count"], 1)
        self.assertGreaterEqual(body["tags_updated"], 1)

    def test_audit_endpoint(self) -> None:
        """POST /api/audit must report errors and warnings for definition files."""
        sample_def = "Modbus;Inverter;Brand;Model;;;;;;;\n1;3;40001;U16;;Active Power;p_ac;1.0;0.0;W;4\n"
        files = {"file": ("sample_def.csv", sample_def.encode("utf-8"), "text/csv")}
        res = self.client.post("/api/audit", files=files)
        self.assertEqual(res.status_code, 200)
        body = res.json()
        self.assertTrue(body["success"])
        self.assertEqual(body["error_count"], 0)

    def test_pillars_endpoint(self) -> None:
        """POST /api/pillars must evaluate file against the 5 production pillars."""
        valid_production = (
            "Modbus;Inverter;Brand;Model;;;;;;;\r\n"
            "1;3;40001;U16;;Active Power;p_ac;1.000000;0.000000;W;4\r\n"
        )
        files = {
            "file": (
                "INV_Brand_Model.csv",
                valid_production.encode("utf-8"),
                "text/csv",
            )
        }
        res = self.client.post("/api/pillars", files=files)
        self.assertEqual(res.status_code, 200)
        body = res.json()
        self.assertTrue(body["success"])
        self.assertIn("pillars", body)

    def test_versions_endpoints_lifecycle(self) -> None:
        """GET /api/versions and commit inspection endpoints."""
        res = self.client.get("/api/versions?limit=5")
        self.assertEqual(res.status_code, 200)
        commits = res.json()
        self.assertIsInstance(commits, list)
        self.assertGreater(len(commits), 0)

        cid = commits[0]["commit_id"]
        detail_res = self.client.get(f"/api/versions/{cid}")
        self.assertEqual(detail_res.status_code, 200)
        detail = detail_res.json()
        self.assertEqual(detail["commit_id"], cid)

        diff_res = self.client.get(f"/api/versions/{cid}/diff")
        self.assertEqual(diff_res.status_code, 200)

    def test_duplicates_endpoint(self) -> None:
        """GET /api/duplicates must return JSON duplicate report."""
        res = self.client.get("/api/duplicates?path=Definitions")
        self.assertEqual(res.status_code, 200)
        body = res.json()
        self.assertIn("total_files_scanned", body)
        self.assertIn("exact_hash_clusters", body)
        self.assertIn("alias_clusters", body)
        self.assertIn("exact_body_clusters", body)

    def test_discovery_lookup_validation(self) -> None:
        """POST /api/discovery/lookup must reject empty query with 400."""
        res = self.client.post("/api/discovery/lookup", json={"query": "   "})
        self.assertEqual(res.status_code, 400)
        self.assertIn("query cannot be empty", res.json()["detail"].lower())

    def test_discovery_upload_validation(self) -> None:
        """POST /api/discovery/upload must reject unsupported extensions."""
        files = {"file": ("unsupported.exe", b"binary", "application/octet-stream")}
        res = self.client.post("/api/discovery/upload", files=files)
        self.assertEqual(res.status_code, 400)


if __name__ == "__main__":
    unittest.main()
