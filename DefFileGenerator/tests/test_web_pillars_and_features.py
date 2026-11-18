import io
import unittest
from pathlib import Path
from fastapi.testclient import TestClient
from web.app import app


class TestWebPillarsAndNewFeatures(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_pillars_check_valid_file(self):
        csv_content = (
            "modbusRTU;Inverter;ValidMfg;ValidModel;;;;;;;\n"
            "1;3;1001;U16;;Active Power;P_AC;1.000000;0.000000;W;4\n"
        )
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        response = self.client.post(
            "/api/pillars",
            files={"file": ("INV_ValidMfg_ValidModel.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertTrue(data["is_valid"])
        self.assertEqual(data["register_count"], 1)
        self.assertTrue(data["pillars"]["naming"])
        self.assertTrue(data["pillars"]["addressing"])
        self.assertTrue(data["pillars"]["function_codes"])
        self.assertTrue(data["pillars"]["action_codes"])
        self.assertTrue(data["pillars"]["tags"])
        self.assertTrue(data["pillars"]["encoding_schema"])

    def test_pillars_check_invalid_naming(self):
        csv_content = (
            "modbusRTU;Inverter;ValidMfg;ValidModel;;;;;;;\n"
            "1;3;1001;U16;;Active Power;P_AC;1.000000;0.000000;W;4\n"
        )
        file_obj = io.BytesIO(csv_content.encode("utf-8"))
        # Invalid prefix 'XYZ'
        response = self.client.post(
            "/api/pillars",
            files={"file": ("XYZ_InvalidName_Model.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertFalse(data["is_valid"])
        self.assertFalse(data["pillars"]["naming"])

    def test_audit_trail_endpoint(self):
        response = self.client.get("/api/audit-trail?limit=10")
        self.assertEqual(response.status_code, 200)
        events = response.json()
        self.assertIsInstance(events, list)
        if events:
            first = events[0]
            self.assertIn("action", first)
            self.assertIn("timestamp", first)

    def test_knowledge_stats_endpoint(self):
        response = self.client.get("/api/knowledge-stats")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertIn("total_mappings", data)
        self.assertIn("unique_tags", data)
        self.assertGreater(data["total_mappings"], 1000)

    def test_library_filtering(self):
        # Test category filter
        response = self.client.get("/api/library?category=INV")
        self.assertEqual(response.status_code, 200)
        items = response.json()
        self.assertGreater(len(items), 0)
        for it in items[:10]:
            self.assertEqual(it["prefix"], "INV")

        # Test query search
        response_huawei = self.client.get("/api/library?q=Huawei")
        self.assertEqual(response_huawei.status_code, 200)
        huawei_items = response_huawei.json()
        self.assertGreater(len(huawei_items), 0)
        for it in huawei_items:
            self.assertIn("huawei", it["filename"].lower())


if __name__ == "__main__":
    unittest.main()
