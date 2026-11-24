"""
Comprehensive test suite for enhanced Web UI API endpoints:
- Dynamic Health & Catalog Statistics
- Definition In-Browser Inspection & Parsing
- In-Browser Definition Re-validation & Editing
- SunSpec Relevance Evaluation & Gate in Web API
- SunSpec Web Generation & Multi-Profile Conversion
- Theme Switching and DOM Elements
"""

import io
import unittest

from fastapi.testclient import TestClient

from web.app import app


class TestWebEnhancedFeatures(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client = TestClient(app)

    def test_health_check_dynamic_stats(self):
        """Verifies GET /api/health provides dynamic production and library statistics."""
        response = self.client.get("/api/health")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(data["status"], "ok")
        self.assertIn("version", data)
        self.assertIn("stats", data)
        stats = data["stats"]
        self.assertGreaterEqual(stats["production_models"], 1000)
        self.assertGreaterEqual(stats["total_certified_models"], 1000)

    def test_library_inspect_existing_definition(self):
        """Verifies GET /api/library/{filename}/inspect returns parsed registers and 5-pillars status."""
        # Test with a known certified production file
        filename = "BAT_X7_BMS.csv"
        response = self.client.get(f"/api/library/{filename}/inspect")
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertEqual(data["filename"], filename)
        self.assertEqual(data["category"], "Battery")
        self.assertTrue(data["pillars_valid"])
        self.assertGreater(data["register_count"], 0)
        self.assertIsInstance(data["registers"], list)
        self.assertEqual(len(data["registers"]), data["register_count"])

        # Check register structure
        first_reg = data["registers"][0]
        self.assertIn("address", first_reg)
        self.assertIn("name", first_reg)
        self.assertIn("tag", first_reg)
        self.assertIn("data_type", first_reg)
        self.assertIn("factor", first_reg)
        self.assertIn("action", first_reg)

    def test_library_inspect_not_found(self):
        """Verifies GET /api/library/{filename}/inspect returns 404 for missing files."""
        response = self.client.get("/api/library/NONEXISTENT_MODEL_9999.csv/inspect")
        self.assertEqual(response.status_code, 404)

    def test_library_inspect_invalid_extension(self):
        """Verifies GET /api/library/{filename}/inspect rejects non-csv files."""
        response = self.client.get("/api/library/malicious_payload.exe/inspect")
        self.assertEqual(response.status_code, 400)

    def test_definition_revalidate_valid_payload(self):
        """Verifies POST /api/definition/revalidate certifies edited registers against 5 Pillars."""
        payload = {
            "filename": "INV_TestVendor_SmartInverter.csv",
            "protocol": "Modbus",
            "category": "Inverter",
            "manufacturer": "TestVendor",
            "model": "SmartInverter",
            "auto_tag": True,
            "registers": [
                {
                    "index": 1,
                    "reg_type": "3",
                    "address": 40001,
                    "data_type": "U16",
                    "operation": "",
                    "name": "Grid Frequency",
                    "tag": "F_AC",
                    "factor": 0.01,
                    "offset": 0.0,
                    "unit": "Hz",
                    "action": "4",
                },
                {
                    "index": 2,
                    "reg_type": "3",
                    "address": 40002,
                    "data_type": "U32",
                    "operation": "",
                    "name": "Total Active Power",
                    "tag": "P_AC",
                    "factor": 1.0,
                    "offset": 0.0,
                    "unit": "W",
                    "action": "4",
                },
            ],
        }
        response = self.client.post("/api/definition/revalidate", json=payload)
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertTrue(data["is_valid"])
        self.assertEqual(data["register_count"], 2)
        self.assertTrue(data["pillars"]["naming"])
        self.assertTrue(data["pillars"]["addressing"])
        self.assertTrue(data["pillars"]["encoding_schema"])
        self.assertIn("P_AC", data["csv_content"])

    def test_definition_revalidate_empty_registers_rejected(self):
        """Verifies POST /api/definition/revalidate rejects empty register arrays."""
        payload = {
            "filename": "INV_Empty_Empty.csv",
            "protocol": "Modbus",
            "category": "Inverter",
            "manufacturer": "Empty",
            "model": "Empty",
            "registers": [],
        }
        response = self.client.post("/api/definition/revalidate", json=payload)
        self.assertEqual(response.status_code, 400)

    def test_sunspec_evaluate_proprietary_skipped(self):
        """Verifies POST /api/sunspec/evaluate skips pure proprietary Modbus."""
        payload = {
            "query": "Huawei SUN2000 Holding Registers 30000 to 35000",
            "mode": "auto",
            "category": "Inverter",
        }
        response = self.client.post("/api/sunspec/evaluate", data=payload)
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertFalse(data["should_generate"])
        self.assertTrue(data["is_skipped"])
        self.assertIn("proprietary", data["reason"].lower())

    def test_sunspec_evaluate_explicit_models_pass(self):
        """Verifies POST /api/sunspec/evaluate detects models and plans target files."""
        text_with_sunspec = (
            "SunSpec Alliance Information Models Supported:\n"
            "Model 1: Common Model (Length = 65)\n"
            "Model 103: Three Phase Inverter\n"
            "Model 160: Multiple MPPT Inverter Extension\n"
        )
        payload = {
            "query": text_with_sunspec,
            "mode": "auto",
            "manufacturer": "Fronius",
            "model": "Symo",
            "category": "Inverter",
            "base_address": 40000,
        }
        response = self.client.post("/api/sunspec/evaluate", data=payload)
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertTrue(data["should_generate"])
        self.assertIn(1, data["detected_models"])
        self.assertIn(103, data["detected_models"])
        self.assertGreater(len(data["targets"]), 0)
        self.assertEqual(data["targets"][0]["profile"], "Legacy")

    def test_sunspec_evaluate_forced_mode(self):
        """Verifies POST /api/sunspec/evaluate in force mode generates target even without explicit models."""
        payload = {
            "query": "Generic Inverter Device",
            "mode": "force",
            "manufacturer": "ForcedMfg",
            "model": "ForcedModel",
            "category": "Inverter",
        }
        response = self.client.post("/api/sunspec/evaluate", data=payload)
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertTrue(data["should_generate"])
        self.assertTrue(data["is_forced"])
        self.assertGreater(len(data["targets"]), 0)

    def test_sunspec_generate_endpoint(self):
        """Verifies POST /api/sunspec/generate compiles certified 11-column CSV definition."""
        payload = {
            "manufacturer": "WebTestMfg",
            "model": "WebTestModel",
            "category": "Inverter",
            "base_address": 40000,
            "common_model_len": 65,
            "profile": "DER700",
            "topology": "ThreePhase",
            "model_ids": [1, 701, 702],
            "mppt_ports": 0,
        }
        response = self.client.post("/api/sunspec/generate", json=payload)
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertTrue(data["is_valid"])
        self.assertGreater(data["register_count"], 0)
        self.assertIn("SunSpec", data["filename"])
        self.assertIn("DER700", data["filename"])
        self.assertIsInstance(data["preview"], list)

    def test_convert_with_sunspec_force(self):
        """Verifies POST /api/convert generates SunSpec definitions alongside native definition when requested."""
        csv_sample = (
            "Register,Description,Type,Scale\n"
            "40001,Power,int16,0.1\n"
            "40002,Voltage,int16,0.1\n"
        )
        file_obj = io.BytesIO(csv_sample.encode("utf-8"))
        data = {
            "manufacturer": "HybridVendor",
            "model": "MultiProtocol",
            "protocol": "Modbus",
            "category": "Inverter",
            "address_offset": 0,
            "tag": True,
            "sunspec_mode": "force",
            "sunspec_base": 40000,
            "sunspec_ports": 0,
        }
        response = self.client.post(
            "/api/convert",
            data=data,
            files={"file": ("vendor_spec.csv", file_obj, "text/csv")},
        )
        self.assertEqual(response.status_code, 200)
        res_data = response.json()
        self.assertTrue(res_data["success"])
        self.assertTrue(res_data["sunspec_generated"])
        self.assertGreater(len(res_data["sunspec_files"]), 0)
        first_sunspec = res_data["sunspec_files"][0]
        self.assertIn("SunSpec", first_sunspec["filename"])
        self.assertTrue(first_sunspec["is_valid"])

    def test_frontend_contains_new_elements(self):
        """Verifies index.html renders SunSpec Studio tab, Theme button, and Inspect modal."""
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.text

        # Header theme toggle
        self.assertIn('id="themeToggleBtn"', html)
        self.assertIn('id="headerModelsBadge"', html)

        # Tabs
        self.assertIn('id="tabBtnSunspec"', html)
        self.assertIn('id="tab-sunspec"', html)
        self.assertIn("SunSpec Protocol Studio", html)

        # Inspect modal
        self.assertIn('id="inspectModal"', html)
        self.assertIn('id="inspectTable"', html)
        self.assertIn('id="inspectEditModeBtn"', html)
        self.assertIn('id="inspectSaveBtn"', html)

        # Convert form SunSpec controls
        self.assertIn('id="sunspecModeSelect"', html)
        self.assertIn('id="convertFileSwitcher"', html)


if __name__ == "__main__":
    unittest.main()
