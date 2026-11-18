import io
import unittest
from unittest.mock import AsyncMock, patch

from fastapi.testclient import TestClient

from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    CanonicalRegisterRecord,
    DocumentCandidate,
    FetchOutcome,
    FetchResult,
    IdentityCandidate,
    RegisterDataType,
    ValidationOutcome,
    ValidationStatus,
)
from DefFileGenerator.discovery.orchestrator import AutonomousPipelineResult
from web.app import app


class TestWebDiscoveryEndpoints(unittest.TestCase):
    def setUp(self):
        self.client = TestClient(app)

    def test_lookup_empty_query_rejected(self):
        response = self.client.post(
            "/api/discovery/lookup",
            json={"query": "   ", "target_format": "webdyn_sunpm", "ocr_fallback": True},
        )
        self.assertEqual(response.status_code, 400)
        data = response.json()
        self.assertIn("cannot be empty", data["detail"])

    @patch("DefFileGenerator.discovery.orchestrator.WorkflowOrchestrator.run", new_callable=AsyncMock)
    def test_lookup_successful_mocked(self, mock_run):
        mock_run.return_value = AutonomousPipelineResult(
            success=True,
            status=ValidationStatus.VALIDATED_OFFLINE,
            query="Huawei SUN2000-100KTL",
            identity=IdentityCandidate(
                manufacturer="Huawei",
                model_family="SUN2000",
                exact_model="SUN2000-100KTL",
                confidence_score=1.0,
            ),
            document=DocumentCandidate(
                source_url="https://solar.huawei.com/doc.pdf",
                title="SUN2000 Modbus Protocol",
                publisher="Huawei",
            ),
            fetch=FetchResult(
                outcome=FetchOutcome.SUCCESS,
                sha256="0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef",
                content_type="application/pdf",
                final_url="https://solar.huawei.com/doc.pdf",
            ),
            model=CanonicalEquipmentModel(
                manufacturer="Huawei",
                model="SUN2000-100KTL",
                category="Inverter",
                total_registers=2,
                registers=[
                    CanonicalRegisterRecord(
                        id="reg_1",
                        name="Model",
                        address=30000,
                        data_type=RegisterDataType.STRING,
                        factor=1.0,
                        unit="",
                        action_code=4,
                    ),
                    CanonicalRegisterRecord(
                        id="reg_2",
                        name="Active Power",
                        address=32064,
                        data_type=RegisterDataType.U32,
                        factor=0.001,
                        unit="kW",
                        action_code=4,
                    ),
                ],
            ),
            validation=ValidationOutcome(
                status=ValidationStatus.VALIDATED_OFFLINE,
                pillar_statuses={"pillar_1": True, "pillar_2": True},
            ),
            csv_content="modbusRTU;Inverter;Huawei;SUN2000-100KTL;;;;;;;\n1;3;30000;STR;;Model;n;0;0;1.000000;4\n",
            evidence_report="# Evidence Ledger\n",
        )

        response = self.client.post(
            "/api/discovery/lookup",
            json={"query": "Huawei SUN2000-100KTL", "target_format": "webdyn_sunpm", "ocr_fallback": True},
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertEqual(data["identity"]["manufacturer"], "Huawei")
        self.assertEqual(data["register_count"], 2)
        self.assertIn("modbusRTU", data["csv_content"])

    def test_upload_invalid_file_extension(self):
        file_obj = io.BytesIO(b"malicious script")
        response = self.client.post(
            "/api/discovery/upload",
            files={"file": ("malicious.exe", file_obj, "application/octet-stream")},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Unsupported file type", response.json()["detail"])

    @patch("DefFileGenerator.discovery.orchestrator.WorkflowOrchestrator.run", new_callable=AsyncMock)
    def test_upload_valid_csv_file(self, mock_run):
        mock_run.return_value = AutonomousPipelineResult(
            success=True,
            status=ValidationStatus.VALIDATED_OFFLINE,
            query="local_registers.csv",
            identity=IdentityCandidate(
                manufacturer="TestMfg",
                model_family="TestSeries",
                exact_model="TestModel",
                confidence_score=1.0,
            ),
            model=CanonicalEquipmentModel(
                manufacturer="TestMfg",
                model="TestModel",
                category="Inverter",
                total_registers=1,
                registers=[
                    CanonicalRegisterRecord(
                        id="reg_1",
                        name="Voltage",
                        address=40001,
                        data_type=RegisterDataType.U16,
                        factor=0.1,
                        unit="V",
                        action_code=4,
                    )
                ],
            ),
            csv_content="modbusRTU;Inverter;TestMfg;TestModel;;;;;;;\n1;3;40001;U16;;Voltage;n;0;0;0.100000;4\n",
        )

        sample_csv = "Address,Name,Type\n40001,Voltage,U16\n"
        file_obj = io.BytesIO(sample_csv.encode("utf-8"))
        response = self.client.post(
            "/api/discovery/upload",
            files={"file": ("local_registers.csv", file_obj, "text/csv")},
            data={"target_format": "webdyn_sunpm", "ocr_fallback": "true"},
        )
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertTrue(data["success"])
        self.assertEqual(data["identity"]["manufacturer"], "TestMfg")
        self.assertEqual(data["register_count"], 1)


if __name__ == "__main__":
    unittest.main()
