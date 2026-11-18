import os
import unittest
from pathlib import Path
from unittest.mock import AsyncMock, patch

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
from DefFileGenerator.main import main


class TestCliDiscoveryCommand(unittest.TestCase):
    def test_discover_help(self):
        with self.assertRaises(SystemExit) as cm:
            main(["discover", "--help"])
        self.assertEqual(cm.exception.code, 0)

    @patch(
        "DefFileGenerator.discovery.orchestrator.WorkflowOrchestrator.run", new_callable=AsyncMock
    )
    def test_discover_success_flow(self, mock_run, tmp_path_factory=None):
        out_csv = "test_cli_discovered_output.csv"
        out_report = "test_cli_discovered_report.md"
        try:
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
                    sha256="abc123sha",
                    content_type="application/pdf",
                    final_url="https://solar.huawei.com/doc.pdf",
                ),
                model=CanonicalEquipmentModel(
                    manufacturer="Huawei",
                    model="SUN2000-100KTL",
                    category="Inverter",
                    total_registers=1,
                    registers=[
                        CanonicalRegisterRecord(
                            id="reg_1",
                            name="Active Power",
                            address=32064,
                            data_type=RegisterDataType.U32,
                            factor=0.001,
                            unit="kW",
                            action_code=4,
                        )
                    ],
                ),
                validation=ValidationOutcome(status=ValidationStatus.VALIDATED_OFFLINE),
                csv_content="modbusRTU;Inverter;Huawei;SUN2000-100KTL;;;;;;;\n1;3;32064;U32;;Active Power;n;0;0;0.001000;4\n",
                evidence_report="# Evidence Ledger for Huawei SUN2000\n",
            )

            main(
                [
                    "discover",
                    "Huawei SUN2000-100KTL",
                    "-o",
                    out_csv,
                    "--report",
                    out_report,
                    "--force",
                ]
            )

            self.assertTrue(os.path.isfile(out_csv))
            with open(out_csv, encoding="utf-8") as f:
                content = f.read()
            self.assertIn("modbusRTU;Inverter;Huawei", content)
            self.assertIn("32064;U32", content)

            self.assertTrue(os.path.isfile(out_report))
            with open(out_report, encoding="utf-8") as f:
                report_content = f.read()
            self.assertIn("# Evidence Ledger", report_content)

        finally:
            if os.path.exists(out_csv):
                os.remove(out_csv)
            if os.path.exists(out_report):
                os.remove(out_report)

    @patch(
        "DefFileGenerator.discovery.orchestrator.WorkflowOrchestrator.run", new_callable=AsyncMock
    )
    def test_discover_auto_naming(self, mock_run):
        expected_auto_name = "INV_GoodWe_GW5000.csv"
        try:
            mock_run.return_value = AutonomousPipelineResult(
                success=True,
                status=ValidationStatus.VALIDATED_OFFLINE,
                query="GoodWe GW5000",
                identity=IdentityCandidate(
                    manufacturer="GoodWe",
                    model_family="GW",
                    exact_model="GW5000",
                ),
                model=CanonicalEquipmentModel(
                    manufacturer="GoodWe",
                    model="GW5000",
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
                csv_content="modbusRTU;Inverter;GoodWe;GW5000;;;;;;;\n1;3;40001;U16;;Voltage;n;0;0;0.100000;4\n",
            )

            main(["discover", "GoodWe GW5000", "--force"])
            self.assertTrue(os.path.isfile(expected_auto_name))
        finally:
            if os.path.exists(expected_auto_name):
                os.remove(expected_auto_name)

    @patch(
        "DefFileGenerator.discovery.orchestrator.WorkflowOrchestrator.run", new_callable=AsyncMock
    )
    def test_discover_failure_exits_with_code_1(self, mock_run):
        mock_run.return_value = AutonomousPipelineResult(
            success=False,
            status=ValidationStatus.NOT_FOUND,
            query="UnknownEquipment XYZ999",
            error_message="Could not identify manufacturer or model",
        )

        with self.assertRaises(SystemExit) as cm:
            main(["discover", "UnknownEquipment XYZ999"])
        self.assertEqual(cm.exception.code, 1)

    def test_discover_missing_fallback_document_exits_code_1(self):
        with self.assertRaises(SystemExit) as cm:
            main(["discover", "Huawei SUN2000", "--document", "non_existent_file.pdf"])
        self.assertEqual(cm.exception.code, 1)

    @patch(
        "DefFileGenerator.discovery.orchestrator.WorkflowOrchestrator.run", new_callable=AsyncMock
    )
    def test_discover_rejects_overwrite_without_force(self, mock_run):
        existing_file = "test_already_exists.csv"
        Path(existing_file).write_text("existing content", encoding="utf-8")
        try:
            mock_run.return_value = AutonomousPipelineResult(
                success=True,
                status=ValidationStatus.VALIDATED_OFFLINE,
                query="Sample Inverter",
                identity=IdentityCandidate(
                    manufacturer="Sample", model_family="Inv", exact_model="100"
                ),
                model=CanonicalEquipmentModel(
                    manufacturer="Sample", model="100", category="Inverter"
                ),
                csv_content="sample csv content",
            )

            with self.assertRaises(SystemExit) as cm:
                main(["discover", "Sample Inverter", "-o", existing_file])
            self.assertEqual(cm.exception.code, 1)
        finally:
            if os.path.exists(existing_file):
                os.remove(existing_file)


if __name__ == "__main__":
    unittest.main()
