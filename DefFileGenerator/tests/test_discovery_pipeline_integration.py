"""
Integration tests for the complete Autonomous Modbus Discovery Pipeline (Phase B).
"""

from unittest.mock import AsyncMock, patch

import pytest

from DefFileGenerator.discovery.models import (
    DocumentCandidate,
    ExportTarget,
    FetchOutcome,
    FetchResult,
    ValidationStatus,
)
from DefFileGenerator.discovery.orchestrator import WorkflowOrchestrator


@pytest.mark.anyio
async def test_end_to_end_autonomous_pipeline_with_csv_manual():
    orchestrator = WorkflowOrchestrator()
    # Sample manufacturer table
    mock_table_csv = (
        b"Address;Signal Name;Data Type;Scale;Unit;Access\n"
        b"0x1000;Active Power;U32;0.1;kW;R\n"
        b"4098;Reactive Power;S16;1;kVAR;R\n"
        b"4099;Grid Frequency;U16;0.01;Hz;R\n"
        b"4100;Active Power Setpoint;U16;1;%;R/W\n"
    )

    mock_fetch = FetchResult(
        outcome=FetchOutcome.SUCCESS,
        file_bytes=mock_table_csv,
        sha256="abcd1234efgh5678mockhash",
        content_type="text/csv",
        final_url="https://solar.huawei.com/download/SUN2000_Modbus.csv",
    )

    with patch.object(orchestrator.fetcher, "fetch", new_callable=AsyncMock) as mock_fetch_call:
        mock_fetch_call.return_value = mock_fetch

        result = await orchestrator.run(
            "Huawei SUN2000-100KTL-M1", target=ExportTarget.WEBDYN_SUNPM
        )

        assert result.success
        assert result.status == ValidationStatus.VALIDATED_OFFLINE
        assert result.csv_content is not None
        assert result.evidence_report is not None

        # Check CSV format
        lines = [line.strip() for line in result.csv_content.split("\n") if line.strip()]
        assert len(lines) >= 5

        # Header check: 11 columns
        header = lines[0].split(";")
        assert len(header) == 11
        assert header[0] == "Modbus"
        assert header[1] == "Inverter"
        assert header[2] == "Huawei"

        # First register check: address 0x1000 -> 4096, U32, CoefA = 0.100000
        row1 = lines[1].split(";")
        assert row1[0] == "1"  # 1-based index
        assert row1[1] == "3"  # Read FC 3
        assert row1[2] == "4096"  # Address 4096
        assert row1[3] == "U32"  # U32 type
        assert row1[5] == "Active Power"
        assert row1[7] == "0.100000"  # CoefA
        assert row1[8] == "0.000000"  # CoefB
        assert row1[9] == "kW"

        # Fourth register (Active Power Setpoint): is_writable -> Action 3
        row4 = lines[4].split(";")
        assert row4[10] == "3"

        # Check evidence report
        assert "abcd1234efgh5678mockhash" in result.evidence_report
        assert "Five Apollo Pillars Compliance" in result.evidence_report


@pytest.mark.anyio
async def test_manual_override_document_flow():
    orchestrator = WorkflowOrchestrator()
    # User provides manual document bytes directly
    manual_csv = (
        b"Address,Parameter Name,Type,Unit,Access\n"
        b"100,Inverter State,U16,,R\n"
        b"101,Fault Code,U16,,R\n"
    )

    result = await orchestrator.run(
        query="Sungrow SG110CX",
        override_file_bytes=manual_csv,
        override_filename="manual_input.csv",
    )

    assert result.success
    assert result.identity is not None
    assert result.identity.manufacturer == "Sungrow"
    assert result.model is not None
    assert len(result.model.registers) == 2
    assert "SG110CX" in (result.csv_content or "")


@pytest.mark.anyio
async def test_ssrf_blocked_in_discovery_pipeline():
    orchestrator = WorkflowOrchestrator()
    # Malicious or misconfigured candidate resolving to loopback
    with patch.object(orchestrator.discovery, "discover", new_callable=AsyncMock) as mock_disc:
        mock_disc.return_value = [
            DocumentCandidate(
                source_url="http://127.0.0.1:8000/internal/manual.pdf",
                title="Malicious Internal Link",
                publisher="Hacker",
                relevance_rank=1,
            )
        ]

        result = await orchestrator.run("Generic Inverter")
        assert not result.success
        assert result.status == ValidationStatus.BLOCKED_ACCESS
        assert "SSRF" in (result.error_message or "")
