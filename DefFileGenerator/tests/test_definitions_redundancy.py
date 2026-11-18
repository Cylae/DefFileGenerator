#!/usr/bin/env python3
"""
Integrity and Redundancy Verification Suite.
Verifies:
  1. Quadruple redundant synchronization of certified definitions
  2. Cryptographic checksums and existence of source manuals in `sources/MANIFEST.json`
  3. Audit log persistence in `logs/audit_trail/audit_log.jsonl`
  4. Traceability matrix consistency in `docs/audit/TRACEABILITY_MATRIX.md`
"""

import hashlib
import json
from pathlib import Path


def test_sources_manifest_integrity():
    """Verify that sources/MANIFEST.json matches reality on disk."""
    manifest_file = Path("sources/MANIFEST.json")
    assert manifest_file.is_file(), "Missing sources/MANIFEST.json"

    data = json.loads(manifest_file.read_text(encoding="utf-8"))
    assert "documents" in data
    assert data["total_documents"] >= 4

    for doc in data["documents"]:
        if "files" in doc:
            for subfile in doc["files"]:
                sub_path = Path(doc["relative_path"]) / subfile["filename"]
                assert sub_path.is_file(), f"Document file missing on disk: {sub_path}"
                assert sub_path.stat().st_size == subfile["size_bytes"], f"Size mismatch for {sub_path}"
                assert hashlib.sha256(sub_path.read_bytes()).hexdigest() == subfile["sha256"], f"SHA mismatch for {sub_path}"
        else:
            rel_path = Path(doc["relative_path"])
            assert rel_path.is_file(), f"Document file missing on disk: {rel_path}"
            actual_size = rel_path.stat().st_size
            assert actual_size == doc["size_bytes"], f"Size mismatch for {rel_path}"

            actual_sha = hashlib.sha256(rel_path.read_bytes()).hexdigest()
            assert actual_sha == doc["sha256"], f"SHA-256 hash mismatch for {rel_path}"


def test_audit_log_trail_persistence():
    """Verify that audit logs exist, are parseable JSONL, and record operations."""
    jsonl_log = Path("logs/audit_trail/audit_log.jsonl")
    assert jsonl_log.is_file(), "Missing logs/audit_trail/audit_log.jsonl"

    events = []
    with open(jsonl_log, encoding="utf-8") as f:
        for line in f:
            if line.strip():
                events.append(json.loads(line))

    assert len(events) >= 2, "Expected at least 2 historical audit events"
    manufacturers = {e.get("manufacturer") for e in events}
    assert "Solplanet" in manufacturers
    assert "Cyberpower" in manufacturers

    # Verify markdown ledger exists and is non-empty
    ledger_md = Path("docs/audit/EQUIPMENT_AUDIT_LOG.md")
    assert ledger_md.is_file(), "Missing docs/audit/EQUIPMENT_AUDIT_LOG.md"
    assert ledger_md.stat().st_size > 500


def test_traceability_matrix_completeness():
    """Verify that docs/audit/TRACEABILITY_MATRIX.md covers required equipments."""
    trace_file = Path("docs/audit/TRACEABILITY_MATRIX.md")
    assert trace_file.is_file(), "Missing docs/audit/TRACEABILITY_MATRIX.md"
    text = trace_file.read_text(encoding="utf-8")

    assert "CyberPower" in text
    assert "Solplanet" in text
    assert "FoxESS" in text
    assert "MarsEnergy" in text
    assert "COM_Solplanet_Ai-Logger.csv" in text
    assert "INV_Cyberpower_GENERIC.csv" in text
    assert "INV_Cyberpower_CPSPV120K-140K.csv" in text


def test_certified_definitions_redundancy():
    """Verify quadruple redundancy between production, Definitions, tagged, and generated folders."""
    sample_definitions = [
        "COM_Solplanet_Ai-Logger.csv",
        "INV_Cyberpower_GENERIC.csv",
        "INV_Cyberpower_CPSPV120K-140K.csv",
    ]
    folders = ["Definitions", "tagged_definitions", "generated_definitions"]

    for fname in sample_definitions:
        prod_path = Path("production") / fname
        assert prod_path.is_file(), f"Missing production file: {prod_path}"
        prod_bytes = prod_path.read_bytes()

        for fld in folders:
            copy_path = Path(fld) / fname
            assert copy_path.is_file(), f"Missing redundant copy in {fld}/{fname}"
            assert copy_path.read_bytes() == prod_bytes, f"Byte mismatch in {fld}/{fname}"
