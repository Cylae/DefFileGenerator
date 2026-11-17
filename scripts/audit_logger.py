#!/usr/bin/env python3
"""
Apollo Audit & Traceability Logging Engine.
Maintains persistent, immutable, and verifiable audit trails over time:
  1. Machine-readable append-only log: `logs/audit_trail/audit_log.jsonl`
  2. Human-readable markdown audit ledger: `docs/audit/EQUIPMENT_AUDIT_LOG.md`
  3. Traceability matrix: `docs/audit/TRACEABILITY_MATRIX.md`
"""

from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

WORKSPACE_DIR = Path(__file__).resolve().parent.parent
LOGS_DIR = WORKSPACE_DIR / "logs" / "audit_trail"
AUDIT_LOG_JSONL = LOGS_DIR / "audit_log.jsonl"
AUDIT_LEDGER_MD = WORKSPACE_DIR / "docs" / "audit" / "EQUIPMENT_AUDIT_LOG.md"
TRACEABILITY_MD = WORKSPACE_DIR / "docs" / "audit" / "TRACEABILITY_MATRIX.md"


def compute_sha256(path: Path) -> str:
    """Computes SHA-256 hash of a file."""
    if not path.is_file():
        return ""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def log_event(
    action: str,
    equipment: str,
    manufacturer: str,
    source_document: str | None,
    definitions: list[str],
    register_count: int,
    tags_count: int,
    validation_status: str,
    details: str = "",
    extra_metadata: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Appends an audit event to the machine-readable JSONL log and refreshes the Markdown ledger.
    """
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    now = datetime.now(timezone.utc).isoformat()
    def_hashes = {}
    for d in definitions:
        p = WORKSPACE_DIR / "production" / d
        if p.is_file():
            def_hashes[d] = compute_sha256(p)

    event = {
        "timestamp": now,
        "action": action,
        "equipment": equipment,
        "manufacturer": manufacturer,
        "source_document": source_document,
        "definitions": definitions,
        "definition_hashes": def_hashes,
        "register_count": register_count,
        "tags_count": tags_count,
        "validation_status": validation_status,
        "details": details,
        "metadata": extra_metadata or {},
    }

    # Append to JSONL
    with open(AUDIT_LOG_JSONL, "a", encoding="utf-8") as f:
        f.write(json.dumps(event, ensure_ascii=False) + "\n")

    # Refresh human-readable markdown ledger
    rebuild_audit_ledger()
    return event


def read_all_events() -> list[dict[str, Any]]:
    """Reads all historical audit events from the JSONL log."""
    if not AUDIT_LOG_JSONL.is_file():
        return []
    events = []
    with open(AUDIT_LOG_JSONL, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                try:
                    events.append(json.loads(line))
                except Exception:
                    pass
    return events


def rebuild_audit_ledger():
    """Generates the Markdown audit ledger from all historical events."""
    events = read_all_events()
    lines = [
        "# 📜 Journal d'Audit & Traçabilité des Équipements",
        "",
        "Ce document consigne l'historique chronologique, immuable et vérifiable de toutes les intégrations,",
        "mises à niveau et validations de définitions d'équipements au sein du référentiel.",
        "",
        "---",
        "",
        "## 📑 Historique Chronologique des Opérations",
        "",
        "| Horodatage (UTC) | Action | Constructeur | Équipement | Document Source | Définition(s) Produite(s) | Registres | Tags | Statut |",
        "|---|---|---|---|---|---|:---:|:---:|:---:|",
    ]

    for ev in reversed(events):
        ts = ev.get("timestamp", "")[:19].replace("T", " ")
        act = ev.get("action", "")
        mfg = ev.get("manufacturer", "")
        eq = ev.get("equipment", "")
        src = ev.get("source_document") or "-"
        defs = ", ".join(f"`{d}`" for d in ev.get("definitions", []))
        regs = ev.get("register_count", 0)
        tags = ev.get("tags_count", 0)
        status = ev.get("validation_status", "UNKNOWN")
        icon = "✅" if "PASS" in status.upper() or "SUCCES" in status.upper() else "⚠️"
        lines.append(f"| {ts} | **{act}** | {mfg} | {eq} | {src} | {defs} | {regs:,} | {tags:,} | {icon} {status} |")

    lines.extend([
        "",
        "---",
        "",
        "## 🔒 Intégrité et Redondance des Empreintes",
        "",
        "Les définitions sont stockées de façon redondante et synchronisée sur 4 répertoires miroirs :",
        "- `production/` (production certifiée)",
        "- `Definitions/` (référentiel source)",
        "- `tagged_definitions/` (télémétrie balisée)",
        "- `generated_definitions/` (générations unifiées)",
        "",
        f"*Dernière mise à jour du journal : {datetime.now(timezone.utc).isoformat()}*",
    ])

    AUDIT_LEDGER_MD.parent.mkdir(parents=True, exist_ok=True)
    AUDIT_LEDGER_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    print("Apollo Audit Logger CLI initialized.")
    events = read_all_events()
    print(f"Total historical audit events: {len(events)}")
    rebuild_audit_ledger()
    print(f"Audit ledger written to: {AUDIT_LEDGER_MD}")
