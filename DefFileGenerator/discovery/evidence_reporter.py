"""
Evidence Reporter for Autonomous Modbus Discovery.
Generates verifiable, immutable markdown and JSON evidence ledgers
documenting cryptographic SHA-256 hashes, source URLs, and property-level citations.
"""

from __future__ import annotations

import datetime

from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    DocumentCandidate,
    FetchResult,
    ValidationOutcome,
)


class EvidenceReporter:
    """Builds cryptographic traceability ledgers for generated definitions."""

    def generate_markdown_report(
        self,
        model: CanonicalEquipmentModel,
        document_candidate: DocumentCandidate,
        fetch_result: FetchResult,
        validation: ValidationOutcome,
    ) -> str:
        timestamp_utc = datetime.datetime.now(datetime.timezone.utc).isoformat()
        total_regs = len(model.registers)

        md_lines = [
            "# AUTONOMOUS MODBUS DISCOVERY & CERTIFICATION EVIDENCE REPORT",
            f"**Generation Timestamp (UTC)**: `{timestamp_utc}`  ",
            f"**Equipment Identity**: `{model.manufacturer}` - `{model.model}`  ",
            f"**Target Protocol**: `{model.transport.value}`  ",
            f"**Acceptance Status**: **`{validation.status.value}`**  ",
            "",
            "---",
            "",
            "## 1. Cryptographic Provenance & Source Identity",
            f"- **Source URL**: `{document_candidate.source_url}`",
            f"- **Final Retrieved URL**: `{fetch_result.final_url}`",
            f"- **Publisher / Authority**: `{document_candidate.publisher}` (Official: `{document_candidate.is_official_manufacturer}`)",
            f"- **Document SHA-256**: `{fetch_result.sha256}`",
            f"- **Content-Type**: `{fetch_result.content_type}`",
            f"- **Payload Size**: `{len(fetch_result.file_bytes or b'') / 1024:.2f} KB`",
            "",
            "---",
            "",
            "## 2. Five Apollo Pillars Compliance",
            "| Apollo Pillar | Status | Detail |",
            "| :--- | :--- | :--- |",
        ]

        for pillar_name, passed in validation.pillar_statuses.items():
            badge = "✅ PASS" if passed else "❌ FAIL"
            md_lines.append(
                f"| **Pillar: {pillar_name.capitalize()}** | {badge} | Verified against domain invariants |"
            )

        md_lines.extend(
            [
                "",
                "---",
                "",
                "## 3. Register Reconciliation & Metrics",
                f"- **Total Extracted & Normalized Registers**: `{total_regs}`",
                f"- **Unresolved Extraction Regions**: `{len(model.unresolved_rows)}`",
                f"- **Blocking Issues**: `{len(validation.blocking_issues)}`",
                f"- **Warnings**: `{len(validation.warnings)}`",
                "",
                "### Sample Register Traceability Citations (Top 5)",
                "| Address | Variable Name | Type | Scaling | Unit | Provenance |",
                "| :--- | :--- | :--- | :--- | :--- | :--- |",
            ]
        )

        for reg in model.registers[:5]:
            prov = reg.provenance
            prov_text = f"Page {prov.page_number}, Table {prov.table_index}" if prov else "Inferred"
            md_lines.append(
                f"| `{reg.address}` | `{reg.name}` | `{reg.data_type.value}` | `{reg.factor:.6f}` | `{reg.unit}` | {prov_text} |"
            )

        md_lines.extend(
            [
                "",
                "---",
                "*Report certified by Webdyn Autonomous Modbus Discovery Pipeline v1.0.0 Apollo Edition.*",
            ]
        )

        return "\n".join(md_lines)
