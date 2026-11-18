"""
Semantic Validator for Autonomous Modbus Discovery.
Validates intermediate CanonicalEquipmentModel against the 5 Apollo Pillars,
detects multi-register overlaps, checks address ranges, and assigns acceptance status.
"""

from __future__ import annotations

import logging

from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    ValidationOutcome,
    ValidationStatus,
)

logger = logging.getLogger(__name__)


class SemanticValidator:
    """Evaluates canonical models against Modbus industrial domain invariants."""

    def validate(self, model: CanonicalEquipmentModel) -> ValidationOutcome:
        blocking_issues: list[str] = []
        warnings: list[str] = []

        # Pillar 1: Mandatory metadata
        pillar_naming = bool(model.manufacturer.strip() and model.model.strip())
        if not pillar_naming:
            blocking_issues.append("Missing manufacturer or model identifier")

        # Pillar 2: Addressing bounds and overlap detection
        pillar_addressing = True
        occupied_spans: list[tuple[int, int, str]] = []  # (start_addr, end_addr, reg_name)

        for reg in model.registers:
            if reg.address < 0 or reg.address > 65535:
                blocking_issues.append(
                    f"Register '{reg.name}' address {reg.address} out of 0-65535 range"
                )
                pillar_addressing = False

            start_addr = reg.address
            end_addr = reg.address + reg.register_count - 1

            for prev_start, prev_end, prev_name in occupied_spans:
                # Check for overlap
                if max(start_addr, prev_start) <= min(end_addr, prev_end):
                    # Multi-register span collision
                    if start_addr != prev_start or reg.register_count > 1:
                        blocking_issues.append(
                            f"Illegal address overlap between '{reg.name}' [{start_addr}..{end_addr}] "
                            f"and '{prev_name}' [{prev_start}..{prev_end}]"
                        )
                        pillar_addressing = False

            occupied_spans.append((start_addr, end_addr, reg.name))

        # Pillar 3: Function code & Action code validation
        pillar_fc = True
        pillar_action = True
        for reg in model.registers:
            if reg.read_fc not in (1, 2, 3, 4):
                blocking_issues.append(f"Invalid read function code {reg.read_fc} for '{reg.name}'")
                pillar_fc = False
            if reg.action_code not in (1, 2, 3, 4):
                blocking_issues.append(f"Invalid action code {reg.action_code} for '{reg.name}'")
                pillar_action = False

        # Pillar 4: Telemetry completeness
        pillar_tags = True
        if len(model.registers) == 0:
            blocking_issues.append("Model contains zero registers")
            pillar_tags = False

        pillar_statuses = {
            "naming": pillar_naming,
            "addressing": pillar_addressing,
            "function_codes": pillar_fc,
            "action_codes": pillar_action,
            "telemetry": pillar_tags,
        }

        reconciliation = {
            "total_registers": len(model.registers),
            "unresolved_rows": len(model.unresolved_rows),
            "blocking_issues": len(blocking_issues),
            "warnings": len(warnings),
        }

        # Status assignment
        if not all(pillar_statuses.values()) or blocking_issues:
            status = ValidationStatus.FAILED_VALIDATION
        elif len(model.unresolved_rows) > 0:
            status = ValidationStatus.REVIEW_REQUIRED
        else:
            status = ValidationStatus.VALIDATED_OFFLINE

        return ValidationOutcome(
            status=status,
            pillar_statuses=pillar_statuses,
            blocking_issues=blocking_issues,
            warnings=warnings,
            reconciliation=reconciliation,
        )
