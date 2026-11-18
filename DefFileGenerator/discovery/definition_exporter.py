"""
Target Definition Exporter for Autonomous Modbus Discovery.
Generates target-certified definition CSV files for WebdynSunPM and Webdyn Classic.
Preserves 11-column structure, sequential 1-based indexing, decimal coefficients, and UTF-8 encoding.
"""

from __future__ import annotations

import csv
import io

from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    Endianness,
    ExportTarget,
    RegisterDataType,
)


class DefinitionExporter:
    """Compiles canonical equipment representations into target CSV formats."""

    def export(
        self,
        model: CanonicalEquipmentModel,
        target: ExportTarget = ExportTarget.WEBDYN_SUNPM,
    ) -> str:
        """
        Exports CanonicalEquipmentModel to standard 11-column CSV format.
        """
        output = io.StringIO()
        writer = csv.writer(output, delimiter=";", lineterminator="\n")

        # 1. Header Row
        transport_str = (
            model.transport.value if hasattr(model.transport, "value") else str(model.transport)
        )
        header = [
            transport_str,
            model.category or "Inverter",
            model.manufacturer or "Unknown",
            model.model or "Generic",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ]
        writer.writerow(header)

        # 2. Variable Rows
        seen_names: dict[str, int] = {}

        for idx, reg in enumerate(model.registers, start=1):
            # Target Data Type with Endianness formatting
            dtype_str = self._format_data_type(reg.data_type, reg.endianness, target)

            # Deterministic duplicate name deduplication
            base_name = reg.name.strip()
            count = seen_names.get(base_name, 0)
            seen_names[base_name] = count + 1
            unique_name = base_name if count == 0 else f"{base_name}_{count + 1}"

            # Tag Attribution Policy: Decision 3 confirmed blank unless 100% verified canonical
            tag_str = reg.canonical_sunpm_tag or ""

            # Explicit 6-decimal coefficient formatting
            coef_a_str = f"{reg.factor:.6f}"
            coef_b_str = f"{reg.offset:.6f}"

            row = [
                str(idx),
                str(reg.read_fc),
                str(reg.address),
                dtype_str,
                "",  # Info4
                unique_name,
                tag_str,
                coef_a_str,
                coef_b_str,
                reg.unit or "",
                str(reg.action_code),
            ]
            writer.writerow(row)

        return output.getvalue()

    def _format_data_type(
        self,
        dtype: RegisterDataType,
        endianness: Endianness,
        target: ExportTarget,
    ) -> str:
        base = dtype.value
        # Word swapped Little-Endian formatting for 32-bit types in WebdynSunPM
        if endianness in (Endianness.LITTLE_ENDIAN, Endianness.MID_LITTLE_ENDIAN):
            if base in ("U32", "S32", "FLOAT32"):
                return f"{base}_W"
        return base
