"""
SunSpec Definition File Generator.

Compiles resolved SunSpec models and points into strictly valid Webdyn 11-column
definition files compliant with the Apollo 5 Pillars:
  - Exact Modbus address chaining: next_addr = current_addr + 2 + L
  - Dynamic support for Common Model 1 length variants (L=65 vs L=66)
  - Base 0 (40000) and Base 1 (40001) addressing
  - Accurate Fixed Scale Factor application (10^SF) or explicit Dynamic SF exposure
  - Strict filtering of padding registers and unimplemented points
  - Expanding repeating blocks for Model 160 (MPPT strings) and Model 714 (DC ports)
  - Documenting sentinel null-values (0x8000, 0xFFFF, 0x80000000, 0xFFFFFFFF)
  - Atomic persistence via staged_text_output with CRLF and UTF-8 without BOM
  - Validation against production_governor.verify_file_pillars
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from DefFileGenerator.io_utils import staged_text_output
from DefFileGenerator.production_governor import verify_file_pillars
from DefFileGenerator.sunspec.constants import (
    BASE_ADDRESS_0,
    BASE_ADDRESS_1,
    COMMON_MODEL_LEN_ALIGNED,
    COMMON_MODEL_LEN_STANDARD,
    SENTINEL_VALUES,
)
from DefFileGenerator.sunspec.evaluator import SunSpecGenerationTarget
from DefFileGenerator.sunspec.models_registry import (
    SUNSPEC_MODELS_REGISTRY,
    SunSpecModelDef,
    SunSpecPoint,
    get_sunspec_model,
)
from DefFileGenerator.sunspec.pics_parser import PICSDocument

logger = logging.getLogger(__name__)

# SunSpec primitive data type to Webdyn column 3 type mapping
TYPE_MAPPING: dict[str, str] = {
    "uint16": "U16",
    "int16": "I16",
    "uint32": "U32",
    "int32": "I32",
    "uint64": "U64",
    "int64": "I64",
    "float32": "F32",
    "sunssf": "I16",
    "enum16": "U16",
    "bitfield16": "U16",
    "bitfield32": "U32",
    "string": "STRING",
    "string8": "STRING",
    "string16": "STRING",
    "string32": "STRING",
}


@dataclass
class SunSpecGenerationResult:
    """Outcome of generating a single SunSpec definition file."""
    output_path: Path
    register_count: int
    is_valid: bool
    issues: list[str] = field(default_factory=list)
    applied_fixed_scale_factors: dict[str, float] = field(default_factory=dict)
    dynamic_scale_factors: list[str] = field(default_factory=list)
    excluded_unimplemented_count: int = 0
    warnings: list[str] = field(default_factory=list)


class SunSpecDefinitionGenerator:
    """Compiles SunSpec models into 11-column Webdyn definition CSV files."""

    def generate(
        self,
        target: SunSpecGenerationTarget,
        output_dir: str | Path = "production",
        pics_doc: PICSDocument | None = None,
        force_overwrite: bool = False,
    ) -> SunSpecGenerationResult:
        """
        Generates a validated Webdyn definition CSV from a SunSpecGenerationTarget.

        Args:
            target: The targeted generation specification.
            output_dir: Output directory (e.g. 'production' or 'generated_definitions').
            pics_doc: Optional parsed PICS document providing implementation flags & fixed SF.
            force_overwrite: Overwrites destination if it exists.

        Returns:
            SunSpecGenerationResult with status, counts, and validation outcome.
        """
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        filename = target.output_filename or target.get_default_filename()
        out_file = out_dir / filename

        if out_file.exists() and not force_overwrite:
            logger.info(f"Skipping existing SunSpec definition: {out_file.name}")
            return SunSpecGenerationResult(
                output_path=out_file,
                register_count=0,
                is_valid=True,
                warnings=[f"File {out_file.name} already exists. Use force_overwrite=True."],
            )

        # 1. Compile register rows
        rows, meta = self._compile_rows(target, pics_doc)

        if not rows:
            raise ValueError(f"No valid SunSpec registers could be compiled for target {filename}")

        # 2. Write CSV atomically using staged_text_output
        self._write_csv(out_file, target, rows)

        # 3. Rigorous validation against Apollo 5 Pillars
        is_valid, issues = verify_file_pillars(out_file, strict_crlf=True)
        if not is_valid:
            logger.error(f"SunSpec generation validation failed for {out_file.name}: {issues}")

        return SunSpecGenerationResult(
            output_path=out_file,
            register_count=len(rows),
            is_valid=is_valid,
            issues=issues,
            applied_fixed_scale_factors=meta["applied_fixed_sf"],
            dynamic_scale_factors=meta["dynamic_sf"],
            excluded_unimplemented_count=meta["excluded_unimplemented"],
            warnings=meta["warnings"],
        )

    def _compile_rows(
        self,
        target: SunSpecGenerationTarget,
        pics_doc: PICSDocument | None,
    ) -> tuple[list[list[str]], dict[str, Any]]:
        """
        Compiles the complete sequential row set for all models in target.
        Enforces strict address chaining: next_addr = current_addr + 2 + L.
        """
        rows: list[list[str]] = []
        applied_fixed_sf: dict[str, float] = {}
        dynamic_sf: list[str] = []
        excluded_unimplemented = 0
        warnings: list[str] = []
        used_tags: set[str] = set()

        base_addr = target.base_address
        current_model_addr = base_addr + 2  # Model 1 starts 2 registers after base address (SunS marker)
        l_common = target.common_model_length

        read_fc = "3" if base_addr in (BASE_ADDRESS_0, BASE_ADDRESS_1) else "4"

        for mid in target.model_ids:
            mdef = get_sunspec_model(mid)
            if mdef is None:
                warnings.append(f"Model ID {mid} is not registered in official registry. Skipping.")
                continue

            num_repeats = target.repeating_counts.get(mid, 0)
            if num_repeats == 0 and mid == 160:
                num_repeats = 2  # Default to 2 MPPT strings if not specified
            elif num_repeats == 0 and mid == 714:
                num_repeats = 2  # Default to 2 DC ports if not specified

            # Calculate actual model length L
            if mid == 1:
                actual_l = l_common
            elif mdef.is_variable_length:
                actual_l = mdef.standard_length + num_repeats * mdef.repeating_block_length
            else:
                actual_l = mdef.standard_length

            # Get points resolved for this model
            points = mdef.get_points(num_repeats=num_repeats, l_common=l_common)

            # Build scale factor lookup for this model from PICS
            model_fixed_sf: dict[str, int | float] = {}
            if pics_doc:
                for (doc_mid, sf_sym), val in pics_doc.fixed_scale_factors.items():
                    if doc_mid == mid:
                        model_fixed_sf[sf_sym] = val

            # First add Model ID and Length L registers for observability
            # Model ID point (at current_model_addr)
            rows.append(
                [
                    str(len(rows) + 1),
                    read_fc,
                    str(current_model_addr),
                    "U16",
                    "",
                    f"SunSpec Model {mid} ID",
                    "",
                    "1.000000",
                    "0.000000",
                    "",
                    "4",
                ]
            )
            # Model Length L point (at current_model_addr + 1)
            rows.append(
                [
                    str(len(rows) + 1),
                    read_fc,
                    str(current_model_addr + 1),
                    "U16",
                    "",
                    f"SunSpec Model {mid} Length (L={actual_l})",
                    "",
                    "1.000000",
                    "0.000000",
                    "",
                    "4",
                ]
            )

            payload_base_addr = current_model_addr + 2

            for pt in points:
                # 1. Filter padding registers
                if pt.is_pad:
                    continue

                # 2. Filter points marked unimplemented in PICS
                if pics_doc and (mid, pt.symbol) in pics_doc.unimplemented_points:
                    excluded_unimplemented += 1
                    continue

                pt_addr = payload_base_addr + pt.offset

                # Format Webdyn address (add string length suffix if string type)
                addr_str = str(pt_addr)
                if pt.data_type.startswith("string"):
                    byte_len = pt.register_count * 2
                    addr_str = f"{pt_addr}_{byte_len}"

                w_type = TYPE_MAPPING.get(pt.data_type.lower(), "U16")

                # 3. Factor & Scale Factor resolution
                factor_str = "1.000000"
                desc_suffix = ""

                if pt.scale_factor_point:
                    if pt.scale_factor_point in model_fixed_sf:
                        sf_val = model_fixed_sf[pt.scale_factor_point]
                        try:
                            factor_float = 10.0 ** float(sf_val)
                            factor_str = f"{factor_float:.6f}"
                            applied_fixed_sf[f"{mid}.{pt.symbol}"] = factor_float
                        except (ValueError, OverflowError):
                            factor_str = "1.000000"
                    else:
                        # Dynamic scale factor: must be calculated at runtime from SF register
                        factor_str = "1.000000"
                        desc_suffix = f" [Dynamic SF: {pt.scale_factor_point}]"
                        if f"{mid}.{pt.scale_factor_point}" not in dynamic_sf:
                            dynamic_sf.append(f"{mid}.{pt.scale_factor_point}")
                            warnings.append(
                                f"Model {mid} point '{pt.symbol}' references dynamic scale factor "
                                f"'{pt.scale_factor_point}'. Emitted factor 1.000000 and exposed SF register."
                            )

                # 4. Canonical SunPM Tag gatekeeping (unique per file)
                assigned_tag = ""
                if pt.sunpm_tag and pt.sunpm_tag not in used_tags:
                    assigned_tag = pt.sunpm_tag
                    used_tags.add(pt.sunpm_tag)

                # 5. Document sentinel nulls in description
                sentinel_note = ""
                if pt.data_type in SENTINEL_VALUES:
                    s_val = SENTINEL_VALUES[pt.data_type]
                    if isinstance(s_val, int):
                        sentinel_note = f" (Null={s_val})"

                full_desc = f"{pt.description}{desc_suffix}{sentinel_note}".strip()

                action_code = pt.action_code
                if action_code == "3" or (pt.access == "RW" and action_code == "4"):
                    action_code = "8"
                elif action_code not in ("0", "1", "2", "4", "6", "7", "8", "9", "10"):
                    action_code = "4"

                row = [
                    str(len(rows) + 1),  # Sequential index
                    read_fc,             # Function Code (3)
                    addr_str,            # Modbus Address
                    w_type,              # Data Type
                    "",                  # Info4
                    full_desc,           # Label / Name
                    assigned_tag,        # Canonical SunPM Tag
                    factor_str,          # Multiplier Factor (10^SF)
                    "0.000000",          # Additive Offset
                    pt.unit,             # Engineering Unit
                    action_code,         # Action Code (4: Measurement, 8: Param/Setting, 1: Alarm)
                ]
                rows.append(row)

            # Advance to next model start address in the chain
            current_model_addr = current_model_addr + 2 + actual_l

        meta = {
            "applied_fixed_sf": applied_fixed_sf,
            "dynamic_sf": dynamic_sf,
            "excluded_unimplemented": excluded_unimplemented,
            "warnings": warnings,
        }
        return rows, meta

    def _write_csv(
        self,
        out_file: Path,
        target: SunSpecGenerationTarget,
        rows: list[list[str]],
    ) -> None:
        """
        Writes definition CSV file atomically using CRLF (\r\n) and UTF-8 without BOM.
        """
        category_str = "Inverter" if target.prefix == "INV" else ("Meter" if target.prefix == "MET" else "Battery")
        header = f"Modbus;{category_str};{target.manufacturer};{target.model_name};;;;;;;\r\n"

        with staged_text_output(str(out_file), encoding="utf-8") as fp:
            fp.write(header)
            for row in rows:
                fp.write(";".join(row) + "\r\n")
