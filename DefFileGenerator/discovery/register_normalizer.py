"""
Modbus Semantic Normalizer for Autonomous Modbus Discovery.
Transforms raw extracted document tables into typed CanonicalEquipmentModel instances.
Handles radix parsing, 0-based vs 1-based conventions, data types, word order, and SunPM tags.
"""

from __future__ import annotations

import logging
import re

from DefFileGenerator.discovery.document_extractor import ExtractedTable, disambiguate_ocr_token
from DefFileGenerator.discovery.models import (
    CanonicalEquipmentModel,
    CanonicalRegisterRecord,
    Endianness,
    IdentityCandidate,
    ProtocolTransport,
    ProvenanceRecord,
    RegisterDataType,
    RegisterSpace,
)
from DefFileGenerator.knowledge_miner import infer_sunpm_tag, is_canonical_sunpm_tag

logger = logging.getLogger(__name__)

TYPE_MAP: dict[str, tuple[RegisterDataType, Endianness, int]] = {
    "u16": (RegisterDataType.U16, Endianness.BIG_ENDIAN, 1),
    "uint16": (RegisterDataType.U16, Endianness.BIG_ENDIAN, 1),
    "s16": (RegisterDataType.S16, Endianness.BIG_ENDIAN, 1),
    "int16": (RegisterDataType.S16, Endianness.BIG_ENDIAN, 1),
    "i16": (RegisterDataType.S16, Endianness.BIG_ENDIAN, 1),
    "u32": (RegisterDataType.U32, Endianness.BIG_ENDIAN, 2),
    "uint32": (RegisterDataType.U32, Endianness.BIG_ENDIAN, 2),
    "s32": (RegisterDataType.S32, Endianness.BIG_ENDIAN, 2),
    "int32": (RegisterDataType.S32, Endianness.BIG_ENDIAN, 2),
    "i32": (RegisterDataType.S32, Endianness.BIG_ENDIAN, 2),
    "u32_w": (RegisterDataType.U32, Endianness.LITTLE_ENDIAN, 2),
    "s32_w": (RegisterDataType.S32, Endianness.LITTLE_ENDIAN, 2),
    "float32": (RegisterDataType.FLOAT32, Endianness.BIG_ENDIAN, 2),
    "float": (RegisterDataType.FLOAT32, Endianness.BIG_ENDIAN, 2),
    "float32_w": (RegisterDataType.FLOAT32, Endianness.LITTLE_ENDIAN, 2),
    "u64": (RegisterDataType.U64, Endianness.BIG_ENDIAN, 4),
    "string": (RegisterDataType.STRING, Endianness.BIG_ENDIAN, 1),
    "bitfield16": (RegisterDataType.BITFIELD16, Endianness.BIG_ENDIAN, 1),
    "bitfield32": (RegisterDataType.BITFIELD32, Endianness.BIG_ENDIAN, 2),
}


class RegisterNormalizer:
    """Normalizes raw extracted tables into canonical Modbus equipment definitions."""

    def __init__(self, tag_attribution_strict: bool = True) -> None:
        self.tag_attribution_strict = tag_attribution_strict

    def normalize(
        self,
        tables: list[ExtractedTable],
        identity: IdentityCandidate,
        document_sha256: str = "",
        document_title: str = "",
    ) -> CanonicalEquipmentModel:
        """
        Processes extracted tables into a single CanonicalEquipmentModel.
        """
        canonical_model = CanonicalEquipmentModel(
            manufacturer=identity.manufacturer,
            model=identity.exact_model,
            hardware_revisions=[identity.hardware_revision] if identity.hardware_revision else [],
            supported_firmwares=[identity.firmware_version] if identity.firmware_version else [],
            transport=ProtocolTransport.MODBUS_RTU,
            category="Inverter",
            registers=[],
        )

        seen_addresses: set[tuple[int, int]] = set()  # (fc, address)

        for table in tables:
            headers = [h.lower() for h in table.headers]
            addr_col = self._find_column(headers, ["address", "adresse", "reg", "register", "modbus address", "offset", "register address"])
            name_col = self._find_column(headers, ["name", "nom", "parameter", "paramètre", "description", "signal", "point name", "variable"])
            type_col = self._find_column(headers, ["type", "data type", "datatype", "format"])
            scale_col = self._find_column(headers, ["scale", "gain", "ratio", "factor", "facteur", "coeff", "multiplier"])
            unit_col = self._find_column(headers, ["unit", "unité", "units"])
            action_col = self._find_column(headers, ["action", "r/w", "access", "rw", "read/write"])

            if addr_col is None or name_col is None:
                # If cannot identify address/name columns, record as unresolved table
                canonical_model.unresolved_rows.append({"table": table.dict(), "reason": "Missing address or name column"})
                continue

            for _row_idx, row in enumerate(table.rows):
                if len(row) <= max(addr_col, name_col):
                    continue

                raw_addr = row[addr_col].strip()
                raw_name = row[name_col].strip()
                if not raw_addr or not raw_name:
                    continue

                # Ignore section dividers or repeated header rows
                if raw_addr.lower() in ("address", "adresse", "reg", "register") or raw_name.lower() in ("name", "description", "parameter"):
                    continue

                # 1. Address Parsing and Disambiguation
                cleaned_addr_str = disambiguate_ocr_token(raw_addr, is_address=True) if table.is_ocr else raw_addr
                parsed_addr, addr_base_prov = self._parse_address(cleaned_addr_str)
                if parsed_addr is None or parsed_addr < 0 or parsed_addr > 65535:
                    canonical_model.unresolved_rows.append({"row": row, "reason": f"Invalid address '{raw_addr}'"})
                    continue

                # 2. Type & Endianness Parsing
                raw_type = row[type_col].strip() if type_col is not None and len(row) > type_col else "U16"
                dtype, endianness, reg_count = self._parse_type(raw_type)

                # 3. Scaling / Factor & Unit Parsing
                raw_scale = row[scale_col].strip() if scale_col is not None and len(row) > scale_col else "1"
                factor, offset = self._parse_scaling(raw_scale)
                unit = row[unit_col].strip() if unit_col is not None and len(row) > unit_col else ""

                # 4. Action & Function Code
                raw_action = row[action_col].strip() if action_col is not None and len(row) > action_col else "R"
                is_writable = any(w in raw_action.lower() for w in ("w", "write", "r/w", "rw"))
                action_code = self._determine_action_code(raw_name, is_writable)

                # 5. SunPM Tag Attribution (Decision 3: strict canonical only or blank)
                sunpm_tag: str | None = None
                inferred = infer_sunpm_tag(raw_name, equipment_type=canonical_model.category)
                if inferred and is_canonical_sunpm_tag(inferred):
                    sunpm_tag = inferred

                provenance = ProvenanceRecord(
                    document_sha256=document_sha256,
                    document_title=document_title or "Extracted Manual",
                    page_number=table.page_number,
                    table_index=table.table_index,
                    sheet_name=table.sheet_name,
                    raw_text_snippet=f"{raw_addr} | {raw_name} | {raw_type}",
                    extraction_confidence=table.confidence,
                )

                reg_record = CanonicalRegisterRecord(
                    id=f"REG_{parsed_addr}_{dtype.value}",
                    name=raw_name,
                    address=parsed_addr,
                    address_base_provenance=addr_base_prov,
                    register_space=RegisterSpace.HOLDING_REGISTER,
                    read_fc=3,
                    write_fc=16 if is_writable else None,
                    data_type=dtype,
                    register_count=reg_count,
                    endianness=endianness,
                    factor=factor,
                    offset=offset,
                    unit=unit,
                    canonical_sunpm_tag=sunpm_tag,
                    action_code=action_code,
                    is_writable=is_writable,
                    provenance=provenance,
                )

                # Deduplicate by address
                key = (3, parsed_addr)
                if key not in seen_addresses:
                    seen_addresses.add(key)
                    canonical_model.registers.append(reg_record)

        # Sort registers by address ascending
        canonical_model.registers.sort(key=lambda r: r.address)
        canonical_model.total_registers = len(canonical_model.registers)

        return canonical_model

    def _find_column(self, headers: list[str], candidates: list[str]) -> int | None:
        for idx, h in enumerate(headers):
            for cand in candidates:
                if cand in h:
                    return idx
        return None

    def _parse_address(self, addr_str: str) -> tuple[int | None, str]:
        addr_clean = addr_str.strip()
        # Check hex notation (0x... or ...H)
        hex_match = re.match(r"^(?:0x([0-9a-fA-F]+)|([0-9a-fA-F]+)h)$", addr_clean, re.IGNORECASE)
        if hex_match:
            hex_val = hex_match.group(1) or hex_match.group(2)
            try:
                return int(hex_val, 16), "hex_radix"
            except ValueError:
                return None, "invalid"

        # Check numeric
        cleaned_num = re.sub(r"[^\d]", "", addr_clean)
        if not cleaned_num:
            return None, "invalid"

        val = int(cleaned_num)
        # Check standard 5-digit PLC addressing (40001 -> 0, 30001 -> 0)
        if 40001 <= val <= 49999:
            return val - 40001, "inferred_plc_4xxxx"
        if 30001 <= val <= 39999:
            return val - 30001, "inferred_plc_3xxxx"

        return val, "documented_0_based"

    def _parse_type(self, raw_type: str) -> tuple[RegisterDataType, Endianness, int]:
        low = raw_type.lower().strip()
        for key, val in TYPE_MAP.items():
            if key in low:
                return val
        return RegisterDataType.U16, Endianness.BIG_ENDIAN, 1

    def _parse_scaling(self, raw_scale: str) -> tuple[float, float]:
        clean = raw_scale.strip().replace(",", ".")
        # Fraction check like 1/100
        if "/" in clean:
            parts = clean.split("/")
            try:
                num = float(parts[0])
                den = float(parts[1])
                if den != 0:
                    return num / den, 0.0
            except ValueError:
                pass

        try:
            val = float(clean)
            if val != 0:
                # If documented as 10 or 100 multiplier, check if it means 0.1 or 0.01
                if val in (10.0, 100.0, 1000.0):
                    return 1.0 / val, 0.0
                return val, 0.0
        except ValueError:
            pass

        return 1.0, 0.0

    def _determine_action_code(self, name: str, is_writable: bool) -> int:
        n_low = name.lower()
        if any(a in n_low for a in ("alarm", "fault", "error", "warning", "défaut")):
            return 1
        if any(e in n_low for e in ("status", "state", "event", "statut", "mode")):
            return 2
        if is_writable or any(p in n_low for p in ("setpoint", "setting", "limit", "param", "consigne")):
            return 3
        return 4
