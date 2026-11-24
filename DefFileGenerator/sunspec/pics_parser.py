"""
SunSpec PICS (Protocol Implementation Conformance Statement) Parser.

Extracts SunSpec model tables, point implementation statuses (supported vs.
unimplemented), fixed scale factors, Common Model 1 length variants (L=65 vs L=66),
and repeating block counts (Model 160 MPPT strings, Model 714 DC ports) from:
  1. SunSpec Dashboard Excel PICS files (.xlsx, .xlsm, .xltx)
  2. Proprietary vendor SunSpec implementation workbooks (ABB, Delta, GoodWe, Solis, etc.)
  3. Intermediate register CSV maps and structured dictionaries
"""

from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from DefFileGenerator.sunspec.constants import (
    BASE_ADDRESS_0,
    BASE_ADDRESS_1,
    COMMON_MODEL_LEN_ALIGNED,
    COMMON_MODEL_LEN_STANDARD,
    SUNS_MAGIC_HEX,
    SUNS_MAGIC_STR,
)

logger = logging.getLogger(__name__)

RE_NUMERIC_SHEET = re.compile(r"^(?:model[_\s]*)?(\d{1,5})$", re.IGNORECASE)
RE_MODULE_POINT = re.compile(r"^(?:module|string|port|mppt)[._\s]*(\d+)[._\s]+(.+)$", re.IGNORECASE)

UNIMPLEMENTED_KEYWORDS = frozenset(
    {
        "unimplemented",
        "not supported",
        "not_supported",
        "unsupported",
        "n/a",
        "na",
        "no",
        "false",
        "disabled",
        "reserved",
        "none",
    }
)


@dataclass
class PICSPoint:
    """Parsed point implementation descriptor from a PICS or register map."""
    offset: int
    address: int
    symbol: str
    label: str = ""
    data_type: str = ""
    units: str = ""
    scale_factor_point: str | None = None
    value: Any = None
    is_supported: bool = True
    fixed_scale_factor: int | float | None = None


@dataclass
class PICSModel:
    """Parsed SunSpec model sheet or section."""
    model_id: int
    address: int
    length: int
    name: str = ""
    points: dict[str, PICSPoint] = field(default_factory=dict)
    repeating_count: int | None = None
    is_supported: bool = True

    @property
    def supported_points_count(self) -> int:
        return sum(1 for p in self.points.values() if p.is_supported)


@dataclass
class PICSDocument:
    """Aggregated parsed information from a SunSpec PICS or register document."""
    file_path: Path | None = None
    manufacturer: str = ""
    product_name: str = ""
    model_name: str = ""
    category: str = "Inverter"
    base_address: int = BASE_ADDRESS_0
    common_model_length: int = COMMON_MODEL_LEN_STANDARD
    has_suns_marker: bool = False
    models: dict[int, PICSModel] = field(default_factory=dict)
    unimplemented_points: set[tuple[int, str]] = field(default_factory=set)
    fixed_scale_factors: dict[tuple[int, str], int | float] = field(default_factory=dict)
    repeating_counts: dict[int, int] = field(default_factory=dict)
    raw_metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def has_models(self) -> bool:
        return bool(self.models)

    @property
    def model_ids(self) -> list[int]:
        return sorted(self.models.keys())

    @property
    def has_der700(self) -> bool:
        return any(701 <= mid <= 714 for mid in self.models)

    @property
    def has_legacy_inverter(self) -> bool:
        return any(mid in (101, 102, 103) for mid in self.models)

    @property
    def has_storage(self) -> bool:
        return any(mid in (124, 802, 803, 804, 805) for mid in self.models)

    @property
    def has_meter(self) -> bool:
        return any(mid in (201, 202, 203, 204, 211, 212, 213, 214) for mid in self.models)


class PICSParser:
    """Parser for SunSpec PICS documents and Excel workbooks."""

    def parse_file(self, file_path: str | Path) -> PICSDocument:
        """Parses a PICS document file (.xlsx, .xlsm, .csv, .json)."""
        p = Path(file_path)
        if not p.is_file():
            return PICSDocument(file_path=p)

        ext = p.suffix.lower()
        if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
            return self._parse_excel(p)
        elif ext == ".csv":
            return self._parse_csv(p)
        elif ext == ".json":
            return self._parse_json(p)
        else:
            logger.debug(f"PICSParser skipping non-PICS extension: {ext}")
            return PICSDocument(file_path=p)

    def _parse_excel(self, file_path: Path) -> PICSDocument:
        try:
            import openpyxl
        except ImportError:
            logger.error("openpyxl is required to parse Excel PICS workbooks.")
            return PICSDocument(file_path=file_path)

        doc = PICSDocument(file_path=file_path)

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        except Exception as e:
            logger.warning(f"Could not open workbook {file_path}: {e}")
            return doc

        # 1. Parse Metadata Sheets
        self._extract_metadata(wb, doc)

        # 2. Parse Model Sheets
        for sheet_name in wb.sheetnames:
            m = RE_NUMERIC_SHEET.match(sheet_name.strip())
            if not m:
                continue

            try:
                model_id = int(m.group(1))
            except ValueError:
                continue

            ws = wb[sheet_name]
            pics_model = self._parse_model_sheet(ws, model_id, doc)
            if pics_model is not None:
                doc.models[model_id] = pics_model

        # 3. Reconcile Global Properties
        if 1 in doc.models:
            m1 = doc.models[1]
            if m1.length in (COMMON_MODEL_LEN_STANDARD, COMMON_MODEL_LEN_ALIGNED):
                doc.common_model_length = m1.length
            # If address of Model 1 is 40002, base address is 40000; if 40003, base is 40001
            if m1.address == 40002:
                doc.base_address = BASE_ADDRESS_0
                doc.has_suns_marker = True
            elif m1.address == 40003:
                doc.base_address = BASE_ADDRESS_1
                doc.has_suns_marker = True
            elif m1.address == 30002:
                doc.base_address = 30000
                doc.has_suns_marker = True

            # Extract manufacturer / model strings if present in Model 1 points
            if not doc.manufacturer and "Mn" in m1.points:
                val = m1.points["Mn"].value
                if val and str(val).strip().lower() not in UNIMPLEMENTED_KEYWORDS:
                    doc.manufacturer = str(val).strip()
            if not doc.model_name and "Md" in m1.points:
                val = m1.points["Md"].value
                if val and str(val).strip().lower() not in UNIMPLEMENTED_KEYWORDS:
                    doc.model_name = str(val).strip()

        # Update repeating counts
        for mid, pm in doc.models.items():
            if pm.repeating_count is not None:
                doc.repeating_counts[mid] = pm.repeating_count

        wb.close()
        return doc

    def _extract_metadata(self, wb: Any, doc: PICSDocument) -> None:
        """Extracts brand, model name, and device category from informational sheets."""
        for name in ("Certification Info", "Testing Info", "Summary", "Index", "Instructions"):
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            for row in ws.iter_rows(values_only=True):
                if not row or len(row) < 2:
                    continue
                k = str(row[0]).strip().lower() if row[0] is not None else ""
                v = str(row[1]).strip() if row[1] is not None else ""
                if not k or not v:
                    continue
                doc.raw_metadata[k] = v
                if "manufacturer" in k and not doc.manufacturer:
                    doc.manufacturer = v
                elif ("product" in k or "model" in k or "device" in k) and not doc.model_name:
                    doc.model_name = v
                elif "category" in k or "type" in k:
                    if "meter" in v.lower():
                        doc.category = "Meter"
                    elif "battery" in v.lower() or "storage" in v.lower():
                        doc.category = "Battery"

    def _parse_model_sheet(self, ws: Any, model_id: int, doc: PICSDocument) -> PICSModel | None:
        """Parses an individual model sheet into a PICSModel."""
        header_row_idx = None
        col_map: dict[str, int] = {}

        # Scan first 10 rows for table header
        rows = list(ws.iter_rows(values_only=True, max_row=200))
        for r_idx, row in enumerate(rows):
            if not row:
                continue
            row_str = [str(c).strip().lower() if c is not None else "" for c in row]
            if any("address" in c or "register" in c for c in row_str) and any(
                "point" in c or "name" in c or "symbol" in c or "label" in c for c in row_str
            ):
                header_row_idx = r_idx
                for c_idx, col_name in enumerate(row_str):
                    if "address" in col_name or "register" in col_name:
                        col_map["address"] = c_idx
                    elif "point" in col_name or "symbol" in col_name:
                        col_map["symbol"] = c_idx
                    elif "name" in col_name and "symbol" not in col_map:
                        col_map["symbol"] = c_idx
                    elif "label" in col_name:
                        col_map["label"] = c_idx
                    elif "type" in col_name:
                        col_map["type"] = c_idx
                    elif "value" in col_name and "subvalue" not in col_name:
                        col_map["value"] = c_idx
                    elif "unit" in col_name:
                        col_map["units"] = c_idx
                    elif "sf" in col_name or "scale" in col_name:
                        col_map["sf"] = c_idx
                    elif "r/w" in col_name or "access" in col_name:
                        col_map["rw"] = c_idx
                break

        if header_row_idx is None or "address" not in col_map or "symbol" not in col_map:
            # Fallback for simple positional templates: col0=addr, col1=point, col3=value
            col_map = {"address": 0, "symbol": 1, "value": 3}
            header_row_idx = 0

        model_start_addr: int | None = None
        model_length: int = 0
        points: dict[str, PICSPoint] = {}
        max_repeat_idx = 0

        for row in rows[header_row_idx + 1 :]:
            if not row:
                continue

            addr_raw = row[col_map["address"]] if col_map["address"] < len(row) else None
            sym_raw = row[col_map["symbol"]] if col_map["symbol"] < len(row) else None
            val_raw = row[col_map["value"]] if "value" in col_map and col_map["value"] < len(row) else None
            lbl_raw = row[col_map["label"]] if "label" in col_map and col_map["label"] < len(row) else ""
            type_raw = row[col_map["type"]] if "type" in col_map and col_map["type"] < len(row) else ""
            unit_raw = row[col_map["units"]] if "units" in col_map and col_map["units"] < len(row) else ""
            sf_raw = row[col_map["sf"]] if "sf" in col_map and col_map["sf"] < len(row) else None

            # Skip blank rows or subvalue rows without address
            if addr_raw is None and sym_raw is None:
                continue

            try:
                addr = int(addr_raw) if addr_raw is not None else 0
            except (ValueError, TypeError):
                addr = 0

            symbol = str(sym_raw).strip() if sym_raw is not None else ""
            if not symbol and lbl_raw:
                symbol = str(lbl_raw).strip()
            if not symbol:
                continue

            val_str = str(val_raw).strip().lower() if val_raw is not None else ""
            is_supported = val_str not in UNIMPLEMENTED_KEYWORDS

            # Check header ID and L
            if symbol.upper() == "ID":
                model_start_addr = addr
                continue
            elif symbol.upper() == "L":
                try:
                    model_length = int(val_raw) if val_raw is not None else 0
                except (ValueError, TypeError):
                    model_length = 0
                continue

            # Scale Factor detection
            fixed_sf: int | float | None = None
            if symbol.endswith("_SF") or str(type_raw).lower() == "sunssf":
                try:
                    if val_raw is not None and not isinstance(val_raw, str):
                        fixed_sf = int(val_raw)
                    elif val_str and val_str not in UNIMPLEMENTED_KEYWORDS:
                        fixed_sf = int(val_str)
                except (ValueError, TypeError):
                    pass

                if fixed_sf is not None:
                    doc.fixed_scale_factors[(model_id, symbol)] = fixed_sf

            # Repeating block index tracking (module.1.DCA or port.1.DCV)
            m_repeat = RE_MODULE_POINT.match(symbol)
            if m_repeat:
                rep_idx = int(m_repeat.group(1))
                if rep_idx > max_repeat_idx:
                    max_repeat_idx = rep_idx
                point_symbol = m_repeat.group(2)
            else:
                point_symbol = symbol

            if not is_supported:
                doc.unimplemented_points.add((model_id, symbol))

            pt = PICSPoint(
                offset=0,  # Computed subsequently relative to model_start_addr
                address=addr,
                symbol=symbol,
                label=str(lbl_raw) if lbl_raw else "",
                data_type=str(type_raw) if type_raw else "",
                units=str(unit_raw) if unit_raw else "",
                scale_factor_point=str(sf_raw).strip() if sf_raw else None,
                value=val_raw,
                is_supported=is_supported,
                fixed_scale_factor=fixed_sf,
            )
            points[symbol] = pt

        if model_start_addr is None:
            # Infer from first point address minus 2 (for ID + L)
            first_addr = min((p.address for p in points.values() if p.address > 0), default=40002)
            model_start_addr = max(0, first_addr - 2)

        # Compute offsets for points relative to payload start
        for pt in points.values():
            if pt.address >= model_start_addr + 2:
                pt.offset = pt.address - (model_start_addr + 2)

        repeating_count: int | None = None
        if max_repeat_idx > 0:
            repeating_count = max_repeat_idx
        elif model_id == 160 and model_length >= 28:
            repeating_count = max(0, (model_length - 8) // 20)
        elif model_id == 714 and model_length >= 24:
            repeating_count = max(0, (model_length - 8) // 16)

        return PICSModel(
            model_id=model_id,
            address=model_start_addr,
            length=model_length,
            points=points,
            repeating_count=repeating_count,
            is_supported=any(p.is_supported for p in points.values()),
        )

    def _parse_csv(self, file_path: Path) -> PICSDocument:
        """Parses CSV representations of PICS or register mappings."""
        doc = PICSDocument(file_path=file_path)
        try:
            with open(file_path, encoding="utf-8", errors="replace") as f:
                lines = f.readlines()
        except Exception as e:
            logger.warning(f"Failed to read CSV {file_path}: {e}")
            return doc

        # Look for SunS magic marker or Model ID headers
        for line in lines:
            line_str = line.strip()
            if SUNS_MAGIC_STR in line_str or str(SUNS_MAGIC_HEX) in line_str:
                doc.has_suns_marker = True
                break

        return doc

    def _parse_json(self, file_path: Path) -> PICSDocument:
        """Parses JSON or SMDX model descriptions."""
        import json

        doc = PICSDocument(file_path=file_path)
        try:
            with open(file_path, encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            logger.warning(f"Failed to read JSON {file_path}: {e}")
            return doc

        if isinstance(data, dict):
            if "manufacturer" in data:
                doc.manufacturer = str(data["manufacturer"])
            if "model" in data:
                doc.model_name = str(data["model"])
            if "models" in data and isinstance(data["models"], list):
                for m in data["models"]:
                    if isinstance(m, dict) and "id" in m:
                        mid = int(m["id"])
                        doc.models[mid] = PICSModel(
                            model_id=mid,
                            address=int(m.get("address", 40002)),
                            length=int(m.get("length", 0)),
                        )
        return doc
