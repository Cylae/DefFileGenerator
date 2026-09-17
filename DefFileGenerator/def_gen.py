#!/usr/bin/env python3
import argparse
import csv
import functools
import itertools
import logging
import math
import os
import re
import sys
import tempfile
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import Any, Optional, Union


def peek_generator(iterable: Optional[Iterable]) -> tuple[bool, Iterator]:
    """
    Checks if an iterable is non-empty without fully consuming it.
    Returns (has_data, original_iterator).
    """
    if iterable is None:
        return False, iter([])
    it = iter(iterable)
    try:
        first = next(it)
    except StopIteration:
        return False, iter([])
    return True, itertools.chain([first], it)


# Pre-compiled regex patterns for optimization
RE_TYPE_NUMERIC = re.compile(r"^([UI](8|16|32|64)|F(32|64))(_(W|B|WB))?$", re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r"^STR(\d+)$", re.IGNORECASE)
RE_ADDR_STRING = re.compile(r"^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)$", re.IGNORECASE)
RE_ADDR_BITS = re.compile(r"^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)_(\d+)$", re.IGNORECASE)
RE_ADDR_INT = re.compile(r"^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)$", re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r"^([UI](16|8)(_(W|B|WB))?|BITS)$", re.IGNORECASE)
RE_COUNT_32 = re.compile(r"^([UI]32(_(W|B|WB))?|F32(_(W|B|WB))?|IP)$", re.IGNORECASE)
RE_COUNT_64 = re.compile(r"^([UI]64(_(W|B|WB))?|F64(_(W|B|WB))?)$", re.IGNORECASE)

_CLEAN_TYPE_RE = re.compile(r"[^a-z0-9_]+")


@dataclass
class CSVHeaderConfig:
    manufacturer: str = ""
    model: str = ""
    protocol: str = "modbusRTU"
    category: str = "Inverter"
    forced_write: str = ""


@dataclass
class GeneratorConfig:
    input_file: Optional[str] = None
    output: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    protocol: str = "modbusRTU"
    category: str = "Inverter"
    forced_write: str = ""
    template: bool = False
    template_mode: str = "input"  # 'input' or 'definition'
    address_offset: int = 0


@dataclass
class ValidationIssue:
    line: int
    severity: str  # "ERROR" or "WARNING"
    code: str
    field: str
    message: str


@dataclass
class ValidationReport:
    is_valid: bool
    register_count: int
    issues: list[ValidationIssue]
    stats: dict[str, Any]


class Generator:
    def __init__(self, strict: bool = False) -> None:
        self.register_type_map = {
            "coil": "1",
            "coils": "1",
            "discrete input": "2",
            "discrete register": "2",
            "discrete registers": "2",
            "discrete": "2",
            "holding register": "3",
            "holding": "3",
            "input register": "4",
            "input": "4",
        }
        self.allowed_actions = ["0", "1", "2", "4", "6", "7", "8", "9"]

    @staticmethod
    def sanitize_csv_field(val: Any) -> str:
        """Precludes CSV injection by prepending an apostrophe if needed."""
        if val is None:
            return ""
        if isinstance(val, (int, float)):
            return str(val)
        s = str(val)
        if not s:
            return ""
        s = "".join(
            ch
            for ch in s
            if (ord(ch) in (9, 10, 13) or (ord(ch) >= 32 and ord(ch) != 127))
            and not (0xD800 <= ord(ch) <= 0xDFFF)
        )
        if not s:
            return ""

        # Check for prefix characters that trigger injection
        if s[0] in (
            "\t",
            "\r",
            "\n",
            "\u00a0",
            "\ufeff",
            "\u200b",
            "\u200e",
            "\u200f",
            "\u2028",
            "\u2029",
            "\u0085",
        ) or s.startswith(("\uff1d", "\uff0b", "\uff0d", "\uff20")):
            prefix = s[0]
            rest = s[1:].replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
            return "'" + prefix + rest

        # Replace internal newlines and carriage returns with a space to prevent multiline CSV records
        s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")

        stripped = s.lstrip(" \t\n\r\v\f\u00a0\ufeff\u200b\u200e\u200f\u2028\u2029\u0085")
        if not stripped:
            return s

        if stripped[0] in ("=", "+", "-", "@", "|", "%", "\uff1d", "\uff0b", "\uff0d", "\uff20"):
            if "_" in stripped or "inf" in stripped.lower() or "nan" in stripped.lower():
                return "'" + s
            try:
                f = float(stripped)
                if math.isfinite(f):
                    if s.startswith(" ") and not s.startswith("  ") and stripped.startswith("-"):
                        return "'" + s
                    return s
            except ValueError:
                pass
            return "'" + s

        return s

    @staticmethod
    @functools.lru_cache(maxsize=2048)
    def normalize_type(dtype: Any) -> str:
        if not dtype:
            return "U16"
        t = str(dtype).lower().strip()
        # Collapse internal whitespace for fragmented PDF font extractions (e.g. "u1 6" -> "u16", "st r" -> "str")
        compact_t = re.sub(r"\s+", "", t)
        if compact_t in ("str", "string"):
            return "STRING"
        if compact_t in ("u16", "u32", "i16", "i32", "s16", "s32", "u64", "i64", "f32", "f64"):
            if compact_t.startswith("s"):
                return f"I{compact_t[1:]}".upper()
            return compact_t.upper()

        for _ in range(2):
            t = re.sub(r"([a-zA-Z0-9])\s+(\d+)\b", r"\1\2", t)

        suffix = ""
        if any(x in t for x in ["_wb", "swap", "big endian"]):
            suffix = "_WB"
        elif any(x in t for x in ["_b", "big"]):
            suffix = "_B"
        elif any(x in t for x in ["_w", "word"]):
            suffix = "_W"

        # Handle "string 20" -> "STR20"
        str_match = re.search(r"string\s*(\d+)", t)
        if str_match:
            return f"STR{str_match.group(1)}{suffix}"

        # Manufacturer tables (e.g. Huawei) often use a bare "STR" type code with
        # the actual character/byte length given in a separate Length/Quantity
        # column rather than embedded in the type itself.
        if t == "str":
            return "STRING"

        # A full-register status/alarm bitmask ("Bitfield16", "Bitfield1 6" once a
        # PDF line-wrap has been collapsed, "Bitfield32", ...) is transported as a
        # plain unsigned integer of matching width, not the discrete "BITS" type.
        clean_bitfield_t = re.sub(r"_(wb|b|w)$|\b(swap|big endian|big|word)\b", "", t).replace(
            " ", ""
        )
        bitfield_match = re.match(r"^bitfield(8|16|32|64)$", clean_bitfield_t)
        if bitfield_match:
            return f"U{bitfield_match.group(1)}{suffix}"

        # Mapping ordered by specificity
        synonyms = [
            (r"unsigned\s*(?:int(?:eger)?)?\s*64|uint64|\bu64\b|\bint64u\b", "U64"),
            (r"signed\s*(?:int(?:eger)?)?\s*64|sint64|\bint64\b|\bi64\b|\bs64\b|\bint64s\b", "I64"),
            (r"unsigned\s*(?:int(?:eger)?)?\s*32|uint32|\bu32\b|\bint32u\b", "U32"),
            (r"signed\s*(?:int(?:eger)?)?\s*32|sint32|\bint32\b|\bi32\b|\bs32\b|\bint32s\b", "I32"),
            (r"unsigned\s*(?:int(?:eger)?)?\s*16|uint16|\bu16\b|\bint16u\b", "U16"),
            (r"signed\s*(?:int(?:eger)?)?\s*16|sint16|\bint16\b|\bi16\b|\bs16\b|\bint16s\b", "I16"),
            (r"unsigned\s*(?:int(?:eger)?)?\s*8|uint8|\bu8\b|\bint8u\b", "U8"),
            (r"signed\s*(?:int(?:eger)?)?\s*8|sint8|\bint8\b|\bi8\b|\bs8\b|\bint8s\b", "I8"),
            (r"float64|double|\bf64\b|64\s*[-_]?\s*bit\s*ieee\s*[-_]?\s*754", "F64"),
            (r"float32|float|\bf32\b|(?:32\s*[-_]?\s*bit\s*)?ieee\s*[-_]?\s*754", "F32"),
            (r"\b(bit32|bitmap32|bits32|bit16|bitmap16|bits16)\b", "BITS"),
            (r"^(?:32\s*[-_]?\s*bit\s*)hex$|^hex32$", "U32"),
            (r"^(?:16\s*[-_]?\s*bit\s*)?hex(?:16)?$", "U16"),
            (r"^unsigned\s*(?:int(?:eger)?)?$|^uint$|^unsigned$", "U16"),
            (r"^signed\s*(?:int(?:eger)?)?$|^sint$|^int$", "I16"),
            (r"\bdate\s*time\b|\bdatetime\b", "U32"),
            (r"\b(?:ip4|ipv4)\b", "IP"),
            (r"\b(?:ip6|ipv6)\b", "IPV6"),
        ]
        for pattern, replacement in synonyms:
            if re.search(pattern, t):
                return f"{replacement}{suffix}"

        str_pattern_match = re.match(r"^(?:string|str)[\*x_]?\s*(\d+)$", t)
        if str_pattern_match:
            return f"STR{str_pattern_match.group(1)}"

        if re.search(r"string", t):
            return f"STRING{suffix}"

        if t.startswith("str") and t[3:].isdigit():
            return t.upper()

        t = _CLEAN_TYPE_RE.sub("", t)
        return t.upper() if t else "U16"

    @staticmethod
    def validate_type(dtype: str) -> bool:
        dtype_upper = str(dtype).upper()
        if dtype_upper in ["STRING", "BITS", "IP", "IPV6", "MAC"]:
            return True
        if RE_TYPE_NUMERIC.match(dtype_upper):
            return True
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return True
        return False

    @staticmethod
    def normalize_address_val(addr_part: Any) -> str:
        if isinstance(addr_part, int):
            return str(addr_part)
        s = str(addr_part).strip()
        if not s:
            return ""

        # Handle range notation if present: e.g. "31657~31658", "0x8232 ~ 0x82FE", "40001-40002", "30001..30002"
        # Extract the start address before the range separator
        range_match = re.match(r"^([^\s~.]+?)\s*(?:~|\.\.|\s+-\s+|-(?=[0-9a-fA-FxX]))", s)
        if range_match:
            s = range_match.group(1).strip()

        if s.isdigit() and (not s.startswith("0") or s == "0"):
            return s
        addr_part = s
        addr_part = re.sub(r"(?<=\d),(?=\d{3}(?!\d))", "", addr_part)
        if not addr_part:
            return ""
        is_neg = addr_part.startswith("-")
        clean_addr = addr_part[1:] if is_neg else addr_part

        if clean_addr.lower().startswith("0x"):
            try:
                val = int(clean_addr, 16)
                return str(-val if is_neg else val)
            except ValueError:
                return addr_part
        elif clean_addr.lower().endswith("h"):
            try:
                val = int(clean_addr[:-1], 16)
                return str(-val if is_neg else val)
            except ValueError:
                return addr_part
        try:
            return str(int(addr_part, 0))
        except ValueError:
            pass
        if re.match(r"^-?[0-9A-Fa-f]+$", addr_part):
            try:
                return str(int(addr_part, 16))
            except ValueError:
                return addr_part
        return addr_part

    @staticmethod
    def validate_address(address: str, dtype: str, strict: bool = True) -> bool:
        """Validates the address format based on type and Modbus range (0-65535)."""
        dtype_upper = dtype.upper()
        if RE_TYPE_STR_CONV.match(dtype_upper):
            dtype_upper = "STRING"

        is_valid_format = False
        if dtype_upper == "STRING":
            is_valid_format = RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == "BITS":
            is_valid_format = RE_ADDR_BITS.match(address) is not None
        else:
            is_valid_format = RE_ADDR_INT.match(address) is not None

        if not is_valid_format:
            return False

        try:
            parts = address.split("_")
            base_addr_str = Generator.normalize_address_val(parts[0])
            base_addr = int(base_addr_str)
            if not (0 <= base_addr <= 65535):
                logging.warning(f"Address {base_addr} is out of Modbus range (0-65535)")
                if strict:
                    return False

            if dtype_upper == "BITS":
                start_bit = int(parts[1])
                bit_length = int(parts[2])
                if bit_length <= 0 or start_bit < 0 or start_bit + bit_length > 16:
                    logging.warning(
                        f"BITS address '{address}' must describe a non-empty slice inside bits 0-15"
                    )
                    return False
        except (ValueError, IndexError):
            return False
        return True

    @staticmethod
    def get_register_count(dtype: str, address: str) -> int:
        dtype_upper = dtype.upper()
        if RE_COUNT_16_8.match(dtype_upper):
            return 1
        elif RE_COUNT_32.match(dtype_upper):
            return 2
        elif RE_COUNT_64.match(dtype_upper):
            return 4
        elif dtype_upper == "MAC":
            return 3
        elif dtype_upper == "IPV6":
            return 8
        elif dtype_upper == "STRING":
            try:
                return math.ceil(int(address.split("_")[1]) / 2)
            except (IndexError, ValueError):
                return 0
        return 1

    @staticmethod
    def _parse_numeric(val: Any, default: float = 0.0) -> float:
        if val is None or str(val).strip() == "":
            return default
        s = str(val).strip()
        if "/" in s:
            try:
                parts = s.split("/")
                res = float(parts[0]) / float(parts[1])
                return res if math.isfinite(res) else default
            except (ValueError, ZeroDivisionError, IndexError):
                return default
        if "," in s and "." in s:
            if s.find(",") < s.find("."):
                s = s.replace(",", "")
            else:
                s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            if re.match(r"^-?\d{1,3}(,\d{3})+$", s):
                s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        try:
            res = float(s)
            return res if math.isfinite(res) else default
        except ValueError:
            return default

    @staticmethod
    def apply_address_offset(
        address: Any, offset: int, line_num: Optional[int] = None, name: Optional[str] = None
    ) -> str:
        if not address:
            return ""
        if offset == 0 and isinstance(address, (int, str)):
            s_addr = str(address).strip()
            if "_" not in s_addr:
                norm = Generator.normalize_address_val(s_addr)
                if norm.isdigit():
                    return norm
        parts = str(address).split("_")
        norm_parts = [Generator.normalize_address_val(p) for p in parts]
        try:
            base_addr = int(norm_parts[0]) + offset
            if base_addr < 0:
                msg = f"Address offset {offset} results in negative address {base_addr}"
                if name:
                    msg += f" for '{name}'"
                if line_num:
                    logging.warning(f"Line {line_num}: {msg}")
                else:
                    logging.warning(msg)
            norm_parts[0] = str(base_addr)
        except (ValueError, IndexError):
            pass
        return "_".join(norm_parts)

    def _process_name_and_tag(
        self,
        name: str,
        tag: str,
        line_num: int,
        seen_names: dict[str, int],
        seen_tags: dict[str, int],
    ) -> str:
        if name:
            if name in seen_names:
                logging.warning(
                    f"Line {line_num}: Duplicate Name '{name}' detected. Previous at line {seen_names[name]}."
                )
            else:
                seen_names[name] = line_num
        if not tag and name:
            base_tag = re.sub(r"[^a-z0-9_]", "", name.lower().replace(" ", "_"))
            base_tag = re.sub(r"_+", "_", base_tag).strip("_")
            if not base_tag or not base_tag[0].isalpha():
                base_tag = f"v_{base_tag}" if base_tag else "var"
            tag = base_tag
            counter = 1
            while tag in seen_tags:
                tag = f"{base_tag}_{counter}"
                counter += 1
        if tag:
            if tag in seen_tags:
                logging.warning(
                    f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous at line {seen_tags[tag]}."
                )
            else:
                seen_tags[tag] = line_num
        return tag

    def _determine_info1(self, reg_type_str: str, line_num: int = 0) -> str:
        """Maps RegisterType string to Webdyn Info1 code."""
        if reg_type_str is None:
            return "3"
        lt = str(reg_type_str).lower().strip()
        if not lt:
            return "3"
        if lt in self.register_type_map:
            return self.register_type_map[lt]
        elif lt in ["1", "2", "3", "4"]:
            return lt
        if line_num:
            logging.warning(
                f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to 3."
            )
        return "3"

    @staticmethod
    def _bit_slice(address: str, dtype: str) -> tuple[int, int] | None:
        if dtype.upper() != "BITS":
            return None
        try:
            _, start, length = address.split("_")
            bit_start = int(start)
            return bit_start, bit_start + int(length) - 1
        except (ValueError, IndexError):
            return None

    def _check_address_overlap(
        self,
        info1: str,
        address: str,
        dtype: str,
        name: str,
        line_num: int,
        address_usage: dict[str, dict[str, Any]],
        warned_lines: set[tuple[int, int]],
    ) -> bool:
        """Checks for address overlaps using O(log N) binary search on intervals. Returns True if overlap detected."""
        try:
            addr_part = address.split("_")[0]
            start_addr = int(self.normalize_address_val(addr_part))
            reg_count = self.get_register_count(dtype, address)
            end_addr = start_addr + reg_count - 1

            if info1 not in address_usage:
                address_usage[info1] = {"intervals": [], "max_len": 0}

            usage = address_usage[info1]
            intervals = usage["intervals"]
            max_len = usage["max_len"]

            is_bits = dtype.upper() == "BITS"
            current_bits = self._bit_slice(address, dtype)
            overlap_detected = False

            import bisect

            idx = bisect.bisect_left(intervals, (start_addr, -1, -1, "", "", -1, -1))

            def slices_overlap(u_type: str, u_start: int, u_bit_start: int, u_bit_end: int) -> bool:
                if not (is_bits and u_type == "BITS" and start_addr == u_start):
                    return True
                if current_bits is None or u_bit_start < 0 or u_bit_end < 0:
                    return True
                bit_start, bit_end = current_bits
                return max(bit_start, u_bit_start) <= min(bit_end, u_bit_end)

            for j in range(idx, len(intervals)):
                u_start, u_end, u_line, u_name, u_type, u_bit_start, u_bit_end = intervals[j]
                if u_start > end_addr:
                    break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if not slices_overlap(u_type, u_start, u_bit_start, u_bit_end):
                        continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        logging.warning(
                            f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line})."
                        )
                        warned_lines.add(warn_key)

            for j in range(idx - 1, -1, -1):
                u_start, u_end, u_line, u_name, u_type, u_bit_start, u_bit_end = intervals[j]
                if start_addr - u_start > max_len:
                    break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if not slices_overlap(u_type, u_start, u_bit_start, u_bit_end):
                        continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        logging.warning(
                            f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line})."
                        )
                        warned_lines.add(warn_key)

            bit_start, bit_end = current_bits if current_bits is not None else (-1, -1)
            bisect.insort(
                intervals,
                (start_addr, end_addr, line_num, name, dtype.upper(), bit_start, bit_end),
            )
            if reg_count > max_len:
                usage["max_len"] = reg_count
            return overlap_detected
        except (ValueError, IndexError):
            return False

    @staticmethod
    def _calculate_coefficients(
        factor_str: Any, offset_str: Any, scale_factor_str: Any
    ) -> tuple[str, str]:
        factor = Generator._parse_numeric(factor_str, default=1.0)
        offset = Generator._parse_numeric(offset_str, default=0.0)
        try:
            scale_val = int(float(scale_factor_str)) if scale_factor_str else 0
            if abs(scale_val) > 100:
                scale_val = 0
        except (ValueError, OverflowError):
            scale_val = 0

        try:
            val_a = factor * (10**scale_val)
            if not math.isfinite(val_a):
                coef_a = "1.000000"
            else:
                coef_a = f"{val_a:.6f}"
        except (OverflowError, ValueError):
            coef_a = "1.000000"

        try:
            if not math.isfinite(offset):
                coef_b = "0.000000"
            else:
                coef_b = f"{offset:.6f}"
        except (OverflowError, ValueError):
            coef_b = "0.000000"

        return coef_a, coef_b

    def process_rows(
        self, rows: Iterable[dict[str, Any]], address_offset: int = 0
    ) -> Iterator[dict[str, Any]]:
        seen_names: dict[str, int] = {}
        seen_tags: dict[str, int] = {}
        address_usage: dict[str, dict[str, Any]] = {}
        warned_lines: set[tuple[int, int]] = set()
        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v):
                continue
            norm_row = {
                k.lower().strip(): (str(v).strip() if v is not None else "") for k, v in row.items()
            }
            name, tag, reg_type_str, address = (
                norm_row.get("name", ""),
                norm_row.get("tag", ""),
                norm_row.get("registertype", ""),
                norm_row.get("address", ""),
            )
            dtype_raw, factor, offset, unit = (
                norm_row.get("type", ""),
                norm_row.get("factor", ""),
                norm_row.get("offset", ""),
                norm_row.get("unit", ""),
            )
            action, scale_factor_str = norm_row.get("action", ""), norm_row.get("scalefactor", "")

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue
            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logging.warning(
                    f"Line {line_num}: Invalid Type '{dtype_raw}' (normalized to '{dtype}'). Skipping."
                )
                continue
            match_str = RE_TYPE_STR_CONV.match(dtype)
            if match_str:
                dtype = "STRING"
                if "_" not in address:
                    address = f"{address}_{match_str.group(1)}"
            elif dtype == "BITS" and "_" not in address:
                address = f"{address}_0_16"
            address = Generator.apply_address_offset(address, address_offset, line_num, name)
            if dtype == "STRING" and "_" in address:
                address = re.sub(r"[^\d_]+$", "", address)
            if not self.validate_address(address, dtype):
                logging.warning(
                    f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping."
                )
                continue
            tag = self._process_name_and_tag(name, tag, line_num, seen_names, seen_tags)
            info1 = self._determine_info1(reg_type_str, line_num)
            self._check_address_overlap(
                info1, address, dtype, name, line_num, address_usage, warned_lines
            )
            coef_a, coef_b = self._calculate_coefficients(factor, offset, scale_factor_str)

            # Action normalization with intelligent defaulting
            act_str = str(action).strip().upper()
            if not act_str:
                norm_action = "4" if info1 in ["2", "4"] else "1"
            elif act_str in ["R", "READ", "RO", "READ-ONLY", "READ ONLY", "4"]:
                norm_action = "4"
            elif act_str in [
                "RW",
                "W",
                "WRITE",
                "READ/WRITE",
                "READ-WRITE",
                "R/W",
                "WO",
                "WRITE-ONLY",
                "WRITE ONLY",
                "1",
            ]:
                norm_action = "1"
            elif act_str in self.allowed_actions:
                norm_action = act_str
            else:
                norm_action = "4" if info1 in ["2", "4"] else "1"

            yield {
                "Info1": info1,
                "Info2": address,
                "Info3": dtype.upper(),
                "Info4": "",
                "Name": name,
                "Tag": tag,
                "CoefA": coef_a,
                "CoefB": coef_b,
                "Unit": unit,
                "Action": norm_action,
            }

    def validate_csv_detailed(
        self, filepath: str, strict: bool = False, strict_overlap: Optional[bool] = None
    ) -> ValidationReport:
        """Validates an existing WebdynSunPM definition file and returns a structured report."""
        if not os.path.exists(filepath):
            msg = f"File not found: {filepath}"
            logging.error(msg)
            return ValidationReport(
                is_valid=False,
                register_count=0,
                issues=[
                    ValidationIssue(
                        line=0,
                        severity="ERROR",
                        code="FILE_NOT_FOUND",
                        field="filepath",
                        message=msg,
                    )
                ],
                stats={"errors": 1, "warnings": 0, "types": {}, "registers": 0},
            )

        if strict_overlap is None:
            strict_overlap = strict

        valid = True
        issues: list[ValidationIssue] = []
        seen_tags: dict[str, int] = {}
        address_usage: dict[str, dict[str, Any]] = {}
        warned_lines: set[tuple[int, int]] = set()
        type_counts: dict[str, int] = {}
        total_registers = 0

        try:
            with open(filepath, "rb") as f:
                header_bytes = f.read(4)
                encoding = (
                    "utf-16" if header_bytes.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
                )

            with open(filepath, encoding=encoding) as f:
                reader = csv.reader(f, delimiter=";")
                header = next(reader, None)
                if not header or len(header) < 2 or not any(header):
                    msg = "Invalid WebdynSunPM header."
                    logging.error(msg)
                    return ValidationReport(
                        is_valid=False,
                        register_count=0,
                        issues=[
                            ValidationIssue(
                                line=1,
                                severity="ERROR",
                                code="INVALID_HEADER",
                                field="header",
                                message=msg,
                            )
                        ],
                        stats={"errors": 1, "warnings": 0, "types": {}, "registers": 0},
                    )

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row):
                        continue
                    if len(row) < 11:
                        msg = f"Line {line_num}: Row has insufficient columns ({len(row)}/11)."
                        logging.warning(msg)
                        issues.append(
                            ValidationIssue(
                                line=line_num,
                                severity="ERROR" if strict else "WARNING",
                                code="INSUFFICIENT_COLUMNS",
                                field="row",
                                message=msg,
                            )
                        )
                        if strict:
                            valid = False
                        continue

                    total_registers += 1
                    # row format: Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                    info1, info2, info3, name, tag = row[1], row[2], row[3], row[5], row[6]

                    # Validate Info1
                    info1_str = str(info1).strip()
                    if info1_str not in ("1", "2", "3", "4"):
                        msg = f"Line {line_num}: Invalid Info1 '{info1}' (expected 1, 2, 3, or 4)."
                        logging.warning(msg)
                        issues.append(
                            ValidationIssue(
                                line=line_num,
                                severity="ERROR" if strict else "WARNING",
                                code="INVALID_INFO1",
                                field="Info1",
                                message=msg,
                            )
                        )
                        if strict:
                            valid = False
                    else:
                        type_counts[info1_str] = type_counts.get(info1_str, 0) + 1

                    # Validate Tag uniqueness (fatal)
                    if tag:
                        if tag in seen_tags:
                            msg = f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (previously at line {seen_tags[tag]})."
                            logging.error(msg)
                            issues.append(
                                ValidationIssue(
                                    line=line_num,
                                    severity="ERROR",
                                    code="DUPLICATE_TAG",
                                    field="Tag",
                                    message=msg,
                                )
                            )
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Validate Type
                    if not self.validate_type(info3):
                        msg = f"Line {line_num}: Invalid Type '{info3}'."
                        logging.warning(msg)
                        issues.append(
                            ValidationIssue(
                                line=line_num,
                                severity="ERROR",
                                code="INVALID_TYPE",
                                field="Info3",
                                message=msg,
                            )
                        )
                        valid = False

                    # Validate Address format and range (fatal)
                    if not self.validate_address(info2, info3, strict=strict):
                        msg = f"Line {line_num}: Invalid address or range '{info2}' for Type '{info3}'."
                        logging.error(msg)
                        issues.append(
                            ValidationIssue(
                                line=line_num,
                                severity="ERROR",
                                code="INVALID_ADDRESS",
                                field="Info2",
                                message=msg,
                            )
                        )
                        valid = False

                    # Validate Action
                    action_val = str(row[10]).strip()
                    if action_val and action_val not in self.allowed_actions:
                        msg = f"Line {line_num}: Invalid Action '{action_val}'."
                        logging.warning(msg)
                        issues.append(
                            ValidationIssue(
                                line=line_num,
                                severity="ERROR" if strict else "WARNING",
                                code="INVALID_ACTION",
                                field="Action",
                                message=msg,
                            )
                        )
                        if strict:
                            valid = False

                    # Validate CoefA and CoefB
                    for coef_idx, coef_name in ((7, "CoefA"), (8, "CoefB")):
                        c_val = str(row[coef_idx]).strip()
                        if c_val:
                            try:
                                f_val = float(c_val)
                                if not math.isfinite(f_val):
                                    raise ValueError()
                            except ValueError:
                                msg = f"Line {line_num}: Non-numeric {coef_name} '{c_val}'."
                                logging.warning(msg)
                                issues.append(
                                    ValidationIssue(
                                        line=line_num,
                                        severity="ERROR" if strict else "WARNING",
                                        code="INVALID_COEF",
                                        field=coef_name,
                                        message=msg,
                                    )
                                )
                                if strict:
                                    valid = False

                    prev_warned = len(warned_lines)
                    self._check_address_overlap(
                        info1, info2, info3, name, line_num, address_usage, warned_lines
                    )
                    if len(warned_lines) > prev_warned:
                        issues.append(
                            ValidationIssue(
                                line=line_num,
                                severity="ERROR" if strict_overlap else "WARNING",
                                code="ADDRESS_OVERLAP",
                                field="Info2",
                                message=f"Line {line_num}: Address overlap detected for '{name}' at {info2}.",
                            )
                        )

            if strict_overlap and warned_lines:
                msg = "Address overlaps detected. Validation failed."
                logging.error(msg)
                valid = False

            err_count = sum(1 for i in issues if i.severity == "ERROR")
            warn_count = sum(1 for i in issues if i.severity == "WARNING")
            stats = {
                "errors": err_count,
                "warnings": warn_count,
                "types": type_counts,
                "registers": total_registers,
            }
            return ValidationReport(
                is_valid=valid,
                register_count=total_registers,
                issues=issues,
                stats=stats,
            )

        except (OSError, csv.Error) as e:
            msg = f"Error reading definition file: {e}"
            logging.error(msg)
            return ValidationReport(
                is_valid=False,
                register_count=0,
                issues=[
                    ValidationIssue(
                        line=0, severity="ERROR", code="IO_ERROR", field="file", message=msg
                    )
                ],
                stats={"errors": 1, "warnings": 0, "types": {}, "registers": 0},
            )

    def validate_csv(
        self, filepath: str, strict: bool = False, strict_overlap: Optional[bool] = None
    ) -> bool:
        """Validates an existing WebdynSunPM definition file."""
        report = self.validate_csv_detailed(filepath, strict=strict, strict_overlap=strict_overlap)
        return report.is_valid

    @staticmethod
    def write_output_csv(
        output: Union[str, Any, None],
        processed_rows: Iterable[dict[str, Any]],
        config: Union[CSVHeaderConfig, GeneratorConfig, str, None] = None,
        *args: Any,
        **kwargs: Any,
    ) -> None:
        """Centralized method to write the WebdynSunPM CSV format atomically for paths."""
        if isinstance(config, (CSVHeaderConfig, GeneratorConfig)):
            mfg_val = getattr(config, "manufacturer", "") or ""
            model_val = getattr(config, "model", "") or ""
            protocol_val = getattr(config, "protocol", "modbusRTU") or "modbusRTU"
            category_val = getattr(config, "category", "Inverter") or "Inverter"
            forced_write_val = getattr(config, "forced_write", "") or ""
        elif isinstance(config, str):
            mfg_val = config
            model_val = str(args[0]) if args else str(kwargs.get("model", ""))
            protocol_val = (
                str(args[1]) if len(args) > 1 else str(kwargs.get("protocol", "modbusRTU"))
            ) or "modbusRTU"
            category_val = (
                str(args[2]) if len(args) > 2 else str(kwargs.get("category", "Inverter"))
            ) or "Inverter"
            forced_write_val = (
                str(args[3]) if len(args) > 3 else str(kwargs.get("forced_write", ""))
            ) or ""
        else:
            mfg_val = str(kwargs.get("manufacturer", "")) or ""
            model_val = str(kwargs.get("model", "")) or ""
            protocol_val = str(kwargs.get("protocol", "modbusRTU")) or "modbusRTU"
            category_val = str(kwargs.get("category", "Inverter")) or "Inverter"
            forced_write_val = str(kwargs.get("forced_write", "")) or ""
        type_counts = {"1": 0, "2": 0, "3": 0, "4": 0}
        type_labels = {"1": "Coils", "2": "Discrete", "3": "Holding", "4": "Input"}
        outfile: Any = None
        temp_path: Optional[str] = None
        try:
            if isinstance(output, str):
                target_path = os.path.abspath(output)
                target_dir = os.path.dirname(target_path)
                fd, temp_path = tempfile.mkstemp(
                    prefix=f".{os.path.basename(target_path)}.",
                    suffix=".tmp",
                    dir=target_dir,
                )
                outfile = os.fdopen(fd, "w", newline="", encoding="utf-8-sig")
            elif output is None:
                outfile = sys.stdout
            else:
                outfile = output

            writer = csv.writer(outfile, delimiter=";", lineterminator="\n")

            header_row = [
                Generator.sanitize_csv_field(protocol_val),
                Generator.sanitize_csv_field(category_val),
                Generator.sanitize_csv_field(mfg_val),
                Generator.sanitize_csv_field(model_val),
                Generator.sanitize_csv_field(forced_write_val),
                "",
                "",
                "",
                "",
                "",
                "",
            ]
            writer.writerow(header_row)

            total = 0
            for index, row in enumerate(processed_rows, start=1):
                writer.writerow(
                    [
                        str(index),
                        Generator.sanitize_csv_field(row["Info1"]),
                        Generator.sanitize_csv_field(row["Info2"]),
                        Generator.sanitize_csv_field(row["Info3"]),
                        Generator.sanitize_csv_field(row["Info4"]),
                        Generator.sanitize_csv_field(row["Name"]),
                        Generator.sanitize_csv_field(row["Tag"]),
                        Generator.sanitize_csv_field(row["CoefA"]),
                        Generator.sanitize_csv_field(row["CoefB"]),
                        Generator.sanitize_csv_field(row["Unit"]),
                        Generator.sanitize_csv_field(row["Action"]),
                    ]
                )
                type_counts[row["Info1"]] = type_counts.get(row["Info1"], 0) + 1
                total += 1

            if isinstance(output, str):
                outfile.flush()
                os.fsync(outfile.fileno())
                outfile.close()
                outfile = None
                assert temp_path is not None
                os.replace(temp_path, os.path.abspath(output))
                temp_path = None

            summary = ", ".join([f"{type_labels[k]}: {v}" for k, v in type_counts.items() if v > 0])
            if summary:
                logging.info(f"Generated {total} registers ({summary})")
        except (OSError, csv.Error) as e:
            logging.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and outfile is not None:
                outfile.close()
            if temp_path is not None:
                try:
                    os.unlink(temp_path)
                except FileNotFoundError:
                    pass


def generate_template(output_file: Optional[str], mode: str = "input") -> None:
    """Generates a sample template CSV file."""
    if mode == "definition":
        headers = [
            "#Index",
            "Info1",
            "Info2",
            "Info3",
            "Info4",
            "Name",
            "Tag",
            "CoefA",
            "CoefB",
            "Unit",
            "Action",
        ]
        rows = [
            [
                "modbusRTU",
                "Inverter",
                "SampleManufacturer",
                "SampleModel",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ],
            [
                "1",
                "3",
                "40001",
                "U16",
                "",
                "Active Power",
                "active_power",
                "1.000000",
                "0.000000",
                "W",
                "4",
            ],
            ["2", "3", "40002", "U16", "", "Voltage", "voltage", "0.100000", "0.000000", "V", "4"],
        ]
        delimiter = ";"
    else:
        headers = [
            "Name",
            "Tag",
            "RegisterType",
            "Address",
            "Type",
            "Factor",
            "Offset",
            "Unit",
            "Action",
            "ScaleFactor",
        ]
        rows = [
            [
                "Example Variable",
                "example_tag",
                "Holding Register",
                "30001",
                "U16",
                "1",
                "0",
                "V",
                "4",
                "0",
            ],
            [
                "Convenience String",
                "str_tag",
                "Holding Register",
                "30030",
                "STR20",
                "",
                "",
                "",
                "4",
                "",
            ],
        ]
        delimiter = ","

    outfile = None
    try:
        if output_file:
            outfile = open(output_file, "w", newline="", encoding="utf-8")
            writer = csv.writer(outfile, delimiter=delimiter)
        else:
            writer = csv.writer(sys.stdout, delimiter=delimiter)

        writer.writerow(headers)
        writer.writerows(rows)
    except OSError as e:
        logging.error(f"Error generating template: {e}")
    finally:
        if output_file and outfile:
            outfile.close()


def run_generator(
    config: GeneratorConfig, input_data: Optional[Iterable[dict[str, Any]]] = None
) -> None:
    generator = Generator()
    if config.template:
        mode = config.template_mode
        if input_data is not None:
            mode = "definition"
        generate_template(config.output, mode=mode)
        return

    manufacturer = config.manufacturer or "Manufacturer"
    model = config.model or "Model"

    if input_data is None:
        if not config.input_file:
            logging.error("input_file or input_data is required.")
            return
        if not os.path.exists(config.input_file):
            logging.error(f"Input file not found: {config.input_file}")
            return

    try:
        hdr_cfg = CSVHeaderConfig(
            manufacturer=manufacturer,
            model=model,
            protocol=config.protocol,
            category=config.category,
            forced_write=config.forced_write,
        )
        if input_data is not None:
            processed_rows = generator.process_rows(input_data, config.address_offset)
            generator.write_output_csv(
                config.output,
                processed_rows,
                hdr_cfg,
            )
        else:
            if not config.input_file:
                logging.error("input_file or input_data is required.")
                return
            if not os.path.exists(config.input_file):
                logging.error(f"Input file not found: {config.input_file}")
                return
            with open(config.input_file, mode="rb") as f:
                header_bytes = f.read(4)
                encoding = (
                    "utf-16" if header_bytes.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
                )

            with open(config.input_file, encoding=encoding) as csvfile:
                snippet = csvfile.read(2048)
                csvfile.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                except csv.Error:
                    dialect = csv.excel
                reader = csv.DictReader(csvfile, dialect=dialect)
                processed_rows = generator.process_rows(reader, config.address_offset)

                generator.write_output_csv(
                    config.output,
                    processed_rows,
                    hdr_cfg,
                )
    except (OSError, csv.Error, ValueError, TypeError, KeyError) as e:
        logging.error(f"An error occurred during generation: {e}")


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s", force=True)
    parser = argparse.ArgumentParser(description="Generate WebdynSunPM Modbus definition file.")
    parser.add_argument("input_file", nargs="?", help="Input simplified CSV.")
    parser.add_argument("-o", "--output", help="Output CSV.")
    parser.add_argument("--manufacturer", help="Manufacturer name.")
    parser.add_argument("--model", help="Model name.")
    parser.add_argument("--protocol", default="modbusRTU")
    parser.add_argument("--category", default="Inverter")
    parser.add_argument("--forced-write", default="")
    parser.add_argument("--template", action="store_true")
    parser.add_argument("--template-mode", choices=["input", "definition"], default="input")
    parser.add_argument("--address-offset", type=int, default=0)
    args = parser.parse_args()
    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        template=args.template,
        template_mode=args.template_mode,
        address_offset=args.address_offset,
    )
    run_generator(config)


if __name__ == "__main__":
    main()
