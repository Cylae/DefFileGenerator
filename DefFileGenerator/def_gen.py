#!/usr/bin/env python3
"""
WebdynSunPM Definition File Generator and Validator.

This module provides the core domain logic for generating, processing, and
validating WebdynSunPM definition CSV files from extracted Modbus registers.

Key Capabilities:
    - Normalizes non-standard vendor data types into valid WebdynSunPM type codes.
    - Resolves addresses, address offsets, and bit-level slicing.
    - Detects address collisions and overlaps using an O(log N) bisect interval search.
    - Sanitizes all CSV text fields against formula injection and multiline splitting.
    - Writes valid definition files atomically using temporary files and fsync.
    - Validates generated and existing definition CSV files against strict Modbus rules.
"""

from __future__ import annotations

import argparse
import bisect
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
from dataclasses import dataclass, field
from typing import Any

# -----------------------------------------------------------------------------
# Modbus & WebdynSunPM Domain Constants
# -----------------------------------------------------------------------------

# WebdynSunPM Modbus Info1 Function Codes
MODBUS_COIL = "1"  # Read/Write 1-bit Coil (Modbus FC 01, 05, 15)
MODBUS_DISCRETE = "2"  # Read-Only 1-bit Discrete Input (Modbus FC 02)
MODBUS_HOLDING = "3"  # Read/Write 16-bit Holding Register (Modbus FC 03, 06, 16)
MODBUS_INPUT = "4"  # Read-Only 16-bit Input Register (Modbus FC 04)

# Modbus Standard Address Boundaries
MIN_MODBUS_ADDRESS = 0
MAX_MODBUS_ADDRESS = 65535
DEFAULT_BITS_REGISTER_WIDTH = 16  # Modbus registers are 16-bit words

# WebdynSunPM Action Permissions (Field 11)
ACTION_WRITE_ONLY = "1"
ACTION_READ_ONLY = "4"

# Allowed Action Codes accepted by WebdynSunPM firmware
ALLOWED_ACTIONS = frozenset({"0", "1", "2", "4", "6", "7", "8", "9"})

# -----------------------------------------------------------------------------
# Pre-compiled Regular Expressions (High-Throughput Parsing)
# -----------------------------------------------------------------------------

RE_TYPE_NUMERIC = re.compile(r"^([UI](8|16|32|64)|F(32|64))(_(W|B|WB))?$", re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r"^STR(\d+)$", re.IGNORECASE)
RE_ADDR_STRING = re.compile(r"^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)$", re.IGNORECASE)
RE_ADDR_BITS = re.compile(r"^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)_(\d+)$", re.IGNORECASE)
RE_ADDR_INT = re.compile(r"^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)$", re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r"^([UI](16|8)(_(W|B|WB))?|BITS)$", re.IGNORECASE)
RE_COUNT_32 = re.compile(r"^([UI]32(_(W|B|WB))?|F32(_(W|B|WB))?|IP)$", re.IGNORECASE)
RE_COUNT_64 = re.compile(r"^([UI]64(_(W|B|WB))?|F64(_(W|B|WB))?)$", re.IGNORECASE)

_CLEAN_TYPE_RE = re.compile(r"[^a-z0-9_]+")
_ADDR_RANGE_RE = re.compile(r"^([^\s~.]+?)\s*(?:~|\.\.|\s+-\s+|-(?=[0-9a-fA-FxX]))")
_ADDR_GROUPING_COMMA_RE = re.compile(r"(?<=\d),(?=\d{3}(?!\d))")
_ADDR_HEX_MATCH_RE = re.compile(r"^-?[0-9A-Fa-f]+$")
_ADDR_CLEAN_BITFIELD_RE = re.compile(r"_(wb|b|w)$|\b(swap|big endian|big|word)\b")
_ADDR_BITFIELD_MATCH_RE = re.compile(r"^bitfield(8|16|32|64)$")
_ADDR_STRING_MATCH_RE = re.compile(r"^(?:string|str)[\*x_]?\s*(\d+)$")
_COLLAPSE_WHITESPACE_RE = re.compile(r"\s+")
_DIGIT_SPLIT_RE = re.compile(r"([a-zA-Z0-9])\s+(\d+)\b")
_WORD_SWAP_RE = re.compile(
    r"(_w\b|_w$|\bword\s*swap\b|\bword-swap\b|(?:float\d*|f\d+|u\d+|i\d+|int\d*|uint\d*)\s+word\b)"
)
_TAG_BASE_RE = re.compile(r"[^a-z0-9_]")
_MULTI_UNDERSCORE_RE = re.compile(r"_+")
_THOUSANDS_COMMA_RE = re.compile(r"^-?\d{1,3}(,\d{3})+$")
_TRAILING_NON_DIGIT_RE = re.compile(r"[^\d_]+$")
_STR_MATCH_RE = re.compile(r"string\s*(\d+)")

# Pre-compiled Type Synonyms Table (ordered by specificity)
TYPE_SYNONYMS_COMPILED: tuple[tuple[re.Pattern[str], str], ...] = (
    (
        re.compile(r"unsigned\s*(?:int(?:eger)?)?\s*64|uint64|\bu64\b|\bint64u\b"),
        "U64",
    ),
    (
        re.compile(r"signed\s*(?:int(?:eger)?)?\s*64|sint64|\bint64\b|\bi64\b|\bs64\b|\bint64s\b"),
        "I64",
    ),
    (re.compile(r"\bunsigned\s*long(?:\s*int)?\b|\bulong\b|\buint32_t\b"), "U32"),
    (re.compile(r"\bsigned\s*long(?:\s*int)?\b|\bslong\b|\blong(?:\s*int)?\b|\bint32_t\b"), "I32"),
    (re.compile(r"unsigned\s*(?:int(?:eger)?)?\s*32|uint32|\bu32\b|\bint32u\b"), "U32"),
    (
        re.compile(r"signed\s*(?:int(?:eger)?)?\s*32|sint32|\bint32\b|\bi32\b|\bs32\b|\bint32s\b"),
        "I32",
    ),
    (re.compile(r"\bunsigned\s*short(?:\s*int)?\b|\bushort\b|\buint16_t\b"), "U16"),
    (re.compile(r"\bsigned\s*short(?:\s*int)?\b|\bsshort\b|\bshort(?:\s*int)?\b"), "I16"),
    (re.compile(r"unsigned\s*(?:int(?:eger)?)?\s*16|uint16|\bu16\b|\bint16u\b"), "U16"),
    (
        re.compile(r"signed\s*(?:int(?:eger)?)?\s*16|sint16|\bint16\b|\bi16\b|\bs16\b|\bint16s\b"),
        "I16",
    ),
    (re.compile(r"\bunsigned\s*char\b|\buchar\b|\buint8_t\b"), "U8"),
    (re.compile(r"\bsigned\s*char\b|\bschar\b"), "I8"),
    (re.compile(r"unsigned\s*(?:int(?:eger)?)?\s*8|uint8|\bu8\b|\bint8u\b"), "U8"),
    (re.compile(r"signed\s*(?:int(?:eger)?)?\s*8|sint8|\bint8\b|\bi8\b|\bs8\b|\bint8s\b"), "I8"),
    (re.compile(r"float64|double|\bf64\b|64\s*[-_]?\s*bit\s*ieee\s*[-_]?\s*754"), "F64"),
    (re.compile(r"float32|float|\bf32\b|(?:32\s*[-_]?\s*bit\s*)?ieee\s*[-_]?\s*754"), "F32"),
    (re.compile(r"\b(bit32|bitmap32|bits32|bit16|bitmap16|bits16)\b"), "BITS"),
    (re.compile(r"\bqword\b"), "U64"),
    (re.compile(r"\bdword\b"), "U32"),
    (re.compile(r"\bword\b|\buword\b"), "U16"),
    (re.compile(r"\bbyte\b|\bubyte\b"), "U8"),
    (re.compile(r"^(?:32\s*[-_]?\s*bit\s*)hex$|^hex32$"), "U32"),
    (re.compile(r"^(?:16\s*[-_]?\s*bit\s*)?hex(?:16)?$"), "U16"),
    (re.compile(r"^unsigned\s*(?:int(?:eger)?)?$|^uint$|^unsigned$"), "U16"),
    (re.compile(r"^signed\s*(?:int(?:eger)?)?$|^sint$|^int$"), "I16"),
    (re.compile(r"\bdate\s*time\b|\bdatetime\b"), "U32"),
    (re.compile(r"\b(?:ip4|ipv4)\b"), "IP"),
    (re.compile(r"\b(?:ip6|ipv6)\b"), "IPV6"),
)

# Characters to strip when inspecting text for CSV formula injection
_CSV_STRIP_CHARS = (
    " \t\n\r\v\f\u00a0\u0085\u00ad\u1680\u180e\ufeff\u3000"
    + "".join(chr(c) for c in range(0x2000, 0x200F + 1))
    + "".join(chr(c) for c in range(0x2028, 0x202F + 1))
    + "".join(chr(c) for c in range(0x205F, 0x206F + 1))
)

# Leading whitespace/zero-width characters that spreadsheet apps ignore before executing formulas
_CSV_PREFIX_TRIGGERS: frozenset[str] = frozenset(
    "\t\r\n\u00a0\u0085\u00ad\u1680\u180e\ufeff\u3000"
    + "".join(chr(c) for c in range(0x2000, 0x200F + 1))
    + "".join(chr(c) for c in range(0x2028, 0x202F + 1))
    + "".join(chr(c) for c in range(0x205F, 0x206F + 1))
)
_CSV_FORMULA_CHARS: frozenset[str] = frozenset(
    {"=", "+", "-", "@", "|", "%", "\uff1d", "\uff0b", "\uff0d", "\uff20"}
)
_CSV_FULLWIDTH_TRIGGERS: tuple[str, ...] = ("\uff1d", "\uff0b", "\uff0d", "\uff20")

MODBUS_VALID_INFO1: frozenset[str] = frozenset(
    {MODBUS_COIL, MODBUS_DISCRETE, MODBUS_HOLDING, MODBUS_INPUT}
)

_EXACT_REG_COUNTS: dict[str, int] = {
    "U16": 1,
    "I16": 1,
    "U8": 1,
    "I8": 1,
    "BITS": 1,
    "U32": 2,
    "I32": 2,
    "F32": 2,
    "IP": 2,
    "U64": 4,
    "I64": 4,
    "F64": 4,
    "MAC": 3,
    "IPV6": 8,
}

_VALID_SPECIAL_TYPES: frozenset[str] = frozenset({"STRING", "BITS", "IP", "IPV6", "MAC"})
_COMMON_VALID_NUMERIC: frozenset[str] = frozenset(
    {
        "U8",
        "U16",
        "U32",
        "U64",
        "I8",
        "I16",
        "I32",
        "I64",
        "F32",
        "F64",
    }
)

_SHORTHAND_TYPE_MAP: dict[str, str] = {
    "word": "U16",
    "uword": "U16",
    "ushort": "U16",
    "unsignedshort": "U16",
    "dword": "U32",
    "udword": "U32",
    "ulong": "U32",
    "unsignedlong": "U32",
    "qword": "U64",
    "uqword": "U64",
    "byte": "U8",
    "ubyte": "U8",
    "uchar": "U8",
    "unsignedchar": "U8",
    "short": "I16",
    "sshort": "I16",
    "signedshort": "I16",
    "long": "I32",
    "slong": "I32",
    "signedlong": "I32",
}

_CANONICAL_TYPE_SET: frozenset[str] = frozenset(
    {
        "u16",
        "u32",
        "i16",
        "i32",
        "s16",
        "s32",
        "u64",
        "i64",
        "f32",
        "f64",
    }
)


# -----------------------------------------------------------------------------
# Helper Utilities
# -----------------------------------------------------------------------------


def peek_generator(
    iterable: Iterable[Any] | None,
) -> tuple[bool, Iterator[Any]]:
    """
    Checks if an iterable contains data without fully consuming it.

    Reconstructs the original stream by prepending the peeked item via itertools.chain.

    Args:
        iterable: Any iterable stream or generator to inspect.

    Returns:
        tuple[bool, Iterator[Any]]: (has_data, reconstructed_iterator).
    """
    if iterable is None:
        return False, iter([])
    it = iter(iterable)
    try:
        first = next(it)
    except StopIteration:
        return False, iter([])
    return True, itertools.chain([first], it)


# -----------------------------------------------------------------------------
# Configuration and Data Models
# -----------------------------------------------------------------------------


@dataclass
class WebdynDefConfig:
    """Configuration options for programmatic definition generation."""

    input_file: str
    output_file: str
    manufacturer: str
    model: str
    protocol: str = "modbusRTU"
    category: str = "Inverter"
    address_offset: int = 0
    strict_validation: bool = True


@dataclass
class CSVHeaderConfig:
    """Header metadata for the first line of a WebdynSunPM definition CSV."""

    manufacturer: str = ""
    model: str = ""
    protocol: str = "modbusRTU"
    category: str = "Inverter"
    forced_write: str = ""


@dataclass
class RegisterEntry:
    """Encapsulates a single Modbus register entry for validation and output."""

    info1: str
    address: str
    dtype: str
    name: str
    line_num: int = 0


@dataclass
class GeneratorConfig:
    """Execution options for CLI and batch generator runs."""

    input_file: str | None = None
    output: str | None = None
    manufacturer: str | None = None
    model: str | None = None
    protocol: str = "modbusRTU"
    category: str = "Inverter"
    forced_write: str = ""
    template: bool = False
    template_mode: str = "input"  # 'input' or 'definition'
    address_offset: int = 0


@dataclass
class ValidationIssue:
    """Details a single error or warning encountered during definition validation."""

    line: int
    severity: str  # "ERROR" or "WARNING"
    code: str
    field: str
    message: str


@dataclass
class ValidationReport:
    """Structured report returned by definition file validation."""

    is_valid: bool
    register_count: int
    issues: list[ValidationIssue]
    stats: dict[str, Any] = field(default_factory=dict)


@functools.lru_cache(maxsize=4096)
def _normalize_address_cached(s: str) -> str:
    """Internal cached helper for resolving compound/hex/range addresses."""
    range_match = _ADDR_RANGE_RE.match(s)
    if range_match:
        s = range_match.group(1).strip()

    if s.endswith(".") and not s.endswith(".."):
        s = s.rstrip(".")

    if s.isdigit() and (not s.startswith("0") or s == "0"):
        return s

    addr_part = s
    # Remove grouping commas (e.g., "30,001" -> "30001")
    addr_part = _ADDR_GROUPING_COMMA_RE.sub("", addr_part)
    if not addr_part:
        return ""

    is_neg = addr_part.startswith("-")
    clean_addr = addr_part[1:] if is_neg else addr_part

    # Check explicit hexadecimal prefix or suffix
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

    # Try parsing plain hexadecimal string without 0x prefix
    if _ADDR_HEX_MATCH_RE.match(addr_part):
        try:
            return str(int(addr_part, 16))
        except ValueError:
            return addr_part

    return addr_part


# -----------------------------------------------------------------------------
# Generator Engine Class
# -----------------------------------------------------------------------------


class Generator:
    """
    Core engine responsible for data normalization, validation, and CSV output generation.

    Attributes:
        register_type_map: Mapping from common Modbus register descriptions to Webdyn Info1 codes.
        allowed_actions: Set of permitted WebdynSunPM register action codes.
    """

    def __init__(self, strict: bool = False) -> None:
        """
        Initializes the Generator instance.

        Args:
            strict: Whether to enforce strict mode (warnings treated as fatal errors).
        """
        self.register_type_map: dict[str, str] = {
            "coil": MODBUS_COIL,
            "coils": MODBUS_COIL,
            "discrete input": MODBUS_DISCRETE,
            "discrete register": MODBUS_DISCRETE,
            "discrete registers": MODBUS_DISCRETE,
            "discrete": MODBUS_DISCRETE,
            "holding register": MODBUS_HOLDING,
            "holding": MODBUS_HOLDING,
            "input register": MODBUS_INPUT,
            "input": MODBUS_INPUT,
        }
        self.allowed_actions: list[str] = ["0", "1", "2", "4", "6", "7", "8", "9"]
        self.strict = strict

    @staticmethod
    def sanitize_csv_field(val: Any) -> str:
        """
        Sanitizes a CSV field value to prevent CSV Formula Injection (DDE attacks)
        and multiline record desynchronization.

        Rules applied:
            1. Strips non-printable ASCII and surrogate codepoints (U+D800 - U+DFFF).
            2. Replaces internal newlines (\r, \n) with a space to avoid multiline rows.
            3. Detects formula invocation characters (=, +, -, @, |, %, and full-width equivalents).
            4. Safely ignores valid finite numeric literals.
            5. Prepends an apostrophe (') to neutralize formula execution in Excel/LibreOffice.

        Args:
            val: Raw value of any type.

        Returns:
            str: Sanitized, injection-safe string representation.
        """
        if val is None:
            return ""
        if isinstance(val, (int, float)):
            return str(val)
        s = str(val)
        if not s:
            return ""

        # Fast-path: 99% of normal industrial modbus registers are clean printable ASCII starting with an alnum
        if s.isascii() and s.isprintable() and s[0].isalnum():
            return s

        # Remove control and invalid surrogate characters
        if not s.isprintable():
            s = "".join(
                ch
                for ch in s
                if (ord(ch) in (9, 10, 13) or (ord(ch) >= 32 and ord(ch) != 127))
                and not (0xD800 <= ord(ch) <= 0xDFFF)
            )
            if not s:
                return ""

        # Check for leading whitespace or zero-width triggers used in CSV injection payloads
        if s[0] in _CSV_PREFIX_TRIGGERS or s.startswith(_CSV_FULLWIDTH_TRIGGERS):
            prefix = s[0]
            rest = s[1:].replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
            return "'" + prefix + rest

        # Replace internal newlines with space to prevent row splitting
        if "\n" in s or "\r" in s:
            s = s.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")

        stripped = s.lstrip(_CSV_STRIP_CHARS)
        if not stripped:
            return s

        # Neutralize spreadsheet formula prefix triggers
        if stripped[0] in _CSV_FORMULA_CHARS:
            if "_" in stripped or "inf" in stripped.lower() or "nan" in stripped.lower():
                return "'" + s
            try:
                f = float(stripped)
                if math.isfinite(f):
                    # Negative numbers with leading space can be misinterpreted in some spreadsheet engines
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
        """
        Normalizes arbitrary manufacturer data type descriptions into WebdynSunPM type codes.

        Handles endianness qualifiers:
            - _WB: Word-swap and Byte-swap (big endian words and bytes)
            - _B: Byte-swap only
            - _W: Word-swap only

        Examples:
            - "uint16" -> "U16"
            - "float32" -> "F32"
            - "float32 big endian" -> "F32_WB"
            - "string 20" -> "STR20"
            - "bitfield16" -> "U16"

        Args:
            dtype: Raw string or object representing the data type.

        Returns:
            str: Normalized WebdynSunPM data type code.
        """
        if not dtype:
            return "U16"
        t = str(dtype).lower().strip()

        # Collapse internal whitespace for fragmented font extractions (e.g. "u1 6" -> "u16")
        compact_t = _COLLAPSE_WHITESPACE_RE.sub("", t)
        if compact_t in ("str", "string"):
            return "STRING"
        if compact_t in _CANONICAL_TYPE_SET:
            if compact_t.startswith("s"):
                return f"I{compact_t[1:]}".upper()
            return compact_t.upper()

        # Normalize spaces before trailing digits (e.g., "uint 16" -> "uint16")
        for _ in range(2):
            t = _DIGIT_SPLIT_RE.sub(r"\1\2", t)

        # Detect endianness / byte-order suffixes
        suffix = ""
        if any(x in t for x in ["_wb", "swap", "big endian"]):
            suffix = "_WB"
        elif any(x in t for x in ["_b", "big"]):
            suffix = "_B"
        elif _WORD_SWAP_RE.search(t):
            suffix = "_W"

        # Direct canonical shorthand matches
        base_shorthand = _SHORTHAND_TYPE_MAP.get(compact_t)
        if base_shorthand:
            return f"{base_shorthand}{suffix}"

        # Handle explicit string length notation (e.g. "string 20" -> "STR20")
        str_match = _STR_MATCH_RE.search(t)
        if str_match:
            return f"STR{str_match.group(1)}{suffix}"

        # Bare "str" maps to STRING; length is typically supplied via quantity column
        if t == "str":
            return "STRING"

        # Bitfield masks (e.g. "Bitfield16") map to unsigned integers of matching width
        clean_bitfield_t = _ADDR_CLEAN_BITFIELD_RE.sub("", t).replace(" ", "")
        bitfield_match = _ADDR_BITFIELD_MATCH_RE.match(clean_bitfield_t)
        if bitfield_match:
            return f"U{bitfield_match.group(1)}{suffix}"

        # Evaluate pre-compiled synonym patterns
        for pattern, replacement in TYPE_SYNONYMS_COMPILED:
            if pattern.search(t):
                return f"{replacement}{suffix}"

        str_pattern_match = _ADDR_STRING_MATCH_RE.match(t)
        if str_pattern_match:
            return f"STR{str_pattern_match.group(1)}"

        if "string" in t:
            return f"STRING{suffix}"

        if t.startswith("str") and t[3:].isdigit():
            return t.upper()

        # Fallback: strip unexpected punctuation and return uppercase
        t = _CLEAN_TYPE_RE.sub("", t)
        return t.upper() if t else "U16"

    @staticmethod
    def validate_type(dtype: str) -> bool:
        """
        Validates if a normalized data type string is recognized by WebdynSunPM.

        Args:
            dtype: Normalized type string (e.g. "U16", "STRING", "STR20", "BITS").

        Returns:
            bool: True if the type is valid, False otherwise.
        """
        dtype_upper = str(dtype).upper()
        if dtype_upper in _COMMON_VALID_NUMERIC or dtype_upper in _VALID_SPECIAL_TYPES:
            return True
        if RE_TYPE_NUMERIC.match(dtype_upper):
            return True
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return True
        return False

    @staticmethod
    def normalize_address_val(addr_part: Any) -> str:
        """
        Converts varied address representations (hexadecimal, range notations, commas)
        into a clean decimal string.

        Supported input formats:
            - Hexadecimal: "0x8232", "8232h", "1F40"
            - Decimal: "40001", "30,001"
            - Range notation: "31657~31658", "0x8232 - 0x82FE" (extracts starting address)

        Args:
            addr_part: Address string or integer.

        Returns:
            str: Normalized base address in decimal string representation.
        """
        if isinstance(addr_part, int):
            return str(addr_part)
        try:
            s = str(addr_part).strip()
        except Exception:
            return ""
        if not s:
            return ""

        # Fast-path: positive decimal integers without range/punctuation (e.g. "40001", "0")
        if s.isdigit() and (s == "0" or not s.startswith("0")):
            return s

        return _normalize_address_cached(s)

    @staticmethod
    def validate_address(address: Any, dtype: str, strict: bool = True) -> bool:
        """
        Validates address syntax and checks Modbus address boundary (0-65535).

        Rules:
            - Non-compound types: single integer ("40001")
            - STRING: compound address with byte length ("40001_20")
            - BITS: compound address with bit offset and length ("40001_0_16")

        Args:
            address: Address string or int to validate.
            dtype: Associated normalized data type.
            strict: If True, out-of-range addresses (not in 0..65535) fail validation.

        Returns:
            bool: True if the address is valid, False otherwise.
        """
        if address is None:
            return False
        addr_str = str(address).strip()
        dtype_upper = dtype.upper()
        if RE_TYPE_STR_CONV.match(dtype_upper):
            dtype_upper = "STRING"

        if dtype_upper == "STRING":
            is_valid_format = RE_ADDR_STRING.match(addr_str) is not None
        elif dtype_upper == "BITS":
            is_valid_format = RE_ADDR_BITS.match(addr_str) is not None
        else:
            is_valid_format = RE_ADDR_INT.match(addr_str) is not None

        if not is_valid_format:
            return False

        try:
            parts = addr_str.split("_")
            base_addr_str = Generator.normalize_address_val(parts[0])
            base_addr = int(base_addr_str)
            if not (MIN_MODBUS_ADDRESS <= base_addr <= MAX_MODBUS_ADDRESS):
                logging.warning(
                    f"Address {base_addr} is out of Modbus range ({MIN_MODBUS_ADDRESS}-{MAX_MODBUS_ADDRESS})"
                )
                if strict:
                    return False

            if dtype_upper == "STRING":
                str_len = int(parts[1])
                if str_len <= 0:
                    logging.warning(f"STRING address '{address}' must have positive byte length")
                    return False

            if dtype_upper == "BITS":
                start_bit = int(parts[1])
                bit_length = int(parts[2])
                if (
                    bit_length <= 0
                    or start_bit < 0
                    or start_bit + bit_length > DEFAULT_BITS_REGISTER_WIDTH
                ):
                    logging.warning(
                        f"BITS address '{address}' must describe a non-empty slice inside bits 0-15"
                    )
                    return False
        except (ValueError, IndexError):
            return False

        return True

    @staticmethod
    def get_register_count(dtype: str, address: str) -> int:
        """
        Computes how many 16-bit Modbus registers are occupied by the given data type.

        Width Mapping:
            - U8, I8, U16, I16, BITS: 1 register (16 bits)
            - U32, I32, F32, IP: 2 registers (32 bits)
            - U64, I64, F64: 4 registers (64 bits)
            - MAC: 3 registers (48 bits)
            - IPV6: 8 registers (128 bits)
            - STRING: ceil(byte_length / 2)

        Args:
            dtype: Normalized data type.
            address: Address string (provides byte length for STRING types).

        Returns:
            int: Number of 16-bit registers occupied.
        """
        dtype_upper = dtype.upper()
        count = _EXACT_REG_COUNTS.get(dtype_upper)
        if count is not None:
            return count
        if RE_COUNT_16_8.match(dtype_upper):
            return 1
        elif RE_COUNT_32.match(dtype_upper):
            return 2
        elif RE_COUNT_64.match(dtype_upper):
            return 4
        elif dtype_upper == "STRING" or dtype_upper.startswith("STR"):
            try:
                if "_" in address:
                    return math.ceil(int(address.split("_")[1]) / 2)
            except (IndexError, ValueError):
                pass
            match_str = RE_TYPE_STR_CONV.match(dtype_upper)
            if match_str:
                try:
                    return math.ceil(int(match_str.group(1)) / 2)
                except ValueError:
                    pass
            return 0
        return 1

    @staticmethod
    def _parse_numeric(val: Any, default: float = 0.0) -> float:
        """
        Extracts a numeric float value from diverse documentation string formats.

        Handles:
            - Fractions (e.g. "1/100" -> 0.01)
            - Decimal commas (European notation: "0,1" -> 0.1)
            - Thousands separators ("1,000" -> 1000.0)

        Args:
            val: Raw value to parse.
            default: Value returned if parsing fails.

        Returns:
            float: Extracted finite floating point number.
        """
        if val is None or str(val).strip() == "":
            return default
        s = str(val).strip()

        # Parse fractions (e.g. 1/10, 1/100)
        if "/" in s:
            try:
                parts = s.split("/")
                if len(parts) != 2:
                    return default
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
            if _THOUSANDS_COMMA_RE.match(s):
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
        address: Any,
        offset: int,
        line_num: int | None = None,
        name: str | None = None,
    ) -> str:
        """
        Applies a numeric offset to the base register address while preserving subfield suffixes.

        For example:
            - apply_address_offset("40001", -1) -> "40000"
            - apply_address_offset("40001_20", -1) -> "40000_20"
            - apply_address_offset("40001_0_16", 10) -> "40011_0_16"

        Args:
            address: Raw or formatted address.
            offset: Integer offset to add to the base address.
            line_num: Optional source row number for logging warnings.
            name: Optional register name for logging warnings.

        Returns:
            str: Modified address string with offset applied.
        """
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
        """
        Generates and deduplicates the Webdyn variable Tag from the register Name.

        WebdynSunPM requires Tags to be unique, lowercase identifiers starting with a letter.

        Args:
            name: Human-readable register name.
            tag: Explicit tag if already provided.
            line_num: Current row number for reporting duplicates.
            seen_names: Dictionary tracking previous name occurrences.
            seen_tags: Dictionary tracking previous tag occurrences.

        Returns:
            str: Normalized unique tag.
        """
        if name:
            if name in seen_names:
                logging.warning(
                    f"Line {line_num}: Duplicate Name '{name}' detected. Previous at line {seen_names[name]}."
                )
            else:
                seen_names[name] = line_num

        if not tag and name:
            base_tag = _TAG_BASE_RE.sub("", name.lower().replace(" ", "_"))
            base_tag = _MULTI_UNDERSCORE_RE.sub("_", base_tag).strip("_")
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
        """
        Maps a register type string to the Webdyn Info1 Modbus function code.

        Returns:
            - "1": Coil (0x)
            - "2": Discrete Input (1x)
            - "3": Holding Register (4x) [Default]
            - "4": Input Register (3x)
        """
        if reg_type_str is None:
            return MODBUS_HOLDING
        lt = str(reg_type_str).lower().strip()
        if not lt:
            return MODBUS_HOLDING
        if lt in self.register_type_map:
            return self.register_type_map[lt]
        elif lt in MODBUS_VALID_INFO1:
            return lt
        if line_num:
            logging.warning(
                f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to {MODBUS_HOLDING}."
            )
        return MODBUS_HOLDING

    @staticmethod
    def _bit_slice(address: str, dtype: str) -> tuple[int, int] | None:
        """
        Extracts the (bit_start, bit_end) range for BITS data types within a 16-bit register.

        Returns:
            tuple[int, int] | None: Inclusive (start_bit, end_bit) slice, or None if not BITS.
        """
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
        info1: str | RegisterEntry,
        address: Any = None,
        dtype: str = "",
        name: str = "",
        line_num: int = 0,
        address_usage: dict[str, dict[str, Any]] | None = None,
        warned_lines: set[tuple[int, int]] | None = None,
    ) -> bool:
        """
        Checks for address collisions across registers using an interval binary search (O(log N)).

        Supports both modern RegisterEntry dataclass objects and legacy positional arguments.

        Collision Rules:
            - Two registers overlap if they share the same Info1 (Modbus address space)
              and their numerical address ranges [start, start + length - 1] intersect.
            - Exception: Multiple BITS entries packed into the same 16-bit word do NOT
              conflict if their bit slices [bit_start, bit_end] are disjoint.

        Args:
            info1: Info1 function code string or a RegisterEntry instance.
            address: Address string (ignored if info1 is a RegisterEntry).
            dtype: Data type string.
            name: Register variable name.
            line_num: Current line number.
            address_usage: State dictionary storing sorted intervals per Info1 space.
            warned_lines: Set tracking reported collision line pairs to prevent duplicate warnings.

        Returns:
            bool: True if an address overlap is detected, False otherwise.
        """
        if isinstance(info1, RegisterEntry):
            entry = info1
            if isinstance(address, dict):
                address_usage = address
            if isinstance(dtype, set):
                warned_lines = dtype
            info1_val = entry.info1
            address_val = entry.address
            dtype_val = entry.dtype
            name_val = entry.name
            line_num_val = entry.line_num
        else:
            info1_val = str(info1)
            address_val = str(address) if address is not None else ""
            dtype_val = str(dtype)
            name_val = str(name)
            line_num_val = int(line_num)

        if address_usage is None:
            address_usage = {}
        if warned_lines is None:
            warned_lines = set()

        try:
            addr_part = address_val.split("_")[0]
            start_addr = int(self.normalize_address_val(addr_part))
            reg_count = self.get_register_count(dtype_val, address_val)
            end_addr = start_addr + reg_count - 1

            if info1_val not in address_usage:
                address_usage[info1_val] = {"intervals": [], "max_len": 0}

            usage = address_usage[info1_val]
            intervals = usage["intervals"]
            max_len = usage["max_len"]

            is_bits = dtype_val.upper() == "BITS"
            current_bits = self._bit_slice(address_val, dtype_val)
            overlap_detected = False

            # Binary search for interval insertion position
            idx = bisect.bisect_left(intervals, (start_addr, -1, -1, "", "", -1, -1))

            def slices_overlap(u_type: str, u_start: int, u_bit_start: int, u_bit_end: int) -> bool:
                """Checks if two entries in the same word have overlapping bit slices."""
                if not (is_bits and u_type == "BITS" and start_addr == u_start):
                    return True
                if current_bits is None or u_bit_start < 0 or u_bit_end < 0:
                    return True
                bit_start, bit_end = current_bits
                return max(bit_start, u_bit_start) <= min(bit_end, u_bit_end)

            # Scan forward intervals
            for j in range(idx, len(intervals)):
                u_start, u_end, u_line, u_name, u_type, u_bit_start, u_bit_end = intervals[j]
                if u_start > end_addr:
                    break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if not slices_overlap(u_type, u_start, u_bit_start, u_bit_end):
                        continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num_val, u_line)))
                    if warn_key not in warned_lines:
                        logging.warning(
                            f"Line {line_num_val}: Address overlap detected for '{name_val}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line})."
                        )
                        warned_lines.add(warn_key)

            # Scan backward intervals within maximum span
            for j in range(idx - 1, -1, -1):
                u_start, u_end, u_line, u_name, u_type, u_bit_start, u_bit_end = intervals[j]
                if start_addr - u_start > max_len:
                    break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if not slices_overlap(u_type, u_start, u_bit_start, u_bit_end):
                        continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num_val, u_line)))
                    if warn_key not in warned_lines:
                        logging.warning(
                            f"Line {line_num_val}: Address overlap detected for '{name_val}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line})."
                        )
                        warned_lines.add(warn_key)

            bit_start, bit_end = current_bits if current_bits is not None else (-1, -1)
            bisect.insort(
                intervals,
                (
                    start_addr,
                    end_addr,
                    line_num_val,
                    name_val,
                    dtype_val.upper(),
                    bit_start,
                    bit_end,
                ),
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
        """
        Calculates Webdyn linear scaling coefficients (CoefA and CoefB).

        The WebdynSunPM data logging engine applies the polynomial conversion:
            Physical Value = CoefA * Raw_Value + CoefB

        Calculation:
            CoefA = Factor * (10 ^ ScaleFactor)
            CoefB = Offset

        Returns:
            tuple[str, str]: (CoefA, CoefB) formatted to 6 decimal places.
        """
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
        """
        Streams and normalizes raw register records into validated WebdynSunPM dictionary rows.

        Args:
            rows: Iterable stream of extracted register dictionaries.
            address_offset: Global numerical offset to shift all register addresses.

        Yields:
            dict[str, Any]: Standardized Webdyn register row dictionary containing:
                - Info1: Modbus register type code ("1", "2", "3", "4")
                - Info2: Modbus address string
                - Info3: Normalized Webdyn type
                - Info4: Reserved field (empty string)
                - Name: Human-readable variable description
                - Tag: Normalized unique variable identifier
                - CoefA: Scaling multiplier factor
                - CoefB: Scaling offset bias
                - Unit: Engineering measurement unit
                - Action: Access code ("1"=Write/RW, "4"=Read-only)
        """
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
            action, scale_factor_str = (
                norm_row.get("action", ""),
                norm_row.get("scalefactor", ""),
            )

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
                address = _TRAILING_NON_DIGIT_RE.sub("", address)

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

            # Action normalization with intelligent defaulting:
            # Inputs and Discrete Inputs default to Read-Only ("4"), Coils/Holding default to Write/RW ("1")
            act_str = str(action).strip().upper()
            if not act_str:
                norm_action = (
                    ACTION_READ_ONLY
                    if info1 in (MODBUS_DISCRETE, MODBUS_INPUT)
                    else ACTION_WRITE_ONLY
                )
            elif act_str in ("R", "READ", "RO", "READ-ONLY", "READ ONLY", "4"):
                norm_action = ACTION_READ_ONLY
            elif act_str in (
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
            ):
                norm_action = ACTION_WRITE_ONLY
            elif act_str in self.allowed_actions:
                norm_action = act_str
            else:
                norm_action = (
                    ACTION_READ_ONLY
                    if info1 in (MODBUS_DISCRETE, MODBUS_INPUT)
                    else ACTION_WRITE_ONLY
                )

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
        self,
        filepath: str,
        strict: bool = False,
        strict_overlap: bool | None = None,
    ) -> ValidationReport:
        """
        Validates an existing WebdynSunPM definition CSV file and returns a detailed report.

        Checks:
            - UTF-8-sig / UTF-16 encoding with valid semicolon-separated columns.
            - Header format with protocol and manufacturer fields.
            - Minimum column count (11 columns).
            - Valid Info1 Modbus function codes (1, 2, 3, 4).
            - Fatal uniqueness check for variable Tags.
            - Supported Webdyn data types.
            - Address syntax and 0..65535 boundary compliance.
            - Finite numeric CoefA and CoefB values.
            - Address interval collision detection.

        Args:
            filepath: Path to the definition CSV file.
            strict: If True, treat all warnings as validation errors.
            strict_overlap: If True, address overlaps fail validation. Defaults to strict.

        Returns:
            ValidationReport: Summary report containing boolean validity, issue counts, and register statistics.
        """
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
                    # Webdyn format: Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                    info1, info2, info3, name, tag = (
                        row[1],
                        row[2],
                        row[3],
                        row[5],
                        row[6],
                    )

                    # Validate Info1
                    info1_str = str(info1).strip()
                    if info1_str not in MODBUS_VALID_INFO1:
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

                    # Validate Tag uniqueness (fatal error)
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

                    # Validate Address format and boundary range (fatal error)
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

                    # Validate Action permission
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

                    # Validate CoefA and CoefB floats
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
                        line=0,
                        severity="ERROR",
                        code="IO_ERROR",
                        field="file",
                        message=msg,
                    )
                ],
                stats={"errors": 1, "warnings": 0, "types": {}, "registers": 0},
            )

    def validate_csv(
        self,
        filepath: str,
        strict: bool = False,
        strict_overlap: bool | None = None,
    ) -> bool:
        """
        Validates an existing WebdynSunPM definition file and returns True if valid.

        Args:
            filepath: Path to definition CSV file.
            strict: If True, warnings are treated as fatal.
            strict_overlap: If True, address overlaps cause failure.

        Returns:
            bool: True if the file passed validation, False otherwise.
        """
        report = self.validate_csv_detailed(filepath, strict=strict, strict_overlap=strict_overlap)
        return report.is_valid

    @staticmethod
    def write_output_csv(
        output: str | Any | None,
        processed_rows: Iterable[dict[str, Any]],
        config: CSVHeaderConfig | GeneratorConfig | str | None = None,
        *args: Any,
        **kwargs: Any,
    ) -> None:
        """
        Writes WebdynSunPM definition CSV output atomically with formula injection sanitization.

        When output is a string filepath, the file is written to a temporary sibling file
        in the target directory, synced to disk via os.fsync, and atomically replaced into place.

        Args:
            output: Filepath string, open file object, or None (defaults to sys.stdout).
            processed_rows: Stream of normalized register dictionaries.
            config: Header configuration object or manufacturer string for legacy positional callers.
            *args: Optional legacy positional parameters (model, protocol, category, forced_write).
            **kwargs: Optional keyword parameters for header metadata.
        """
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
        type_labels = {
            "1": "Coils",
            "2": "Discrete",
            "3": "Holding",
            "4": "Input",
        }
        outfile: Any = None
        temp_path: str | None = None

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
                if temp_path is not None:
                    os.replace(temp_path, os.path.abspath(output))
                    temp_path = None

            summary = ", ".join([f"{type_labels[k]}: {v}" for k, v in type_counts.items() if v > 0])
            if summary:
                logging.info(f"Generated {total} registers ({summary})")
        except (OSError, csv.Error) as e:
            logging.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and outfile is not None:
                try:
                    outfile.close()
                except OSError:
                    pass
            if temp_path is not None:
                try:
                    os.unlink(temp_path)
                except OSError:
                    pass


# -----------------------------------------------------------------------------
# Template Generation & Top-Level Execution
# -----------------------------------------------------------------------------


def generate_template(output_file: str | None, mode: str = "input") -> None:
    """
    Generates a sample template CSV file.

    Modes:
        - "input": Generates a simplified, human-friendly register map CSV.
        - "definition": Generates a valid raw WebdynSunPM semicolon-delimited CSV.

    Args:
        output_file: Path to destination file, or None to print to standard output.
        mode: "input" or "definition".
    """
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
            [
                "2",
                "3",
                "40002",
                "U16",
                "",
                "Voltage",
                "voltage",
                "0.100000",
                "0.000000",
                "V",
                "4",
            ],
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
    config: GeneratorConfig,
    input_data: Iterable[dict[str, Any]] | None = None,
) -> None:
    """
    Executes definition generation from a GeneratorConfig object and optional input stream.

    Args:
        config: Execution configuration.
        input_data: Optional pre-extracted register stream. If None, reads from config.input_file.
    """
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


def main() -> None:
    """Command-line entrypoint for the standalone def_gen script."""
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
