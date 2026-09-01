#!/usr/bin/env python3
"""
WebdynSunPM Definition File Generator Core Module.

Provides register normalization, address range and overlap validation,
coefficient calculation, data type parsing, and WebdynSunPM CSV definition generation.
"""

import argparse
import bisect
import csv
import sys
import tempfile
import logging
import re
import math
import itertools
from functools import lru_cache
from typing import Dict, Optional, Any, Union, Tuple, Set, Iterator, Iterable
import os
from dataclasses import dataclass

def peek_generator(iterable: Optional[Iterable]) -> Tuple[bool, Iterator]:
    """
    Checks if an iterable is non-empty without fully consuming it.

    Args:
        iterable: An optional iterable object to inspect.

    Returns:
        Tuple[bool, Iterator]: A tuple containing a boolean flag indicating if data exists,
                               and an iterator yielding the full sequence without loss.
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
RE_TYPE_NUMERIC = re.compile(r'^([UI](8|16|32|64)|F(32|64))(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32(_(W|B|WB))?|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64(_(W|B|WB))?)$', re.IGNORECASE)

_CLEAN_TYPE_RE = re.compile(r'[^a-z0-9_]+')

# --- CSV / formula-injection hardening -------------------------------------
# A spreadsheet strips leading whitespace before deciding whether a cell is a
# formula, and normalises full-width punctuation to ASCII. Guarding only the
# raw first byte is therefore bypassable (cf. OWASP WSTG, CVE-2021-41270).
_CSV_LEADING_WHITESPACE = " \t\r\n\x0b\x0c\u00a0"
_CSV_FORMULA_TRIGGERS = frozenset(
    "=+-@|\t\r\n\x0b\x0c\u00a0\uff1d\uff0b\uff0d\uff20"
)
_RE_SIGNED_NUMBER = re.compile(r'^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$')

# --- Type normalisation fast paths -----------------------------------------
_RE_TYPE_STRING_N = re.compile(r'string\s*(\d+)')

# Synonym table ordered by specificity. The first pattern that matches anywhere
# in the token wins, so "floatdouble" must resolve to F64 via the float64|double
# rule before the float32|float rule is considered. Patterns are precompiled
# once here instead of re-resolved from re's internal cache on every call.
_TYPE_SYNONYMS = tuple(
    (re.compile(pattern), code)
    for pattern, code in (
        (r'unsigned integer 64|unsigned int 64|uint64', 'U64'),
        (r'signed integer 64|signed int 64|int64', 'I64'),
        (r'unsigned integer 32|unsigned int 32|uint32', 'U32'),
        (r'signed integer 32|signed int 32|int32', 'I32'),
        (r'unsigned integer 16|unsigned int 16|uint16', 'U16'),
        (r'signed integer 16|signed int 16|int16', 'I16'),
        (r'unsigned integer 8|unsigned int 8|uint8', 'U8'),
        (r'signed integer 8|signed int 8|int8', 'I8'),
        (r'float64|double', 'F64'),
        (r'float32|float', 'F32'),
        (r'string\s*(\d+)', None),
        (r'string', 'STRING'),
    )
)

_SUFFIX_WB = ('_wb', 'swap', 'big endian')
_SUFFIX_B = ('_b', 'big')
_SUFFIX_W = ('_w', 'word')

# Leading-zero decimals ("0021") are deliberately NOT fast-pathed: int(x, 0)
# rejects them, and the legacy contract falls through to hex interpretation.
_RE_PLAIN_DECIMAL = re.compile(r'^-?(0|[1-9]\d*)$')
_RE_HEX_WORD = re.compile(r'^[0-9A-Fa-f]+$')
_RE_THOUSANDS = re.compile(r'(?<=\d),(?=\d{3}(?!\d))')

_TYPE_CACHE_SIZE = 4096

_ACTION_READ = frozenset(('R', 'READ', 'RO', 'READ-ONLY', 'READ ONLY', '4'))
_ACTION_WRITE = frozenset(
    ('RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', 'R/W',
     'WO', 'WRITE-ONLY', 'WRITE ONLY', '1')
)
_READONLY_INFO1 = frozenset(('2', '4'))
_NUMERIC_INFO1 = frozenset(('1', '2', '3', '4'))

@dataclass
class GeneratorConfig:
    """Configuration settings for WebdynSunPM definition file generation."""
    input_file: Optional[str] = None
    output: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    protocol: str = 'modbusRTU'
    category: str = 'Inverter'
    forced_write: str = ''
    template: bool = False
    template_mode: str = 'input'
    address_offset: int = 0

class Generator:
    """Core generator class for processing Modbus register maps and building definition CSVs."""
    def __init__(self) -> None:
        self.register_type_map = {
            'coil': '1',
            'coils': '1',
            'discrete input': '2',
            'discrete': '2',
            'holding register': '3',
            'holding': '3',
            'input register': '4',
            'input': '4'
        }
        # Action codes accepted by WebdynSunPM. '10' designates a 'Constant'
        # variable, introduced in firmware 5.2.02: a value not read from the
        # device over Modbus but declared statically in the definition file.
        self.allowed_actions = frozenset(
            ('0', '1', '2', '4', '6', '7', '8', '9', '10')
        )

    @staticmethod
    def sanitize_csv_field(val: Any) -> str:
        """
        Precludes CSV/formula injection by prepending an apostrophe when needed.

        A spreadsheet client evaluates a cell as a formula when its first
        *significant* character is a trigger. Leading whitespace (space, tab,
        CR, LF, VT, FF, NBSP) is stripped by the client before that test, and
        full-width Unicode punctuation is normalised to ASCII, so both classes
        of prefix are treated as triggers here. The DDE pipe is included per
        OWASP guidance.

        Genuine finite signed numbers (``-10.5``, ``+25``, ``1.5e3``) are
        preserved verbatim so negative coefficients and signed Modbus
        addresses stay usable downstream. Non-finite literals (``-inf``,
        ``+nan``) and underscore-grouped forms (``-1_000``) are escaped: they
        are accepted by ``float()`` but are not valid definition payloads.
        """
        if val is None:
            return ""
        s = str(val)
        if not s:
            return s
        stripped = s.lstrip(_CSV_LEADING_WHITESPACE)
        if s[0] not in _CSV_FORMULA_TRIGGERS and stripped[:1] not in _CSV_FORMULA_TRIGGERS:
            return s
        if s == stripped and _RE_SIGNED_NUMBER.match(s):
            try:
                if math.isfinite(float(s)):
                    return s
            except ValueError:
                pass
        return "'" + s

    @staticmethod
    @lru_cache(maxsize=_TYPE_CACHE_SIZE)
    def _normalize_type_cached(t: str) -> str:
        """Pure, memoised core of :meth:`normalize_type` keyed on the lowered token."""
        suffix = ''
        if any(x in t for x in _SUFFIX_WB):
            suffix = '_WB'
        elif any(x in t for x in _SUFFIX_B):
            suffix = '_B'
        elif any(x in t for x in _SUFFIX_W):
            suffix = '_W'

        str_match = _RE_TYPE_STRING_N.search(t)
        if str_match:
            return f"STR{str_match.group(1)}{suffix}"

        for pattern, code in _TYPE_SYNONYMS:
            m = pattern.search(t)
            if m is None:
                continue
            if code is None:
                return f"STR{m.group(1)}{suffix}"
            return f"{code}{suffix}"

        if t.startswith('str') and t[3:].isdigit():
            return t.upper()

        cleaned = _CLEAN_TYPE_RE.sub('', t)
        return cleaned.upper() if cleaned else 'U16'

    @staticmethod
    def normalize_type(dtype: Any) -> str:
        """
        Normalises a free-form data type token to its WebdynSunPM code.

        Recognises vendor spellings (``unsigned int 32``, ``float64``),
        endianness hints (``_wb``, ``swap``, ``big endian``, ``word``) and
        sized strings (``string 20`` -> ``STR20``). Results are memoised
        because register maps reuse very few distinct tokens across many rows.
        """
        if not dtype:
            return 'U16'
        return Generator._normalize_type_cached(str(dtype).lower().strip())

    @staticmethod
    def validate_type(dtype: str) -> bool:
        """Validates the data type."""
        dtype_upper = str(dtype).upper()
        # Base types
        base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC']
        if dtype_upper in base_types:
            return True

        # Numeric types (Int/Float) with optional suffixes
        if RE_TYPE_NUMERIC.match(dtype_upper):
            return True

        # STR<n> syntax (e.g., STR20)
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return True

        return False

    @staticmethod
    @lru_cache(maxsize=_TYPE_CACHE_SIZE)
    def _normalize_address_cached(addr_part: str) -> str:
        """Pure, memoised core of :meth:`normalize_address_val`."""
        if ',' in addr_part:
            addr_part = _RE_THOUSANDS.sub('', addr_part)
        if not addr_part:
            return ""
        if _RE_PLAIN_DECIMAL.match(addr_part):
            return str(int(addr_part))
        low = addr_part.lower()
        if low.startswith('0x'):
            try:
                return str(int(addr_part, 16))
            except ValueError:
                return addr_part
        if low.endswith('h'):
            try:
                return str(int(addr_part[:-1], 16))
            except ValueError:
                return addr_part
        try:
            return str(int(addr_part, 0))
        except ValueError:
            pass
        if _RE_HEX_WORD.match(addr_part):
            try:
                return str(int(addr_part, 16))
            except ValueError:
                return addr_part
        return addr_part

    @staticmethod
    def normalize_address_val(addr_part: Any) -> str:
        """
        Converts one address component to its canonical decimal string.

        Accepts decimal, ``0x``-prefixed hex, ``h``-suffixed hex, bare hex
        words, negative values, and thousands-separated digits. Unparseable
        input is returned unchanged so the caller's validator can reject it.
        """
        return Generator._normalize_address_cached(str(addr_part).strip())

    @staticmethod
    def validate_address(address: str, dtype: str, strict: bool = True) -> bool:
        """Validates the address format based on type and Modbus range (0-65535)."""
        dtype_upper = dtype.upper()
        if RE_TYPE_STR_CONV.match(dtype_upper):
            dtype_upper = 'STRING'

        is_valid_format = False
        if dtype_upper == 'STRING':
            is_valid_format = RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == 'BITS':
            is_valid_format = RE_ADDR_BITS.match(address) is not None
        else:
            is_valid_format = RE_ADDR_INT.match(address) is not None

        if not is_valid_format:
            return False

        try:
            # Handle compound addresses by splitting and normalizing the base part
            parts = address.split('_')
            base_addr_str = Generator.normalize_address_val(parts[0])
            base_addr = int(base_addr_str, 0)
            if not (0 <= base_addr <= 65535):
                logging.warning(f"Address {base_addr} is out of standard Modbus range (0-65535)")
                if strict:
                    return False
        except (ValueError, IndexError):
            return False
        return True

    @staticmethod
    @lru_cache(maxsize=_TYPE_CACHE_SIZE)
    def _register_count_for_type(dtype_upper: str) -> Optional[int]:
        """Memoised width lookup for non-STRING types; ``None`` means STRING."""
        if RE_COUNT_16_8.match(dtype_upper):
            return 1
        if RE_COUNT_32.match(dtype_upper):
            return 2
        if RE_COUNT_64.match(dtype_upper):
            return 4
        if dtype_upper == 'MAC':
            return 3
        if dtype_upper == 'IPV6':
            return 8
        if dtype_upper == 'STRING':
            return None
        return 1

    @staticmethod
    def get_register_count(dtype: str, address: str) -> int:
        """
        Returns how many 16-bit Modbus registers a value of ``dtype`` occupies.

        STRING widths depend on the compound address suffix (``<addr>_<len>``)
        and are computed per call; all other widths are memoised on the type.
        """
        dtype_upper = dtype.upper()
        count = Generator._register_count_for_type(dtype_upper)
        if count is not None:
            return count
        try:
            return math.ceil(int(address.split('_')[1]) / 2)
        except (IndexError, ValueError):
            return 0

    @staticmethod
    def _parse_numeric(val: Any, default: float = 0.0) -> float:
        if val is None or str(val).strip() == '': return default
        s = str(val).strip()
        if '/' in s:
            try:
                parts = s.split('/')
                return float(parts[0]) / float(parts[1])
            except (ValueError, ZeroDivisionError, IndexError): return default
        if ',' in s and '.' in s:
            if s.find(',') < s.find('.'): s = s.replace(',', '')
            else: s = s.replace('.', '').replace(',', '.')
        elif ',' in s:
            if re.match(r'^-?\d{1,3}(,\d{3})+$', s): s = s.replace(',', '')
            else: s = s.replace(',', '.')
        try: return float(s)
        except ValueError: return default

    @staticmethod
    def apply_address_offset(address: Any, offset: int, line_num: Optional[int] = None, name: Optional[str] = None) -> str:
        """Applies an integer offset to a register address (simple or compound)."""
        if not address:
            return ""
        parts = str(address).split('_')
        # Normalize each part individually to avoid int() consuming underscores
        norm_parts = [Generator.normalize_address_val(p) for p in parts]
        try:
            base_addr = int(norm_parts[0]) + offset
            if base_addr < 0:
                msg = f"Address offset {offset} results in negative address {base_addr}"
                if name: msg += f" for '{name}'"
                if line_num: logging.warning(f"Line {line_num}: {msg}")
                else: logging.warning(msg)
            norm_parts[0] = str(base_addr)
        except (ValueError, IndexError): pass
        return '_'.join(norm_parts)

    def _process_name_and_tag(self, name: str, tag: str, line_num: int, seen_names: Dict[str, int], seen_tags: Dict[str, int]) -> str:
        if name:
            if name in seen_names: logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous at line {seen_names[name]}.")
            else: seen_names[name] = line_num
        if not tag and name:
            base_tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
            base_tag = re.sub(r'_+', '_', base_tag).strip('_')
            if not base_tag or not base_tag[0].isalpha(): base_tag = f"v_{base_tag}" if base_tag else "var"
            tag = base_tag
            counter = 1
            while tag in seen_tags:
                tag = f"{base_tag}_{counter}"
                counter += 1
        if tag:
            if tag in seen_tags: logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous at line {seen_tags[tag]}.")
            else: seen_tags[tag] = line_num
        return tag

    def _determine_info1(self, reg_type_str: str, line_num: int = 0) -> str:
        """Maps RegisterType string to Webdyn Info1 code."""
        if reg_type_str is None:
            return '3'

        lt = str(reg_type_str).lower().strip()
        if not lt:
            return '3'
        if lt in self.register_type_map:
            return self.register_type_map[lt]
        elif str(reg_type_str).strip() in _NUMERIC_INFO1:
            return str(reg_type_str).strip()

        if line_num:
            logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to 3.")
        return '3'

    def _check_address_overlap(self, info1: str, address: str, dtype: str, name: str, line_num: int, address_usage: Dict[str, Dict[str, Any]], warned_lines: Set[Tuple[int, int]]) -> None:
        """Checks for address overlaps using O(log N) binary search on intervals."""
        try:
            addr_part = address.split('_')[0]
            start_addr = int(Generator.normalize_address_val(addr_part))
            reg_count = self.get_register_count(dtype, address)
            end_addr = start_addr + reg_count - 1

            if info1 not in address_usage:
                address_usage[info1] = {'intervals': [], 'max_len': 0}

            usage = address_usage[info1]
            intervals = usage['intervals']
            max_len = usage['max_len']

            is_bits = (dtype.upper() == 'BITS')

            idx = bisect.bisect_left(intervals, (start_addr, -1, -1, '', ''))

            # Check to the right
            for j in range(idx, len(intervals)):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if u_start > end_addr:
                    break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if is_bits and u_type == 'BITS' and start_addr == u_start:
                        continue
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        overlap_start = max(start_addr, u_start)
                        logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {overlap_start}. Overlaps with '{u_name}' (Line {u_line}).")
                        warned_lines.add(warn_key)

            # Check to the left
            for j in range(idx - 1, -1, -1):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if start_addr - u_start > max_len:
                    break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if is_bits and u_type == 'BITS' and start_addr == u_start:
                        continue
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        overlap_start = max(start_addr, u_start)
                        logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {overlap_start}. Overlaps with '{u_name}' (Line {u_line}).")
                        warned_lines.add(warn_key)

            bisect.insort(intervals, (start_addr, end_addr, line_num, name, dtype.upper()))
            if reg_count > max_len:
                usage['max_len'] = reg_count

        except (ValueError, IndexError):
            pass

    @staticmethod
    def _calculate_coefficients(factor_str: Any, offset_str: Any, scale_factor_str: Any) -> Tuple[str, str]:
        factor = Generator._parse_numeric(factor_str, default=1.0)
        offset = Generator._parse_numeric(offset_str, default=0.0)
        try: scale_val = int(float(scale_factor_str)) if scale_factor_str else 0
        except ValueError: scale_val = 0
        coef_a = "{:.6f}".format(factor * (10 ** scale_val))
        coef_b = "{:.6f}".format(offset)
        return coef_a, coef_b

    def process_rows(self, rows: Iterable[Dict[str, Any]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        seen_names, seen_tags, address_usage, warned_lines = {}, {}, {}, set()
        for line_num, row in enumerate(rows, start=2):
            if not any(row.values()):
                continue
            norm_row = {k.lower().strip(): (str(v).strip() if v is not None else '') for k, v in row.items()}
            name, tag, reg_type_str, address = norm_row.get('name', ''), norm_row.get('tag', ''), norm_row.get('registertype', ''), norm_row.get('address', '')
            dtype_raw, factor, offset, unit = norm_row.get('type', ''), norm_row.get('factor', ''), norm_row.get('offset', ''), norm_row.get('unit', '')
            action, scale_factor_str = norm_row.get('action', ''), norm_row.get('scalefactor', '')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue
            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype_raw}' (normalized to '{dtype}'). Skipping.")
                continue
            match_str = RE_TYPE_STR_CONV.match(dtype)
            if match_str:
                dtype = 'STRING'
                if '_' not in address: address = f"{address}_{match_str.group(1)}"
            address = Generator.apply_address_offset(address, address_offset, line_num, name)
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping.")
                continue
            tag = self._process_name_and_tag(name, tag, line_num, seen_names, seen_tags)
            info1 = self._determine_info1(reg_type_str, line_num)
            self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)
            coef_a, coef_b = self._calculate_coefficients(factor, offset, scale_factor_str)

            # Action normalisation with intelligent defaulting.
            act_str = action.upper()
            if act_str in _ACTION_READ:
                norm_action = '4'
            elif act_str in _ACTION_WRITE:
                norm_action = '1'
            elif act_str in self.allowed_actions:
                norm_action = act_str
            else:
                norm_action = '4' if info1 in _READONLY_INFO1 else '1'

            yield {
                'Info1': info1,
                'Info2': address,
                'Info3': dtype.upper(),
                'Info4': '',
                'Name': name,
                'Tag': tag,
                'CoefA': coef_a,
                'CoefB': coef_b,
                'Unit': unit,
                'Action': norm_action
            }

    def validate_csv(self, filepath: str, strict: bool = True) -> bool:
        """Validates an existing WebdynSunPM definition file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, 'r', encoding=encoding) as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header or len(header) < 2 or not any(header):
                    logging.error("Invalid WebdynSunPM header.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row): continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Row has insufficient columns.")
                        continue

                    # row format: Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                    info1, info2, info3, name, tag = row[1], row[2], row[3], row[5], row[6]

                    # Validate Tag uniqueness (fatal)
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (previously at line {seen_tags[tag]}).")
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Validate Type (warning/fatal?)
                    if not self.validate_type(info3):
                        logging.warning(f"Line {line_num}: Invalid Type '{info3}'.")
                        valid = False

                    # Validate Address format and range (fatal)
                    if not self.validate_address(info2, info3, strict=strict):
                        logging.error(f"Line {line_num}: Invalid address or range '{info2}' for Type '{info3}'.")
                        valid = False

                    self._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)

            if strict and warned_lines:
                logging.error("Address overlaps detected (strict=True). Validation failed.")
                valid = False

            return valid

        except (OSError, csv.Error) as e:
            logging.error(f"Error reading definition file: {e}")
            return False

    @staticmethod
    def write_output_csv(output: Union[str, Any, None], processed_rows: Iterable[Dict[str, Any]], manufacturer: str, model: str,
                        protocol: str = 'modbusRTU', category: str = 'Inverter', forced_write: str = '') -> None:
        """
        Writes the WebdynSunPM definition CSV.

        When ``output`` is a path the file is written atomically: content goes
        to a temporary file on the same filesystem and is then moved into
        place, so a crash or a mid-stream error can never leave a truncated
        CSV where a valid production definition used to be. Streams and
        ``None`` (stdout) are written directly.
        """
        type_counts = {'1': 0, '2': 0, '3': 0, '4': 0}
        type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}
        sanitize = Generator.sanitize_csv_field

        def _emit(handle: Any) -> int:
            writer = csv.writer(handle, delimiter=';', lineterminator='\n')
            writer.writerow([
                sanitize(protocol), sanitize(category), sanitize(manufacturer),
                sanitize(model), sanitize(forced_write), '', '', '', '', '', ''
            ])
            written = 0
            for index, row in enumerate(processed_rows, start=1):
                writer.writerow([
                    str(index), sanitize(row['Info1']), sanitize(row['Info2']),
                    sanitize(row['Info3']), sanitize(row['Info4']), sanitize(row['Name']),
                    sanitize(row['Tag']), sanitize(row['CoefA']), sanitize(row['CoefB']),
                    sanitize(row['Unit']), sanitize(row['Action']),
                ])
                info1 = row['Info1']
                type_counts[info1] = type_counts.get(info1, 0) + 1
                written += 1
            return written

        total = 0
        try:
            if isinstance(output, str):
                directory = os.path.dirname(os.path.abspath(output)) or '.'
                tmp_fd, tmp_path = tempfile.mkstemp(
                    dir=directory, prefix='.webdyn-', suffix='.csv.tmp'
                )
                try:
                    with os.fdopen(tmp_fd, 'w', newline='', encoding='utf-8-sig') as tmp:
                        total = _emit(tmp)
                    os.replace(tmp_path, output)
                except BaseException:
                    try:
                        os.unlink(tmp_path)
                    except OSError:
                        pass
                    raise
            elif output is None:
                total = _emit(sys.stdout)
            else:
                total = _emit(output)

            summary = ", ".join(
                f"{type_labels[k]}: {v}" for k, v in type_counts.items() if v > 0
            )
            if summary:
                logging.info(f"Generated {total} registers ({summary})")
        except (OSError, csv.Error) as e:
            logging.error(f"Error writing output CSV: {e}")

def generate_template(output_file: Optional[str], mode: str = 'input') -> None:
    if mode == 'definition':
        headers = ['#Index', 'Info1', 'Info2', 'Info3', 'Info4', 'Name', 'Tag', 'CoefA', 'CoefB', 'Unit', 'Action']
        rows = [
            ['modbusRTU', 'Inverter', 'SampleManufacturer', 'SampleModel', '', '', '', '', '', '', ''],
            ['1', '3', '40001', 'U16', '', 'Active Power', 'active_power', '1.000000', '0.000000', 'W', '4'],
            ['2', '3', '40002', 'U16', '', 'Voltage', 'voltage', '0.100000', '0.000000', 'V', '4']
        ]
        delimiter = ';'
    else:
        headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]
        delimiter = ','

    try:
        if output_file:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=delimiter)
                writer.writerow(headers)
                writer.writerows(rows)
        else:
            writer = csv.writer(sys.stdout, delimiter=delimiter)
            writer.writerow(headers)
            writer.writerows(rows)
    except OSError as e:
        logging.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig, input_data: Optional[Iterable[Dict[str, Any]]] = None) -> None:
    generator = Generator()
    if config.template:
        mode = config.template_mode
        if input_data is not None:
            mode = 'definition'
        generate_template(config.output, mode=mode)
        return

    manufacturer = config.manufacturer or 'Manufacturer'
    model = config.model or 'Model'

    if input_data is None:
        if not config.input_file:
            logging.error("input_file or input_data is required.")
            return
        if not os.path.exists(config.input_file):
            logging.error(f"Input file not found: {config.input_file}")
            return

    try:
        if input_data is not None:
            processed_rows = generator.process_rows(input_data, config.address_offset)
            generator.write_output_csv(config.output, processed_rows, manufacturer, model,
                                       config.protocol, config.category, config.forced_write)
        else:
            with open(config.input_file, mode='rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(config.input_file, mode='r', encoding=encoding) as csvfile:
                snippet = csvfile.read(2048)
                csvfile.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                except csv.Error:
                    dialect = csv.excel
                reader = csv.DictReader(csvfile, dialect=dialect)
                processed_rows = generator.process_rows(reader, config.address_offset)

                generator.write_output_csv(config.output, processed_rows, manufacturer, model,
                                           config.protocol, config.category, config.forced_write)
    except (OSError, csv.Error, ValueError, TypeError, KeyError) as e:
        logging.error(f"An error occurred during generation: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM Modbus definition file.')
    parser.add_argument('input_file', nargs='?', help='Input simplified CSV.')
    parser.add_argument('-o', '--output', help='Output CSV.')
    parser.add_argument('--manufacturer', help='Manufacturer name.')
    parser.add_argument('--model', help='Model name.')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true')
    parser.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    parser.add_argument('--address-offset', type=int, default=0)

    args = parser.parse_args()
    config = GeneratorConfig(
        input_file=args.input_file, output=args.output,
        manufacturer=args.manufacturer, model=args.model,
        protocol=args.protocol, category=args.category,
        forced_write=args.forced_write, template=args.template,
        template_mode=args.template_mode,
        address_offset=args.address_offset
    )
    run_generator(config)

if __name__ == "__main__":
    main()