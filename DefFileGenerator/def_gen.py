#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
import itertools
from typing import Dict, List, Optional, Any, Union, Tuple, Set, Iterator, Iterable
import os
import itertools
from dataclasses import dataclass

def peek_generator(iterable: Optional[Iterable]) -> Tuple[bool, Iterable]:
    """
    Checks if an iterable is empty without consuming it.
    Returns (has_data, original_iterable).
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

def peek_generator(iterable: Optional[Iterable]) -> Tuple[bool, Iterator]:
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

@dataclass
class GeneratorConfig:
    input_file: str = None
    output: str = None
    manufacturer: str = None
    model: str = None
    protocol: str = 'modbusRTU'
    category: str = 'Inverter'
    forced_write: str = ''
    template: bool = False
    template_mode: str = 'input'
    address_offset: int = 0
    template_mode: str = 'input'

class Generator:
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
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing WebdynSunPM definition file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader)
                if len(header) < 4:
                    logging.error("Invalid header in definition file.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row): continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns. Skipping.")
                        continue

                    # row format: Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                    info1, info2, info3, name, tag = row[1], row[2], row[3], row[5], row[6]

                    # Validate Tag uniqueness
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (Previously at line {seen_tags[tag]}).")
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Validate Type
                    if not self.validate_type(info3):
                        logging.warning(f"Line {line_num}: Invalid Type '{info3}'.")
                        valid = False

                    # Validate Address
                    if not self.validate_address(info2, info3):
                        logging.warning(f"Line {line_num}: Invalid Address '{info2}' for Type '{info3}'.")
                        valid = False
                    else:
                        self._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)

            # If address overlaps are warnings, we don't necessarily set valid=False,
            # but duplicate tags ARE fatal errors in Webdyn.
            return valid

        except (OSError, csv.Error) as e:
            logging.error(f"Error reading definition file: {e}")
            return False

    @staticmethod
    def normalize_type(dtype):
        if not dtype: return 'U16'
        t = str(dtype).lower().strip()
        suffix = ''
        if any(x in t for x in ['_wb', 'swap', 'big endian']): suffix = '_WB'
        elif any(x in t for x in ['_b', 'big']): suffix = '_B'
        elif any(x in t for x in ['_w', 'word']): suffix = '_W'

        # Handle "string 20" -> "STR20"
        str_match = re.search(r'string\s*(\d+)', t)
        if str_match:
            return f"STR{str_match.group(1)}{suffix}"

        # Mapping ordered by specificity (longer strings first)
        synonyms = [
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
            (r'string\s*(\d+)', r'STR\1'),
            (r'string', 'STRING'),
        ]
        for pattern, replacement in synonyms:
            if re.search(pattern, t):
                if r'\1' in replacement:
                    res = re.sub(pattern, replacement, t).upper()
                    return f"{res}{suffix}"
                return f"{replacement}{suffix}"

        # Handle STR<n> explicitly if it comes in as raw type
        if t.startswith('str') and t[3:].isdigit():
            return t.upper()

        t = _CLEAN_TYPE_RE.sub('', t)
        return t.upper() if t else 'U16'

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
    def normalize_address_val(addr_part: Any) -> str:
        """Converts a single address part (possibly hex) to decimal string."""
        addr_part = str(addr_part).strip()
        addr_part = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_part)
        if not addr_part: return ""
        if addr_part.lower().startswith('0x'):
            try: return str(int(addr_part, 16))
            except ValueError: return addr_part
        elif addr_part.lower().endswith('h'):
            try:
                return str(int(addr_part[:-1], 16))
            except ValueError:
                return addr_part

        # Handle decimal (including negative)
        try:
            return str(int(addr_part, 0))
        except ValueError:
            pass

        # If it's a raw hex word (e.g. "A0")
        if re.match(r'^[0-9A-Fa-f]+$', addr_part):
            try: return str(int(addr_part, 16))
            except ValueError: return addr_part
        return addr_part

    @staticmethod
    def validate_address(address: str, dtype: str) -> bool:
        """Validates the address format based on type and Modbus range (0-65535)."""
        dtype_upper = dtype.upper()
        if RE_TYPE_STR_CONV.match(dtype_upper):
            dtype_upper = 'STRING'

        is_valid_format = False
        if dtype_upper == 'STRING':
            is_valid_format = RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == 'BITS':
            match = RE_ADDR_BITS.match(address)
        else:
            match = RE_ADDR_INT.match(address)

        try:
            # Handle compound addresses by splitting and normalizing the base part
            parts = address.split('_')
            base_addr_str = Generator.normalize_address_val(parts[0])
            base_addr = int(base_addr_str, 0)
            if not (0 <= base_addr <= 65535):
                logging.warning(f"Address {base_addr} is out of Modbus range (0-65535)")
                return False
        return is_valid

    def validate_csv(self, filepath: str) -> bool:
        """Performs comprehensive validation of a definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        success = True
        try:
            with open(filepath, 'rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, 'r', encoding=encoding) as f:
                # Determine if it's a Webdyn definition file (lineterminator=; and first row has protocol)
                snippet = f.read(1024)
                f.seek(0)

                is_webdyn = ';' in snippet and any(p in snippet for p in ['modbusRTU', 'modbusTCP'])

                if is_webdyn:
                    reader = csv.reader(f, delimiter=';')
                    next(reader) # Skip global header
                    for line_num, row in enumerate(reader, start=2):
                        if not row or len(row) < 11: continue
                        # Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
                        dtype = row[3]
                        addr = row[2]
                        name = row[5]
                        if not self.validate_type(dtype):
                            logging.error(f"Line {line_num}: Invalid type '{dtype}' for '{name}'")
                            success = False
                        if not self.validate_address(addr, dtype):
                            logging.error(f"Line {line_num}: Invalid address '{addr}' for '{name}'")
                            success = False
                else:
                    # Assume simplified CSV with headers
                    f.seek(0)
                    reader = csv.DictReader(f)
                    for line_num, row in enumerate(reader, start=2):
                        dtype = row.get('Type', '')
                        addr = row.get('Address', '')
                        name = row.get('Name', '')
                        if not dtype or not addr: continue
                        if not self.validate_type(dtype):
                            logging.error(f"Line {line_num}: Invalid type '{dtype}' for '{name}'")
                            success = False
                        if not self.validate_address(addr, dtype):
                            logging.error(f"Line {line_num}: Invalid address '{addr}' for '{name}'")
                            success = False
        except Exception as e:
            logging.error(f"Validation error: {e}")
            return False

        return success

    def validate_csv(self, filepath: str) -> bool:
        """Validates a WebdynSunPM definition file for correct formatting and range constraints."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header or len(header) < 4:
                    logging.error("Invalid WebdynSunPM header.")
                    return False

                success = True
                for line_num, row in enumerate(reader, start=2):
                    if not row: continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Row has fewer than 11 columns.")
                        success = False
                        continue

                    addr = row[2]
                    dtype = row[3]
                    if not self.validate_address(addr, dtype):
                        logging.warning(f"Line {line_num}: Invalid address/type combination '{addr}' ['{dtype}'].")
                        success = False
                return success
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV: {e}")
            return False

    @staticmethod
    def get_register_count(dtype: str, address: str) -> int:
        dtype_upper = dtype.upper()
        if RE_COUNT_16_8.match(dtype_upper): return 1
        elif RE_COUNT_32.match(dtype_upper): return 2
        elif RE_COUNT_64.match(dtype_upper): return 4
        elif dtype_upper == 'MAC': return 3
        elif dtype_upper == 'IPV6': return 8
        elif dtype_upper == 'STRING':
            try: return math.ceil(int(address.split('_')[1]) / 2)
            except (IndexError, ValueError): return 0
        return 1

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

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found for validation: {filepath}")
            return False

        is_valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, mode='rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, mode='r', encoding=encoding) as f:
                reader = csv.reader(f, delimiter=';')
                # Skip header row
                try:
                    next(reader)
                except StopIteration:
                    logging.error("Empty definition file.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row: continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns ({len(row)}/11). Skipping.")
                        continue

                    # Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action
                    info1 = row[1].strip()
                    address = row[2].strip()
                    dtype = row[3].strip()
                    name = row[5].strip()
                    tag = row[6].strip()

                    # 1. Validate Address Format and Range
                    if not self.validate_address(address, dtype):
                        logging.error(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'.")
                        is_valid = False

                    # 2. Check for Duplicate Tags (Fatal)
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Duplicate Tag '{tag}' detected. Fatal error. (Previous at line {seen_tags[tag]})")
                            is_valid = False
                        else:
                            seen_tags[tag] = line_num

                    # 3. Check for Overlaps (Warning)
                    self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            return is_valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error during validation: {e}")
            return False

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
        elif str(reg_type_str).strip() in ['1', '2', '3', '4']:
            return str(reg_type_str).strip()

        if line_num:
            logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to 3.")
        return '3'

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing WebdynSunPM definition file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header:
                    logging.error("Empty file.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or len(row) < 11:
                        if any(row):
                            logging.warning(f"Line {line_num}: Invalid row format (expected at least 11 columns).")
                        continue

                    # Header/Meta mapping from process_rows:
                    # 0:Index, 1:Info1, 2:Info2, 3:Info3, 4:Info4, 5:Name, 6:Tag, 7:CoefA, 8:CoefB, 9:Unit, 10:Action
                    info1 = row[1]
                    address = row[2]
                    dtype = row[3]
                    name = row[5]
                    tag = row[6]

                    if not tag:
                        logging.error(f"Line {line_num}: Missing Tag.")
                        valid = False
                    elif tag in seen_tags:
                        logging.error(f"Line {line_num}: Duplicate Tag '{tag}' (Fatal). Previously at line {seen_tags[tag]}.")
                        valid = False
                    else:
                        seen_tags[tag] = line_num

                    if not self.validate_address(address, dtype):
                        logging.error(f"Line {line_num}: Invalid Address '{address}' for type '{dtype}'.")
                        valid = False

                    self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            return valid
        except Exception as e:
            logging.error(f"Error validating CSV: {e}")
            return False

    def _check_address_overlap(self, info1: str, address: str, dtype: str, name: str, line_num: int, address_usage: Dict[str, Dict[int, List[Tuple[int, str, str]]]], warned_lines: Set[Tuple[int, int]]) -> None:
        try:
            addr_part = address.split('_')[0]
            start_addr = int(Generator.normalize_address_val(addr_part))
            reg_count = self.get_register_count(dtype, address)
            if info1 not in address_usage: address_usage[info1] = {}
            is_bits = (dtype.upper() == 'BITS')
            for i in range(reg_count):
                curr_addr = start_addr + i
                if curr_addr in address_usage[info1]:
                    for u_line, u_name, u_type in address_usage[info1][curr_addr]:
                        if is_bits and u_type == 'BITS' and curr_addr == start_addr: continue
                        warn_key = tuple(sorted((line_num, u_line)))
                        if warn_key not in warned_lines:
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {curr_addr}. Overlaps with '{u_name}' (Line {u_line}).")
                            warned_lines.add(warn_key)
                else: address_usage[info1][curr_addr] = []
                address_usage[info1][curr_addr].append((line_num, name, dtype.upper()))
        except (ValueError, IndexError): pass

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, mode='rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, mode='r', encoding=encoding) as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                try:
                    header = next(reader)
                except StopIteration:
                    logging.error(f"Empty file: {filepath}")
                    return False

                if len(header) < 4:
                    logging.error(f"Invalid Webdyn header in {filepath}")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row):
                        continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns (expected 11, got {len(row)}). Skipping.")
                        continue

                    info1, info2, info3 = row[1].strip(), row[2].strip(), row[3].strip()
                    name, tag = row[5].strip(), row[6].strip()

                    # Compound address handling and validation
                    addr_parts = info2.split('_')
                    norm_addr = '_'.join([Generator.normalize_address_val(p) for p in addr_parts])

                    if not self.validate_address(norm_addr, info3):
                        valid = False

                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Duplicate Tag '{tag}' (Fatal). Previous occurrence at line {seen_tags[tag]}.")
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    self._check_address_overlap(info1, norm_addr, info3, name, line_num, address_usage, warned_lines)

                    if info1 not in ['1', '2', '3', '4']:
                        logging.warning(f"Line {line_num}: Invalid RegisterType code '{info1}'")
        except (OSError, csv.Error) as e:
            logging.error(f"Error reading {filepath} for validation: {e}")
            return False
        return valid

    @staticmethod
    def validate_csv(filepath: str) -> bool:
        """
        Validates an existing Webdyn definition CSV file.
        Checks for duplicate tags, address overlaps, and basic format.
        """
        if not os.path.exists(filepath):
            logging.error(f"Validation failed: File not found {filepath}")
            return False

        seen_tags = {}
        address_usage = {} # Info1 -> dict of address -> list of (line, name, type)
        warned_lines = set()
        is_valid = True

        generator = Generator()

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header:
                    logging.error("Validation failed: Empty file")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or len(row) < 11:
                        continue

                    # Webdyn format: Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                    # Indices: 0; 1; 2; 3; 4; 5; 6; 7; 8; 9; 10
                    info1 = row[1].strip()
                    address = row[2].strip()
                    dtype = row[3].strip()
                    name = row[5].strip()
                    tag = row[6].strip()

                    # Tag uniqueness
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Duplicate Tag '{tag}' (previously at line {seen_tags[tag]})")
                            is_valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Address range and format validation
                    if address and dtype:
                        if not Generator.validate_address(address, dtype):
                            is_valid = False

                    # Address overlap check
                    if info1 and address and dtype:
                        generator._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            # If any overlaps were found (which are logged as warnings by _check_address_overlap),
            # for strict validation we might want to consider them errors.
            # But the requirement said "log a warning if the normalized address value falls outside this standard range"
            # and "duplicate tag detection (which is a fatal error)".
            # Let's check warned_lines.
            if warned_lines:
                # Based on the memory, address overlaps are warnings, but duplicate tags are fatal.
                pass

            return is_valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error during validation: {e}")
            return False

    @staticmethod
    def _calculate_coefficients(factor_str: Any, offset_str: Any, scale_factor_str: Any) -> Tuple[str, str]:
        factor = Generator._parse_numeric(factor_str, default=1.0)
        offset = Generator._parse_numeric(offset_str, default=0.0)
        try: scale_val = int(float(scale_factor_str)) if scale_factor_str else 0
        except ValueError: scale_val = 0
        coef_a = "{:.6f}".format(factor * (10 ** scale_val))
        coef_b = "{:.6f}".format(offset)
        return coef_a, coef_b

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"Validation target not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'rb') as f:
                header = f.read(4)
                encoding = 'utf-16' if header.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, mode='r', encoding=encoding) as f:
                reader = csv.reader(f, delimiter=';')
                # Skip header
                try:
                    next(reader)
                except StopIteration:
                    return False

                success = True
                for line_num, row in enumerate(reader, start=2):
                    if not row or len(row) < 11:
                        if any(row):
                            logging.warning(f"Line {line_num}: Row has insufficient columns (expected 11).")
                        continue

                    # row index [0]=Index, [1]=Info1, [2]=Info2, [3]=Info3, [4]=Info4, [5]=Name, [6]=Tag, [10]=Action
                    info1, address, dtype, tag, name = row[1], row[2], row[3], row[6], row[5]

                    # Validate Address and Range
                    if not self.validate_address(address, dtype):
                        valid = False
                        logging.error(f"Line {line_num}: Invalid address or out of range: {address}")

                    # Validate Tag Uniqueness
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (previously line {seen_tags[tag]})")
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Check Overlaps
                    self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            return valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error during CSV validation: {e}")
            return False

        return valid

    def process_rows(self, rows: Iterable[Dict[str, Any]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        seen_names, seen_tags, address_usage, warned_lines = {}, {}, {}, set()
        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v): continue
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

            # Action normalization with intelligent defaulting
            act_str = str(action).strip().upper()
            if not act_str:
                # Intelligent action defaulting
                norm_action = '4' if info1 in ['2', '4'] else '1'
            elif act_str in ['R', 'READ', 'RO', 'READ-ONLY', 'READ ONLY', '4']:
                norm_action = '4'
            elif act_str in ['RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', 'R/W', 'WO', 'WRITE-ONLY', 'WRITE ONLY', '1']:
                norm_action = '1'
            elif act_str in self.allowed_actions:
                norm_action = act_str
            else:
                norm_action = '4' if info1 in ['2', '4'] else '1'

            yield {'Info1': info1, 'Info2': address, 'Info3': dtype.upper(), 'Info4': '', 'Name': name, 'Tag': tag, 'CoefA': coef_a, 'CoefB': coef_b, 'Unit': unit, 'Action': norm_action}

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition file for common errors."""
        if not os.path.exists(filepath):
            logging.error(f"File not found for validation: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                # Skip header
                try:
                    next(reader)
                except StopIteration:
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns ({len(row)}). Skipping.")
                        continue

                    # row format: index, info1, info2, info3, info4, name, tag, coefa, coefb, unit, action
                    info1, info2, info3, name, tag = row[1], row[2], row[3], row[5], row[6]

                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Fatal error - Duplicate Tag '{tag}' (previous at line {seen_tags[tag]}).")
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    if not self.validate_address(info2, info3):
                        valid = False

                    self._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)

        except (OSError, csv.Error) as e:
            logging.error(f"Error during validation: {e}")
            return False

        return valid

    @staticmethod
    def validate_csv(filepath: str) -> bool:
        """
        Validates an existing Webdyn definition file.
        Checks for duplicates, overlaps, and basic formatting.
        """
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        is_valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()
        gen = Generator()

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader)
                if len(header) < 4:
                    logging.error("Invalid definition file header.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or len(row) < 11:
                        if any(row): logging.warning(f"Line {line_num}: Row has insufficient columns.")
                        continue

                    idx, info1, info2, info3, info4, name, tag, coef_a, coef_b, unit, action = row[:11]

                    if not gen.validate_address(info2, info3):
                        is_valid = False

                    if tag in seen_tags:
                        logging.error(f"Line {line_num}: Fatal - Duplicate Tag '{tag}' (Previously at line {seen_tags[tag]}).")
                        is_valid = False
                    seen_tags[tag] = line_num

                    gen._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)

            return is_valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV {filepath}: {e}")
            return False

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition file (CSV)."""
        if not os.path.exists(filepath):
            logging.error(f"File not found for validation: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header or len(header) < 4:
                    logging.error(f"Invalid Webdyn header in {filepath}")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row): continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns. Skipping.")
                        continue

                    # Index (0), Info1 (1), Info2 (2), Info3 (3), Info4 (4), Name (5), Tag (6), CoefA (7), CoefB (8), Unit (9), Action (10)
                    info1, addr, dtype, tag, name = row[1], row[2], row[3], row[6], row[5]

                    # 1. Tag uniqueness (Fatal)
                    if tag in seen_tags:
                        logging.error(f"Line {line_num}: Fatal - Duplicate Tag '{tag}'. Already seen at line {seen_tags[tag]}.")
                        valid = False
                    seen_tags[tag] = line_num

                    # 2. Address Range (0-65535)
                    norm_addr = self.normalize_address_val(addr.split('_')[0])
                    try:
                        addr_int = int(norm_addr, 0)
                        if addr_int < 0 or addr_int > 65535:
                            logging.warning(f"Line {line_num}: Address {addr_int} is out of standard Modbus range (0-65535).")
                    except ValueError:
                        logging.error(f"Line {line_num}: Invalid address format '{addr}'.")
                        valid = False

                    # 3. Address Overlap
                    self._check_address_overlap(info1, addr, dtype, name, line_num, address_usage, warned_lines)

            # Overlap warnings in _check_address_overlap do not necessarily invalidate,
            # but duplicate tags do.
            return valid

        except (OSError, csv.Error) as e:
            logging.error(f"Error during validation: {e}")
            return False

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        is_valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            # Handle potential encoding issues
            with open(filepath, 'rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, 'r', encoding=encoding) as f:
                reader = csv.reader(f, delimiter=';')
                try:
                    header = next(reader)
                except StopIteration:
                    logging.error("Definition file is empty.")
                    return False

                if len(header) < 4:
                    logging.error("Invalid Webdyn definition header.")
                    return False

                logging.info(f"Validating definition: {header[2]} - {header[3]}")

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row): continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns (found {len(row)}, expected 11). Skipping.")
                        continue

                    # index, info1, info2, info3, info4, name, tag, coefA, coefB, unit, action
                    info1, info2, info3, name, tag = row[1], row[2], row[3], row[5], row[6]

                    # Tag uniqueness (FATAL for Webdyn)
                    if tag in seen_tags:
                        logging.error(f"Line {line_num}: Duplicate Tag '{tag}' (already used at line {seen_tags[tag]}).")
                        is_valid = False
                    else:
                        seen_tags[tag] = line_num

                    # Type validation
                    if not self.validate_type(info3):
                        logging.warning(f"Line {line_num}: Invalid Type '{info3}'.")

                    # Address validation
                    if not self.validate_address(info2, info3):
                        logging.error(f"Line {line_num}: Invalid Address format '{info2}' for type '{info3}'.")
                        is_valid = False

                    # Address overlap check
                    self._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)

        except (OSError, csv.Error) as e:
            logging.error(f"Error reading {filepath}: {e}")
            return False

        return is_valid

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                # Skip header row
                try:
                    next(reader)
                except StopIteration:
                    logging.error(f"{filepath} is empty.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row: continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns (expected 11, found {len(row)}).")
                        continue

                    info1, address, dtype, name, tag = row[1], row[2], row[3], row[5], row[6]

                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Fatal - Duplicate Tag '{tag}' detected.")
                            valid = False
                        seen_tags[tag] = line_num

                    if not self.validate_type(dtype):
                        logging.warning(f"Line {line_num}: Invalid Type '{dtype}'.")
                        valid = False

                    if not self.validate_address(address, dtype):
                        logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'.")
                        valid = False

                    self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            return valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV {filepath}: {e}")
            return False

    def validate_csv(self, filepath: str) -> bool:
        """Deep validation of an existing WebdynSunPM definition file."""
        self.valid = True
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header or len(header) < 4:
                    logging.error(f"Invalid Webdyn definition header in {filepath}")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(row):
                        continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Row has insufficient columns ({len(row)}/11). Skipping.")
                        continue

                    # Index; Info1; Info2; Info3; Info4; Name; Tag; CoefA; CoefB; Unit; Action
                    info1 = row[1].strip()
                    address = row[2].strip()
                    dtype = row[3].strip()
                    name = row[5].strip()
                    tag = row[6].strip()

                    # Tag validation
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Duplicate Tag '{tag}' (Fatal error). Previous at line {seen_tags[tag]}.")
                            self.valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Address validation
                    norm_addr = Generator.normalize_address_val(address.split('_')[0])
                    try:
                        addr_val = int(norm_addr)
                        if addr_val < 0 or addr_val > 65535:
                            logging.warning(f"Line {line_num}: Address {addr_val} is out of standard Modbus range (0-65535).")
                    except ValueError:
                        pass

                    if not self.validate_address(address, dtype):
                        logging.warning(f"Line {line_num}: Invalid address format '{address}' for type '{dtype}'.")
                        self.valid = False

                    self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            return self.valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV: {e}")
            return False

    @staticmethod
    def write_output_csv(output: Union[str, Any, None], processed_rows: Iterable[Dict[str, Any]], manufacturer: str, model: str,
                        protocol: str = 'modbusRTU', category: str = 'Inverter', forced_write: str = '') -> None:
        """Centralized method to write the WebdynSunPM CSV format."""
        type_summary = {'1': 0, '2': 0, '3': 0, '4': 0}
        type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}
        try:
            if isinstance(output, str):
                outfile = open(output, 'w', newline='', encoding='utf-8-sig')
            elif output is None:
                outfile = sys.stdout
            else:
                outfile = output

            header_row = [
                Generator.sanitize_csv_field(protocol),
                Generator.sanitize_csv_field(category),
                Generator.sanitize_csv_field(manufacturer),
                Generator.sanitize_csv_field(model),
                Generator.sanitize_csv_field(forced_write),
                '', '', '', '', '', ''
            ]
            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            stats = {'1': 0, '2': 0, '3': 0, '4': 0}
            type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}

            for index, row in enumerate(processed_rows, start=1):
                type_summary[row['Info1']] = type_summary.get(row['Info1'], 0) + 1
                writer.writerow([
                    str(index), i1, row['Info2'], row['Info3'], row['Info4'],
                    row['Name'], row['Tag'], row['CoefA'], row['CoefB'], row['Unit'], row['Action']
                ])
                type_counts[row['Info1']] = type_counts.get(row['Info1'], 0) + 1

            summary = ", ".join([f"{type_labels[k]}: {v}" for k, v in counts.items() if v > 0])
            if summary:
                logging.info(f"Generated {index} registers ({summary})")

            summary_str = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in type_summary.items() if v > 0])
            if summary_str:
                logging.info(f"Register Summary: {summary_str}")

            summary = ", ".join([f"{type_labels[k]}: {v}" for k, v in type_counts.items() if v > 0])
            if summary:
                logging.info(f"Registers summary: {summary}")

            summary = f"Generated {total} registers: Coils={counts['1']}, Discrete={counts['2']}, Holding={counts['3']}, Input={counts['4']}"
            if isinstance(output, str):
                summary = ", ".join([f"{type_labels[k]}: {v}" for k, v in type_counts.items() if v > 0])
                logging.info(f"Definition file generated at {output}. Summary: {summary}")
        except (OSError, csv.Error) as e:
            logging.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and 'outfile' in locals() and not outfile.closed: outfile.close()

    def validate_csv(self, filepath: str) -> bool:
        valid = True
        seen_tags, address_usage, warned_lines = {}, {}, set()
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header or len(header) < 4:
                    logging.error(f"Invalid Webdyn definition header in {filepath}")
                    return False
                for line_num, row in enumerate(reader, start=2):
                    if not row: continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns (found {len(row)}, expected 11). Skipping.")
                        continue
                    info1, info2, info3, name, tag = row[1], row[2], row[3], row[5], row[6]
                    if not self.validate_address(info2, info3):
                        logging.warning(f"Line {line_num}: Invalid address '{info2}' for type '{info3}'.")
                        valid = False
                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: FATAL - Duplicate Tag '{tag}' (previously at line {seen_tags[tag]}).")
                            valid = False
                        else: seen_tags[tag] = line_num
                    self._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)
            if valid and not warned_lines: logging.info(f"Validation successful for {filepath}")
            elif valid: logging.info(f"Validation finished with warnings for {filepath}")
            else: logging.error(f"Validation failed for {filepath}")
            return valid
        except Exception as e:
            logging.error(f"Error validating {filepath}: {e}")
            return False

def generate_template(output_file: Optional[str], mode: str = 'input') -> None:
    if mode == 'definition':
        headers = ['#Index', 'Info1', 'Info2', 'Info3', 'Info4', 'Name', 'Tag', 'CoefA', 'CoefB', 'Unit', 'Action']
        rows = [
            ['modbusRTU', 'Inverter', 'SampleManufacturer', 'SampleModel', '', '', '', '', '', '', ''],
            ['1', '3', '40001', 'U16', '', 'Active Power', 'active_power', '1.000000', '0.000000', 'W', '4'],
            ['2', '3', '40002', 'U16', '', 'Voltage', 'voltage', '0.100000', '0.000000', 'V', '4']
        ]
    else:
        headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]

    def validate_csv(self, filepath: str) -> bool:
        """Validates an existing Webdyn definition file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        is_valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader, None)
                if not header:
                    logging.error("Empty definition file.")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns (expected at least 11, got {len(row)}). Skipping.")
                        continue

                    # row format: [Index, Info1, Info2, Info3, Info4, Name, Tag, CoefA, CoefB, Unit, Action]
                    _, info1, info2, info3, _, name, tag, _, _, _, _ = row[:11]

                    if tag:
                        if tag in seen_tags:
                            logging.error(f"Line {line_num}: Duplicate Tag '{tag}' (Fatal). Already seen at line {seen_tags[tag]}.")
                            is_valid = False
                        else:
                            seen_tags[tag] = line_num

                    if not self.validate_address(info2, info3):
                        is_valid = False

                    self._check_address_overlap(info1, info2, info3, name, line_num, address_usage, warned_lines)

            return is_valid
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV: {e}")
            return False

def generate_template(output_file: Optional[str]) -> None:
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    rows = [['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'], ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']]
    try:
        if output_file:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=delimiter)
                if mode != 'definition':
                    writer.writerow(headers)
                writer.writerows(rows)
        else:
            writer = csv.writer(sys.stdout, delimiter=delimiter)
            if mode != 'definition':
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

    mfg = config.manufacturer or 'Manufacturer'
    model = config.model or 'Model'

    if input_data is None:
        if not config.input_file:
            logging.error("input_file or input_data is required.")
            return
        if not os.path.exists(config.input_file):
            logging.error(f"Input file not found: {config.input_file}")
            return

    manufacturer = config.manufacturer or 'Manufacturer'
    model = config.model or 'Model'

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
                snippet = csvfile.read(2048); csvfile.seek(0)
                try: dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                except csv.Error: dialect = csv.excel
                reader = csv.DictReader(csvfile, dialect=dialect)
                processed_rows = generator.process_rows(reader, config.address_offset)

                generator.write_output_csv(config.output, processed_rows, config.manufacturer, config.model,
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

if __name__ == "__main__": main()
