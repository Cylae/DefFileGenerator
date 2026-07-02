#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
from typing import Dict, List, Optional, Any, Union, Tuple, Set, Iterator, Iterable
import os
from dataclasses import dataclass

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
    template_mode: str = 'input'  # 'input' or 'definition'
    address_offset: int = 0
    template_mode: str = 'input'

class Generator:
    def __init__(self) -> None:
        # RegisterType mapping to Info1
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
        # Allowed Action codes
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    def validate_csv(self, filepath: str) -> bool:
        """Validates a WebdynSunPM definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        try:
            with open(filepath, mode='rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, mode='r', encoding=encoding) as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                try:
                    header = next(reader)
                except StopIteration:
                    logging.error(f"File {filepath} is empty.")
                    return False

                if len(header) < 4:
                    logging.error(f"Invalid header in {filepath}. Expected at least 4 columns (Protocol, Category, Mfg, Model).")
                    return False

                success = True
                for line_num, row in enumerate(reader, start=2):
                    if not row or not any(str(c).strip() for c in row):
                        continue
                    if len(row) < 11:
                        logging.error(f"Line {line_num}: Invalid row length. Expected 11 columns, got {len(row)}.")
                        success = False
                        continue

                    # row[1] is Info1 (RegisterType), row[2] is Info2 (Address), row[3] is Info3 (Type)
                    info1, address, dtype = row[1], row[2], row[3]

                    if not self.validate_type(dtype):
                        logging.error(f"Line {line_num}: Invalid Type '{dtype}'.")
                        success = False

                    # For validation, we use the same logic as generation
                    if not self.validate_address(address, dtype):
                        # validate_address already logs warnings if range is invalid
                        success = False

                return success
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV {filepath}: {e}")
            return False

    @staticmethod
    def normalize_type(dtype):
        """Standardizes common type synonyms while preserving suffixes."""
        if not dtype:
            return 'U16'
        t = str(dtype).lower().strip()

        # Endianness suffix detection
        suffix = ''
        if any(x in t for x in ['_wb', 'swap', 'big endian']):
            suffix = '_WB'
        elif any(x in t for x in ['_b', 'big']):
            suffix = '_B'
        elif any(x in t for x in ['_w', 'word']):
            suffix = '_W'

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

        t = _CLEAN_TYPE_RE.sub('', t)
        return t.upper() if t else 'U16'

    @staticmethod
    def validate_type(dtype: str) -> bool:
        """Validates the data type."""
        if not dtype:
            return False
        dtype_upper = dtype.upper()
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
    def normalize_address_val(addr_part):
        """Converts a single address part (possibly hex) to decimal string."""
        addr_part = str(addr_part).strip()
        # Remove thousands separators if they exist
        addr_part = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_part)

        if not addr_part:
            return ""

        # Support hex with 0x prefix or h suffix
        if addr_part.lower().startswith('0x'):
            try:
                return str(int(addr_part, 16))
            except ValueError:
                return addr_part
        elif addr_part.lower().endswith('h'):
            try:
                return str(int(addr_part[:-1], 16))
            except ValueError:
                return addr_part

        # Handle decimal (including negative)
        try:
            return str(int(addr_part))
        except ValueError:
            pass

        # If it's a raw hex word (e.g. "A0")
        if re.match(r'^[0-9A-Fa-f]+$', addr_part):
            try:
                return str(int(addr_part, 16))
            except ValueError:
                return addr_part

        return addr_part

    @staticmethod
    def validate_address(address: str, dtype: str) -> bool:
        """Validates the address format based on type and ensures it is within Modbus range."""
        dtype_upper = dtype.upper()
        is_string = (dtype_upper == 'STRING' or RE_TYPE_STR_CONV.match(dtype_upper))

        # Check basic format
        if dtype_upper == 'STRING' or RE_TYPE_STR_CONV.match(dtype_upper):
            if not RE_ADDR_STRING.match(address):
                return False
        elif dtype_upper == 'BITS':
            is_valid = RE_ADDR_BITS.match(address) is not None
        else:
            is_valid = RE_ADDR_INT.match(address) is not None

        # Check address range (0-65535)
        try:
            base_addr_str = address.split('_')[0]
            norm_addr = Generator.normalize_address_val(base_addr_str)
            addr_val = int(norm_addr, 0)
            if addr_val < 0 or addr_val > 65535:
                logging.warning(f"Address {addr_val} is out of standard Modbus range (0-65535).")
                return False
        return is_valid

    @staticmethod
    def get_register_count(dtype: str, address: str) -> int:
        """Calculates the number of registers used by the type."""
        dtype_upper = dtype.upper()

        if RE_COUNT_16_8.match(dtype_upper):
            return 1
        elif RE_COUNT_32.match(dtype_upper):
            return 2
        elif RE_COUNT_64.match(dtype_upper):
            return 4
        elif dtype_upper == 'MAC':
            return 3
        elif dtype_upper == 'IPV6':
            return 8
        elif dtype_upper == 'STRING':
            try:
                parts = address.split('_')
                length = int(parts[1])
                return math.ceil(length / 2)
            except (IndexError, ValueError):
                return 0
        return 1

    @staticmethod
    def _parse_numeric(val: Any, default: float = 0.0) -> float:
        """Robust numeric parsing for scale factors and offsets."""
        if val is None or str(val).strip() == '':
            return default
        s = str(val).strip()

        # Handle fractions
        if '/' in s:
            try:
                parts = s.split('/')
                return float(parts[0]) / float(parts[1])
            except (ValueError, ZeroDivisionError, IndexError):
                return default

        # Heuristic for thousands vs decimal separators
        if ',' in s and '.' in s:
            if s.find(',') < s.find('.'): # 1,234.56
                s = s.replace(',', '')
            else: # 1.234,56
                s = s.replace('.', '').replace(',', '.')
        elif ',' in s:
            # Only comma. Match ^\d{1,3}(,\d{3})+$ as thousands
            if re.match(r'^-?\d{1,3}(,\d{3})+$', s):
                s = s.replace(',', '')
            else:
                s = s.replace(',', '.')

        try:
            return float(s)
        except ValueError:
            return default

    @staticmethod
    def apply_address_offset(address: Any, offset: int, line_num: Optional[int] = None, name: Optional[str] = None) -> str:
        """Applies an integer offset to a register address (simple or compound)."""
        if not address:
            return ""
        # Split by underscore but ensure we handle compound addresses correctly
        parts = str(address).split('_')
        # Normalize each part individually
        norm_parts = [Generator.normalize_address_val(p) for p in parts]

        try:
            base_addr = int(norm_parts[0]) + offset
            if base_addr < 0:
                msg = f"Address offset {offset} results in negative address {base_addr}"
                if name: msg += f" for '{name}'"
                if line_num: logging.warning(f"Line {line_num}: {msg}")
                else: logging.warning(msg)
            norm_parts[0] = str(base_addr)
        except (ValueError, IndexError):
            pass

        return '_'.join(norm_parts)

    def _process_name_and_tag(self, name: str, tag: str, line_num: int, seen_names: Dict[str, int], seen_tags: Dict[str, int]) -> str:
        """Validates name and ensures unique tag generation."""
        if name:
            if name in seen_names:
                logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
            else:
                seen_names[name] = line_num

        if not tag and name:
            # Generate valid tag: alphanumeric and underscores only, start with v_ if needed
            base_tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
            base_tag = re.sub(r'_+', '_', base_tag).strip('_')

            if not base_tag or not base_tag[0].isalpha():
                base_tag = f"v_{base_tag}" if base_tag else "var"

            tag = base_tag
            counter = 1
            while tag in seen_tags:
                tag = f"{base_tag}_{counter}"
                counter += 1

        if tag:
            if tag in seen_tags:
                logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous occurrence at line {seen_tags[tag]}.")
            else:
                seen_tags[tag] = line_num

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

    def _check_address_overlap(self, info1: str, address: str, dtype: str, name: str, line_num: int, address_usage: Dict[str, Dict[int, List[Tuple[int, str, str]]]], warned_lines: Set[Tuple[int, int]]) -> None:
        """Checks for address overlaps using O(N) dictionary lookup."""
        try:
            start_addr = int(address.split('_')[0])
            reg_count = self.get_register_count(dtype, address)

            if info1 not in address_usage:
                address_usage[info1] = {}

            is_bits = (dtype.upper() == 'BITS')
            for i in range(reg_count):
                curr_addr = start_addr + i
                if curr_addr in address_usage[info1]:
                    for u_line, u_name, u_type in address_usage[info1][curr_addr]:
                        # Allow multiple BITS on exactly the same base address
                        if is_bits and u_type == 'BITS' and curr_addr == start_addr:
                            continue

                        warn_key = tuple(sorted((line_num, u_line)))
                        if warn_key not in warned_lines:
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {curr_addr}. Overlaps with '{u_name}' (Line {u_line}).")
                            warned_lines.add(warn_key)
                else:
                    address_usage[info1][curr_addr] = []
                address_usage[info1][curr_addr].append((line_num, name, dtype.upper()))
        except (ValueError, IndexError):
            pass

    def validate_csv(self, filepath: str) -> bool:
        """Validates a WebdynSunPM definition file for correct formatting and ranges."""
        if not os.path.exists(filepath):
            logging.error(f"File not found: {filepath}")
            return False

        success = True
        seen_names = {}
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
                    if not row or not any(row):
                        continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Row too short ({len(row)} columns, expected 11).")
                        success = False
                        continue

                    # row format: index, info1, info2, info3, info4, name, tag, coef_a, coef_b, unit, action
                    _, info1, info2, info3, _, name, tag, _, _, _, action = row[:11]

                    if info1 not in ['1', '2', '3', '4']:
                        logging.warning(f"Line {line_num}: Invalid Info1 (RegisterType) '{info1}'.")
                        success = False

                    dtype = info3.upper()
                    if not self.validate_type(dtype):
                        logging.warning(f"Line {line_num}: Invalid Info3 (Type) '{info3}'.")
                        success = False

                    if not self.validate_address(info2, dtype):
                        success = False

                    self._check_address_overlap(info1, info2, dtype, name, line_num, address_usage, warned_lines)

                    if name in seen_names:
                        logging.warning(f"Line {line_num}: Duplicate Name '{name}'. Previously at line {seen_names[name]}.")
                    seen_names[name] = line_num

                    if tag in seen_tags:
                        logging.warning(f"Line {line_num}: Duplicate Tag '{tag}'. Previously at line {seen_tags[tag]}.")
                    seen_tags[tag] = line_num

                    if action not in self.allowed_actions:
                        logging.warning(f"Line {line_num}: Invalid Action '{action}'.")
                        success = False

            if warned_lines:
                success = False

        except Exception as e:
            logging.error(f"Error validating CSV: {e}")
            return False

        return success

    @staticmethod
    def sanitize_csv_field(field: Any) -> str:
        """Sanitizes a field to prevent CSV Formula Injection."""
        if field is None:
            return ""
        s = str(field)
        if s and s[0] in ('=', '+', '-', '@', '\t', '\r'):
            return "'" + s
        return s

    @staticmethod
    def _calculate_coefficients(factor_str: Any, offset_str: Any, scale_factor_str: Any) -> Tuple[str, str]:
        """Calculates CoefA and CoefB based on input values."""
        factor = Generator._parse_numeric(factor_str, default=1.0)
        offset = Generator._parse_numeric(offset_str, default=0.0)

        try:
            scale_val = int(float(scale_factor_str)) if scale_factor_str else 0
        except ValueError:
            scale_val = 0

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

    def process_rows(self, rows: Iterable[Dict[str, Any]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        """Processes simplified CSV rows into WebdynSunPM format."""
        seen_names = {}
        seen_tags = {}
        address_usage = {} # Info1 -> dict of address -> list of (line, name, type)
        warned_lines = set()

        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v):
                continue

            # Normalize row once
            norm_row = {k.lower().strip(): (str(v).strip() if v is not None else '') for k, v in row.items()}

            name = norm_row.get('name', '')
            tag = norm_row.get('tag', '')
            reg_type_str = norm_row.get('registertype', '')
            address = norm_row.get('address', '')
            dtype_raw = norm_row.get('type', '')
            factor = norm_row.get('factor', '')
            offset = norm_row.get('offset', '')
            unit = norm_row.get('unit', '')
            action = norm_row.get('action', '')
            scale_factor_str = norm_row.get('scalefactor', '')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype_raw}' (normalized to '{dtype}'). Skipping row.")
                continue

            # STR<n> conversion
            match_str = RE_TYPE_STR_CONV.match(dtype)
            if match_str:
                length = int(match_str.group(1))
                dtype = 'STRING'
                if '_' not in address:
                    address = f"{address}_{length}"

            # Apply address offset and normalize
            address = Generator.apply_address_offset(address, address_offset, line_num, name)

            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            tag = self._process_name_and_tag(name, tag, line_num, seen_names, seen_tags)
            info1 = self._determine_info1(reg_type_str, line_num)

            self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            coef_a, coef_b = self._calculate_coefficients(factor, offset, scale_factor_str)

            # Action normalization and intelligent defaulting
            act_str = str(action).strip().upper()
            if not act_str:
                # Default based on Info1: Input (4) and Discrete (2) are RO
                norm_action = '4' if info1 in ['2', '4'] else '1'
            elif act_str in ['R', 'READ', 'RO', 'READ-ONLY', 'READ ONLY', '4']:
                norm_action = '4'
            elif act_str in ['RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', 'R/W', 'WO', 'WRITE-ONLY', 'WRITE ONLY', '1']:
                norm_action = '1'
            elif act_str in self.allowed_actions:
                norm_action = act_str
            else:
                norm_action = '4' if info1 in ['2', '4'] else '1'

            yield {
                'Info1': info1, 'Info2': address, 'Info3': dtype.upper(), 'Info4': '',
                'Name': name, 'Tag': tag, 'CoefA': coef_a, 'CoefB': coef_b,
                'Unit': unit, 'Action': norm_action
            }

    @staticmethod
    def write_output_csv(output: Union[str, Any, None], processed_rows: Iterable[Dict[str, Any]], manufacturer: str, model: str,
                        protocol: str = 'modbusRTU', category: str = 'Inverter', forced_write: str = '') -> None:
        """Centralized method to write the WebdynSunPM CSV format."""
        type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}
        counts = {'1': 0, '2': 0, '3': 0, '4': 0}
        try:
            if isinstance(output, str):
                outfile = open(output, 'w', newline='', encoding='utf-8')
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

            counts = {'1': 0, '2': 0, '3': 0, '4': 0}
            type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}

            for index, row in enumerate(processed_rows, start=1):
                info1 = row['Info1']
                counts[info1] = counts.get(info1, 0) + 1
                writer.writerow([
                    str(index), info1, row['Info2'], row['Info3'], row['Info4'],
                    row['Name'], row['Tag'], row['CoefA'], row['CoefB'], row['Unit'], row['Action']
                ])
                counts[row['Info1']] = counts.get(row['Info1'], 0) + 1

            summary = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in counts.items() if v > 0])
            if summary:
                logging.info(f"Register Summary -> {summary}")

            summary = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in counts.items() if v > 0])
            if summary:
                logging.info(f"Generated {index} registers ({summary})")

            summary_str = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in type_summary.items() if v > 0])
            if isinstance(output, str):
                summary = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in counts.items() if v > 0])
                logging.info(f"Definition file generated at {output}. Summary: {summary}")
        except (OSError, csv.Error) as e:
            logging.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and 'outfile' in locals() and not outfile.closed:
                outfile.close()

    @staticmethod
    def validate_csv(filepath: str) -> bool:
        """Validates a Webdyn definition CSV file."""
        if not os.path.exists(filepath):
            logging.error(f"Validation failed: File {filepath} not found.")
            return False

        success = True
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f, delimiter=';')
                header = next(reader)
                if len(header) < 4:
                    logging.error("Validation failed: Invalid header format (insufficient columns).")
                    return False

                for line_num, row in enumerate(reader, start=2):
                    if not row or all(not cell.strip() for cell in row):
                        continue
                    if len(row) < 11:
                        logging.warning(f"Line {line_num}: Insufficient columns (expected 11, got {len(row)}). skipping.")
                        continue

                    # Column mapping: 1:Info1 (Type), 2:Info2 (Address), 3:Info3 (DataType)
                    info1, address, info3 = row[1], row[2], row[3]
                    if not Generator.validate_address(address, info3):
                        logging.warning(f"Line {line_num}: Invalid address '{address}' for type '{info3}'.")
                        success = False

            return success
        except (OSError, csv.Error) as e:
            logging.error(f"Error validating CSV: {e}")
            return False

def generate_template(output_file: Optional[str], mode: str = 'input') -> None:
    if mode == 'definition':
        # Sample Webdyn Definition CSV (semicolon delimited)
        header = ['modbusRTU', 'Inverter', 'Manufacturer', 'Model', '', '', '', '', '', '', '']
        rows = [
            ['1', '3', '30001', 'U16', '', 'Example Variable', 'example_tag', '1.000000', '0.000000', 'V', '4'],
            ['2', '3', '30030_20', 'STRING', '', 'Convenience String', 'str_tag', '1.000000', '0.000000', '', '4']
        ]
        delimiter = ';'
    else:
        # Sample Simplified Input CSV (comma delimited)
        header = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]
        delimiter = ','

    try:
        if output_file:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=delimiter)
                writer.writerow(header)
                writer.writerows(rows)
        else:
            writer = csv.writer(sys.stdout, delimiter=delimiter)
            writer.writerow(header)
            writer.writerows(rows)
    except OSError as e:
        logging.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig, input_data: Optional[Iterable[Dict[str, Any]]] = None) -> None:
    if config.template:
        generate_template(config.output, config.template_mode)
        return

    if input_data is None:
        if not config.input_file:
            logging.error("input_file or input_data is required.")
            return
        if not os.path.exists(config.input_file):
            logging.error(f"Input file not found: {config.input_file}")
            return

    if not config.manufacturer or not config.model:
        logging.error("manufacturer and model are required.")
        return

    generator = Generator()
    try:
        if input_data is not None:
            processed_rows = generator.process_rows(input_data, config.address_offset)
            generator.write_output_csv(config.output, processed_rows, config.manufacturer, config.model,
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

                # Consume the generator while the input file is still open
                generator.write_output_csv(config.output, processed_rows, config.manufacturer, config.model,
                                           config.protocol, config.category, config.forced_write)
    except (OSError, csv.Error, ValueError) as e:
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
    parser.add_argument('--address-offset', type=int, default=0)

    args = parser.parse_args()
    config = GeneratorConfig(
        input_file=args.input_file, output=args.output,
        manufacturer=args.manufacturer, model=args.model,
        protocol=args.protocol, category=args.category,
        forced_write=args.forced_write, template=args.template,
        address_offset=args.address_offset
    )
    run_generator(config)

if __name__ == "__main__":
    main()
