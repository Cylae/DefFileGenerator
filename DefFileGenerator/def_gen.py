#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
import itertools
import os
import bisect
from typing import Dict, List, Optional, Any, Union, Tuple, Set, Iterator, Iterable
from dataclasses import dataclass

# Setup logger for this module
logger = logging.getLogger('DefFileGenerator.def_gen')

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
    template_mode: str = 'input'
    address_offset: int = 0
    strict_overlap: bool = False

class Generator:
    def __init__(self) -> None:
        self.register_type_map = {
            'coil': '1',
            'coils': '1',
            'discrete input': '2',
            'discrete': '2',
            'discrete registers': '2',
            'discrete register': '2',
            'holding register': '3',
            'holding': '3',
            'input register': '4',
            'input': '4'
        }
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

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

        if t.startswith('str') and t[3:].isdigit():
            return t.upper()

        t = _CLEAN_TYPE_RE.sub('', t)
        return t.upper() if t else 'U16'

    @staticmethod
    def validate_type(dtype: str) -> bool:
        """Validates the data type."""
        dtype_upper = str(dtype).upper()
        base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC']
        if dtype_upper in base_types:
            return True
        if RE_TYPE_NUMERIC.match(dtype_upper):
            return True
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
            try: return str(int(addr_part[:-1], 16))
            except ValueError: return addr_part
        try: return str(int(addr_part, 0))
        except ValueError: pass
        if re.match(r'^[0-9A-Fa-f]+$', addr_part):
            try: return str(int(addr_part, 16))
            except ValueError: return addr_part
        return addr_part

    @staticmethod
    def validate_address(address: str, dtype: str) -> bool:
        """Validates the address format based on type and Modbus range (0-65535)."""
        dtype_upper = dtype.upper()

        is_str_synonym = RE_TYPE_STR_CONV.match(dtype_upper) is not None
        if is_str_synonym:
            dtype_upper = 'STRING'

        is_valid_format = False
        if dtype_upper == 'STRING':
            is_valid_format = RE_ADDR_STRING.match(address) is not None
            if not is_valid_format and is_str_synonym:
                 is_valid_format = RE_ADDR_INT.match(address) is not None
        elif dtype_upper == 'BITS':
            is_valid_format = RE_ADDR_BITS.match(address) is not None
        else:
            is_valid_format = RE_ADDR_INT.match(address) is not None

        if not is_valid_format:
            return False

        try:
            parts = address.split('_')
            base_addr_str = Generator.normalize_address_val(parts[0])
            base_addr = int(base_addr_str, 0)
            if not (0 <= base_addr <= 65535):
                logger.warning(f"Address {base_addr} is out of standard Modbus range (0-65535).")
                return False
        except (ValueError, IndexError):
            return False
        return True

    def validate_csv(self, filepath: str, strict_overlap: bool = True) -> bool:
        """Deep validation of an existing WebdynSunPM definition file or input file."""
        if not os.path.exists(filepath):
            logger.error(f"File not found: {filepath}")
            return False

        is_valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, 'r', encoding=encoding) as f:
                snippet = f.read(4096)
                f.seek(0)
                is_webdyn = ';' in snippet and any(p in snippet for p in ['modbusRTU', 'modbusTCP'])

                if is_webdyn:
                    reader = csv.reader(f, delimiter=';')
                    header = next(reader, None)
                    if not header or len(header) < 4:
                        # For very short files, we might still want to return True if it's not a fatal error
                        # but just a minimalist definition. However, Webdyn needs at least protocol/cat/mfg/model.
                        if header and len(header) > 0:
                            logger.warning(f"Definition file {filepath} has a minimalist header.")
                            return True
                        logger.error(f"Invalid Webdyn definition header in {filepath}")
                        return False

                    for line_num, row in enumerate(reader, start=2):
                        if not row or not any(row) or row[0].startswith('#'): continue
                        if len(row) < 11:
                            logger.warning(f"Line {line_num}: Row has insufficient columns ({len(row)}/11). skipping.")
                            continue

                        info1, address, dtype, name, tag = row[1], row[2], row[3], row[5], row[6]
                        if tag:
                            if tag in seen_tags:
                                logger.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (Previously at line {seen_tags[tag]}).")
                                is_valid = False
                            else:
                                seen_tags[tag] = line_num

                        if not self.validate_address(address, dtype):
                            is_valid = False

                        overlap = self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)
                        if overlap and strict_overlap: is_valid = False
                else:
                    # Input file
                    try:
                        dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                    except csv.Error:
                        dialect = csv.excel
                    f.seek(0)
                    reader = csv.DictReader(f, dialect=dialect)
                    if not reader.fieldnames or len(reader.fieldnames) < 4:
                         logger.warning(f"Input file {filepath} has fewer than 4 columns.")
                         return True

                    for line_num, row in enumerate(reader, start=2):
                        name = row.get('Name') or row.get('name')
                        address = row.get('Address') or row.get('address')
                        dtype_raw = row.get('Type') or row.get('type')
                        if not name or not address: continue

                        dtype = self.normalize_type(dtype_raw)
                        if not self.validate_address(address, dtype):
                            is_valid = False
            return is_valid
        except Exception as e:
            logger.error(f"Error during validation: {e}")
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

    @staticmethod
    def apply_address_offset(address: Any, offset: int, line_num: Optional[int] = None, name: Optional[str] = None) -> str:
        """Applies an integer offset to a register address (simple or compound)."""
        if not address: return ""
        parts = str(address).split('_')
        norm_parts = [Generator.normalize_address_val(p) for p in parts]
        try:
            base_addr = int(norm_parts[0]) + offset
            if base_addr < 0:
                msg = f"Address offset {offset} results in negative address {base_addr}"
                if name: msg += f" for '{name}'"
                if line_num: logger.warning(f"Line {line_num}: {msg}")
                else: logger.warning(msg)
            norm_parts[0] = str(base_addr)
        except (ValueError, IndexError): pass
        return '_'.join(norm_parts)

    def _check_address_overlap(self, info1: str, address: str, dtype: str, name: str, line_num: int, address_usage: Dict[str, Dict[str, Any]], warned_lines: Set[Tuple[int, int]]) -> bool:
        """Checks for address overlaps using O(log N) binary search on intervals."""
        overlap_found = False
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

            # Check right and left
            for j in itertools.chain(range(idx, len(intervals)), range(idx - 1, -1, -1)):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if j >= idx and u_start > end_addr: break
                if j < idx and start_addr - u_start > max_len: break

                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if is_bits and u_type == 'BITS' and start_addr == u_start: continue
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        logger.warning(f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line}).")
                        warned_lines.add(warn_key)
                        overlap_found = True

            bisect.insort(intervals, (start_addr, end_addr, line_num, name, dtype.upper()))
            if reg_count > max_len: usage['max_len'] = reg_count
        except (ValueError, IndexError): pass
        return overlap_found

    @staticmethod
    def sanitize_csv_field(field: Any) -> str:
        s = str(field)
        if s and s[0] in ['=', '+', '-', '@', '\t', '\r']:
            return "'" + s
        return s

    def process_rows(self, rows: Iterable[Dict[str, Any]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        seen_names, seen_tags, address_usage, warned_lines = {}, {}, {}, set()
        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v): continue
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
                logger.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logger.warning(f"Line {line_num}: Invalid Type '{dtype_raw}' (normalized to '{dtype}'). Skipping.")
                continue

            match_str = RE_TYPE_STR_CONV.match(dtype)
            if match_str:
                dtype = 'STRING'
                if '_' not in address: address = f"{address}_{match_str.group(1)}"

            address = Generator.apply_address_offset(address, address_offset, line_num, name)
            if not self.validate_address(address, dtype):
                logger.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping.")
                continue

            info1 = self._determine_info1(reg_type_str, line_num)

            # Tag/Name processing
            if name:
                if name in seen_names: logger.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous at line {seen_names[name]}.")
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
                if tag in seen_tags: logger.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous at line {seen_tags[tag]}.")
                else: seen_tags[tag] = line_num

            self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            # Coefficients
            f_val = Generator._parse_numeric(factor, 1.0)
            o_val = Generator._parse_numeric(offset, 0.0)
            try: s_val = int(float(scale_factor_str)) if scale_factor_str else 0
            except ValueError: s_val = 0
            coef_a = "{:.6f}".format(f_val * (10 ** s_val))
            coef_b = "{:.6f}".format(o_val)

            # Action
            act_str = str(action).strip().upper()
            if not act_str: norm_action = '4' if info1 in ['2', '4'] else '1'
            elif act_str in ['R', 'READ', 'RO', 'READ-ONLY', 'READ ONLY', '4']: norm_action = '4'
            elif act_str in ['RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', 'R/W', 'WO', 'WRITE-ONLY', 'WRITE ONLY', '1']: norm_action = '1'
            elif act_str in self.allowed_actions: norm_action = act_str
            else: norm_action = '4' if info1 in ['2', '4'] else '1'

            yield {
                'Info1': info1, 'Info2': address, 'Info3': dtype.upper(), 'Info4': '',
                'Name': name, 'Tag': tag, 'CoefA': coef_a, 'CoefB': coef_b, 'Unit': unit, 'Action': norm_action
            }

    def _determine_info1(self, reg_type_str: str, line_num: int = 0) -> str:
        if reg_type_str is None: return '3'
        lt = str(reg_type_str).lower().strip()
        if not lt: return '3'
        if lt in self.register_type_map: return self.register_type_map[lt]
        if lt in ['1', '2', '3', '4']: return lt
        if line_num: logger.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to 3.")
        return '3'

    @staticmethod
    def write_output_csv(output: Union[str, Any, None], processed_rows: Iterable[Dict[str, Any]], manufacturer: str, model: str,
                        protocol: str = 'modbusRTU', category: str = 'Inverter', forced_write: str = '') -> None:
        type_counts = {'1': 0, '2': 0, '3': 0, '4': 0}
        type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}
        last_index = 0
        try:
            if isinstance(output, str): outfile = open(output, 'w', newline='', encoding='utf-8-sig')
            elif output is None: outfile = sys.stdout
            else: outfile = output

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow([
                Generator.sanitize_csv_field(protocol), Generator.sanitize_csv_field(category),
                Generator.sanitize_csv_field(manufacturer), Generator.sanitize_csv_field(model),
                Generator.sanitize_csv_field(forced_write), '', '', '', '', '', ''
            ])

            for index, row in enumerate(processed_rows, start=1):
                last_index = index
                type_counts[row['Info1']] = type_counts.get(row['Info1'], 0) + 1
                writer.writerow([
                    str(index),
                    Generator.sanitize_csv_field(row['Info1']),
                    Generator.sanitize_csv_field(row['Info2']),
                    Generator.sanitize_csv_field(row['Info3']),
                    Generator.sanitize_csv_field(row['Info4']),
                    Generator.sanitize_csv_field(row['Name']),
                    Generator.sanitize_csv_field(row['Tag']),
                    Generator.sanitize_csv_field(row['CoefA']),
                    Generator.sanitize_csv_field(row['CoefB']),
                    Generator.sanitize_csv_field(row['Unit']),
                    Generator.sanitize_csv_field(row['Action'])
                ])

            summary = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in type_counts.items() if v > 0])
            if summary: logger.info(f"Generated {last_index} registers ({summary})")
        except Exception as e:
            logger.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and 'outfile' in locals() and not outfile.closed: outfile.close()

def generate_template(output_file: Optional[str], mode: str = 'input') -> None:
    if mode == 'definition':
        headers = []
        rows = [
            ['modbusRTU', 'Inverter', 'SampleManufacturer', 'SampleModel', '', '', '', '', '', '', ''],
            ['#Index', 'Info1', 'Info2', 'Info3', 'Info4', 'Name', 'Tag', 'CoefA', 'CoefB', 'Unit', 'Action'],
            ['1', '3', '40001', 'U16', '', 'Active Power', 'active_power', '1.000000', '0.000000', 'W', '4'],
            ['2', '3', '40002', 'U16', '', 'Voltage', 'voltage', '0.100000', '0.000000', 'V', '4']
        ]
    else:
        headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]

    try:
        f = open(output_file, 'w', newline='', encoding='utf-8') if output_file else sys.stdout
        writer = csv.writer(f, delimiter=';' if mode == 'definition' else ',')
        if headers: writer.writerow(headers)
        writer.writerows(rows)
        if output_file: f.close()
    except OSError as e:
        logger.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig, input_data: Optional[Iterable[Dict[str, Any]]] = None) -> None:
    generator = Generator()
    if config.template:
        generate_template(config.output, mode=config.template_mode)
        return

    if input_data is None:
        if not config.input_file or not os.path.exists(config.input_file):
            logger.error(f"Input file not found: {config.input_file}")
            return
        try:
            with open(config.input_file, mode='rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(config.input_file, mode='r', encoding=encoding) as csvfile:
                snippet = csvfile.read(2048); csvfile.seek(0)
                try: dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                except csv.Error: dialect = csv.excel
                reader = csv.DictReader(csvfile, dialect=dialect)
                input_data = list(reader)
        except Exception as e:
            logger.error(f"Error reading input file: {e}")
            return

    processed_rows = generator.process_rows(input_data, config.address_offset)
    generator.write_output_csv(config.output, processed_rows, config.manufacturer or 'Manufacturer',
                               config.model or 'Model', config.protocol, config.category, config.forced_write)

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
    parser.add_argument('--strict-overlap', action='store_true')

    args = parser.parse_args()
    config = GeneratorConfig(
        input_file=args.input_file, output=args.output,
        manufacturer=args.manufacturer, model=args.model,
        protocol=args.protocol, category=args.category,
        forced_write=args.forced_write, template=args.template,
        template_mode=args.template_mode, address_offset=args.address_offset,
        strict_overlap=args.strict_overlap
    )
    run_generator(config)

if __name__ == "__main__": main()
