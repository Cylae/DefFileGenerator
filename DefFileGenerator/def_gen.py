#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
import itertools
import os
from dataclasses import dataclass

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
        return True, itertools.chain([first], it)
    except StopIteration:
        return False, iter([])

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
    input_file: Optional[str] = None
    output: Optional[str] = None
    manufacturer: Optional[str] = None
    model: Optional[str] = None
    protocol: str = 'modbusRTU'
    category: str = 'Inverter'
    forced_write: str = ''
    template: bool = False
    template_mode: str = 'input'  # 'input' or 'definition'
    address_offset: int = 0

class Generator:
    """Core generator class for processing Modbus register maps and building definition CSVs."""

    def __init__(self) -> None:
        self.register_type_map = {
            "coil": "1",
            "coils": "1",
            "discrete input": "2",
            "discrete": "2",
            "holding register": "3",
            "holding": "3",
            "input register": "4",
            "input": "4",
        }
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    @staticmethod
    def normalize_type(dtype: Any) -> str:
        if not dtype: return 'U16'
        t = str(dtype).lower().strip()
        suffix = ''
        if any(x in t for x in ['_wb', 'swap', 'big endian']): suffix = '_WB'
        elif any(x in t for x in ['_b', 'big']): suffix = '_B'
        elif any(x in t for x in ['_w', 'word']): suffix = '_W'

        str_match = _RE_TYPE_STRING_N.search(t)
        if str_match:
            return f"STR{str_match.group(1)}{suffix}"

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

        # Handle normalization for STRING
        if t == 'string': return 'STRING'

        t = _CLEAN_TYPE_RE.sub('', t)
        return t.upper() if t else 'U16'

    @staticmethod
    def validate_type(dtype: str) -> bool:
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
        try:
            return str(int(addr_part, 0))
        except ValueError:
            pass
        if re.match(r'^[0-9A-Fa-f]+$', addr_part):
            try: return str(int(addr_part, 16))
            except ValueError: return addr_part
        return addr_part

    @staticmethod
    def validate_address(address: str, dtype: str) -> bool:
        dtype_upper = dtype.upper()
        if RE_TYPE_STR_CONV.match(dtype_upper):
            dtype_upper = "STRING"

        if dtype_upper == 'STRING':
            if not RE_ADDR_STRING.match(address): return False
        elif dtype_upper == 'BITS':
            is_valid_format = RE_ADDR_BITS.match(address) is not None
        else:
            is_valid_format = RE_ADDR_INT.match(address) is not None

        if not is_valid_format:
            return False

        try:
            parts = address.split('_')
            base_addr_str = Generator.normalize_address_val(parts[0])
            base_addr = int(base_addr_str)
            if not (0 <= base_addr <= 65535):
                logging.warning(f"Address {base_addr} is out of standard Modbus range (0-65535).")
                return False
        except (ValueError, IndexError):
            return False
        return True

    def validate_csv(self, filepath: str) -> bool:
        """Deep validation of an existing WebdynSunPM definition file."""
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
                if not header or len(header) < 2:
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
                            logging.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (previously at line {seen_tags[tag]}).")
                            valid = False
                        else:
                            seen_tags[tag] = line_num

                    # Address validation
                    if not self.validate_address(address, dtype):
                        valid = False

                    if self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines):
                        valid = False

            return valid
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
        elif dtype_upper == 'STRING' or dtype_upper.startswith('STR'):
            try:
                if '_' in address:
                    return math.ceil(int(address.split('_')[1]) / 2)
                elif RE_TYPE_STR_CONV.match(dtype_upper):
                    return math.ceil(int(RE_TYPE_STR_CONV.match(dtype_upper).group(1)) / 2)
                return 0
            except (IndexError, ValueError): return 0
        return 1

    @staticmethod
    def apply_address_offset(address: Any, offset: int, line_num: Optional[int] = None, name: Optional[str] = None) -> str:
        if not address: return ""
        parts = str(address).split('_')
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

    def _check_address_overlap(self, info1: str, address: str, dtype: str, name: str, line_num: int, address_usage: Dict[str, Dict[str, Any]], warned_lines: Set[Tuple[int, int]]) -> bool:
        """Checks for address overlaps using O(log N) binary search on intervals."""
        try:
            addr_part = address.split('_')[0]
            start_addr = int(Generator.normalize_address_val(addr_part))
            reg_count = self.get_register_count(dtype, address)
            end_addr = start_addr + reg_count - 1

            if info1 not in address_usage:
                address_usage[info1] = {"intervals": [], "max_len": 0}

            usage = address_usage[info1]
            intervals = usage['intervals']
            max_len = usage['max_len']
            is_bits = (dtype.upper() == 'BITS')
            overlap_detected = False

            import bisect
            idx = bisect.bisect_left(intervals, (start_addr, -1, -1, '', ''))
            overlap_detected = False

            overlap_found = False
            # Check right
            for j in range(idx, len(intervals)):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if u_start > end_addr: break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if is_bits and u_type == 'BITS' and start_addr == u_start:
                        continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line}).")
                        warned_lines.add(warn_key)

            # Check left
            for j in range(idx - 1, -1, -1):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if start_addr - u_start > max_len: break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if is_bits and u_type == 'BITS' and start_addr == u_start:
                        continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line}).")
                        warned_lines.add(warn_key)

            bisect.insort(intervals, (start_addr, end_addr, line_num, name, dtype.upper()))
            if reg_count > max_len:
                usage['max_len'] = reg_count
            return overlap_detected
        except (ValueError, IndexError):
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

    def process_rows(self, rows: Iterable[Dict[str, Any]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        seen_names, seen_tags, address_usage, warned_lines = {}, {}, {}, set()
        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v): continue
            norm_row = {k.lower().strip(): (str(v).strip() if v is not None else '') for k, v in row.items()}
            name, tag, reg_type_str, address = norm_row.get('name', ''), norm_row.get('tag', ''), norm_row.get('registertype', ''), norm_row.get('address', '')
            dtype_raw, factor, offset, unit = norm_row.get('type', ''), norm_row.get('factor', ''), norm_row.get('offset', ''), norm_row.get('unit', '')
            action, scale_factor_str = norm_row.get('action', ''), norm_row.get('scalefactor', '')

            if not name and not address: continue
            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype_raw}'. Skipping.")
                continue

            match_str = RE_TYPE_STR_CONV.match(dtype)
            if match_str:
                dtype = 'STRING'
                if '_' not in address: address = f"{address}_{match_str.group(1)}"

            address = self.apply_address_offset(address, address_offset, line_num, name)
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}'. Skipping.")
                continue

            # Tag generation
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
                if tag in seen_tags: logging.warning(f"Line {line_num}: Duplicate Tag '{tag}'.")
                else: seen_tags[tag] = line_num

            info1 = self._determine_info1(reg_type_str, line_num)
            self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            coef_a, coef_b = self._calculate_coefficients(factor, offset, scale_factor_str)

            act_str = str(action).strip().upper()
            if not act_str:
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

    @staticmethod
    def sanitize_csv_field(field: Any) -> str:
        s = str(field)
        if s.startswith(('=', '+', '-', '@')):
            return "'" + s
        return s

    @staticmethod
    def write_output_csv(output: Union[str, Any, None], processed_rows: Iterable[Dict[str, Any]], manufacturer: str, model: str,
                        protocol: str = 'modbusRTU', category: str = 'Inverter', forced_write: str = '') -> None:
        """Centralized method to write the WebdynSunPM CSV format."""
        type_summary = {'1': 0, '2': 0, '3': 0, '4': 0}
        type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}
        total_index = 0
        try:
            if isinstance(output, str): outfile = open(output, 'w', newline='', encoding='utf-8-sig')
            elif output is None: outfile = sys.stdout
            else: outfile = output

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                total_index = index
                type_summary[row['Info1']] = type_summary.get(row['Info1'], 0) + 1
                writer.writerow([
                    str(index), row['Info1'], row['Info2'], row['Info3'], row['Info4'],
                    row['Name'], row['Tag'], row['CoefA'], row['CoefB'], row['Unit'], row['Action']
                ])

            summary_str = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in type_summary.items() if v > 0])
            if summary_str:
                logging.info(f"Generated {total_index} registers ({summary_str})")
        except (OSError, csv.Error) as e:
            logging.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and 'outfile' in locals() and not outfile.closed: outfile.close()

    @staticmethod
    def sanitize_csv_field(field: Any) -> str:
        s = str(field)
        if s.startswith(('=', '+', '-', '@')):
            return "'" + s
        return s

def generate_template(output_file: Optional[str], mode: str = 'input') -> None:
    if mode == 'definition':
        # Webdyn definition format
        headers = ['#Index', 'Info1', 'Info2', 'Info3', 'Info4', 'Name', 'Tag', 'CoefA', 'CoefB', 'Unit', 'Action']
        rows = [
            ['modbusRTU', 'Inverter', 'SampleMfg', 'SampleModel', '', '', '', '', '', '', ''],
            ['1', '3', '40001', 'U16', '', 'Active Power', 'active_power', '1.000000', '0.000000', 'W', '4']
        ]
        delimiter = ';'
    else:
        # Standard input format
        headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]
    try:
        if output_file:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';' if mode == 'definition' else ',')
                if mode == 'definition':
                    writer.writerow(headers)
                elif mode == 'input':
                    writer.writerow(headers)
                writer.writerows(rows)
        else:
            writer = csv.writer(sys.stdout, delimiter=';' if mode == 'definition' else ',')
            if mode == 'definition':
                writer.writerow(headers)
            elif mode == 'input':
                writer.writerow(headers)
            writer.writerows(rows)
    except OSError as e:
        logging.error(f"Error generating template: {e}")
    finally:
        if output_file and outfile: outfile.close()


def run_generator(
    config: GeneratorConfig, input_data: Optional[Iterable[dict[str, Any]]] = None
) -> None:
    generator = Generator()
    if config.template:
        mode = config.template_mode
        if input_data is not None: mode = 'definition'
        generate_template(config.output, mode=mode)
        return

    mfg = config.manufacturer or 'Manufacturer'
    model = config.model or 'Model'

    try:
        if input_data is not None:
            processed_rows = generator.process_rows(input_data, config.address_offset)
            generator.write_output_csv(config.output, processed_rows, mfg, model,
                                       config.protocol, config.category, config.forced_write)
        else:
            if not config.input_file or not os.path.exists(config.input_file):
                logging.error(f"Input file not found: {config.input_file}")
                return
            with open(config.input_file, mode='rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(config.input_file, 'r', encoding=encoding) as csvfile:
                snippet = csvfile.read(2048); csvfile.seek(0)
                try: dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                except csv.Error:
                    dialect = csv.excel
                    dialect.delimiter = ';' if ';' in snippet else ','
                reader = csv.DictReader(csvfile, dialect=dialect)
                processed_rows = generator.process_rows(reader, config.address_offset)
                generator.write_output_csv(config.output, processed_rows, mfg, model,
                                           config.protocol, config.category, config.forced_write)
    except Exception as e:
        logging.error(f"An error occurred during generation: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM definition file.')
    parser.add_argument('input_file', nargs='?')
    parser.add_argument('-o', '--output')
    parser.add_argument('--manufacturer')
    parser.add_argument('--model')
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
        template_mode=args.template_mode, address_offset=args.address_offset
    )
    run_generator(config)

if __name__ == "__main__":
    main()
