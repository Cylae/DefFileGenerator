#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
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
    address_offset: int = 0

class Generator:
    def __init__(self):
        # RegisterType mapping to Info1
        self.register_type_map = {
            'coil': '1',
            'coils': '1',
            'discrete input': '2',
            'holding register': '3',
            'holding': '3',
            'input register': '4',
            'input': '4'
        }
        # Allowed Action codes
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

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
        ]

        for pattern, replacement in synonyms:
            if re.search(pattern, t):
                return f"{replacement}{suffix}"

        t = _CLEAN_TYPE_RE.sub('', t)
        return t.upper() if t else 'U16'

    @staticmethod
    def validate_type(dtype):
        """Validates the data type."""
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
    def validate_address(address, dtype):
        """Validates the address format based on type."""
        dtype_upper = dtype.upper()

        if dtype_upper == 'STRING':
            return RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == 'BITS':
            return RE_ADDR_BITS.match(address) is not None
        else:
            return RE_ADDR_INT.match(address) is not None

    @staticmethod
    def get_register_count(dtype, address):
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
    def _get_val(row, key):
        """Helper to get value from row case-insensitively."""
        for k, v in row.items():
            if k.lower().strip() == key.lower():
                return str(v).strip() if v is not None else ''
        return ''

    @staticmethod
    def _parse_numeric(val, default=0.0):
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

    def apply_address_offset(self, address, offset, line_num=None, name=None):
        """Applies an integer offset to a register address (simple or compound)."""
        if not address:
            return ""
        parts = address.split('_')
        # Normalize each part individually
        norm_parts = [self.normalize_address_val(p) for p in parts]

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

    def _process_name_and_tag(self, name, tag, line_num, seen_names, seen_tags):
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

    def _determine_info1(self, reg_type_str, line_num):
        """Maps RegisterType string to Webdyn Info1 code."""
        if not reg_type_str:
            return '3'

        lt = reg_type_str.lower()
        if lt in self.register_type_map:
            return self.register_type_map[lt]
        elif reg_type_str in ['1', '2', '3', '4']:
            return reg_type_str

        logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to 3.")
        return '3'

    def _check_address_overlap(self, info1, address, dtype, name, line_num, address_usage, warned_lines):
        """Checks for address overlaps, allowing multiple BITS at same address."""
        try:
            start_addr = int(address.split('_')[0])
            reg_count = self.get_register_count(dtype, address)
            end_addr = start_addr + reg_count - 1

            if info1 not in address_usage:
                address_usage[info1] = {} # addr -> list of (line, name, type, start, end)

            is_bits = (dtype.upper() == 'BITS')

            for addr in range(start_addr, end_addr + 1):
                if addr in address_usage[info1]:
                    for u_line, u_name, u_type, u_start, u_end in address_usage[info1][addr]:
                        # Allow multiple BITS on exactly the same base address
                        if is_bits and u_type == 'BITS' and start_addr == u_start:
                            continue

                        pair = tuple(sorted((line_num, u_line)))
                        if pair not in warned_lines:
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' ({start_addr}-{end_addr}). Overlaps with '{u_name}' (Line {u_line}, {u_start}-{u_end}).")
                            warned_lines.add(pair)

                if addr not in address_usage[info1]:
                    address_usage[info1][addr] = []
                address_usage[info1][addr].append((line_num, name, dtype.upper(), start_addr, end_addr))
        except (ValueError, IndexError):
            pass

    @staticmethod
    def _calculate_coefficients(factor_str, offset_str, scale_factor_str):
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

    def process_rows(self, rows, address_offset=0):
        """Processes simplified CSV rows into WebdynSunPM format."""
        processed_rows = []
        seen_names = {}
        seen_tags = {}
        address_usage = {} # Info1 -> dict(addr -> list of usages)
        warned_lines = set()

        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v):
                continue

            name = self._get_val(row, 'Name')
            tag = self._get_val(row, 'Tag')
            reg_type_str = self._get_val(row, 'RegisterType')
            address = self._get_val(row, 'Address')
            dtype_raw = self._get_val(row, 'Type')
            factor = self._get_val(row, 'Factor')
            offset = self._get_val(row, 'Offset')
            unit = self._get_val(row, 'Unit')
            action = self._get_val(row, 'Action')
            scale_factor_str = self._get_val(row, 'ScaleFactor')

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
            address = self.apply_address_offset(address, address_offset, line_num, name)

            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            tag = self._process_name_and_tag(name, tag, line_num, seen_names, seen_tags)
            info1 = self._determine_info1(reg_type_str, line_num)

            self._check_address_overlap(info1, address, dtype, name, line_num, address_usage, warned_lines)

            coef_a, coef_b = self._calculate_coefficients(factor, offset, scale_factor_str)

            # Action normalization
            act_str = str(action).strip().upper()
            if not act_str:
                norm_action = '1'
            elif act_str in ['R', 'READ', '4']:
                norm_action = '4'
            elif act_str in ['RW', 'W', 'WRITE', '1']:
                norm_action = '1'
            elif act_str in self.allowed_actions:
                norm_action = act_str
            else:
                norm_action = '1'

            processed_rows.append({
                'Info1': info1, 'Info2': address, 'Info3': dtype.upper(), 'Info4': '',
                'Name': name, 'Tag': tag, 'CoefA': coef_a, 'CoefB': coef_b,
                'Unit': unit, 'Action': norm_action
            })
        return processed_rows

    @staticmethod
    def write_output_csv(output, processed_rows, manufacturer, model,
                        protocol='modbusRTU', category='Inverter', forced_write=''):
        """Centralized method to write the WebdynSunPM CSV format."""
        try:
            if isinstance(output, str):
                outfile = open(output, 'w', newline='', encoding='utf-8')
            elif output is None:
                outfile = sys.stdout
            else:
                outfile = output

            header_row = [protocol, category, manufacturer, model, forced_write, '', '', '', '', '', '']
            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                writer.writerow([
                    str(index), row['Info1'], row['Info2'], row['Info3'], row['Info4'],
                    row['Name'], row['Tag'], row['CoefA'], row['CoefB'], row['Unit'], row['Action']
                ])

            if isinstance(output, str):
                outfile.close()
                logging.info(f"Definition file generated at {output}")
        except Exception as e:
            logging.error(f"Error writing output CSV: {e}")

def generate_template(output_file):
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    rows = [
        ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
        ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
    ]
    try:
        if output_file:
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(rows)
        else:
            writer = csv.writer(sys.stdout)
            writer.writerow(headers)
            writer.writerows(rows)
    except Exception as e:
        logging.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig):
    if config.template:
        generate_template(config.output)
        return

    if not config.input_file or not config.manufacturer or not config.model:
        logging.error("input_file, manufacturer, and model are required.")
        return

    generator = Generator()
    try:
        with open(config.input_file, mode='rb') as f:
            content = f.read()
            encoding = 'utf-16' if content.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

        with open(config.input_file, mode='r', encoding=encoding) as csvfile:
            content = csvfile.read(2048)
            csvfile.seek(0)
            try:
                dialect = csv.Sniffer().sniff(content, delimiters=";,")
            except csv.Error:
                dialect = csv.excel

            reader = csv.DictReader(csvfile, dialect=dialect)
            processed_rows = generator.process_rows(reader, config.address_offset)

        generator.write_output_csv(config.output, processed_rows, config.manufacturer, config.model,
                                   config.protocol, config.category, config.forced_write)
    except Exception as e:
        logging.error(f"An error occurred: {e}")

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
