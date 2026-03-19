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
RE_TAG_CLEAN = re.compile(r'[^a-zA-Z0-9]')
RE_ADDR_VAL = re.compile(r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])', re.IGNORECASE)

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

    def normalize_type(self, dtype):
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

    def validate_type(self, dtype):
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

    def normalize_address_val(self, addr_part):
        """Converts a single address part (possibly hex) to decimal string."""
        addr_part = str(addr_part).strip()
        # Remove thousands separators if they exist
        addr_part = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_part)

        if not addr_part:
            return ""

        match = RE_ADDR_VAL.fullmatch(addr_part)
        if not match:
            return addr_part

        val = match.group(1)
        # Support hex with 0x prefix or h suffix
        if val.lower().startswith('0x'):
            try:
                return str(int(val, 16))
            except ValueError:
                return val
        elif val.lower().endswith('h'):
            try:
                return str(int(val[:-1], 16))
            except ValueError:
                return val

        # Handle decimal (including negative)
        try:
            return str(int(val))
        except ValueError:
            pass

        # If it's a raw hex word (e.g. "A0")
        if re.match(r'^[0-9A-Fa-f]+$', val):
            try:
                return str(int(val, 16))
            except ValueError:
                return val

        return val

    def apply_address_offset(self, address, offset):
        """Applies an integer offset to a Modbus address (simple or complex)."""
        if not address or offset == 0:
            return address
        parts = address.split('_')
        try:
            base_addr = int(self.normalize_address_val(parts[0])) + offset
            parts[0] = str(base_addr)
            return '_'.join(parts)
        except (ValueError, IndexError):
            return address

    def validate_address(self, address, dtype):
        """Validates the address format based on type."""
        dtype_upper = dtype.upper()

        if dtype_upper == 'STRING':
            return RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == 'BITS':
            return RE_ADDR_BITS.match(address) is not None
        else:
            return RE_ADDR_INT.match(address) is not None

    def get_register_count(self, dtype, address):
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

    def normalize_action(self, action):
        """Normalizes action codes and synonyms."""
        act_str = str(action).strip().upper()
        if not act_str:
            return '1'
        if act_str in ['R', 'READ', '4']:
            return '4'
        if act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        if act_str in self.allowed_actions:
            return act_str
        return '1'

    def _get_val(self, row_clean, key):
        """Internal helper for robust dictionary lookup."""
        return str(row_clean.get(key.lower(), '')).strip()

    def _normalize_type_and_address(self, row_clean, line_num, address_offset):
        """Normalizes data type and address, handling STR<n> and offsets."""
        dtype_raw = self._get_val(row_clean, 'Type')
        dtype = self.normalize_type(dtype_raw)

        if not self.validate_type(dtype):
            logging.warning(f"Line {line_num}: Invalid Type '{dtype_raw}' (normalized to '{dtype}'). Skipping row.")
            return None

        address = self._get_val(row_clean, 'Address')

        # Handle STR<n> conversion
        match_str = RE_TYPE_STR_CONV.match(dtype)
        if match_str:
            length = int(match_str.group(1))
            dtype = 'STRING'
            if '_' not in address:
                address = f"{address}_{length}"

        if address:
            address = self.apply_address_offset(address, address_offset)
            try:
                base_addr = int(address.split('_')[0])
                if base_addr < 0:
                     logging.warning(f"Line {line_num}: Address offset {address_offset} results in negative address {base_addr} for variable.")
            except (ValueError, IndexError):
                pass

        if not self.validate_address(address, dtype):
            logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
            return None

        return dtype, address

    def _process_name_and_tag(self, row_clean, line_num, seen_names, seen_tags):
        """Processes Name and Tag, handling duplicates and auto-generation."""
        name = self._get_val(row_clean, 'Name')
        tag = self._get_val(row_clean, 'Tag')

        if name:
            if name in seen_names:
                logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
            else:
                seen_names[name] = line_num

        if not tag and name:
            base_tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
            tag = base_tag if base_tag else "var"
            counter = 1
            original_tag = tag
            while tag in seen_tags:
                tag = f"{original_tag}_{counter}"
                counter += 1

        if tag:
            if tag in seen_tags:
                logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous occurrence at line {seen_tags[tag]}.")
            else:
                seen_tags[tag] = line_num

        return name, tag

    def _determine_info1(self, row_clean, line_num):
        """Maps RegisterType string to Webdyn Info1 code."""
        reg_type_str = self._get_val(row_clean, 'RegisterType')
        info1 = '3'
        if reg_type_str:
            lt = reg_type_str.lower()
            if lt in self.register_type_map:
                info1 = self.register_type_map[lt]
            elif reg_type_str in ['1', '2', '3', '4']:
                info1 = reg_type_str
            else:
                logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to 3.")
        return info1

    def _check_address_overlap(self, name, address, dtype, info1, line_num, address_usage):
        """Performs O(1) overlap detection using the address_usage dictionary."""
        try:
            start_addr = int(address.split('_')[0])
            reg_count = self.get_register_count(dtype, address)

            if info1 not in address_usage:
                address_usage[info1] = {}

            is_bits = (dtype.upper() == 'BITS')

            for i in range(reg_count):
                curr_addr = start_addr + i
                if curr_addr in address_usage[info1]:
                    u_line, u_name, u_dtype = address_usage[info1][curr_addr]
                    if not (is_bits and u_dtype == 'BITS' and start_addr == curr_addr):
                         logging.warning(f"Line {line_num}: Address overlap detected for '{name}' ({start_addr}). Overlaps with '{u_name}' (Line {u_line}).")
                         return
                else:
                    address_usage[info1][curr_addr] = (line_num, name, dtype.upper())
        except (ValueError, IndexError):
            pass

    def _calculate_coefficients(self, row_clean):
        """Calculates CoefA and CoefB based on Factor, Offset, and ScaleFactor."""
        factor = self._get_val(row_clean, 'Factor')
        scale_factor_str = self._get_val(row_clean, 'ScaleFactor')
        try:
            val_factor = float(factor) if factor else 1.0
            val_scale = int(float(scale_factor_str)) if scale_factor_str else 0
            coef_a = "{:.6f}".format(val_factor * (10 ** val_scale))
        except ValueError:
            coef_a = "1.000000"

        offset = self._get_val(row_clean, 'Offset')
        try:
            val_offset = float(offset) if offset else 0.0
            coef_b = "{:.6f}".format(val_offset)
        except ValueError:
            coef_b = "0.000000"

        return coef_a, coef_b

    def process_rows(self, rows, address_offset=0):
        """Processes simplified CSV rows into WebdynSunPM format."""
        processed_rows = []
        seen_names = {}
        seen_tags = {}
        address_usage = {}

        for line_num, row in enumerate(rows, start=2):
            row_clean = {str(k).lower().strip(): v for k, v in row.items()}

            if not any(v for v in row.values() if v):
                continue

            # 1. Normalize Type and Address
            res = self._normalize_type_and_address(row_clean, line_num, address_offset)
            if not res: continue
            dtype, address = res

            # 2. Process Name and Tag
            name, tag = self._process_name_and_tag(row_clean, line_num, seen_names, seen_tags)
            if not name and not address: continue

            # 3. Determine Info1
            info1 = self._determine_info1(row_clean, line_num)

            # 4. Check Address Overlap
            self._check_address_overlap(name, address, dtype, info1, line_num, address_usage)

            # 5. Calculate Coefficients
            coef_a, coef_b = self._calculate_coefficients(row_clean)

            # 6. Normalize Action
            action = self.normalize_action(self._get_val(row_clean, 'Action'))

            unit = self._get_val(row_clean, 'Unit')

            processed_rows.append({
                'Info1': info1, 'Info2': address, 'Info3': dtype.upper(), 'Info4': '',
                'Name': name, 'Tag': tag, 'CoefA': coef_a, 'CoefB': coef_b,
                'Unit': unit, 'Action': action
            })
        return processed_rows

    @staticmethod
    def write_output_csv(output, processed_rows, manufacturer, model,
                        protocol='modbusRTU', category='Inverter', forced_write=''):
        """Centralized method to write the WebdynSunPM CSV format."""
        try:
            if not output:
                outfile = sys.stdout
            elif isinstance(output, str):
                outfile = open(output, 'w', newline='', encoding='utf-8')
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

    @staticmethod
    def generate_template(output_file):
        """Generates a sample definition template CSV."""
        headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]
        try:
            f = open(output_file, 'w', newline='', encoding='utf-8') if output_file else sys.stdout
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)
            if output_file:
                f.close()
        except Exception as e:
            logging.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig):
    if config.template:
        Generator.generate_template(config.output)
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
