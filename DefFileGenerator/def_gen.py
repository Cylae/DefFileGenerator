#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Union

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^(-?\d+|[0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^(-?\d+|[0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^(-?\d+|[0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

# Robust address normalization regex
RE_ADDR_VAL = re.compile(r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])')

@dataclass
class GeneratorConfig:
    manufacturer: str
    model: str
    output: Optional[str] = None
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
    def normalize_action(action):
        """Centralized action normalization."""
        if not action or not str(action).strip():
            return '1' # Default per spec
        act_str = str(action).strip().upper()
        if act_str in ['R', 'READ', '4']:
            return '4'
        if act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        # Check if it's already a valid numeric code
        valid_codes = ['0', '1', '2', '4', '6', '7', '8', '9']
        if act_str in valid_codes:
            return act_str
        return '1' # Default fallback

    def normalize_type(self, dtype):
        """Normalizes data type synonyms with ordered specificity."""
        if not dtype:
            return 'U16'
        t = str(dtype).lower().strip()

        # Remove spaces but preserve suffixes
        t = t.replace(' ', '')

        # Common synonyms mapping ordered by specificity
        mappings = [
            (r'unsignedint64', 'U64'),
            (r'signedint64', 'I64'),
            (r'unsignedint32', 'U32'),
            (r'signedint32', 'I32'),
            (r'unsignedint16', 'U16'),
            (r'signedint16', 'I16'),
            (r'unsignedint8', 'U8'),
            (r'signedint8', 'I8'),
            (r'unsignedinteger64', 'U64'),
            (r'signedinteger64', 'I64'),
            (r'unsignedinteger32', 'U32'),
            (r'signedinteger32', 'I32'),
            (r'unsignedinteger16', 'U16'),
            (r'signedinteger16', 'I16'),
            (r'unsignedinteger8', 'U8'),
            (r'signedinteger8', 'I8'),
            (r'float64', 'F64'),
            (r'double', 'F64'),
            (r'float32', 'F32'),
            (r'float', 'F32'),
            (r'uint64', 'U64'),
            (r'int64', 'I64'),
            (r'uint32', 'U32'),
            (r'int32', 'I32'),
            (r'uint16', 'U16'),
            (r'int16', 'I16'),
            (r'uint8', 'U8'),
            (r'int8', 'I8'),
        ]

        for pattern, replacement in mappings:
            if re.search(pattern, t):
                # Check for suffixes (check longest suffix first)
                suffix = ''
                for s in ['_WB', '_W', '_B']:
                    if s.lower() in t:
                        suffix = s
                        break
                return replacement + suffix

        # Handle STR<n>
        if RE_TYPE_STR_CONV.match(t):
            return t.upper()

        return t.upper() if t else 'U16'

    def validate_type(self, dtype):
        """Validates the data type."""
        dtype_upper = dtype.upper()
        # Base types
        base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
        if dtype_upper in base_types:
            return True

        # Integer types with optional suffixes
        if RE_TYPE_INT.match(dtype_upper):
            return True

        # STR<n> syntax (e.g., STR20)
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return True

        return False

    def normalize_address_val(self, addr_part):
        """Converts a single address part (possibly hex) to decimal string."""
        addr_part = str(addr_part).strip().replace(',', '')
        if not addr_part:
            return ""

        # Use robust regex to find candidate hex/dec word
        match = RE_ADDR_VAL.search(addr_part)
        if not match:
            return addr_part

        val = match.group(1)

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

        # If it contains A-F and no leading/trailing context suggests it's not hex
        if any(c in val.upper() for c in 'ABCDEF'):
            try:
                return str(int(val, 16))
            except ValueError:
                return val

        return val

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
        return 1 # Default

    def process_rows(self, rows, address_offset=0):
        """Processes simplified CSV rows into WebdynSunPM format."""
        processed_rows = []
        seen_names = {}
        seen_tags = {}
        # Tracks used addresses per register type (Info1)
        used_addresses_by_type = {}

        for line_num, row in enumerate(rows, start=2):
            # Skip empty rows
            if not any(v for v in row.values() if v):
                continue

            def get_val(key):
                val = row.get(key)
                if val is not None:
                    return str(val).strip()
                # Case-insensitive fallback
                for k, v in row.items():
                    if k.lower() == key.lower():
                        return str(v).strip()
                return ''

            # Extract values
            name = get_val('Name')
            tag = get_val('Tag')
            reg_type_str = get_val('RegisterType')
            address = get_val('Address')
            dtype_input = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Normalize and validate Type
            dtype = self.normalize_type(dtype_input)
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype_input}' (normalized: '{dtype}'). Skipping row.")
                continue

            # Handle STR<n> conversion
            dtype_upper = dtype.upper()
            match_str = RE_TYPE_STR_CONV.match(dtype_upper)
            if match_str:
                try:
                    length = int(match_str.group(1))
                    dtype = 'STRING'
                    if '_' not in address:
                        address = f"{address}_{length}"
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid STR format '{dtype_upper}'. Skipping row.")
                    continue

            # Normalize Address
            if address:
                addr_parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]

                # Apply address_offset to the base address
                try:
                    base_addr = int(norm_parts[0])
                    new_base_addr = base_addr + address_offset
                    if new_base_addr < 0:
                         logging.warning(f"Line {line_num}: Resulting address {new_base_addr} is negative (base {base_addr} + offset {address_offset}).")
                    norm_parts[0] = str(new_base_addr)
                except ValueError:
                    pass

                address = '_'.join(norm_parts)

            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            # Global Check: Duplicate Name
            if name:
                if name in seen_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
                else:
                    seen_names[name] = line_num

            # Automatic Tag generation
            if not tag and name:
                base_tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
                if not base_tag:
                    base_tag = "var"
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

            # Map RegisterType to Info1
            info1 = '3' # Default
            if reg_type_str:
                lower_type = reg_type_str.lower()
                if lower_type in self.register_type_map:
                    info1 = self.register_type_map[lower_type]
                elif reg_type_str in ['1', '2', '3', '4']:
                    info1 = reg_type_str
                else:
                    logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to Holding Register (3).")

            # Address Overlap Calculation
            try:
                parts = address.split('_')
                start_addr = int(parts[0])
                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1
                current_range = (start_addr, end_addr, line_num, name, dtype.upper())

                if info1 not in used_addresses_by_type:
                    used_addresses_by_type[info1] = []

                is_bits = (dtype.upper() == 'BITS')
                for used_start, used_end, used_line, used_name, used_type in used_addresses_by_type[info1]:
                    if max(start_addr, used_start) <= min(end_addr, used_end):
                        if not (is_bits and used_type == 'BITS'):
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). Overlaps with '{used_name}' (Line {used_line}, Addr: {used_start}-{used_end}) in register type {info1}.")

                used_addresses_by_type[info1].append(current_range)
            except ValueError:
                logging.warning(f"Line {line_num}: Could not calculate register range for address '{address}'.")

            # CoefA / CoefB
            try:
                val_factor = float(factor) if factor and str(factor).strip() else 1.0
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using 1.0.")
                val_factor = 1.0
            try:
                val_scale = int(float(scale_factor_str)) if scale_factor_str and str(scale_factor_str).strip() else 0
            except ValueError:
                 logging.warning(f"Line {line_num}: Invalid ScaleFactor '{scale_factor_str}'. Using 0.")
                 val_scale = 0

            final_coef_a_val = val_factor * (10 ** val_scale)
            coef_a = "{:.6f}".format(final_coef_a_val)

            try:
                val_offset = float(offset) if offset and str(offset).strip() else 0.0
                coef_b = "{:.6f}".format(val_offset)
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Defaulting to 0.000000.")
                coef_b = "0.000000"

            action_norm = self.normalize_action(action)

            processed_row = {
                'Info1': info1,
                'Info2': address,
                'Info3': dtype.upper(),
                'Info4': '',
                'Name': name,
                'Tag': tag,
                'CoefA': coef_a,
                'CoefB': coef_b,
                'Unit': unit,
                'Action': action_norm
            }
            processed_rows.append(processed_row)

        return processed_rows

    @staticmethod
    def write_output_csv(processed_rows, config: GeneratorConfig):
        """Centralized method to write the WebdynSunPM CSV output."""
        try:
            if config.output:
                outfile = open(config.output, 'w', newline='', encoding='utf-8')
            else:
                outfile = sys.stdout

            # Header row
            header_row = [
                config.protocol,
                config.category,
                config.manufacturer,
                config.model,
                config.forced_write,
                '', '', '', '', '', ''
            ]

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                data_row = [
                    str(index),
                    row['Info1'],
                    row['Info2'],
                    row['Info3'],
                    row['Info4'],
                    row['Name'],
                    row['Tag'],
                    row['CoefA'],
                    row['CoefB'],
                    row['Unit'],
                    row['Action']
                ]
                writer.writerow(data_row)

            if config.output:
                outfile.close()
                logging.info(f"Definition file generated at {config.output}")
        except Exception as e:
            logging.error(f"Error writing output file: {e}")

def generate_template(output_file):
    """Generates a template CSV input file."""
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    rows = [
        ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
        ['String Variable', 'string_tag', 'Holding Register', '30010_10', 'String', '', '', '', '4', ''],
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'Bits', '', '', '', '4', ''],
        ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
    ]

    try:
        if output_file:
            f = open(output_file, 'w', newline='', encoding='utf-8')
        else:
            f = sys.stdout

        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)

        if output_file:
            f.close()
            logging.info(f"Template generated at {output_file}")
    except Exception as e:
        logging.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig, input_file: Optional[str] = None):
    if config.template:
        generate_template(config.output)
        return

    if not input_file:
        logging.error("input_file is required")
        return

    if not config.manufacturer or not config.model:
         logging.error("manufacturer and model are required")
         return

    generator = Generator()

    try:
        # Check for UTF-16 encoding (often used by Excel)
        encoding = 'utf-8-sig'
        try:
            with open(input_file, 'rb') as f:
                raw = f.read(2)
                if raw == b'\xff\xfe' or raw == b'\xfe\xff':
                    encoding = 'utf-16'
        except Exception:
            pass

        with open(input_file, mode='r', encoding=encoding) as csvfile:
            try:
                sample = csvfile.read(2048)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(sample, delimiters=";,")
            except csv.Error:
                csvfile.seek(0)
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(csvfile, dialect=dialect)

            if reader.fieldnames:
                reader.fieldnames = [name.strip() for name in reader.fieldnames]
            else:
                logging.error("Input CSV is empty or missing headers.")
                return

            required_columns = ['Name', 'RegisterType', 'Address', 'Type']
            header_map = {h.lower(): h for h in reader.fieldnames}
            missing_columns = [col for col in required_columns if col.lower() not in header_map]

            if missing_columns:
                logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)}")
                return

            processed_rows = generator.process_rows(reader, address_offset=config.address_offset)
            generator.write_output_csv(processed_rows, config)

    except FileNotFoundError:
        logging.error(f"File '{input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Generate WebdynSunPM Modbus definition file from simplified CSV.')
    parser.add_argument('input_file', nargs='?', help='Path to the simplified CSV input file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file. Defaults to stdout.')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name (default: modbusRTU).')
    parser.add_argument('--category', default='Inverter', help='Device category (default: Inverter).')
    parser.add_argument('--manufacturer', help='Manufacturer name.')
    parser.add_argument('--model', help='Model name.')
    parser.add_argument('--forced-write', default='', help='Forced write code (default: empty).')
    parser.add_argument('--template', action='store_true', help='Generate a template input CSV file.')
    parser.add_argument('--address-offset', type=int, default=0, help='Offset to add to all register addresses.')

    args = parser.parse_args()

    config = GeneratorConfig(
        manufacturer=args.manufacturer,
        model=args.model,
        output=args.output,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        template=args.template,
        address_offset=args.address_offset
    )

    run_generator(config, input_file=args.input_file)

if __name__ == "__main__":
    main()
