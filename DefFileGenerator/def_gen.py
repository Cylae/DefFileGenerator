#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
import io
from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|_?\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|_?\d+)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|\d+)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

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
    address_offset: int = 0
    sheet: Optional[str] = None

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
    def normalize_type(dtype: Any) -> str:
        """Standardizes synonyms using specificity-ordered regex substitutions."""
        if not dtype:
            return 'U16'

        t_str = str(dtype).lower().strip()

        # Mapping ordered by specificity (longer strings before shorter ones)
        # to prevent incorrect partial matches.
        mappings = [
            (r'unsigned\s+int(?:eger)?\s*64', 'U64'),
            (r'signed\s+int(?:eger)?\s*64', 'I64'),
            (r'unsigned\s+int(?:eger)?\s*32', 'U32'),
            (r'signed\s+int(?:eger)?\s*32', 'I32'),
            (r'unsigned\s+int(?:eger)?\s*16', 'U16'),
            (r'signed\s+int(?:eger)?\s*16', 'I16'),
            (r'unsigned\s+int(?:eger)?\s*8', 'U8'),
            (r'signed\s+int(?:eger)?\s*8', 'I8'),
            (r'uint64', 'U64'),
            (r'int64', 'I64'),
            (r'uint32', 'U32'),
            (r'int32', 'I32'),
            (r'uint16', 'U16'),
            (r'int16', 'I16'),
            (r'uint8', 'U8'),
            (r'int8', 'I8'),
            (r'float64', 'F64'),
            (r'float32', 'F32'),
            (r'double', 'F64'),
            (r'float', 'F32'),
            (r'string', 'STRING'),
            (r'bits', 'BITS'),
            (r'ipv6', 'IPV6'),
            (r'ip', 'IP'),
            (r'mac', 'MAC')
        ]

        for pattern, replacement in mappings:
            if re.search(pattern, t_str):
                return replacement

        # Fallback to cleaning and uppercasing
        cleaned = _CLEAN_TYPE_RE.sub('', t_str).upper()
        return cleaned if cleaned else 'U16'

    def normalize_action(self, action: Any) -> str:
        """Normalizes 'Action' values by mapping synonyms."""
        if not action or not str(action).strip():
            return '1' # Default per spec

        act_str = str(action).strip().upper()
        if act_str in ['R', 'READ', '4']:
            return '4'
        if act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        if act_str in self.allowed_actions:
            return act_str

        logging.warning(f"Invalid Action '{action}'. Defaulting to '1'.")
        return '1'

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
        addr_part = str(addr_part).strip()
        # Remove thousands separators if present: 40,000 -> 40000
        # But only if it's a numeric comma: (?<=\d),(?=\d{3}(?!\d))
        addr_part = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_part)

        if not addr_part:
            return ""

        # Regex to identify candidate hex, decimal, or negative integer words
        # while avoiding partial matches.
        # This is more robust than simple string checks.
        word_match = re.search(r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])', addr_part)
        if not word_match:
            return addr_part

        val_str = word_match.group(1)

        if val_str.lower().startswith('0x'):
            try:
                return str(int(val_str, 16))
            except ValueError:
                return val_str
        elif val_str.lower().endswith('h'):
            try:
                return str(int(val_str[:-1], 16))
            except ValueError:
                return val_str
        # If it looks like hex (contains A-F and not just digits)
        if any(c in val_str.upper() for c in 'ABCDEF') and not val_str.startswith('-'):
             try:
                return str(int(val_str, 16))
             except ValueError:
                pass

        return val_str

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
        # Dictionary of Info1 -> list of tuples (start_addr, end_addr, line_num, name, type)
        address_usage = {}

        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v):
                continue

            def get_val(key):
                val = row.get(key)
                if val is not None:
                    return str(val).strip()
                for k, v in row.items():
                    if k.lower() == key.lower():
                        return str(v).strip()
                return ''

            name = get_val('Name')
            tag = get_val('Tag')
            reg_type_str = get_val('RegisterType')
            address = get_val('Address')
            dtype = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Normalize type before validation
            dtype = self.normalize_type(dtype)

            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
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

                # Apply address offset to the first part
                try:
                    base_addr = int(norm_parts[0])
                    final_addr = base_addr + address_offset
                    if final_addr < 0:
                        logging.warning(f"Line {line_num}: Resulting address {final_addr} is negative (base {base_addr} + offset {address_offset}).")
                    norm_parts[0] = str(final_addr)
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
            info1 = '3'
            if reg_type_str:
                lower_type = reg_type_str.lower()
                if lower_type in self.register_type_map:
                    info1 = self.register_type_map[lower_type]
                elif reg_type_str in ['1', '2', '3', '4']:
                    info1 = reg_type_str
                else:
                    logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to Holding Register (3).")

            # Overlap detection using dictionary lookup
            try:
                parts = address.split('_')
                start_addr = int(parts[0])
                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1

                if info1 not in address_usage:
                    address_usage[info1] = []

                is_bits = (dtype.upper() == 'BITS')
                for used_start, used_end, used_line, used_name, used_type in address_usage[info1]:
                    if max(start_addr, used_start) <= min(end_addr, used_end):
                        if not (is_bits and used_type == 'BITS' and start_addr == used_start):
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' ({start_addr}-{end_addr}). Overlaps with '{used_name}' (Line {used_line}, {used_start}-{used_end}) in RegisterType {info1}.")

                address_usage[info1].append((start_addr, end_addr, line_num, name, dtype.upper()))
            except (ValueError, IndexError):
                logging.warning(f"Line {line_num}: Could not calculate register range for address '{address}'.")

            # CoefA
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

            # CoefB
            try:
                val_offset = float(offset) if offset and str(offset).strip() else 0.0
                coef_b = "{:.6f}".format(val_offset)
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Defaulting to 0.000000.")
                coef_b = "0.000000"

            action = self.normalize_action(action)

            processed_rows.append({
                'Info1': info1,
                'Info2': address,
                'Info3': dtype.upper(),
                'Info4': '',
                'Name': name,
                'Tag': tag,
                'CoefA': coef_a,
                'CoefB': coef_b,
                'Unit': unit,
                'Action': action
            })

        return processed_rows

    @staticmethod
    def write_output_csv(output_path, processed_rows, config: GeneratorConfig):
        """Centralized CSV writing logic."""
        try:
            if output_path:
                outfile = open(output_path, 'w', newline='', encoding='utf-8')
            else:
                outfile = sys.stdout

            # Header row
            header_row = [
                config.protocol,
                config.category,
                config.manufacturer or '',
                config.model or '',
                config.forced_write or '',
                '', '', '', '', '', ''
            ]

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                writer.writerow([
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
                ])

            if output_path:
                outfile.close()
                logging.info(f"Definition file generated at {output_path}")
        except Exception as e:
            logging.error(f"Error writing output CSV: {e}")

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
        with (open(output_file, 'w', newline='', encoding='utf-8') if output_file else sys.stdout) as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)
            if output_file:
                logging.info(f"Template generated at {output_file}")
    except Exception as e:
        logging.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig):
    if config.template:
        generate_template(config.output)
        return

    if not config.input_file:
        logging.error("input_file is required")
        return

    if not config.manufacturer or not config.model:
         logging.error("--manufacturer and --model are required")
         return

    generator = Generator()

    try:
        # Robust encoding detection
        encoding = 'utf-8-sig'
        try:
            with open(config.input_file, 'rb') as f:
                raw = f.read(4)
                if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
                    encoding = 'utf-16'
        except Exception:
            pass

        with open(config.input_file, mode='r', encoding=encoding) as csvfile:
            content = csvfile.read()
            csvfile.seek(0)

            try:
                dialect = csv.Sniffer().sniff(content[:2048], delimiters=";,")
                csvfile.seek(0)
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

            required = ['Name', 'RegisterType', 'Address', 'Type']
            header_map = {h.lower(): h for h in reader.fieldnames}
            missing = [col for col in required if col.lower() not in header_map]
            if missing:
                logging.error(f"Missing required columns: {', '.join(missing)}")
                return

            processed_rows = generator.process_rows(reader, config.address_offset)
            generator.write_output_csv(config.output, processed_rows, config)

    except FileNotFoundError:
        logging.error(f"File '{config.input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Tool')
    parser.add_argument('input_file', nargs='?', help='Simplified CSV input file.')
    parser.add_argument('-o', '--output', help='Output CSV file.')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol (default: modbusRTU).')
    parser.add_argument('--category', default='Inverter', help='Category (default: Inverter).')
    parser.add_argument('--manufacturer', help='Manufacturer.')
    parser.add_argument('--model', help='Model.')
    parser.add_argument('--forced-write', default='', help='Forced write code.')
    parser.add_argument('--template', action='store_true', help='Generate template.')
    parser.add_argument('--address-offset', type=int, default=0, help='Global address offset.')

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
        address_offset=args.address_offset
    )
    run_generator(config)

if __name__ == "__main__":
    main()
