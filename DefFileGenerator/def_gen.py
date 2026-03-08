#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
from dataclasses import dataclass

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)
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
        """Normalizes manufacturer-specific data types to Webdyn types."""
        if not dtype:
            return 'U16'

        dtype_str = str(dtype).lower().strip()

        # Mapping ordered by specificity (longer strings first)
        type_synonyms = [
            (r'unsigned\s+int(eger)?\s*64', 'U64'),
            (r'signed\s+int(eger)?\s*64', 'I64'),
            (r'unsigned\s+int(eger)?\s*32', 'U32'),
            (r'signed\s+int(eger)?\s*32', 'I32'),
            (r'unsigned\s+int(eger)?\s*16', 'U16'),
            (r'signed\s+int(eger)?\s*16', 'I16'),
            (r'unsigned\s+int(eger)?\s*8', 'U8'),
            (r'signed\s+int(eger)?\s*8', 'I8'),
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
            (r'u64', 'U64'),
            (r'i64', 'I64'),
            (r'u32', 'U32'),
            (r'i32', 'I32'),
            (r'u16', 'U16'),
            (r'i16', 'I16'),
            (r'u8', 'U8'),
            (r'i8', 'I8'),
            (r'string', 'STRING'),
            (r'bits', 'BITS')
        ]

        for pattern, replacement in type_synonyms:
            if re.search(pattern, dtype_str):
                return replacement

        # Clean up common characters like () or space
        dtype_str = re.sub(r'[^a-z0-9_]+', '', dtype_str)
        return dtype_str.upper() if dtype_str else 'U16'

    def normalize_action(self, action):
        """Normalizes Action values to Webdyn codes."""
        if not action or not str(action).strip():
            return '1'

        act_str = str(action).strip().upper()
        if act_str in ['R', 'READ', '4']:
            return '4'
        if act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        if act_str in self.allowed_actions:
            return act_str

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
        # Remove thousands separators
        addr_part = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_part)

        if not addr_part:
            return ""

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

        # Heuristic for hex: contains A-F and not just digits/negative signs
        if any(c in val.upper() for c in 'ABCDEF') and not (val.isdigit() or (val.startswith('-') and val[1:].isdigit())):
            try:
                return str(int(val, 16))
            except ValueError:
                pass

        return val

    def validate_address(self, address, dtype):
        """Validates the address format based on type."""
        dtype_upper = dtype.upper()

        if dtype_upper == 'STRING':
            # Expect Address_Length (e.g., 30000_30 or 0x10_30)
            return RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == 'BITS':
            # Expect Address_StartBit_NbBits (e.g., 30000_0_1 or 0x10_0_1)
            return RE_ADDR_BITS.match(address) is not None
        else:
            # Expect integer address (dec or hex)
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
            # Parse length from address: Address_Length
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
        # Tracks used addresses per register type (Info1) to avoid O(N^2)
        address_usage = {} # info1 -> {addr: (line, name)}

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
            dtype_raw = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action_raw = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Normalize Type and Action early
            dtype = self.normalize_type(dtype_raw)
            action = self.normalize_action(action_raw)

            # Validation: Type
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
                    # Update address to Address_Length if not already formatted
                    if '_' not in address:
                        address = f"{address}_{length}"
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid STR format '{dtype_upper}'. Skipping row.")
                    continue

            # Normalize Address (convert any hex parts to decimal)
            if address:
                addr_parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]

                # Apply Address Offset to the primary address part
                try:
                    base_addr = int(norm_parts[0])
                    final_addr = base_addr + address_offset
                    if final_addr < 0:
                        logging.warning(f"Line {line_num}: Address {base_addr} with offset {address_offset} is negative ({final_addr}).")
                    norm_parts[0] = str(final_addr)
                except ValueError:
                    pass

                address = '_'.join(norm_parts)

            # Validation: Address format based on Type
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
            info1 = '3' # Default to Holding Register
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

                if info1 not in address_usage:
                    address_usage[info1] = {}

                is_bits = (dtype.upper() == 'BITS')

                for a in range(start_addr, start_addr + reg_count):
                    if a in address_usage[info1]:
                        prev_line, prev_name, prev_type = address_usage[info1][a]
                        # Overlap allowed if both are BITS
                        if not (is_bits and prev_type == 'BITS'):
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' at {a}. Overlaps with '{prev_name}' (Line {prev_line}).")
                    else:
                        address_usage[info1][a] = (line_num, name, dtype.upper())
            except ValueError:
                 logging.warning(f"Line {line_num}: Could not calculate register range for address '{address}'.")

            # CoefA calculation
            try:
                val_factor = float(factor) if factor and str(factor).strip() else 1.0
            except ValueError:
                val_factor = 1.0

            try:
                val_scale = int(float(scale_factor_str)) if scale_factor_str and str(scale_factor_str).strip() else 0
            except ValueError:
                 val_scale = 0

            final_coef_a_val = val_factor * (10 ** val_scale)
            coef_a = "{:.6f}".format(final_coef_a_val)

            # CoefB calculation
            try:
                val_offset = float(offset) if offset and str(offset).strip() else 0.0
                coef_b = "{:.6f}".format(val_offset)
            except ValueError:
                coef_b = "0.000000"

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
    def write_output_csv(output_file, processed_rows, config: GeneratorConfig):
        """Centralized method to write the WebdynSunPM definition CSV."""
        try:
            if output_file and isinstance(output_file, str):
                f = open(output_file, 'w', newline='', encoding='utf-8')
            else:
                f = output_file or sys.stdout

            # Header row
            header_row = [
                config.protocol,
                config.category,
                config.manufacturer,
                config.model,
                config.forced_write,
                '', '', '', '', '', ''
            ]

            writer = csv.writer(f, delimiter=';', lineterminator='\n')
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

            if output_file and isinstance(output_file, str):
                f.close()
        except Exception as e:
            logging.error(f"Error writing output file: {e}")
            raise

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
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(rows)
            logging.info(f"Template generated at {output_file}")
        else:
            writer = csv.writer(sys.stdout)
            writer.writerow(headers)
            writer.writerows(rows)
    except Exception as e:
        logging.error(f"Error generating template: {e}")

def run_generator(input_file=None, output=None, manufacturer=None, model=None,
                 protocol='modbusRTU', category='Inverter', forced_write='',
                 template=False, address_offset=0, config=None):
    if not config:
        config = GeneratorConfig(
            input_file=input_file,
            output=output,
            manufacturer=manufacturer,
            model=model,
            protocol=protocol,
            category=category,
            forced_write=forced_write,
            template=template,
            address_offset=address_offset
        )

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
        with open(config.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Sniff for delimiter
            content = csvfile.read(4096)
            csvfile.seek(0)
            try:
                dialect = csv.Sniffer().sniff(content, delimiters=";,")
            except csv.Error:
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(csvfile, dialect=dialect)
            if reader.fieldnames:
                reader.fieldnames = [n.strip() for n in reader.fieldnames]

            processed_rows = generator.process_rows(reader, address_offset=config.address_offset)
            generator.write_output_csv(config.output, processed_rows, config)

            if config.output:
                logging.info(f"Definition file generated at {config.output}")

    except FileNotFoundError:
        logging.error(f"File '{config.input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

    parser = argparse.ArgumentParser(description='Generate WebdynSunPM Modbus definition file from simplified CSV.')
    parser.add_argument('input_file', nargs='?', help='Path to the simplified CSV input file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--manufacturer', help='Manufacturer name.')
    parser.add_argument('--model', help='Model name.')
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true')
    parser.add_argument('--address-offset', type=int, default=0)

    args = parser.parse_args()
    run_generator(
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

if __name__ == "__main__":
    main()
