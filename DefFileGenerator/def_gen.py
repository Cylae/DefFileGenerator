#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
from dataclasses import dataclass

@dataclass
class GeneratorConfig:
    manufacturer: str
    model: str
    input_file: str = None
    output: str = None
    protocol: str = 'modbusRTU'
    category: str = 'Inverter'
    forced_write: str = ''
    template: bool = False
    address_offset: int = 0

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

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
        """Standardizes type synonyms to canonical Webdyn types."""
        if not dtype:
            return 'U16'
        t = str(dtype).strip().lower()

        # Remove common extra words and spaces
        t = t.replace('unsigned ', 'u').replace('signed ', 'i')
        t = t.replace('integer', '').replace('int', '')
        t = t.replace(' ', '')

        # Handle suffixes
        suffix = ""
        for s in ['_wb', '_w', '_b']:
            if t.endswith(s):
                suffix = s.upper()
                t = t[:-len(s)]
                break

        # Standard mappings
        mapping = {
            'float64': 'F64', 'double': 'F64', 'f64': 'F64',
            'float32': 'F32', 'float': 'F32', 'f32': 'F32',
            'string': 'STRING', 'bits': 'BITS',
            'ip': 'IP', 'ipv6': 'IPV6', 'mac': 'MAC'
        }
        if t in mapping:
            return mapping[t] + suffix

        # Handle u8, i16, etc.
        match = re.match(r'^([ui])(\d+)$', t)
        if match:
            return f"{match.group(1).upper()}{match.group(2)}{suffix}"

        return dtype.upper()

    def normalize_action(self, action):
        """Normalizes the action value."""
        if action is None or not str(action).strip():
            return '1'  # Default per spec
        act_str = str(action).strip().upper()
        if act_str in ['R', 'READ', '4']:
            return '4'
        elif act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        elif act_str in self.allowed_actions:
            return act_str
        else:
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
        if addr_part is None:
            return ""
        addr_str = str(addr_part).strip()

        # Remove thousands separators: only commas followed by exactly 3 digits (if it's decimal)
        # However, to be robust for hex as well, we use the regex from memory.
        # But first, let's just handle the base case.
        addr_str = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_str)

        # Regex to identify candidate hex or decimal words
        pattern = re.compile(r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])')
        match = pattern.search(addr_str)

        if not match:
            return addr_str

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
        # If it contains A-F and no other indicators, we try hex if it's not a valid decimal
        if any(c in val.upper() for c in 'ABCDEF'):
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
        # Tracks used addresses per register type (Info1) to check overlaps
        # Info1 -> dict: { register_address -> list of {'line': int, 'name': str, 'is_bits': bool, 'start': int, 'end': int} }
        address_usage = {}

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
            dtype = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Normalize and Validate Type
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
                    # Update address to Address_Length if not already formatted
                    if '_' not in address:
                        address = f"{address}_{length}"
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid STR format '{dtype_upper}'. Skipping row.")
                    continue

            # Normalize Address (convert any hex parts to decimal and remove commas)
            if address:
                addr_parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]

                # Apply address_offset to the base address
                try:
                    base_addr = int(norm_parts[0])
                    new_base_addr = base_addr + address_offset
                    if new_base_addr < 0:
                        logging.warning(f"Line {line_num}: Address {base_addr} becomes {new_base_addr} after offset {address_offset}.")
                    norm_parts[0] = str(new_base_addr)
                except ValueError:
                    pass

                address = '_'.join(norm_parts)

            # Validation: Address format based on Type (checks hex/dec and composite)
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            # Global Check: Duplicate Name
            if name:
                if name in seen_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
                else:
                    seen_names[name] = line_num

            # Automatic Tag generation and Global Check
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

            # Address Overlap Calculation (Optimized)
            try:
                parts = address.split('_')
                start_addr = int(parts[0])
                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1
                is_bits = (dtype.upper() == 'BITS')

                if info1 not in address_usage:
                    address_usage[info1] = {}

                overlap_found_with = set()
                for addr in range(start_addr, end_addr + 1):
                    if addr in address_usage[info1]:
                        for used in address_usage[info1][addr]:
                            if not (is_bits and used['is_bits']):
                                # Potential overlap. Only log once per overlapping variable.
                                if used['line'] not in overlap_found_with:
                                    logging.warning(
                                        f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). "
                                        f"Overlaps with '{used['name']}' (Line {used['line']}, Addr: {used['start']}-{used['end']}) "
                                        f"in register type {info1}."
                                    )
                                    overlap_found_with.add(used['line'])
                        address_usage[info1][addr].append({
                            'line': line_num, 'name': name, 'is_bits': is_bits,
                            'start': start_addr, 'end': end_addr
                        })
                    else:
                        address_usage[info1][addr] = [{
                            'line': line_num, 'name': name, 'is_bits': is_bits,
                            'start': start_addr, 'end': end_addr
                        }]

            except ValueError:
                logging.warning(f"Line {line_num}: Could not calculate register range for address '{address}'.")

            # Info2: Address
            info2 = address

            # Info3: Type
            info3 = dtype.upper() # Ensure uppercase in output

            # Info4: Empty
            info4 = ''

            # CoefA from Factor and ScaleFactor
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

            # CoefB from Offset
            try:
                val_offset = float(offset) if offset and str(offset).strip() else 0.0
                coef_b = "{:.6f}".format(val_offset)
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Defaulting to 0.000000.")
                coef_b = "0.000000"

            # Action normalization
            action = self.normalize_action(action)

            # Construct processed row
            processed_row = {
                'Info1': info1,
                'Info2': info2,
                'Info3': info3,
                'Info4': info4,
                'Name': name,
                'Tag': tag,
                'CoefA': coef_a,
                'CoefB': coef_b,
                'Unit': unit,
                'Action': action
            }

            processed_rows.append(processed_row)

        return processed_rows

    @staticmethod
    def write_output_csv(outfile, protocol, category, manufacturer, model, forced_write, processed_rows):
        """Centralized method to write the WebdynSunPM output CSV."""
        # Prepare output header row
        header_row = [
            protocol,
            category,
            manufacturer,
            model,
            forced_write,
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
        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(config.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Detect delimiter
            try:
                dialect = csv.Sniffer().sniff(csvfile.read(1024), delimiters=";,")
                csvfile.seek(0)
            except csv.Error:
                # Fallback to comma if detection fails
                csvfile.seek(0)
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(csvfile, dialect=dialect)

            # Normalize headers
            if reader.fieldnames:
                reader.fieldnames = [name.strip() for name in reader.fieldnames]
            else:
                logging.error("Input CSV is empty or missing headers.")
                return

            # Check for required columns
            required_columns = ['Name', 'RegisterType', 'Address', 'Type']
            header_map = {h.lower(): h for h in reader.fieldnames}
            missing_columns = [col for col in required_columns if col.lower() not in header_map]

            if missing_columns:
                logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)}")
                return

            processed_rows = generator.process_rows(reader, address_offset=config.address_offset)

        # Write Output
        if config.output:
            outfile = open(config.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        Generator.write_output_csv(
            outfile,
            config.protocol,
            config.category,
            config.manufacturer,
            config.model,
            config.forced_write,
            processed_rows
        )

        if config.output:
            outfile.close()
            logging.info(f"Definition file generated at {config.output}")

    except FileNotFoundError:
        logging.error(f"File '{config.input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def main():
    # Configure logging
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
    parser.add_argument('--address-offset', type=int, default=0, help='Address offset to apply to all registers.')

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
