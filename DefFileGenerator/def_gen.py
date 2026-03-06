#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F-]+|0x[0-9A-F-]+|[0-9A-F-]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F-]+|0x[0-9A-F-]+|[0-9A-F-]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

class Generator:
    def __init__(self, address_offset=0):
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
        self.address_offset = address_offset

    def normalize_type(self, dtype):
        """Standardizes data type synonyms."""
        if not dtype:
            return 'U16'

        # Initial cleaning: preserve alphanumeric, underscores, and case-insensitive
        # but normalize common terms
        t = str(dtype).lower().strip()
        t = t.replace('unsigned ', 'u').replace('signed ', 'i')
        t = t.replace('uint', 'u').replace('int', 'i')
        t = t.replace(' ', '') # Remove all spaces

        # Specific mappings ordered by specificity
        if 'float64' in t or 'double' in t:
            return 'F64'
        if 'float32' in t or 'float' in t:
            return 'F32'

        # Regex-based standardized mapping for U/I types
        # Match u16, ui16, i32, etc. and preserve suffixes
        match = re.match(r'^([ui]+)(\d+)(_.*)?$', t)
        if match:
            prefix = 'U' if 'u' in match.group(1) else 'I'
            bits = match.group(2)
            suffix = match.group(3).upper() if match.group(3) else ''
            return f"{prefix}{bits}{suffix}"

        # Fallback to general cleaning while preserving case for BITS, STRING etc if they matched standard
        t_clean = re.sub(r'[^a-z0-9_]+', '', str(dtype).lower())
        if t_clean == 'string': return 'STRING'
        if t_clean == 'bits': return 'BITS'
        if t_clean == 'ip': return 'IP'
        if t_clean == 'ipv6': return 'IPV6'
        if t_clean == 'mac': return 'MAC'

        # Check for STR<n>
        if RE_TYPE_STR_CONV.match(dtype):
            return dtype.upper()

        return dtype.upper()

    def normalize_action(self, action):
        """Standardizes action codes."""
        if not action or not str(action).strip():
            return '1'
        a = str(action).upper().strip()
        if a in ['R', 'READ', '4']:
            return '4'
        if a in ['RW', 'W', 'WRITE', '1']:
            return '1'
        if a in self.allowed_actions:
            return a
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

    def normalize_address_val(self, addr_val):
        """Converts a single address part (possibly hex) to decimal string."""
        addr_val = str(addr_val).strip().replace(',', '')
        if not addr_val:
            return ""

        # 1. Explicit hex formats
        if addr_val.lower().startswith('0x'):
            try: return str(int(addr_val, 16))
            except ValueError: pass
        if addr_val.lower().endswith('h'):
            try: return str(int(addr_val[:-1], 16))
            except ValueError: pass

        # 2. Extract potential number/hex
        # Use lookbehind/lookahead to match candidate words that are not parts of other words
        # (avoiding single hex letters in words like 'Reg')
        pattern = r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])'
        match = re.search(pattern, addr_val)
        if match:
            val = match.group(1)
            # Re-check explicit hex in case it was found inside messy string
            if val.lower().startswith('0x'):
                try: return str(int(val, 16))
                except ValueError: pass
            if val.lower().endswith('h'):
                try: return str(int(val[:-1], 16))
                except ValueError: pass

            # If it's a decimal number
            if re.match(r'^-?\d+$', val):
                return str(int(val))

            # If it's hex (contains A-F)
            if any(c in val.upper() for c in 'ABCDEF'):
                try: return str(int(val, 16))
                except ValueError: pass

            # Fallback for plain decimal if no A-F but didn't match \d+ (shouldn't happen with above regex)
            try:
                return str(int(val))
            except ValueError:
                pass

        # 3. Final fallback: whole string if it's just alphanumeric hex
        if re.match(r'^[0-9A-Fa-f]+$', addr_val):
             try: return str(int(addr_val, 16))
             except ValueError: pass

        return addr_val

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
            # Expect integer address (dec or hex, possibly negative)
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

    def process_rows(self, rows):
        """Processes simplified CSV rows into WebdynSunPM format."""
        processed_rows = []
        seen_names = {}
        seen_tags = {}
        # Tracks used addresses per register type (Info1)
        # Dictionary of Info1 -> list of tuples (start_addr, end_addr, line_num, name, type)
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
            raw_dtype = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            raw_action = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Normalize Type before validation
            dtype = self.normalize_type(raw_dtype)

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

            # Normalize Address (convert any hex parts to decimal and remove commas)
            if address:
                addr_parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]
                address = '_'.join(norm_parts)

            # Validation: Address format based on Type (checks hex/dec and composite)
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            # Apply address offset
            if address:
                try:
                    parts = address.split('_')
                    raw_start_addr = int(parts[0])
                    start_addr = raw_start_addr - self.address_offset
                    if start_addr < 0:
                        logging.warning(f"Line {line_num}: Address {raw_start_addr} with offset {self.address_offset} results in negative address {start_addr}")
                    parts[0] = str(start_addr)
                    address = '_'.join(parts)
                except ValueError:
                    pass

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

            # Address Overlap Calculation
            try:
                # Parse start address
                parts = address.split('_')
                start_addr = int(parts[0])

                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1

                current_range = (start_addr, end_addr, line_num, name, dtype.upper())

                if info1 not in used_addresses_by_type:
                    used_addresses_by_type[info1] = []

                # Check overlap against used_addresses for this register type
                is_bits = (dtype.upper() == 'BITS')

                for used_start, used_end, used_line, used_name, used_type in used_addresses_by_type[info1]:
                    # Check if ranges overlap
                    if max(start_addr, used_start) <= min(end_addr, used_end):
                        # Overlap detected
                        # Allowed only if both are BITS and same start address
                        is_overlap_allowed = is_bits and (used_type == 'BITS')

                        if not is_overlap_allowed:
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). Overlaps with '{used_name}' (Line {used_line}, Addr: {used_start}-{used_end}) in register type {info1}.")

                used_addresses_by_type[info1].append(current_range)

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
            action = self.normalize_action(raw_action)

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

def run_generator(input_file, output=None, manufacturer=None, model=None,
                 protocol='modbusRTU', category='Inverter', forced_write='',
                 template=False, address_offset=0):
    if template:
        generate_template(output)
        return

    if not input_file:
        logging.error("input_file is required")
        return

    if not manufacturer or not model:
         logging.error("--manufacturer and --model are required")
         return

    generator = Generator(address_offset=address_offset)

    try:
        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Detect delimiter
            try:
                content = csvfile.read(1024)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(content, delimiters=";,")
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

            processed_rows = generator.process_rows(reader)

        # Write Output
        if output:
            outfile = open(output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

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

        if output:
            outfile.close()
            logging.info(f"Definition file generated at {output}")

    except FileNotFoundError:
        logging.error(f"File '{input_file}' not found.")
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
    parser.add_argument('--address-offset', type=int, default=0, help='Subtract this value from all addresses.')

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
