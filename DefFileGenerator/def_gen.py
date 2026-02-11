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
RE_ADDR_STRING = re.compile(r'^(\d+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^(\d+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^(\d+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
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
        # Data type aliases mapping to standard Webdyn types
        self.type_aliases = {
            'UINT8': 'U8',
            'INT8': 'I8',
            'UINT16': 'U16',
            'INT16': 'I16',
            'UINT32': 'U32',
            'INT32': 'I32',
            'UINT64': 'U64',
            'INT64': 'I64',
            'FLOAT': 'F32',
            'FLOAT32': 'F32',
            'DOUBLE': 'F64',
            'FLOAT64': 'F64'
        }
        # Allowed Action codes
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    def normalize_type(self, dtype):
        """Normalizes the data type to Webdyn standard or supported convenience types."""
        if not dtype:
            return 'U16'

        # Pre-process common words and spaces
        dtype_str = str(dtype).lower().strip()
        dtype_str = dtype_str.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')
        dtype_upper = dtype_str.upper()

        # Check aliases
        if dtype_upper in self.type_aliases:
            return self.type_aliases[dtype_upper]

        # Check for patterns like uint16, int32 etc.
        match_int = re.match(r'^(U|I|UINT|INT)(\d+)$', dtype_upper)
        if match_int:
            prefix = 'U' if match_int.group(1).startswith('U') else 'I'
            bits = match_int.group(2)
            return f"{prefix}{bits}"

        # Handle STR<n> - Preserve it as it's a valid convenience type for input
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return dtype_upper

        return dtype_upper

    def validate_type(self, dtype):
        """Validates the data type (accepts Webdyn types and STR<n>)."""
        dtype_upper = dtype.upper()
        # Base types
        base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
        if dtype_upper in base_types:
            return True

        # Integer types with optional suffixes
        if RE_TYPE_INT.match(dtype_upper):
            return True

        # STR<n> convenience type
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return True

        return False

    def normalize_action(self, action):
        """Normalizes action code from synonyms."""
        if not action:
            return '1' # Default per spec

        act_str = str(action).upper().strip()
        if act_str in ['R', 'READ']:
            return '4'
        elif act_str in ['RW', 'W', 'WRITE']:
            return '1'
        elif act_str in self.allowed_actions:
            return act_str

        return '1' # Default fallback

    def normalize_address_val(self, addr_part):
        """Converts a single address part (possibly hex) to decimal string."""
        if addr_part is None:
            return ""
        addr_part = str(addr_part).strip()
        if not addr_part:
            return ""

        # Remove commas (sometimes found in documentation like 40,001)
        if ',' in addr_part and '.' not in addr_part:
            addr_part = addr_part.replace(',', '')

        # Hex detection: 0x prefix, h suffix, or contains A-F and no other non-digit chars
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
        elif any(c in 'ABCDEFabcdef' for c in addr_part) and all(c in '0123456789ABCDEFabcdef' for c in addr_part):
            try:
                return str(int(addr_part, 16))
            except ValueError:
                return addr_part

        return addr_part

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
            dtype = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name:
                if address:
                    name = f"Register {address}"
                    logging.info(f"Line {line_num}: Missing Name, using 'Register {address}'")
                else:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                    continue

            # Normalization: Type
            original_dtype = dtype
            dtype = self.normalize_type(dtype)

            # Validation: Type
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{original_dtype}'. Skipping row.")
                continue

            # Handle STR<n> conversion and address update
            match_str = RE_TYPE_STR_CONV.match(dtype.upper())
            if match_str:
                try:
                    length = int(match_str.group(1))
                    dtype = 'STRING' # Convert to standard Webdyn type
                    # Update address to Address_Length if not already formatted
                    if address and '_' not in str(address):
                        address = f"{address}_{length}"
                except ValueError:
                    pass

            # Validation: Address format based on Type (checks hex/dec and composite)
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            # Normalize Address (convert any hex parts to decimal)
            if address:
                addr_parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]
                address = '_'.join(norm_parts)

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
            if not factor:
                val_factor = 1.0
            else:
                try:
                    val_factor = float(factor)
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using 1.0.")
                    val_factor = 1.0

            if not scale_factor_str:
                val_scale = 0
            else:
                try:
                    val_scale = int(float(scale_factor_str))
                except ValueError:
                     logging.warning(f"Line {line_num}: Invalid ScaleFactor '{scale_factor_str}'. Using 0.")
                     val_scale = 0

            final_coef_a_val = val_factor * (10 ** val_scale)
            # Use g formatting to avoid loss of precision on very small numbers but maintain readability
            if 0 < abs(final_coef_a_val) < 0.000001:
                coef_a = "{:.10f}".format(final_coef_a_val).rstrip('0').rstrip('.')
            else:
                coef_a = "{:.6f}".format(final_coef_a_val)

            # CoefB from Offset
            if not offset:
                coef_b = "0.000000"
            else:
                try:
                    coef_b = "{:.6f}".format(float(offset))
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Using as is.")
                    coef_b = offset

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
                 template=False):
    if template:
        generate_template(output)
        return

    if not input_file:
        logging.error("input_file is required")
        return

    if not manufacturer or not model:
         logging.error("--manufacturer and --model are required")
         return

    generator = Generator()

    try:
        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(input_file, mode='r', encoding='utf-8-sig') as csvfile:
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

    args = parser.parse_args()
    run_generator(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        template=args.template
    )

if __name__ == "__main__":
    main()
