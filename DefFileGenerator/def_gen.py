#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def generate_template(output_file):
    """Generates a template CSV input file."""
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action']
    rows = [
        ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4'],
        ['String Variable', 'string_tag', 'Holding Register', '30010_10', 'String', '', '', '', '4'],
        ['String Convenience', 'str_conv_tag', 'Holding Register', '30050', 'STR20', '', '', '', '4'],
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'Bits', '', '', '', '4']
    ]

    try:
        if output_file:
            f = open(output_file, 'w', newline='', encoding='utf-8')
        else:
            f = sys.stdout

        writer = csv.writer(f, delimiter=';')
        writer.writerow(headers)
        writer.writerows(rows)

        if output_file:
            f.close()
            logging.info(f"Template generated at {output_file}")
    except Exception as e:
        logging.error(f"Error generating template: {e}")

def validate_type(dtype):
    """Validates the data type."""
    dtype_upper = dtype.upper()
    # Base types
    base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
    if dtype_upper in base_types:
        return True

    # Integer types with optional suffixes
    # U8, U16, U32, U64, I8, I16, I32, I64
    # Suffixes: _W, _B, _WB
    # Regex: ^[UI](8|16|32|64)(_(W|B|WB))?$
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype_upper):
        return True

    # STR<n> convenience type
    if re.match(r'^STR\d+$', dtype_upper):
        return True

    return False

def validate_address(address, dtype):
    """Validates the address format based on type."""
    dtype_upper = dtype.upper()

    # STR<n> expects integer address (will be converted) or already Address_Length
    if re.match(r'^STR\d+$', dtype_upper):
        if re.match(r'^\d+$', address):
            return True
        # Also allow if user already put length
        return re.match(r'^\d+_\d+$', address) is not None

    if dtype_upper == 'STRING':
        # Expect Address_Length (e.g., 30000_30)
        return re.match(r'^\d+_\d+$', address) is not None
    elif dtype_upper == 'BITS':
        # Expect Address_StartBit_NbBits (e.g., 30000_0_1)
        return re.match(r'^\d+_\d+_\d+$', address) is not None
    else:
        # Expect integer address
        return re.match(r'^\d+$', address) is not None

def get_register_count(dtype, address):
    """Calculates the number of registers used by a type."""
    dtype = dtype.upper()

    if dtype == 'STRING':
        # Address is Addr_Len
        match = re.match(r'^\d+_(\d+)$', address)
        if match:
            length = int(match.group(1))
            return math.ceil(length / 2)
        return 0

    if dtype == 'BITS':
        return 1

    if dtype == 'IP':
        return 2
    if dtype == 'IPV6':
        return 8
    if dtype == 'MAC':
        return 3

    # Handle base types with optional suffixes
    base_type = dtype.split('_')[0]

    if base_type in ['U8', 'I8', 'U16', 'I16']:
        return 1
    if base_type in ['U32', 'I32', 'F32']:
        return 2
    if base_type in ['U64', 'I64', 'F64']:
        return 4

    return 1 # Default

class GlobalValidator:
    def __init__(self):
        self.seen_names = set()
        self.seen_tags = set()
        self.used_registers = {} # map register_index -> list of (Name, line_num)

    def check_duplicate_name(self, name, line_num):
        if name in self.seen_names:
            logging.warning(f"Line {line_num}: Duplicate Name '{name}'.")
        self.seen_names.add(name)

    def check_duplicate_tag(self, tag, line_num):
        if tag and tag in self.seen_tags:
            logging.warning(f"Line {line_num}: Duplicate Tag '{tag}'.")
        if tag:
            self.seen_tags.add(tag)

    def check_overlap(self, start_addr, num_regs, name, line_num, dtype):
        for i in range(num_regs):
            reg = start_addr + i
            if reg in self.used_registers:
                users = self.used_registers[reg]
                logging.warning(f"Line {line_num}: Address overlap at register {reg} (Name: '{name}', Type: '{dtype}'). Register already used by: {', '.join([f'{u[0]} (Line {u[1]})' for u in users])}")

            if reg not in self.used_registers:
                self.used_registers[reg] = []
            self.used_registers[reg].append((name, line_num))

def main():
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

    if args.template:
        generate_template(args.output)
        sys.exit(0)

    if not args.input_file:
        parser.error("the following arguments are required: input_file")

    if not args.manufacturer or not args.model:
         parser.error("the following arguments are required: --manufacturer, --model")

    # RegisterType mapping to Info1
    register_type_map = {
        'coil': '1',
        'discrete input': '2',
        'holding register': '3',
        'input register': '4'
    }

    # Allowed Action codes
    allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    validator = GlobalValidator()

    try:
        # Open output file or stdout
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Sniff delimiter
            try:
                sample = csvfile.read(1024)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(sample, delimiters=';,')
                delimiter = dialect.delimiter
            except csv.Error:
                # Fallback to semicolon
                csvfile.seek(0)
                delimiter = ';'

            reader = csv.DictReader(csvfile, delimiter=delimiter)

            # Normalize headers to remove whitespace and handle case insensitivity
            if reader.fieldnames:
                # Create a map of normalized lower case headers to actual headers
                header_map = {name.strip().lower(): name for name in reader.fieldnames}
            else:
                logging.error("Input CSV is empty or missing headers.")
                sys.exit(1)

            # Helper to get value case-insensitively
            def get_col(row, col_name):
                actual_col = header_map.get(col_name.lower())
                if actual_col:
                    return row.get(actual_col, '')
                return ''

            # Check for required columns
            required_columns = ['name', 'registertype', 'address', 'type']
            missing_columns = [col for col in required_columns if col not in header_map]
            if missing_columns:
                logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)} (Case insensitive check)")
                sys.exit(1)

            # Prepare output header row
            # Protocol;Category;Manufacturer;Model;ForcedWriteCode;;;;;;
            header_row = [
                args.protocol,
                args.category,
                args.manufacturer,
                args.model,
                args.forced_write,
                '', '', '', '', '', '' # Fill to 11 columns to match data row structure
            ]

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            index = 1
            for line_num, row in enumerate(reader, start=2): # Start at 2 because header is 1
                # Skip empty rows (if any)
                if not any(row.values()):
                    continue

                # Extract values using case-insensitive lookup
                name = get_col(row, 'name').strip()
                tag = get_col(row, 'tag').strip()
                reg_type_str = get_col(row, 'registertype').strip()
                address = get_col(row, 'address').strip()
                dtype = get_col(row, 'type').strip()
                factor = get_col(row, 'factor').strip()
                offset = get_col(row, 'offset').strip()
                unit = get_col(row, 'unit').strip()
                action = get_col(row, 'action').strip()

                if not name and not address:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                    continue

                # Validation: Type
                if not validate_type(dtype):
                    logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                    continue

                # Validation: Address format based on Type
                if not validate_address(address, dtype):
                    logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                    continue

                # Transformation: STR<n> to STRING
                str_match = re.match(r'^STR(\d+)$', dtype.upper())
                if str_match:
                    length = str_match.group(1)
                    dtype = 'STRING'
                    # If address doesn't have length, add it
                    if re.match(r'^\d+$', address):
                        address = f"{address}_{length}"
                    elif re.match(r'^\d+_\d+$', address):
                         pass

                # Global Validation: Duplicates and Overlap
                validator.check_duplicate_name(name, line_num)
                validator.check_duplicate_tag(tag, line_num)

                try:
                    start_addr = int(address.split('_')[0])
                    num_regs = get_register_count(dtype, address)
                    validator.check_overlap(start_addr, num_regs, name, line_num, dtype)
                except ValueError:
                    logging.warning(f"Line {line_num}: Error calculating register usage for address '{address}'.")


                # Map RegisterType to Info1
                info1 = '3' # Default to Holding Register
                if reg_type_str:
                    lower_type = reg_type_str.lower()
                    if lower_type in register_type_map:
                        info1 = register_type_map[lower_type]
                    elif reg_type_str in ['1', '2', '3', '4']:
                        info1 = reg_type_str
                    else:
                        logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to Holding Register (3).")

                # Info2: Address (supports Address_Length format)
                info2 = address

                # Info3: Type
                info3 = dtype

                # Info4: Empty
                info4 = ''

                # CoefA from Factor
                if not factor:
                    coef_a = "1.000000"
                else:
                    try:
                        coef_a = "{:.6f}".format(float(factor))
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using as is.")
                        coef_a = factor

                # CoefB from Offset
                if not offset:
                    coef_b = "0.000000"
                else:
                    try:
                        coef_b = "{:.6f}".format(float(offset))
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Using as is.")
                        coef_b = offset

                # Action
                if not action:
                    action = '1' # Default per spec
                elif action not in allowed_actions:
                    logging.warning(f"Line {line_num}: Invalid Action '{action}'. Defaulting to '1'.")
                    action = '1'

                # Construct data row
                data_row = [
                    str(index),
                    info1,
                    info2,
                    info3,
                    info4,
                    name,
                    tag,
                    coef_a,
                    coef_b,
                    unit,
                    action
                ]

                writer.writerow(data_row)
                index += 1

        if args.output:
            outfile.close()
            logging.info(f"Definition file generated at {args.output}")

    except FileNotFoundError:
        logging.error(f"File '{args.input_file}' not found.")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
