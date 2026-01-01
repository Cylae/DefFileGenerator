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
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'Bits', '', '', '', '4']
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

def get_type_size(dtype, address_str):
    """Calculates the number of registers used by a type."""
    dtype = dtype.upper()
    if dtype in ['U8', 'I8', 'U16', 'I16', 'BITS']:
        return 1
    elif dtype in ['U32', 'I32', 'F32', 'IP']:
        return 2
    elif dtype in ['U64', 'I64', 'F64']:
        return 4
    elif dtype == 'MAC':
        return 3
    elif dtype == 'IPV6':
        return 8
    elif dtype == 'STRING':
        # Address format: Addr_Length
        # If address_str is just a number (before update), we might not know length yet if it wasn't in STR<n>
        # But we expect address_str to be normalized before calling this.
        match = re.match(r'^(\d+)_(\d+)$', address_str)
        if match:
            length = int(match.group(2))
            return math.ceil(length / 2)
        return 1 # Default fallback
    return 1

class Validator:
    def __init__(self):
        self.names = set()
        self.tags = set()
        # map: address -> (line_num, name, type)
        self.register_map = {}

    def check_duplicate_name(self, name, line_num):
        if name and name in self.names:
            logging.warning(f"Line {line_num}: Duplicate Name '{name}'.")
        if name:
            self.names.add(name)

    def check_duplicate_tag(self, tag, line_num):
        if tag and tag in self.tags:
            logging.warning(f"Line {line_num}: Duplicate Tag '{tag}'.")
        if tag:
            self.tags.add(tag)

    def check_overlap(self, address_str, dtype, name, line_num):
        # Parse start address
        # For BITS: Addr_Start_Nb -> Addr
        # For STRING: Addr_Len -> Addr
        parts = address_str.split('_')
        if not parts[0].isdigit():
            return # Should have been caught by validate_address

        start_addr = int(parts[0])
        size = get_type_size(dtype, address_str)

        for i in range(size):
            addr = start_addr + i
            if addr in self.register_map:
                prev_line, prev_name, prev_type = self.register_map[addr]

                # Allow overlap if both are BITS (usually different bits in same register)
                if dtype.upper() == 'BITS' and prev_type.upper() == 'BITS':
                    continue

                logging.warning(f"Line {line_num}: Address overlap at {addr} (Variable '{name}', Type '{dtype}') "
                                f"conflicts with Line {prev_line} (Variable '{prev_name}', Type '{prev_type}').")
            else:
                self.register_map[addr] = (line_num, name, dtype)

def validate_type(dtype):
    """Validates the data type."""
    # Check for STR<n>
    if re.match(r'^STR\d+$', dtype.upper()):
        return True

    # Base types
    base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
    if dtype.upper() in base_types:
        return True

    # Integer types with optional suffixes
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype.upper()):
        return True

    return False

def validate_address(address, dtype):
    """Validates the address format based on type."""
    dtype_upper = dtype.upper()
    if dtype_upper == 'STRING':
        # Expect Address_Length (e.g., 30000_30)
        return re.match(r'^\d+_\d+$', address) is not None
    elif dtype_upper == 'BITS':
        # Expect Address_StartBit_NbBits (e.g., 30000_0_1)
        return re.match(r'^\d+_\d+_\d+$', address) is not None
    else:
        # Expect integer address
        return re.match(r'^\d+$', address) is not None

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

    validator = Validator()

    try:
        # Open output file or stdout
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Sniffer to detect delimiter
            try:
                sample = csvfile.read(1024)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(sample)
                delimiter = dialect.delimiter
            except csv.Error:
                # Fallback to comma if sniffing fails
                delimiter = ','

            reader = csv.DictReader(csvfile, delimiter=delimiter)

            # Normalize headers to remove whitespace
            if reader.fieldnames:
                reader.fieldnames = [name.strip() for name in reader.fieldnames]
            else:
                logging.error("Input CSV is empty or missing headers.")
                sys.exit(1)

            # Check for required columns
            required_columns = ['Name', 'RegisterType', 'Address', 'Type']
            missing_columns = [col for col in required_columns if col not in reader.fieldnames]
            if missing_columns:
                logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)}")
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

                # Extract values
                name = row.get('Name', '').strip()
                tag = row.get('Tag', '').strip()
                reg_type_str = row.get('RegisterType', '').strip()
                address = row.get('Address', '').strip()
                dtype = row.get('Type', '').strip()
                factor = row.get('Factor', '').strip()
                offset = row.get('Offset', '').strip()
                unit = row.get('Unit', '').strip()
                action = row.get('Action', '').strip()

                if not name and not address:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                    continue

                # Handle STR<n> convenience type
                str_match = re.match(r'^STR(\d+)$', dtype.upper())
                if str_match:
                    length = str_match.group(1)
                    dtype = 'STRING'
                    # Update address if it doesn't have length
                    if '_' not in address:
                        address = f"{address}_{length}"

                # Validation: Type
                if not validate_type(dtype):
                    logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                    continue

                # Validation: Address format based on Type
                if not validate_address(address, dtype):
                    logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                    continue

                # Validator checks (Duplicates, Overlaps)
                validator.check_duplicate_name(name, line_num)
                validator.check_duplicate_tag(tag, line_num)
                validator.check_overlap(address, dtype, name, line_num)

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
