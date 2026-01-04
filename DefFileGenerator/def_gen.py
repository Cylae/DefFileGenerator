#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re

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

def validate_type(dtype):
    """Validates the data type."""
    # Base types
    base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
    if dtype.upper() in base_types:
        return True

    # Integer types with optional suffixes
    # U8, U16, U32, U64, I8, I16, I32, I64
    # Suffixes: _W, _B, _WB
    # Regex: ^[UI](8|16|32|64)(_(W|B|WB))?$
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype):
        return True

    # STR<n> support (e.g., STR20)
    if re.match(r'^STR\d+$', dtype):
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

def get_register_size(dtype, address_str):
    """Calculates the number of registers used by a type."""
    dtype_upper = dtype.upper()

    if dtype_upper == 'BITS':
        return 1 # Bits are packed in registers, checking overlap is complex but base is 1 reg
    elif dtype_upper == 'STRING':
        # Address format: Start_Length
        try:
            _, length = address_str.split('_')
            length = int(length)
            return (length + 1) // 2 # 2 chars per register? Or 1?
            # Standard Modbus string usually 2 chars per register (16-bit).
            # ceil(length / 2)
            return (length + 1) // 2
        except ValueError:
            return 1 # Fallback
    elif dtype_upper in ['U8', 'I8', 'U16', 'I16']:
        return 1
    elif dtype_upper in ['U32', 'I32', 'F32', 'IP']:
        return 2
    elif dtype_upper in ['U64', 'I64', 'F64']:
        return 4
    elif dtype_upper == 'MAC':
        return 3 # 6 bytes
    elif dtype_upper == 'IPV6':
        return 8 # 16 bytes

    # Default fallback
    return 1

def parse_address_start(address, dtype):
    """Extracts the start register address."""
    if '_' in address:
        return int(address.split('_')[0])
    return int(address)

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

    # Global tracking for validation
    used_names = set()
    used_tags = set()
    # address_map maps register address to list of (variable_name, type)
    address_map = {}

    try:
        # Open output file or stdout
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig', newline='') as csvfile:
            # Sniff dialect
            try:
                sample = csvfile.read(1024)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(sample)
                # Ensure delimiters are reasonable (comma or semicolon)
                if dialect.delimiter not in [',', ';']:
                     dialect.delimiter = ','
            except csv.Error:
                # Fallback to comma if sniffing fails
                csvfile.seek(0)
                dialect = 'excel'

            reader = csv.DictReader(csvfile, dialect=dialect)

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
                scale_factor = row.get('ScaleFactor', '').strip()
                offset = row.get('Offset', '').strip()
                unit = row.get('Unit', '').strip()
                action = row.get('Action', '').strip()

                if not name and not address:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                    continue

                # Validation: Type
                if not validate_type(dtype):
                    logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                    continue

                # Handle STR<n> -> STRING conversion and address update
                dtype_upper = dtype.upper()
                str_match = re.match(r'^STR(\d+)$', dtype_upper)
                if str_match:
                    length = str_match.group(1)
                    dtype = 'STRING'
                    # Update address to Address_Length if not already present
                    if '_' not in address:
                        address = f"{address}_{length}"

                # Validation: Address format based on Type
                if not validate_address(address, dtype):
                    logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                    continue

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

                # CoefA from Factor and ScaleFactor
                if not factor:
                    f_val = 1.0
                else:
                    try:
                        f_val = float(factor)
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using 1.0.")
                        f_val = 1.0

                if scale_factor:
                    try:
                        sf_val = float(scale_factor)
                        f_val = f_val * (10 ** sf_val)
                    except ValueError:
                         logging.warning(f"Line {line_num}: Invalid ScaleFactor '{scale_factor}'. Ignoring.")

                coef_a = "{:.6f}".format(f_val)

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

                # Global Validation Checks

                # Duplicate Name
                if name in used_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}'.")
                else:
                    used_names.add(name)

                # Duplicate Tag
                if tag:
                    if tag in used_tags:
                        logging.warning(f"Line {line_num}: Duplicate Tag '{tag}'.")
                    else:
                        used_tags.add(tag)

                # Address Overlap
                try:
                    start_addr = parse_address_start(address, dtype)
                    reg_count = get_register_size(dtype, address)

                    for i in range(reg_count):
                        curr_addr = start_addr + i
                        if curr_addr in address_map:
                            # Check overlap validity
                            existing_vars = address_map[curr_addr]
                            for existing_name, existing_type in existing_vars:
                                if existing_type.upper() == 'BITS' and dtype.upper() == 'BITS':
                                    continue # BITS overlap is allowed
                                else:
                                    logging.warning(f"Line {line_num}: Address overlap detected at {curr_addr}. Variable '{name}' overlaps with '{existing_name}'.")

                            address_map[curr_addr].append((name, dtype))
                        else:
                            address_map[curr_addr] = [(name, dtype)]

                except ValueError:
                    logging.warning(f"Line {line_num}: Could not parse address for overlap check.")

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
