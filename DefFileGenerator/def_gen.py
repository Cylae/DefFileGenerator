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
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    rows = [
        ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', ''],
        ['String Variable', 'string_tag', 'Holding Register', '30010_10', 'String', '', '', '', '4', ''],
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'Bits', '', '', '', '4', ''],
        ['Str20 Variable', 'str20_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
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
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype, re.IGNORECASE):
        return True

    # STR<n> types
    if re.match(r'^STR\d+$', dtype, re.IGNORECASE):
        return True

    return False

def calculate_register_count(dtype, address):
    """Calculates the number of registers used by a type."""
    dtype_upper = dtype.upper()

    # Handle STR<n> by converting to STRING length
    str_match = re.match(r'^STR(\d+)$', dtype_upper)
    if str_match:
        length = int(str_match.group(1))
        return math.ceil(length / 2)

    if dtype_upper == 'STRING':
        # Parse length from Address_Length
        match = re.match(r'^\d+_(\d+)$', address)
        if match:
            length = int(match.group(1))
            return math.ceil(length / 2)
        return 0 # Should be validated before
    elif dtype_upper == 'BITS':
        return 1
    elif dtype_upper in ['U16', 'I16', 'U8', 'I8']: # U8/I8 usually take 1 register in modbus context unless packed
        return 1
    elif dtype_upper in ['U32', 'I32', 'F32', 'IP']:
        return 2
    elif dtype_upper == 'MAC':
        return 3
    elif dtype_upper in ['U64', 'I64', 'F64']:
        return 4
    elif dtype_upper == 'IPV6':
        return 8

    # Handle suffixed types
    if re.match(r'^[UI](8|16)(_(W|B|WB))?$', dtype_upper):
        return 1
    if re.match(r'^[UI]32(_(W|B|WB))?$', dtype_upper):
        return 2
    if re.match(r'^[UI]64(_(W|B|WB))?$', dtype_upper):
        return 4

    return 1 # Default fallback

def validate_address(address, dtype):
    """Validates the address format based on type."""
    dtype_upper = dtype.upper()

    if dtype_upper == 'STRING':
        # Expect Address_Length (e.g., 30000_30)
        return re.match(r'^\d+_\d+$', address) is not None
    elif re.match(r'^STR\d+$', dtype_upper):
         # Expect integer address or Address_Length (though length is in type)
         # We accept just address for STR<n> as we will format it later
         return re.match(r'^\d+(_\d+)?$', address) is not None
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

    # Overlap detection registry: {register_address: [list of variable names]}
    used_registers = {}

    try:
        # Open output file or stdout
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Detect delimiter
            try:
                dialect = csv.Sniffer().sniff(csvfile.read(1024), delimiters=";,")
                csvfile.seek(0)
            except csv.Error:
                # Default to comma if detection fails
                csvfile.seek(0)
                dialect = 'excel' # defaults to comma

            reader = csv.DictReader(csvfile, dialect=dialect)

            # Normalize headers to remove whitespace and handle case insensitivity
            if reader.fieldnames:
                reader.fieldnames = [name.strip() for name in reader.fieldnames]
                # Map common variations to standard names if needed (not strictly required if user follows template)
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
                scale_factor = row.get('ScaleFactor', '').strip()

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

                # Handle STR<n> conversion
                str_match = re.match(r'^STR(\d+)$', dtype, re.IGNORECASE)
                if str_match:
                    str_len = str_match.group(1)
                    dtype = 'STRING' # Normalize to STRING
                    # Update address if it doesn't have length
                    if '_' not in address:
                        address = f"{address}_{str_len}"
                    elif not address.endswith(f"_{str_len}"):
                        logging.warning(f"Line {line_num}: STR{str_len} type specified but address '{address}' has different length suffix. Using address as is.")

                # Overlap Detection
                try:
                    # Parse base address
                    if '_' in address:
                         base_addr = int(address.split('_')[0])
                    else:
                         base_addr = int(address)

                    reg_count = calculate_register_count(dtype if not str_match else f"STR{str_len}", address)

                    for r in range(base_addr, base_addr + reg_count):
                        if r in used_registers:
                            # Check for allowed overlaps (BITS with BITS)
                            # We assume if existing is BITS and current is BITS, it is allowed.
                            # We need to know the type of the existing registration.
                            # Since we just store names, we might want to store type too.
                            # For now, let's just warn if overlap.

                            # Simple check: BITS overlap allowed?
                            # To do this correctly, we need to track Type in used_registers.
                            pass # We will handle inside loop

                        if r not in used_registers:
                            used_registers[r] = []

                        used_registers[r].append({'name': name, 'type': dtype})

                except ValueError:
                    logging.warning(f"Line {line_num}: Error parsing address '{address}' for overlap check.")

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
                    base_factor = 1.0
                else:
                    try:
                        base_factor = float(factor)
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using 1.0.")
                        base_factor = 1.0

                if scale_factor:
                    try:
                        sf = float(scale_factor)
                        base_factor = base_factor * (10 ** sf)
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid ScaleFactor '{scale_factor}'. Ignoring.")

                coef_a = "{:.6f}".format(base_factor)

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

            # Check overlaps after processing
            for addr, users in used_registers.items():
                if len(users) > 1:
                    # Check if all users are BITS
                    all_bits = all(u['type'] == 'BITS' for u in users)
                    if not all_bits:
                        names = [u['name'] for u in users]
                        logging.warning(f"Overlap detected at register {addr}: {', '.join(names)}")


        if args.output:
            outfile.close()
            logging.info(f"Definition file generated at {args.output}")

    except FileNotFoundError:
        logging.error(f"File '{args.input_file}' not found.")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        # print stack trace for debugging
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
