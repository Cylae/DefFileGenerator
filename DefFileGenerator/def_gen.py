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

def calculate_registers(address_str, dtype):
    """
    Calculates the start address and number of registers consumed.
    Returns (start_address, num_registers).
    """
    dtype_upper = dtype.upper()
    start_addr = 0
    num_regs = 0

    if dtype_upper == 'STRING':
        # Format: Address_Length
        match = re.match(r'^(\d+)_(\d+)$', address_str)
        if match:
            start_addr = int(match.group(1))
            length = int(match.group(2))
            # 1 register = 2 bytes
            num_regs = math.ceil(length / 2)
    elif dtype_upper == 'BITS':
        # Format: Address_StartBit_NbBits
        match = re.match(r'^(\d+)_(\d+)_(\d+)$', address_str)
        if match:
            start_addr = int(match.group(1))
            num_regs = 0 # Bits are within a register, handled specially in overlap check
    elif dtype_upper == 'IP':
        start_addr = int(address_str)
        num_regs = 2
    elif dtype_upper == 'IPV6':
        start_addr = int(address_str)
        num_regs = 8
    elif dtype_upper == 'MAC':
        start_addr = int(address_str)
        num_regs = 3
    else:
        # Standard numeric types
        match = re.match(r'^[UI](8|16|32|64)', dtype_upper)
        if match:
            bits = int(match.group(1))
            num_regs = max(1, bits // 16)
        elif dtype_upper in ['F32']:
            num_regs = 2
        elif dtype_upper in ['F64']:
            num_regs = 4

        # Parse simple address
        if re.match(r'^\d+$', address_str):
            start_addr = int(address_str)

    return start_addr, num_regs

def validate_uniqueness(rows):
    """Checks for duplicate Names and Tags."""
    names = {}
    tags = {}

    for i, row in enumerate(rows):
        name = row['Name']
        tag = row['Tag']

        if name:
            if name in names:
                logging.warning(f"Duplicate Name '{name}' found in lines {names[name]} and {i + 1}.")
            names[name] = i + 1

        if tag:
            if tag in tags:
                logging.warning(f"Duplicate Tag '{tag}' found in lines {tags[tag]} and {i + 1}.")
            tags[tag] = i + 1

def check_overlaps(rows):
    """Checks for address overlaps within the same RegisterType (Info1)."""
    # map: Info1 -> { address: [ (line_num, type, name) ] }
    register_maps = {}

    for i, row in enumerate(rows):
        info1 = row['Info1']
        info2 = row['Info2'] # Address string
        dtype = row['Info3']
        name = row['Name']
        line_num = i + 1 # Approximate line number (index + 1)

        if info1 not in register_maps:
            register_maps[info1] = {}

        start_addr, num_regs = calculate_registers(info2, dtype)

        # If num_regs is 0 (BITS), we treat it as checking the base address
        # but we need to handle it carefully.
        # BITS usually share the same register.
        # Check logic:
        # If address X is used by BITS, it can be used by other BITS.
        # If address X is used by BITS, it should NOT be used by U16/U32 etc.
        # If address X is used by U16, it should NOT be used by other types (overlap).

        target_indices = []
        if num_regs == 0 and dtype.upper() == 'BITS':
             target_indices = [start_addr]
        elif num_regs > 0:
             target_indices = range(start_addr, start_addr + num_regs)
        else:
             # Should not happen if validation passed, but just in case
             continue

        for addr in target_indices:
            if addr not in register_maps[info1]:
                register_maps[info1][addr] = []

            # Check for conflict with existing entries
            for entry in register_maps[info1][addr]:
                existing_line, existing_type, existing_name = entry

                # Conflict logic:
                # If both are BITS, it is allowed (same register, different bits) - Warning if exact same?
                # If one is BITS and other is not, WARN.
                # If neither is BITS, WARN (direct overlap).

                is_bits = (dtype.upper() == 'BITS')
                existing_is_bits = (existing_type.upper() == 'BITS')

                if is_bits and existing_is_bits:
                    # Both BITS, allowed.
                    pass
                else:
                    logging.warning(f"Address Overlap detected at Address {addr} (Info1={info1}): "
                                    f"'{name}' (Line {line_num}, Type {dtype}) conflicts with "
                                    f"'{existing_name}' (Line {existing_line}, Type {existing_type}).")

            register_maps[info1][addr].append((line_num, dtype, name))

def main():
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM Modbus definition file from simplified CSV.')
    parser.add_argument('input_file', nargs='?', help='Path to the simplified CSV input file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
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

    try:
        # Open output file or stdout
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile)

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

            # Process all rows first
            processed_rows = []

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

                # Validation: Type
                if not validate_type(dtype):
                    logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                    continue

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

                # Store processed row
                row_data = {
                    'Index': str(index),
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
                processed_rows.append(row_data)
                index += 1

            # Run robustness checks
            validate_uniqueness(processed_rows)
            check_overlaps(processed_rows)

            # Write output
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

            for row in processed_rows:
                data_row = [
                    row['Index'],
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
