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
        ['String Variable', 'string_tag', 'Holding Register', '30010_10', 'STRING', '', '', '', '4'],
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'BITS', '', '', '', '4'],
        ['Simplified String', 'simple_str_tag', 'Holding Register', '30050', 'STR10', '', '', '', '4']
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
    if dtype in base_types:
        return True

    # Integer types with optional suffixes
    # U8, U16, U32, U64, I8, I16, I32, I64
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype):
        return True

    # Float types with optional suffixes
    # F32, F64 with suffixes
    if re.match(r'^F(32|64)(_(W|B|WB))?$', dtype):
        return True

    return False

def validate_address(address, dtype):
    """Validates the address format based on type."""
    if dtype == 'STRING':
        # Expect Address_Length (e.g., 30000_30)
        return re.match(r'^\d+_\d+$', address) is not None
    elif dtype == 'BITS':
        # Expect Address_StartBit_NbBits (e.g., 30000_0_1)
        return re.match(r'^\d+_\d+_\d+$', address) is not None
    else:
        # Expect integer address
        return re.match(r'^\d+$', address) is not None

def get_register_count(dtype, address):
    """Calculates the number of registers used by the type."""
    if dtype == 'STRING':
        # address format: Start_Length
        try:
            length = int(address.split('_')[1])
            return math.ceil(length / 2)
        except IndexError:
            return 0
    elif dtype == 'BITS':
        return 1
    elif dtype == 'IP':
        return 2
    elif dtype == 'MAC':
        return 3
    elif dtype == 'IPV6':
        return 8
    elif dtype in ['U8', 'I8', 'U16', 'I16']:
        return 1
    elif dtype in ['U32', 'I32', 'F32']:
        return 2
    elif dtype in ['U64', 'I64', 'F64']:
        return 4

    # Handle suffix cases (e.g. U32_B, F32_W)
    if re.match(r'^[UIF]32', dtype): return 2
    if re.match(r'^[UIF]64', dtype): return 4
    if re.match(r'^[UI](8|16)', dtype): return 1

    return 1

def parse_start_address(address):
    """Extracts the integer start address."""
    try:
        return int(address.split('_')[0])
    except ValueError:
        return -1

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

    # State for validation
    used_ranges = [] # List of (start, end, line_num)
    seen_names = {} # name -> line_num
    seen_tags = {} # tag -> line_num

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

            # Prepare output header row
            header_row = [
                args.protocol,
                args.category,
                args.manufacturer,
                args.model,
                args.forced_write,
                '', '', '', '', '', ''
            ]

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            index = 1
            for line_num, row in enumerate(reader, start=2):
                # Skip empty rows
                if not any(row.values()):
                    continue

                # Extract values
                name = row.get('Name', '').strip()
                tag = row.get('Tag', '').strip()
                reg_type_str = row.get('RegisterType', '').strip()
                address = row.get('Address', '').strip()
                dtype = row.get('Type', '').strip().upper() # Normalize to uppercase
                factor = row.get('Factor', '').strip()
                offset = row.get('Offset', '').strip()
                unit = row.get('Unit', '').strip()
                action = row.get('Action', '').strip()

                if not name and not address:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                    continue

                # --- Transformation: Handle STR<n> ---
                str_match = re.match(r'^STR(\d+)$', dtype)
                if str_match:
                    length = str_match.group(1)
                    # If address doesn't already have length suffix
                    if '_' not in address:
                        address = f"{address}_{length}"
                    dtype = 'STRING'

                # --- Validation: Type ---
                if not validate_type(dtype):
                    logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                    continue

                # --- Validation: Address format ---
                if not validate_address(address, dtype):
                    logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                    continue

                # --- Validation: Overlap ---
                start_addr = parse_start_address(address)
                count = get_register_count(dtype, address)

                if start_addr != -1 and count > 0:
                    end_addr = start_addr + count - 1
                    overlap = False
                    for r_start, r_end, r_line in used_ranges:
                        if start_addr <= r_end and end_addr >= r_start:
                            logging.warning(f"Line {line_num}: Address range {start_addr}-{end_addr} ('{name}') overlaps with Line {r_line} ({r_start}-{r_end}).")
                            overlap = True
                            break
                    if not overlap:
                        used_ranges.append((start_addr, end_addr, line_num))

                # --- Validation: Duplicates ---
                if name:
                    if name in seen_names:
                        logging.warning(f"Line {line_num}: Duplicate Name '{name}' found. Previous occurrence at Line {seen_names[name]}.")
                    else:
                        seen_names[name] = line_num

                if tag:
                    if tag in seen_tags:
                        logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' found. Previous occurrence at Line {seen_tags[tag]}.")
                    else:
                        seen_tags[tag] = line_num


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

                info2 = address
                info3 = dtype
                info4 = ''

                # CoefA
                if not factor:
                    coef_a = "1.000000"
                else:
                    try:
                        coef_a = "{:.6f}".format(float(factor))
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using as is.")
                        coef_a = factor

                # CoefB
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
                    action = '1'
                elif action not in allowed_actions:
                    logging.warning(f"Line {line_num}: Invalid Action '{action}'. Defaulting to '1'.")
                    action = '1'

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
