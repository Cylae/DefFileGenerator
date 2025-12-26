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

def normalize_type(dtype):
    """Normalizes the data type to match the expected format (e.g., U16, String)."""
    dtype_upper = dtype.upper()

    # Base types mapping (case-insensitive to standard format)
    # The manual lists: U8, U16, U32, U64, I8, I16, I32, I64
    # F32, F64
    # String
    # Bits
    # IP, IPV6, MAC
    # Suffixes: _W, _B, _WB

    # Check for String and Bits first as they are special
    if dtype_upper == 'STRING':
        return 'STRING' # Using uppercase as seen in HUAWEI_V4.csv
    if dtype_upper == 'BITS':
        return 'BITS' # Assuming uppercase based on consistency, though manual uses Title case. Let's stick to what input provides if valid.
    if dtype_upper == 'IP':
        return 'IP'
    if dtype_upper == 'IPV6':
        return 'IPV6'
    if dtype_upper == 'MAC':
        return 'MAC'

    # Numeric types
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype_upper):
        return dtype_upper

    if dtype_upper in ['F32', 'F64']:
        return dtype_upper

    return None

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

    try:
        # Open output file or stdout
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig') as csvfile:
            # Detect dialect (delimiter)
            sample = csvfile.read(1024)
            csvfile.seek(0)
            sniffer = csv.Sniffer()
            try:
                dialect = sniffer.sniff(sample)
            except csv.Error:
                dialect = csv.excel # Fallback to default excel dialect

            reader = csv.DictReader(csvfile, dialect=dialect)

            # Normalize headers to remove whitespace and handle case insensitivity
            if reader.fieldnames:
                original_headers = reader.fieldnames
                normalized_headers = {h.strip().lower(): h for h in original_headers}

                # Helper to get value case-insensitively
                def get_val(row, col_name):
                    key = normalized_headers.get(col_name.lower())
                    if key:
                        return row.get(key, '').strip()
                    return ''
            else:
                logging.error("Input CSV is empty or missing headers.")
                sys.exit(1)

            # Check for required columns
            required_columns_lower = ['name', 'registertype', 'address', 'type']
            present_columns_lower = [h.strip().lower() for h in reader.fieldnames]
            missing_columns = [col for col in required_columns_lower if col not in present_columns_lower]

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
                name = get_val(row, 'Name')
                tag = get_val(row, 'Tag')
                reg_type_str = get_val(row, 'RegisterType')
                address = get_val(row, 'Address')
                dtype_raw = get_val(row, 'Type')
                factor = get_val(row, 'Factor')
                offset = get_val(row, 'Offset')
                unit = get_val(row, 'Unit')
                action = get_val(row, 'Action')

                if not name and not address:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                    continue

                # Normalize and Validate Type
                dtype = normalize_type(dtype_raw)
                if not dtype:
                    logging.warning(f"Line {line_num}: Invalid Type '{dtype_raw}'. Skipping row.")
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
                        val = float(factor)
                        coef_a = "{:.6f}".format(val)
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using as is.")
                        coef_a = factor

                # CoefB from Offset
                if not offset:
                    coef_b = "0.000000"
                else:
                    try:
                        val = float(offset)
                        coef_b = "{:.6f}".format(val)
                    except ValueError:
                        logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Using as is.")
                        coef_b = offset

                # Action
                if not action:
                    action = '1' # Default per spec, but actually Huawei file uses '4' for regular vars and '1' for params?
                    # The manual says '1' is parameter, '4' is instantaneous value.
                    # The prompt says "If Action is empty or missing, it defaults to '1'".
                    # Let's stick to '1' as default if not specified, but usually simplified CSV should specify it.
                    # Wait, looking at HUAWEI_V4.csv, many are '4' (instantaneous).
                    # If the input CSV has action, we use it.
                    pass

                if action and action not in allowed_actions:
                    logging.warning(f"Line {line_num}: Invalid Action '{action}'. Defaulting to '1'.")
                    action = '1'

                if not action:
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
