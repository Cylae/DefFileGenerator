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
    # Base types without suffixes
    base_types_no_suffix = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC']
    if dtype.upper() in base_types_no_suffix:
        return True

    # Types that allow suffixes
    # U8-U64, I8-I64, F32, F64
    # Suffixes: _W, _B, _WB
    # Regex: ^([UI](8|16|32|64)|F(32|64))(_(W|B|WB))?$
    if re.match(r'^([UI](8|16|32|64)|F(32|64))(_(W|B|WB))?$', dtype, re.IGNORECASE):
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
        # Expect integer address for other types
        # IP, IPV6, MAC are just start address
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
        'input register': '4',
        'coils': '1',
        'discrete inputs': '2',
        'holding registers': '3',
        'input registers': '4'
    }

    # Allowed Action codes
    allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    try:
        # Determine delimiter using Sniffer
        with open(args.input_file, mode='r', encoding='utf-8-sig') as f:
            sample = f.read(2048)
            try:
                # Try to sniff with preferred delimiters
                dialect = csv.Sniffer().sniff(sample, delimiters=';,')
                delimiter = dialect.delimiter
            except csv.Error:
                # Fallback to comma if sniffing fails
                logging.warning("Could not determine delimiter. Defaulting to semicolon.")
                delimiter = ';'
            f.seek(0)

            # Use DictReader with inferred delimiter
            reader = csv.DictReader(f, delimiter=delimiter)

            # Normalize headers to remove whitespace and handle case-insensitivity
            if reader.fieldnames:
                # Create a map of normalized headers to actual headers
                header_map = {h.strip().lower(): h for h in reader.fieldnames}
                # Check for required columns (case-insensitive)
                required_columns_lower = ['name', 'registertype', 'address', 'type']
                missing_columns = [col for col in required_columns_lower if col not in header_map]
                if missing_columns:
                    logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)}")
                    sys.exit(1)
            else:
                logging.error("Input CSV is empty or missing headers.")
                sys.exit(1)

            # Open output file or stdout
            if args.output:
                outfile = open(args.output, 'w', newline='', encoding='utf-8')
            else:
                outfile = sys.stdout

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
                # Check for comments or empty rows
                # If using Sniffer, comments starting with # might be read as data if inside quotes or strict CSV
                # But here we iterate DictReader.
                # If a line starts with # in the file, DictReader might process it oddly if it doesn't match columns.
                # However, for robustness, we can check if the 'name' (or first column) starts with #

                # Get values using the header map to ensure we get the right column regardless of case
                def get_val(key):
                    actual_key = header_map.get(key.lower())
                    if actual_key:
                        return row.get(actual_key, '').strip()
                    return ''

                name = get_val('Name')

                # Skip comments (lines starting with # in the Name column, or just empty)
                if not name and not any(row.values()):
                    continue
                if name.startswith('#'):
                    continue

                tag = get_val('Tag')
                reg_type_str = get_val('RegisterType')
                address = get_val('Address')
                dtype = get_val('Type')
                factor = get_val('Factor')
                offset = get_val('Offset')
                unit = get_val('Unit')
                action = get_val('Action')

                if not name:
                    logging.warning(f"Line {line_num}: Skipping row with missing Name.")
                    continue

                if not address:
                    logging.warning(f"Line {line_num}: Skipping row with missing Address.")
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

                # Info2: Address
                info2 = address

                # Info3: Type (Normalize to uppercase)
                info3 = dtype.upper()

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
