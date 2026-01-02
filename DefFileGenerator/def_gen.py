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
        ['Scaled Variable', 'scaled_tag', 'Holding Register', '30030', 'I16', '1', '0', 'V', '4', 'ScaleFactorVar']
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
    dtype_upper = dtype.upper()
    # Base types
    base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
    if dtype_upper in base_types:
        return True

    # STR<n> type
    if re.match(r'^STR\d+$', dtype_upper):
        return True

    # Integer types with optional suffixes
    # U8, U16, U32, U64, I8, I16, I32, I64
    # Suffixes: _W, _B, _WB
    # Regex: ^[UI](8|16|32|64)(_(W|B|WB))?$
    if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype_upper):
        return True

    return False

def validate_address(address, dtype):
    """Validates the address format based on type."""
    dtype_upper = dtype.upper()
    if dtype_upper == 'STRING' or re.match(r'^STR\d+$', dtype_upper):
        # Expect Address_Length (e.g., 30000_30) OR just Address if Type is STR<n>
        if re.match(r'^STR\d+$', dtype_upper) and re.match(r'^\d+$', address):
             return True
        return re.match(r'^\d+_\d+$', address) is not None
    elif dtype_upper == 'BITS':
        # Expect Address_StartBit_NbBits (e.g., 30000_0_1)
        return re.match(r'^\d+_\d+_\d+$', address) is not None
    else:
        # Expect integer address
        return re.match(r'^\d+$', address) is not None

def get_register_count(dtype, address):
    """Calculates the number of registers used by a type."""
    dtype_upper = dtype.upper()

    if dtype_upper == 'BITS':
        return 1

    # Check STR<n> first
    str_match = re.match(r'^STR(\d+)$', dtype_upper)
    if str_match:
        length = int(str_match.group(1))
        return math.ceil(length / 2)

    if dtype_upper == 'STRING':
        # Length is in address: Addr_Length
        try:
            length = int(address.split('_')[1])
            return math.ceil(length / 2)
        except IndexError:
            return 0 # Should be caught by validation

    if dtype_upper in ['U16', 'I16']:
        return 1
    if dtype_upper in ['U32', 'I32', 'F32', 'IP']:
        return 2
    if dtype_upper in ['U64', 'I64', 'F64']:
        return 4
    if dtype_upper == 'MAC':
        return 3
    if dtype_upper == 'IPV6':
        return 8

    # Handle suffixed types
    if re.match(r'^[UI]16', dtype_upper): return 1
    if re.match(r'^[UI]32', dtype_upper): return 2
    if re.match(r'^[UI]64', dtype_upper): return 4

    # U8/I8 usually map to bytes within a register, but we reserve the whole register to avoid overlap issues unless handled specifically
    # However, Modbus usually addresses 16-bit registers.
    # If the user specifies byte access (e.g. 30000_1 for U8), it still effectively uses register 30000.
    if re.match(r'^[UI]8', dtype_upper): return 1

    return 1 # Default fallback

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

    # Track used registers for overlap detection: list of (start, end, type, line_num)
    used_ranges = []

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
                # Read a snippet to detect dialect
                sample = csvfile.read(2048)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(sample)
                delimiter = dialect.delimiter
            except csv.Error:
                # Fallback to comma if detection fails
                csvfile.seek(0)
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

                # Handle STR<n>
                str_match = re.match(r'^STR(\d+)$', dtype.upper())
                if str_match:
                    length = str_match.group(1)
                    dtype = 'STRING'
                    # Update address if it doesn't have length
                    if '_' not in address:
                        address = f"{address}_{length}"

                # Address Overlap Detection
                # Extract base address
                try:
                    if '_' in address:
                        base_addr = int(address.split('_')[0])
                    else:
                        base_addr = int(address)

                    reg_count = get_register_count(dtype, address)
                    if reg_count > 0:
                        start_addr = base_addr
                        end_addr = base_addr + reg_count - 1

                        # Check against used ranges
                        for u_start, u_end, u_type, u_line in used_ranges:
                            # Check overlap: (StartA <= EndB) and (EndA >= StartB)
                            if (start_addr <= u_end) and (end_addr >= u_start):
                                # If both are BITS, overlap is allowed on the same register
                                if dtype.upper() == 'BITS' and u_type == 'BITS' and start_addr == u_start:
                                    continue
                                else:
                                    logging.warning(f"Line {line_num}: Address overlap detected! {name} ({start_addr}-{end_addr}) overlaps with previous entry at line {u_line} ({u_start}-{u_end}).")

                        used_ranges.append((start_addr, end_addr, dtype.upper(), line_num))

                except ValueError:
                    # Address parsing failed, skipping overlap check (validation should have caught it)
                    pass

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

                # Info4: ScaleFactor
                info4 = scale_factor

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
