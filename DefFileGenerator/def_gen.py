#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math

class Generator:
    def __init__(self):
        self.seen_names = {}
        self.seen_tags = {}
        self.used_addresses = []

        # RegisterType mapping to Info1
        self.register_type_map = {
            'coil': '1',
            'discrete input': '2',
            'holding register': '3',
            'input register': '4'
        }

        # Allowed Action codes
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    def validate_type(self, dtype):
        """Validates the data type."""
        dtype_upper = dtype.upper()
        # Base types
        base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
        if dtype_upper in base_types:
            return True

        # Integer types with optional suffixes
        # U8, U16, U32, U64, I8, I16, I32, I64
        # Suffixes: _W, _B, _WB
        if re.match(r'^[UI](8|16|32|64)(_(W|B|WB))?$', dtype_upper):
            return True

        # STR<n> syntax (e.g., STR20)
        if re.match(r'^STR\d+$', dtype_upper):
            return True

        return False

    def validate_address(self, address, dtype):
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

    def get_register_count(self, dtype, address):
        """Calculates the number of registers used by the type."""
        dtype_upper = dtype.upper()

        if dtype_upper in ['U16', 'I16', 'BITS'] or re.match(r'^[UI]16(_(W|B|WB))?$', dtype_upper) or re.match(r'^[UI]8(_(W|B|WB))?$', dtype_upper):
            return 1
        elif dtype_upper in ['U32', 'I32', 'F32', 'IP'] or re.match(r'^[UI]32(_(W|B|WB))?$', dtype_upper):
            return 2
        elif dtype_upper in ['U64', 'I64', 'F64'] or re.match(r'^[UI]64(_(W|B|WB))?$', dtype_upper):
            return 4
        elif dtype_upper == 'MAC':
            return 3
        elif dtype_upper == 'IPV6':
            return 8
        elif dtype_upper == 'STRING':
            # Parse length from address: Address_Length
            try:
                length = int(address.split('_')[1])
                return math.ceil(length / 2)
            except (IndexError, ValueError):
                return 0
        return 1 # Default

    def process_rows(self, reader):
        """
        Processes rows from a CSV reader (iterator of dicts).
        Returns a list of processed row dictionaries.
        """
        processed_rows = []

        # We need to handle case-insensitive headers if not already handled by the caller.
        # However, the DictReader in main() sets fieldnames.
        # Assuming `reader` yields dicts where keys are the header names.

        # Helper to get column value case-insensitively
        # We need the fieldnames to build a map if we want to be robust,
        # but reader is just an iterator here.
        # If passed from csv.DictReader, the keys are fixed.
        # Let's assume the keys are normalized or we search them.

        # To avoid re-scanning keys for every row, we can build a map from the first row?
        # No, let's assume the caller normalized the keys or we search.
        # The original code built a `header_map`.

        # But `reader` here might be a list of dicts for testing.
        # Let's inspect the first row if available, or just use a helper function.

        header_map = None

        for line_num, row in enumerate(reader, start=2): # Start at 2 because header is 1
            # Skip empty rows
            if not any(row.values()):
                continue

            # Lazy init header_map based on keys of the first row
            if header_map is None:
                header_map = {k.lower(): k for k in row.keys()}

            def get_col(r, col_name):
                if col_name in r:
                    return r[col_name]
                if col_name.lower() in header_map:
                    return r[header_map[col_name.lower()]]
                return ''

            # Extract values
            name = get_col(row, 'Name').strip()
            tag = get_col(row, 'Tag').strip()
            reg_type_str = get_col(row, 'RegisterType').strip()
            address = get_col(row, 'Address').strip()
            dtype = get_col(row, 'Type').strip()
            factor = get_col(row, 'Factor').strip()
            offset = get_col(row, 'Offset').strip()
            unit = get_col(row, 'Unit').strip()
            action = get_col(row, 'Action').strip()
            scale_factor_str = get_col(row, 'ScaleFactor').strip()

            if not name and not address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Validation: Type
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                continue

            # Handle STR<n> conversion
            dtype_upper = dtype.upper()
            if re.match(r'^STR\d+$', dtype_upper):
                try:
                    length = int(dtype_upper[3:])
                    dtype = 'STRING'
                    # Update address to Address_Length if not already formatted
                    if '_' not in address:
                        address = f"{address}_{length}"
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid STR format '{dtype_upper}'. Skipping row.")
                    continue

            # Validation: Address format based on Type
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            # Global Check: Duplicate Name
            if name:
                if name in self.seen_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {self.seen_names[name]}.")
                else:
                    self.seen_names[name] = line_num

            # Global Check: Duplicate Tag
            if tag:
                if tag in self.seen_tags:
                    logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous occurrence at line {self.seen_tags[tag]}.")
                else:
                    self.seen_tags[tag] = line_num

            # Address Overlap Calculation
            try:
                # Parse start address
                parts = address.split('_')
                start_addr = int(parts[0])

                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1

                current_range = (start_addr, end_addr, line_num, name, dtype.upper())

                # Check overlap against used_addresses
                is_bits = (dtype.upper() == 'BITS')

                for used_start, used_end, used_line, used_name, used_type in self.used_addresses:
                    # Check if ranges overlap
                    if max(start_addr, used_start) <= min(end_addr, used_end):
                        # Overlap detected
                        # Allowed only if both are BITS and same start address
                        is_overlap_allowed = is_bits and (used_type == 'BITS')

                        if not is_overlap_allowed:
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). Overlaps with '{used_name}' (Line {used_line}, Addr: {used_start}-{used_end}).")

                self.used_addresses.append(current_range)

            except ValueError:
                logging.warning(f"Line {line_num}: Could not calculate register range for address '{address}'.")

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

            # Action
            if not action:
                action = '1' # Default per spec
            elif action not in self.allowed_actions:
                logging.warning(f"Line {line_num}: Invalid Action '{action}'. Defaulting to '1'.")
                action = '1'

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

    if args.template:
        generate_template(args.output)
        sys.exit(0)

    if not args.input_file:
        parser.error("the following arguments are required: input_file")

    if not args.manufacturer or not args.model:
         parser.error("the following arguments are required: --manufacturer, --model")

    try:
        # Use utf-8-sig to handle potential BOM from Excel-saved CSVs
        with open(args.input_file, mode='r', encoding='utf-8-sig') as csvfile:
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

            # Normalize headers check (optional, but handled in process_rows slightly)
            if not reader.fieldnames:
                 logging.error("Input CSV is empty or missing headers.")
                 sys.exit(1)

            # Normalize headers by stripping whitespace
            reader.fieldnames = [name.strip() for name in reader.fieldnames]

            # Check for required columns
            required_columns = ['Name', 'RegisterType', 'Address', 'Type']
            # We need to map actual headers to required ones to check existence?
            # Or just assume standard names. The original code did strict check on normalized or raw headers.
            # Let's trust the DictReader fieldnames for the check.
            # We need to account for potential case differences in the check too.

            header_map = {h.lower(): h for h in reader.fieldnames}
            missing_columns = [col for col in required_columns if col.lower() not in header_map]

            if missing_columns:
                logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)}")
                sys.exit(1)

            # Initialize Generator and process
            generator = Generator()
            processed_rows = generator.process_rows(reader)

        # Write Output
        if args.output:
            outfile = open(args.output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

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
        for row in processed_rows:
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
            index += 1

        if args.output:
            outfile.close()
            logging.info(f"Definition file generated at {args.output}")

    except FileNotFoundError:
        logging.error(f"File '{args.input_file}' not found.")
        sys.exit(1)
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        # import traceback
        # traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
