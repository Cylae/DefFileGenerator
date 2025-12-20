import csv
import argparse
import sys
import os

def parse_arguments():
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM definition file from simplified CSV.')
    parser.add_argument('input_file', help='Path to the simplified input CSV file')
    parser.add_argument('output_file', help='Path to the output definition CSV file')
    parser.add_argument('--protocol', default='modbus', help='Protocol (default: modbus)')
    parser.add_argument('--category', default='', help='Category (e.g., Inverter)')
    parser.add_argument('--manufacturer', default='', help='Manufacturer (e.g., HUAWEI)')
    parser.add_argument('--model', default='', help='Model (e.g., V4)')
    parser.add_argument('--forced-write', default='0', help='Forced Write Code (0 or 1, default: 0)')
    return parser.parse_args()

def map_register_type(reg_type_str):
    reg_type_str = reg_type_str.strip().lower()
    if reg_type_str in ['coil', '1']:
        return 1
    elif reg_type_str in ['discrete input', 'discreteinput', '2']:
        return 2
    elif reg_type_str in ['holding register', 'holdingregister', '3']:
        return 3
    elif reg_type_str in ['input register', 'inputregister', '4']:
        return 4
    else:
        try:
            return int(reg_type_str)
        except ValueError:
            print(f"Warning: Unknown RegisterType '{reg_type_str}', defaulting to 3 (Holding Register)")
            return 3

def main():
    args = parse_arguments()

    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' not found.")
        sys.exit(1)

    # Output header
    header_row_1 = [
        args.protocol,
        args.category,
        args.manufacturer,
        args.model,
        args.forced_write
    ]

    # Expected input columns (normalized)
    expected_cols = {
        'name': 'Name',
        'tag': 'Tag',
        'registertype': 'RegisterType',
        'address': 'Address',
        'type': 'Type',
        'factor': 'Factor',
        'offset': 'Offset',
        'unit': 'Unit',
        'action': 'Action'
    }

    # Output columns
    # Index ; Info1 ; Info2 ; Info3 ; Info4 ; Name ; Tag ; CoefA ; CoefB ; Unit ; Action

    try:
        with open(args.input_file, mode='r', encoding='utf-8-sig') as infile:
            # Detect delimiter
            sample = infile.read(1024)
            infile.seek(0)
            sniffer = csv.Sniffer()
            try:
                dialect = sniffer.sniff(sample)
            except csv.Error:
                dialect = csv.excel # Default fallback

            reader = csv.DictReader(infile, dialect=dialect)

            # Normalize header map
            header_map = {}
            for field in reader.fieldnames:
                norm_field = field.strip().lower()
                header_map[norm_field] = field

            # Check for required columns
            required = ['name', 'registertype', 'address', 'type', 'action']
            missing = []
            for req in required:
                if req not in header_map:
                    missing.append(expected_cols[req])

            if missing:
                print(f"Error: Missing required columns in input CSV: {', '.join(missing)}")
                # Allow user to proceed if they mapped it differently? For now, fail.
                # Actually, let's try to be robust.
                # If 'RegisterType' is missing but 'Info1' exists, we could support that?
                # But task says "The input CSV ... expects columns ...".
                sys.exit(1)

            rows = []
            index = 1
            for row in reader:
                # Get values using normalized keys
                def get_val(key, default=''):
                    real_key = header_map.get(key)
                    if real_key and real_key in row:
                        val = row[real_key]
                        return val.strip() if val else default
                    return default

                name = get_val('name')
                tag = get_val('tag')
                reg_type_raw = get_val('registertype')
                address = get_val('address')
                var_type = get_val('type')
                factor = get_val('factor', '1.0')
                offset = get_val('offset', '0.0')
                unit = get_val('unit')
                action = get_val('action', '4')

                # Processing
                info1 = map_register_type(reg_type_raw)
                info2 = address
                info3 = var_type
                info4 = '' # Empty

                # Format Factor and Offset to ensure they look like floats
                try:
                    coef_a = "{:.6f}".format(float(factor))
                except ValueError:
                    coef_a = "1.000000"

                try:
                    coef_b = "{:.6f}".format(float(offset))
                except ValueError:
                    coef_b = "0.000000"

                output_row = [
                    str(index),
                    str(info1),
                    str(info2),
                    str(info3),
                    str(info4),
                    name,
                    tag,
                    coef_a,
                    coef_b,
                    unit,
                    str(action)
                ]
                rows.append(output_row)
                index += 1

        # Write output
        with open(args.output_file, mode='w', encoding='utf-8', newline='') as outfile:
            writer = csv.writer(outfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)

            # Write Header Row 1
            # Note: The manual shows specific fields separated by semicolons.
            # csv.writer will handle quoting if necessary.
            # However, looking at HUAWEI_V4.csv:
            # modbusRTU;Inverter;HUAWEI;V4;;;;;;;
            # It seems it has empty fields to match the column count of data rows?
            # Data rows have 11 columns. Header row has 5 relevant fields.
            # Let's pad the header row to 11 columns to match `HUAWEI_V4.csv` format if that's what is expected.
            # HUAWEI_V4.csv first line: `modbusRTU;Inverter;HUAWEI;V4;;;;;;;`
            # That is 4 items, followed by 7 empty items = 11 items total.

            padded_header = header_row_1 + [''] * (11 - len(header_row_1))
            writer.writerow(padded_header)

            # Write Data Rows
            writer.writerows(rows)

        print(f"Successfully generated '{args.output_file}' with {len(rows)} variables.")

    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
