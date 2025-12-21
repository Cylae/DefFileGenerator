import argparse
import csv
import sys
import os

def parse_register_type(reg_type):
    """
    Parses register type which can be an integer string ('3') or a description ('Holding Register').
    Returns the integer code as string.
    """
    reg_type = str(reg_type).strip()
    if reg_type.isdigit():
        return reg_type

    mapping = {
        'coil': '1',
        'coils': '1',
        'discrete input': '2',
        'discrete inputs': '2',
        'holding register': '3',
        'holding registers': '3',
        'input register': '4',
        'input registers': '4',
        'input': '4'
    }

    key = reg_type.lower()
    if key in mapping:
        return mapping[key]

    # Default fallback or error?
    # Assume if not found, maybe it's already a code or invalid.
    return reg_type

def main():
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM Modbus definition files.')
    parser.add_argument('input_file', help='Path to the simplified input CSV file.')
    parser.add_argument('output_file', help='Path to the output definition CSV file.')

    parser.add_argument('--protocol', default='modbusRTU', help='Protocol (default: modbusRTU)')
    parser.add_argument('--category', default='Inverter', help='Category (default: Inverter)')
    parser.add_argument('--manufacturer', required=True, help='Manufacturer name')
    parser.add_argument('--model', required=True, help='Model name')
    parser.add_argument('--forced-write', default='', help='Forced Write Code (optional)')

    args = parser.parse_args()

    # Define input columns expected
    input_fieldnames = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action']

    # Read input CSV
    rows = []
    try:
        with open(args.input_file, 'r', newline='', encoding='utf-8') as f:
            # Detect dialect or assume comma?
            # User provided example HUAWEI_V4.csv is output.
            # Simplified input is usually standard CSV (comma).
            reader = csv.DictReader(f)

            # Normalize headers (strip spaces, handle case sensitivity if needed)
            # For now assume headers match exactly.

            for row in reader:
                rows.append(row)
    except Exception as e:
        print(f"Error reading input file: {e}")
        sys.exit(1)

    # Write output CSV
    try:
        with open(args.output_file, 'w', newline='', encoding='utf-8') as f:
            # Use semicolon delimiter
            writer = csv.writer(f, delimiter=';')

            # Write Header Row
            # Protocole;Category;Manufacturer;Model;ForcedWriteCode;;;;;;
            # Note: 11 columns total to match data rows structure usually?
            # HUAWEI_V4.csv has: modbusRTU;Inverter;HUAWEI;V4;;;;;;; (11 cols)
            header_row = [
                args.protocol,
                args.category,
                args.manufacturer,
                args.model,
                args.forced_write,
                '', '', '', '', '', '' # Empty columns to fill 11
            ]
            writer.writerow(header_row)

            # Process and write data rows
            # Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action

            index = 1
            for row in rows:
                # Get values with defaults
                name = row.get('Name', '')
                tag = row.get('Tag', '')
                reg_type_raw = row.get('RegisterType', '3')
                address = row.get('Address', '')
                dtype = row.get('Type', '')
                factor = row.get('Factor', '1.000000')
                offset = row.get('Offset', '0.000000')
                unit = row.get('Unit', '')
                action = row.get('Action', '')
                if not action:
                    action = '1'

                # Process values
                info1 = parse_register_type(reg_type_raw)
                info2 = address # Supports Address_Length
                info3 = dtype # e.g. U32, STRING
                info4 = '' # Always empty per manual/examples

                # Format Factor/Offset to float string with precision if possible?
                # HUAWEI_V4.csv uses 6 decimal places: 1.000000
                try:
                    coef_a = "{:.6f}".format(float(factor))
                except ValueError:
                    coef_a = factor

                try:
                    coef_b = "{:.6f}".format(float(offset))
                except ValueError:
                    coef_b = offset

                out_row = [
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
                writer.writerow(out_row)
                index += 1

    except Exception as e:
        print(f"Error writing output file: {e}")
        sys.exit(1)

    print(f"Successfully generated {args.output_file}")

if __name__ == '__main__':
    main()
