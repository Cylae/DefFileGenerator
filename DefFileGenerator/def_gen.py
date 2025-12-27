#!/usr/bin/env python3
import argparse
import csv
import sys
import os
import re

def parse_arguments():
    parser = argparse.ArgumentParser(description="Generate WebdynSunPM definition files.")
    parser.add_argument("-o", "--output", help="Output file path")
    parser.add_argument("--template", action="store_true", help="Generate sample CSV")
    parser.add_argument("--protocol", default="modbusRTU", help="Protocol (default: modbusRTU)")
    parser.add_argument("--category", default="Inverter", help="Category (default: Inverter)")
    parser.add_argument("--manufacturer", default="Manufacturer", help="Manufacturer")
    parser.add_argument("--model", default="Model", help="Model")
    parser.add_argument("--forced-write", default="", help="Forced Write Code")
    parser.add_argument("input_file", nargs="?", help="Input CSV file")
    return parser.parse_args()

def generate_template(output_file):
    headers = ["Name", "Tag", "RegisterType", "Address", "Type", "Factor", "Offset", "Unit", "Action"]
    rows = [
        ["Example Name", "ExampleTag", "Holding", "40001", "U16", "1", "0", "V", "4"],
        ["Example String", "ExampleStr", "Holding", "40010_10", "STRING", "1", "0", "", "4"]
    ]
    if not output_file:
        output_file = "template.csv"

    try:
        with open(output_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(headers)
            writer.writerows(rows)
        print(f"Template generated at {output_file}")
    except IOError as e:
        print(f"Error writing template file: {e}")

def get_register_type_code(reg_type):
    reg_type = str(reg_type).lower().strip()
    if "coil" in reg_type:
        return 1
    elif "discrete" in reg_type: # discrete or discrete input
        return 2
    elif "holding" in reg_type: # holding or holding register
        return 3
    elif "input" in reg_type: # input or input register
        return 4
    # Default to Holding (3) if unspecified or unknown?
    # Or maybe 3 is safest for Modbus devices usually.
    return 3

def format_float(val, default_val):
    if not val:
        val = default_val
    try:
        f = float(val)
        return "{:.6f}".format(f)
    except ValueError:
        return val

def validate_address(address, dtype):
    address = address.strip()
    dtype = dtype.upper()

    if dtype == "STRING":
        if not re.match(r"^\d+_\d+$", address):
            print(f"Warning: Address '{address}' for type STRING should be in format 'Address_Length'.")
    elif dtype == "BITS":
        if not re.match(r"^\d+_\d+_\d+$", address):
            print(f"Warning: Address '{address}' for type BITS should be in format 'Address_StartBit_NumBits'.")
    else:
        if not re.match(r"^\d+$", address):
             print(f"Warning: Address '{address}' for type {dtype} should be a simple integer.")
    return address

def process_file(args):
    input_path = args.input_file
    output_path = args.output

    if not input_path:
        print("Error: Input file required.")
        sys.exit(1)

    if not output_path:
        base, _ = os.path.splitext(input_path)
        output_path = f"{base}_def.csv"

    rows = []
    try:
        with open(input_path, "r", encoding="utf-8-sig") as f:
            # Detect delimiter
            sample = f.read(2048)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.excel
                dialect.delimiter = ";" # Fallback

            reader = csv.DictReader(f, dialect=dialect)

            # Map headers
            field_map = {}
            if reader.fieldnames:
                for field in reader.fieldnames:
                    f_lower = field.lower().strip()
                    if f_lower == "name": field_map["Name"] = field
                    elif f_lower == "tag": field_map["Tag"] = field
                    elif f_lower == "registertype": field_map["RegisterType"] = field
                    elif f_lower == "address": field_map["Address"] = field
                    elif f_lower == "type": field_map["Type"] = field
                    elif f_lower == "factor": field_map["Factor"] = field
                    elif f_lower == "offset": field_map["Offset"] = field
                    elif f_lower == "unit": field_map["Unit"] = field
                    elif f_lower == "action": field_map["Action"] = field

            index = 1
            for row in reader:
                # Extract values with stripping
                name = row.get(field_map.get("Name"), "").strip()
                tag = row.get(field_map.get("Tag"), "").strip()
                reg_type = row.get(field_map.get("RegisterType"), "Holding").strip()
                address = row.get(field_map.get("Address"), "").strip()
                dtype = row.get(field_map.get("Type"), "U16").strip().upper()
                factor = row.get(field_map.get("Factor"), "1").strip()
                offset = row.get(field_map.get("Offset"), "0").strip()
                unit = row.get(field_map.get("Unit"), "").strip()
                action = row.get(field_map.get("Action"), "1").strip()

                if not action: action = "1"

                # Validation
                address = validate_address(address, dtype)

                info1 = get_register_type_code(reg_type)
                info2 = address
                info3 = dtype
                info4 = "" # Empty

                coef_a = format_float(factor, "1")
                coef_b = format_float(offset, "0")

                # Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                out_row = [
                    index, info1, info2, info3, info4, name, tag, coef_a, coef_b, unit, action
                ]
                rows.append(out_row)
                index += 1
    except IOError as e:
        print(f"Error reading input file: {e}")
        sys.exit(1)

    # Write output
    try:
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=";")
            # Header: Protocole;Category;Manufacturer;Model;ForcedWriteCode
            header = [args.protocol, args.category, args.manufacturer, args.model, args.forced_write]

            while len(header) < 11:
                header.append("")

            writer.writerow(header)
            writer.writerows(rows)

        print(f"Definition file generated at {output_path}")
    except IOError as e:
        print(f"Error writing output file: {e}")
        sys.exit(1)

def main():
    args = parse_arguments()
    if args.template:
        generate_template(args.output)
    else:
        process_file(args)

if __name__ == "__main__":
    main()
