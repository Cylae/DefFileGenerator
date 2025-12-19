import csv
import argparse
import sys

def parse_type(type_str, address):
    """
    Parses the user-friendly type string and returns Info2 suffix and Info3.

    Format examples:
    - uint16 -> U16
    - int32 -> I32
    - string(10) -> STRING, Info2 adds _10
    - bit(2) -> BITS, Info2 adds _2_1 (bit position 2, length 1) -> Wait, manual says Address_BitPos_BitLength
      Let's assume input bit(offset, length) or just bit(offset) for length 1.
    - u8(1) -> U8, Info2 adds _1 (byte offset)

    Returns: (info2_suffix, info3)
    """
    type_str = type_str.strip().upper()

    info2_suffix = ""
    info3 = ""

    if type_str.startswith("STRING"):
        # Expecting STRING(length)
        if "(" in type_str and type_str.endswith(")"):
            try:
                length = type_str.split("(")[1].strip(")")
                info2_suffix = f"_{length}"
                info3 = "STRING"
            except:
                raise ValueError(f"Invalid String format: {type_str}")
        else:
             # Default or error? Let's assume error or require length.
             raise ValueError(f"String type must specify length, e.g., STRING(10). Got: {type_str}")

    elif type_str.startswith("BIT"):
        # Expecting BIT(offset, length) or BIT(offset)
        if "(" in type_str and type_str.endswith(")"):
            content = type_str.split("(")[1].strip(")")
            parts = content.split(",")
            offset = parts[0].strip()
            length = "1"
            if len(parts) > 1:
                length = parts[1].strip()

            info2_suffix = f"_{offset}_{length}"
            info3 = "BITS"
        else:
             raise ValueError(f"Bit type must specify offset, e.g., BIT(0). Got: {type_str}")

    elif type_str.startswith("U8"):
         # Expecting U8(offset)
        if "(" in type_str and type_str.endswith(")"):
            offset = type_str.split("(")[1].strip(")")
            info2_suffix = f"_{offset}"
            info3 = "U8"
        else:
             # Assume offset 0 if not specified? Or high/low byte?
             # Manual says: "U8(offset)".
             # Let's support simple U8 -> U8 with suffix _0 (default?) or error.
             # Given manual example "40000_1" for 2nd byte, "40000_0" for 1st byte.
             # I'll default to _0 if not specified but maybe warn.
             info2_suffix = "_0"
             info3 = "U8"

    elif type_str.startswith("I8"):
        if "(" in type_str and type_str.endswith(")"):
            offset = type_str.split("(")[1].strip(")")
            info2_suffix = f"_{offset}"
            info3 = "I8"
        else:
             info2_suffix = "_0"
             info3 = "I8"

    else:
        # Standard types
        mapping = {
            "UINT16": "U16", "U16": "U16",
            "INT16": "I16", "I16": "I16",
            "UINT32": "U32", "U32": "U32",
            "INT32": "I32", "I32": "I32",
            "UINT64": "U64", "U64": "U64",
            "INT64": "I64", "I64": "I64",
            "FLOAT32": "F32", "F32": "F32",
            "FLOAT64": "F64", "F64": "F64",
            "IP": "IP",
            "IPV6": "IPV6",
            "MAC": "MAC"
        }

        # Handle modifiers like _W, _B
        base_type = type_str
        modifier = ""
        if "_WB" in type_str:
            base_type = type_str.replace("_WB", "")
            modifier = "_WB"
        elif "_W" in type_str:
            base_type = type_str.replace("_W", "")
            modifier = "_W"
        elif "_B" in type_str:
            base_type = type_str.replace("_B", "")
            modifier = "_B"

        if base_type in mapping:
            info3 = mapping[base_type] + modifier
        else:
            # Fallback or keep as is if it matches known types
            info3 = type_str

    return info2_suffix, info3

def map_register_type(reg_type):
    reg_type = reg_type.strip().lower()
    if "holding" in reg_type or "holding_register" in reg_type or reg_type == "3":
        return "3"
    elif "input" in reg_type or "input_register" in reg_type or reg_type == "4":
        return "4"
    elif "coil" in reg_type or reg_type == "1":
        return "1"
    elif "discrete" in reg_type or "discrete_input" in reg_type or reg_type == "2":
        return "2"
    else:
        # Default to Holding Register (3) if unsure, or raise error.
        # But let's be robust and return the value if it's a digit 1-4
        if reg_type in ["1", "2", "3", "4"]:
            return reg_type
        raise ValueError(f"Unknown register type: {reg_type}")

def generate_def_file(input_csv, output_csv, header_info):
    rows = []

    try:
        with open(input_csv, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)

            # Normalize column names (strip spaces, lower case for matching)
            # We expect standard column names but we should be flexible

            for i, row in enumerate(reader):
                try:
                    # Required fields
                    name = row.get("Name", "").strip()
                    tag = row.get("Tag", "").strip()
                    reg_type_str = row.get("RegisterType", "Holding").strip()
                    address = row.get("Address", "").strip()
                    type_str = row.get("Type", "U16").strip()

                    # Optional fields with defaults
                    factor = row.get("Factor", "").strip()
                    if not factor: factor = "1.0"

                    offset = row.get("Offset", "").strip()
                    if not offset: offset = "0.0"

                    unit = row.get("Unit", "").strip()

                    action = row.get("Action", "").strip()
                    if not action: action = "4" # Default to instantaneous value

                    # Processing
                    index = i + 1 # Auto-generate index starting at 1

                    info1 = map_register_type(reg_type_str)

                    info2_suffix, info3 = parse_type(type_str, address)
                    info2 = f"{address}{info2_suffix}"

                    info4 = "" # Scale factor variable name (SunSpec), usually empty for manual files

                    # Format output row
                    # Index;Info1;Info2;Info3;Info4;Name;Tag;CoefA;CoefB;Unit;Action
                    out_row = [
                        str(index),
                        str(info1),
                        str(info2),
                        str(info3),
                        str(info4),
                        name,
                        tag,
                        factor,
                        offset,
                        unit,
                        action
                    ]
                    rows.append(out_row)

                except Exception as e:
                    print(f"Error processing row {i+2}: {row} -> {e}")

    except FileNotFoundError:
        print(f"Error: Input file {input_csv} not found.")
        return

    # Write output
    try:
        with open(output_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')

            # Write Header
            # Protocole;Catégorie;Fabricant;Modèle;Code d’écriture forcé
            # We assume Modbus for now as per task context
            protocol = header_info.get("Protocol", "modbus")
            category = header_info.get("Category", "Inverter")
            manufacturer = header_info.get("Manufacturer", "Generic")
            model = header_info.get("Model", "Device")
            write_code = header_info.get("ForceWrite", "")

            writer.writerow([protocol, category, manufacturer, model, write_code])

            # Write Rows
            for row in rows:
                writer.writerow(row)

        print(f"Successfully generated {output_csv} with {len(rows)} variables.")

    except Exception as e:
        print(f"Error writing output file: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate WebdynSunPM definition file from CSV.")
    parser.add_argument("input_file", help="Path to input CSV file")
    parser.add_argument("output_file", help="Path to output CSV file")
    parser.add_argument("--protocol", default="modbus", help="Protocol (default: modbus)")
    parser.add_argument("--category", default="Inverter", help="Category (default: Inverter)")
    parser.add_argument("--manufacturer", default="Generic", help="Manufacturer (default: Generic)")
    parser.add_argument("--model", default="Device", help="Model (default: Device)")
    parser.add_argument("--force_write", default="", help="Force Write Code (0 or 1)")

    args = parser.parse_args()

    header_info = {
        "Protocol": args.protocol,
        "Category": args.category,
        "Manufacturer": args.manufacturer,
        "Model": args.model,
        "ForceWrite": args.force_write
    }

    generate_def_file(args.input_file, args.output_file, header_info)
