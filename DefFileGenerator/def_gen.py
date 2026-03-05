#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F-]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F-]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

class Generator:
    def __init__(self, address_offset=0):
        self.address_offset = address_offset
        # RegisterType mapping to Info1
        self.register_type_map = {
            'coil': '1',
            'coils': '1',
            'discrete input': '2',
            'holding register': '3',
            'holding': '3',
            'input register': '4',
            'input': '4'
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
        if RE_TYPE_INT.match(dtype_upper):
            return True

        # STR<n> syntax (e.g., STR20)
        if RE_TYPE_STR_CONV.match(dtype_upper):
            return True

        return False

    def normalize_type(self, dtype):
        """Standardizes data type synonyms while preserving suffixes."""
        if not dtype:
            return 'U16'

        d = str(dtype).lower().strip()
        d = d.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')

        # Order by specificity (longer first)
        mapping = [
            ('float64', 'F64'),
            ('float32', 'F32'),
            ('float', 'F32'),
            ('double', 'F64'),
            ('uint64', 'U64'),
            ('uint32', 'U32'),
            ('uint16', 'U16'),
            ('uint8', 'U8'),
            ('int64', 'I64'),
            ('int32', 'I32'),
            ('int16', 'I16'),
            ('int8', 'I8'),
            ('uint', 'U'),
            ('int', 'I'),
            ('u16', 'U16'),
            ('u32', 'U32'),
            ('u64', 'U64'),
            ('u8', 'U8'),
            ('i16', 'I16'),
            ('i32', 'I32'),
            ('i64', 'I64'),
            ('i8', 'I8')
        ]

        for syn, target in mapping:
            if syn in d:
                # If we matched a short synonym like 'u' or 'i', make sure it's followed by a digit or is the end
                if syn in ['u', 'i', 'uint', 'int']:
                    d = re.sub(rf'\b{syn}(\d+)', rf'{target}\1', d)
                    d = re.sub(rf'\b{syn}\b', target, d)
                else:
                    d = d.replace(syn, target)
                break

        # Preserve suffixes and clean up
        d = re.sub(r'[^a-zA-Z0-9_]+', '', d)
        return d.upper()

    def normalize_action(self, action):
        """Normalizes action values by mapping synonyms."""
        if not action or not str(action).strip():
            return '1'

        a = str(action).strip().upper()
        if a in ['R', 'READ', '4']:
            return '4'
        if a in ['RW', 'W', 'WRITE', '1']:
            return '1'
        if a in self.allowed_actions:
            return a
        return '1'

    def normalize_address_val(self, addr_part):
        """Converts a single address part (possibly hex) to decimal string."""
        addr_part = str(addr_part).strip().replace(',', '')
        if not addr_part:
            return ""

        # Priority: explicit hex markers
        hex_match = re.search(r'\b(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h)\b', addr_part)
        if hex_match:
            val = hex_match.group(1)
            if val.lower().startswith('0x'):
                try:
                    return str(int(val, 16))
                except ValueError:
                    return val
            else:
                try:
                    return str(int(val[:-1], 16))
                except ValueError:
                    return val

        # General match for decimal or hex without prefix
        match = re.search(r'\b([0-9A-Fa-f-]+)\b', addr_part)
        if match:
            val = match.group(1)
            # If it contains A-F, it's likely hex
            if any(c in val.upper() for c in 'ABCDEF'):
                try:
                    return str(int(val, 16))
                except ValueError:
                    return val
            return val

        return addr_part

    def validate_address(self, address, dtype):
        """Validates the address format based on type."""
        dtype_upper = dtype.upper()

        if dtype_upper == 'STRING':
            return RE_ADDR_STRING.match(address) is not None
        elif dtype_upper == 'BITS':
            return RE_ADDR_BITS.match(address) is not None
        else:
            return RE_ADDR_INT.match(address) is not None

    def get_register_count(self, dtype, address):
        """Calculates the number of registers used by the type."""
        dtype_upper = dtype.upper()

        if RE_COUNT_16_8.match(dtype_upper):
            return 1
        elif RE_COUNT_32.match(dtype_upper):
            return 2
        elif RE_COUNT_64.match(dtype_upper):
            return 4
        elif dtype_upper == 'MAC':
            return 3
        elif dtype_upper == 'IPV6':
            return 8
        elif dtype_upper == 'STRING':
            try:
                parts = address.split('_')
                length = int(parts[1])
                return math.ceil(length / 2)
            except (IndexError, ValueError):
                return 0
        return 1

    def process_rows(self, rows):
        """Processes simplified CSV rows into WebdynSunPM format."""
        processed_rows = []
        seen_names = {}
        seen_tags = {}
        used_addresses_by_type = {}

        for line_num, row in enumerate(rows, start=2):
            if not any(v for v in row.values() if v):
                continue

            def get_val(key):
                val = row.get(key)
                if val is not None:
                    return str(val).strip()
                for k, v in row.items():
                    if k.lower() == key.lower():
                        return str(v).strip()
                return ''

            name = get_val('Name')
            tag = get_val('Tag')
            reg_type_str = get_val('RegisterType')
            raw_address = get_val('Address')
            raw_dtype = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            raw_action = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not raw_address:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            # Normalize Type
            dtype = self.normalize_type(raw_dtype)

            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                continue

            # Handle STR<n> conversion
            dtype_upper = dtype.upper()
            match_str = RE_TYPE_STR_CONV.match(dtype_upper)
            if match_str:
                try:
                    length = int(match_str.group(1))
                    dtype = 'STRING'
                    if '_' not in raw_address:
                        raw_address = f"{raw_address}_{length}"
                except ValueError:
                    logging.warning(f"Line {line_num}: Invalid STR format '{dtype_upper}'. Skipping row.")
                    continue

            # Normalize Address
            address = raw_address
            if address:
                addr_parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]

                # Apply address offset to the first part
                try:
                    start_addr_val = int(norm_parts[0])
                    final_start_addr = start_addr_val - self.address_offset
                    if final_start_addr < 0:
                        logging.warning(f"Line {line_num}: Address {norm_parts[0]} with offset {self.address_offset} results in negative address {final_start_addr}")
                    norm_parts[0] = str(final_start_addr)
                except (ValueError, IndexError):
                    pass

                address = '_'.join(norm_parts)

            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            # Global Check: Duplicate Name
            if name:
                if name in seen_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
                else:
                    seen_names[name] = line_num

            # Tag generation
            if not tag and name:
                base_tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
                if not base_tag:
                    base_tag = "var"
                tag = base_tag
                counter = 1
                while tag in seen_tags:
                    tag = f"{base_tag}_{counter}"
                    counter += 1

            if tag:
                if tag in seen_tags:
                    logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous occurrence at line {seen_tags[tag]}.")
                else:
                    seen_tags[tag] = line_num

            # Map RegisterType to Info1
            info1 = '3'
            if reg_type_str:
                lower_type = reg_type_str.lower()
                if lower_type in self.register_type_map:
                    info1 = self.register_type_map[lower_type]
                elif reg_type_str in ['1', '2', '3', '4']:
                    info1 = reg_type_str
                else:
                    logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to Holding Register (3).")

            # Overlap Calculation
            try:
                parts = address.split('_')
                start_addr = int(parts[0])
                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1
                current_range = (start_addr, end_addr, line_num, name, dtype.upper())

                if info1 not in used_addresses_by_type:
                    used_addresses_by_type[info1] = []

                is_bits = (dtype.upper() == 'BITS')
                for used_start, used_end, used_line, used_name, used_type in used_addresses_by_type[info1]:
                    if max(start_addr, used_start) <= min(end_addr, used_end):
                        is_overlap_allowed = is_bits and (used_type == 'BITS')
                        if not is_overlap_allowed:
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). Overlaps with '{used_name}' (Line {used_line}, Addr: {used_start}-{used_end}) in register type {info1}.")
                used_addresses_by_type[info1].append(current_range)
            except ValueError:
                logging.warning(f"Line {line_num}: Could not calculate register range for address '{address}'.")

            # CoefA
            try:
                val_factor = float(factor) if factor and str(factor).strip() else 1.0
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid Factor '{factor}'. Using 1.0.")
                val_factor = 1.0

            try:
                val_scale = int(float(scale_factor_str)) if scale_factor_str and str(scale_factor_str).strip() else 0
            except ValueError:
                 logging.warning(f"Line {line_num}: Invalid ScaleFactor '{scale_factor_str}'. Using 0.")
                 val_scale = 0

            final_coef_a_val = val_factor * (10 ** val_scale)
            coef_a = "{:.6f}".format(final_coef_a_val)

            # CoefB
            try:
                val_offset = float(offset) if offset and str(offset).strip() else 0.0
                coef_b = "{:.6f}".format(val_offset)
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid Offset '{offset}'. Defaulting to 0.000000.")
                coef_b = "0.000000"

            # Action normalization
            action = self.normalize_action(raw_action)

            processed_rows.append({
                'Info1': info1,
                'Info2': address,
                'Info3': dtype.upper(),
                'Info4': '',
                'Name': name,
                'Tag': tag,
                'CoefA': coef_a,
                'CoefB': coef_b,
                'Unit': unit,
                'Action': action
            })

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

def run_generator(input_file, output=None, manufacturer=None, model=None,
                 protocol='modbusRTU', category='Inverter', forced_write='',
                 template=False, address_offset=0):
    if template:
        generate_template(output)
        return

    if not input_file:
        logging.error("input_file is required")
        return

    if not manufacturer or not model:
         logging.error("--manufacturer and --model are required")
         return

    generator = Generator(address_offset=address_offset)

    try:
        # Robust encoding detection
        encoding = 'utf-8-sig'
        try:
            with open(input_file, 'rb') as f:
                raw = f.read(2)
                if raw == b'\xff\xfe' or raw == b'\xfe\xff':
                    encoding = 'utf-16'
        except Exception:
            pass

        with open(input_file, mode='r', encoding=encoding) as csvfile:
            try:
                content = csvfile.read(1024)
                csvfile.seek(0)
                dialect = csv.Sniffer().sniff(content, delimiters=";,")
            except Exception:
                csvfile.seek(0)
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(csvfile, dialect=dialect)
            if reader.fieldnames:
                reader.fieldnames = [name.strip() for name in reader.fieldnames]
            else:
                logging.error("Input CSV is empty or missing headers.")
                return

            required_columns = ['Name', 'RegisterType', 'Address', 'Type']
            header_map = {h.lower(): h for h in reader.fieldnames}
            missing_columns = [col for col in required_columns if col.lower() not in header_map]

            if missing_columns:
                logging.error(f"Missing required columns in input CSV: {', '.join(missing_columns)}")
                return

            processed_rows = generator.process_rows(reader)

        if output:
            outfile = open(output, 'w', newline='', encoding='utf-8')
        else:
            outfile = sys.stdout

        header_row = [protocol, category, manufacturer, model, forced_write, '', '', '', '', '', '']
        writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
        writer.writerow(header_row)

        for index, row in enumerate(processed_rows, start=1):
            writer.writerow([
                str(index), row['Info1'], row['Info2'], row['Info3'], row['Info4'],
                row['Name'], row['Tag'], row['CoefA'], row['CoefB'], row['Unit'], row['Action']
            ])

        if output:
            outfile.close()
            logging.info(f"Definition file generated at {output}")

    except FileNotFoundError:
        logging.error(f"File '{input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM Modbus definition file from simplified CSV.')
    parser.add_argument('input_file', nargs='?', help='Path to the simplified CSV input file.')
    parser.add_argument('-o', '--output', help='Path to the output CSV file.')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name.')
    parser.add_argument('--category', default='Inverter', help='Device category.')
    parser.add_argument('--manufacturer', help='Manufacturer name.')
    parser.add_argument('--model', help='Model name.')
    parser.add_argument('--forced-write', default='', help='Forced write code.')
    parser.add_argument('--template', action='store_true', help='Generate template.')
    parser.add_argument('--address-offset', type=int, default=0, help='Address offset to subtract.')

    args = parser.parse_args()
    run_generator(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        template=args.template,
        address_offset=args.address_offset
    )

if __name__ == "__main__":
    main()
