#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
import io
from dataclasses import dataclass

# Pre-compiled regex patterns for optimization
RE_TYPE_INT = re.compile(r'^[UI](8|16|32|64)(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^-?([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

@dataclass
class GeneratorConfig:
    input_file: str = None
    output: str = None
    manufacturer: str = None
    model: str = None
    protocol: str = 'modbusRTU'
    category: str = 'Inverter'
    forced_write: str = ''
    template: bool = False
    address_offset: int = 0

class Generator:
    def __init__(self):
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

    def normalize_type(self, t):
        if not t:
            return 'U16'
        t_str = str(t).lower().strip()

        # Order by specificity to prevent partial matches
        mapping = [
            ('unsigned int 64', 'U64'),
            ('signed integer 64', 'I64'),
            ('unsigned int 32', 'U32'),
            ('signed integer 32', 'I32'),
            ('unsigned int 16', 'U16'),
            ('signed integer 16', 'I16'),
            ('unsigned int 8', 'U8'),
            ('signed integer 8', 'I8'),
            ('float64', 'F64'),
            ('float32', 'F32'),
            ('uint64', 'U64'),
            ('int64', 'I64'),
            ('uint32', 'U32'),
            ('int32', 'I32'),
            ('uint16', 'U16'),
            ('int16', 'I16'),
            ('uint8', 'U8'),
            ('int8', 'I8'),
            ('float', 'F32'),
            ('double', 'F64'),
            ('string', 'STRING'),
            ('bits', 'BITS'),
            ('u64', 'U64'),
            ('i64', 'I64'),
            ('u32', 'U32'),
            ('i32', 'I32'),
            ('u16', 'U16'),
            ('i16', 'I16'),
            ('u8', 'U8'),
            ('i8', 'I8'),
            ('f32', 'F32'),
            ('f64', 'F64')
        ]

        for src, dest in mapping:
            if src in t_str:
                # Preserve suffixes
                suffix_match = re.search(r'(_[WB]+)$', t_str.upper())
                suffix = suffix_match.group(1) if suffix_match else ''
                return dest + suffix

        # Fallback cleaning
        clean_t = re.sub(r'[^a-z0-9_]+', '', t_str)
        return clean_t.upper() if clean_t else 'U16'

    def normalize_address_val(self, addr_part):
        """Converts a single address part (possibly hex) to decimal string."""
        if addr_part is None:
            return ""
        addr_str = str(addr_part).strip()

        # Remove thousands separators
        addr_str = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_str)

        # Regex to identify candidate hex or decimal words
        pattern = re.compile(r'(?<![0-9A-Za-z])(0x[0-9A-Fa-f]+|[0-9A-Fa-f]+h|-?\d+|[0-9A-Fa-f]+)(?![0-9A-Za-z])')
        match = pattern.search(addr_str)

        if not match:
            return addr_str

        val = match.group(1)

        if val.lower().startswith('0x'):
            try:
                return str(int(val, 16))
            except ValueError:
                return val
        elif val.lower().endswith('h'):
            try:
                return str(int(val[:-1], 16))
            except ValueError:
                return val
        # If it contains A-F and no x/h, it might still be hex
        if any(c in val.upper() for c in 'ABCDEF'):
            try:
                return str(int(val, 16))
            except ValueError:
                return val
        return val

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

    def normalize_action(self, action):
        if action is None or action == '':
            return '1'
        act_str = str(action).strip().upper()
        if act_str in ['R', 'READ', '4']:
            return '4'
        if act_str in ['RW', 'W', 'WRITE', '1']:
            return '1'
        if act_str in self.allowed_actions:
            return act_str
        return '1'

    def process_rows(self, rows, address_offset=0):
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
            address_raw = get_val('Address')
            dtype_raw = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action_raw = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address_raw:
                logging.warning(f"Line {line_num}: Skipping row with missing Name and Address.")
                continue

            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid Type '{dtype}'. Skipping row.")
                continue

            # Handle STR<n> conversion
            dtype_upper = dtype.upper()
            match_str = RE_TYPE_STR_CONV.match(dtype_upper)
            if match_str:
                length = int(match_str.group(1))
                dtype = 'STRING'
                if '_' not in address_raw:
                    address_raw = f"{address_raw}_{length}"

            # Normalize Address
            if address_raw:
                addr_parts = address_raw.split('_')
                norm_parts = [self.normalize_address_val(p) for p in addr_parts]

                # Apply address offset to the base address
                try:
                    base_addr = int(norm_parts[0])
                    base_addr += address_offset
                    if base_addr < 0:
                        logging.warning(f"Line {line_num}: Address became negative ({base_addr}) after applying offset {address_offset}.")
                    norm_parts[0] = str(base_addr)
                except ValueError:
                    pass

                address = '_'.join(norm_parts)
            else:
                address = ''

            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid Address '{address}' for Type '{dtype}'. Skipping row.")
                continue

            if name:
                if name in seen_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
                else:
                    seen_names[name] = line_num

            if not tag and name:
                base_tag = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
                tag = base_tag if base_tag else "var"
                counter = 1
                while tag in seen_tags:
                    tag = f"{base_tag}_{counter}"
                    counter += 1

            if tag:
                if tag in seen_tags:
                    logging.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous occurrence at line {seen_tags[tag]}.")
                else:
                    seen_tags[tag] = line_num

            info1 = '3'
            if reg_type_str:
                lower_type = reg_type_str.lower()
                if lower_type in self.register_type_map:
                    info1 = self.register_type_map[lower_type]
                elif reg_type_str in ['1', '2', '3', '4']:
                    info1 = reg_type_str
                else:
                    logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to Holding Register (3).")

            # Address Overlap Check (O(N^2) as noted in memories)
            try:
                parts = address.split('_')
                start_addr = int(parts[0])
                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1

                if info1 not in used_addresses_by_type:
                    used_addresses_by_type[info1] = []

                is_bits = (dtype.upper() == 'BITS')
                for used_start, used_end, used_line, used_name, used_type in used_addresses_by_type[info1]:
                    if max(start_addr, used_start) <= min(end_addr, used_end):
                        if not (is_bits and used_type == 'BITS' and start_addr == used_start):
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). Overlaps with '{used_name}' (Line {used_line}, Addr: {used_start}-{used_end}) in register type {info1}.")

                used_addresses_by_type[info1].append((start_addr, end_addr, line_num, name, dtype.upper()))
            except ValueError:
                pass

            # Coefficients
            try:
                val_factor = float(factor) if factor and str(factor).strip() else 1.0
            except ValueError:
                val_factor = 1.0
            try:
                val_scale = int(float(scale_factor_str)) if scale_factor_str and str(scale_factor_str).strip() else 0
            except ValueError:
                val_scale = 0
            coef_a = "{:.6f}".format(val_factor * (10 ** val_scale))

            try:
                val_offset = float(offset) if offset and str(offset).strip() else 0.0
            except ValueError:
                val_offset = 0.0
            coef_b = "{:.6f}".format(val_offset)

            action = self.normalize_action(action_raw)

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

    @staticmethod
    def write_output_csv(output, config, processed_rows):
        """Writes the WebdynSunPM definition file CSV."""
        header_row = [
            config.protocol,
            config.category,
            config.manufacturer,
            config.model,
            config.forced_write,
            '', '', '', '', '', ''
        ]

        if isinstance(output, str):
            f = open(output, 'w', newline='', encoding='utf-8')
        else:
            f = output

        try:
            writer = csv.writer(f, delimiter=';', lineterminator='\n')
            writer.writerow(header_row)

            for index, row in enumerate(processed_rows, start=1):
                writer.writerow([
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
                ])
        finally:
            if isinstance(output, str):
                f.close()

def generate_template(output_file):
    """Generates a template CSV input file."""
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    rows = [
        ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
        ['String Variable', 'string_tag', 'Holding Register', '30010_10', 'String', '', '', '', '4', ''],
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'Bits', '', '', '', '4', ''],
        ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
    ]

    if output_file:
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)
        logging.info(f"Template generated at {output_file}")
    else:
        writer = csv.writer(sys.stdout)
        writer.writerow(headers)
        writer.writerows(rows)

def run_generator(config: GeneratorConfig):
    if config.template:
        generate_template(config.output)
        return

    if not config.input_file:
        logging.error("input_file is required")
        return

    if not config.manufacturer or not config.model:
         logging.error("--manufacturer and --model are required")
         return

    generator = Generator()

    try:
        # Detect encoding and handle BOM (UTF-16 check per memory)
        with open(config.input_file, 'rb') as f:
            raw = f.read(4)
            if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
                encoding = 'utf-16'
            else:
                encoding = 'utf-8-sig'

        with open(config.input_file, mode='r', encoding=encoding) as csvfile:
            content = csvfile.read(4096)
            csvfile.seek(0)
            try:
                sniffer = csv.Sniffer()
                dialect = sniffer.sniff(content, delimiters=";,")
            except csv.Error:
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(csvfile, dialect=dialect)
            if reader.fieldnames:
                reader.fieldnames = [name.strip() for name in reader.fieldnames]

            processed_rows = generator.process_rows(reader, config.address_offset)

        generator.write_output_csv(config.output or sys.stdout, config, processed_rows)
        if config.output:
            logging.info(f"Definition file generated at {config.output}")

    except FileNotFoundError:
        logging.error(f"File '{config.input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Generate WebdynSunPM definition file.')
    parser.add_argument('input_file', nargs='?', help='Simplified CSV input file.')
    parser.add_argument('-o', '--output', help='Output CSV file.')
    parser.add_argument('--manufacturer', help='Manufacturer name.')
    parser.add_argument('--model', help='Model name.')
    parser.add_argument('--protocol', default='modbusRTU', help='Protocol name.')
    parser.add_argument('--category', default='Inverter', help='Device category.')
    parser.add_argument('--forced-write', default='', help='Forced write code.')
    parser.add_argument('--address-offset', type=int, default=0, help='Offset to apply to addresses.')
    parser.add_argument('--template', action='store_true', help='Generate template.')

    args = parser.parse_args()
    config = GeneratorConfig(
        input_file=args.input_file,
        output=args.output,
        manufacturer=args.manufacturer,
        model=args.model,
        protocol=args.protocol,
        category=args.category,
        forced_write=args.forced_write,
        address_offset=args.address_offset,
        template=args.template
    )
    run_generator(config)

if __name__ == "__main__":
    main()
