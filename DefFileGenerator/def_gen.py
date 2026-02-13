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
RE_ADDR_STRING = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64)$', re.IGNORECASE)

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
        # Allowed Action codes for WebdynSunPM
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    def normalize_type(self, dtype):
        """Normalizes various data type synonyms to standard WebdynSunPM types."""
        if not dtype:
            return 'U16'

        s = str(dtype).lower().strip()
        # Remove common documentation artifacts but keep underscores
        s = re.sub(r'[^a-z0-9_ ]+', '', s)

        # Check for common byte/word swap suffixes
        suffix = ""
        for suf in ['_wb', '_w', '_b']:
            if suf in s:
                suffix = suf.upper()
                s = s.replace(suf, '')
                break

        # Normalize common words
        s = s.replace('unsigned ', 'u').replace('signed ', 'i').replace(' ', '')

        # 1. Check for standard integer patterns (e.g., uint16, int32, u16, i32)
        # We also allow optional words like 'long', 'short', 'word', 'byte'
        m = re.search(r'(u|i|uint|int)(?:long|short|word|byte)?(\d+)', s, re.IGNORECASE)
        if m:
            prefix = 'U' if m.group(1).startswith('u') else 'I'
            bits = m.group(2)
            if bits in ['8', '16', '32', '64']:
                return f"{prefix}{bits}{suffix}"

        # 2. Check for other common types
        type_map = [
            ('uint64', 'U64'), ('int64', 'I64'),
            ('uint32', 'U32'), ('int32', 'I32'),
            ('uint16', 'U16'), ('int16', 'I16'),
            ('uint8', 'U8'), ('int8', 'I8'),
            ('float64', 'F64'), ('float32', 'F32'),
            ('double', 'F64'), ('float', 'F32'),
            ('string', 'STRING'), ('boolean', 'BITS'), ('bool', 'BITS'), ('bits', 'BITS'),
            ('ipv6', 'IPV6'), ('ipv4', 'IP'), ('ip', 'IP'), ('mac', 'MAC')
        ]
        for k, v in type_map:
            if k in s:
                return v + suffix

        # 3. Handle STR<n> pattern
        m_str = RE_TYPE_STR_CONV.search(s)
        if m_str:
            return m_str.group(0).upper()

        # Fallback to uppercase cleaned string
        res = s.upper()
        return res + suffix if res else 'U16' + suffix

    def validate_type(self, dtype):
        """Validates if the data type is supported by WebdynSunPM."""
        d = dtype.upper()
        base_types = ['STRING', 'BITS', 'IP', 'IPV6', 'MAC', 'F32', 'F64']
        if d in base_types:
            return True
        return RE_TYPE_INT.match(d) is not None or RE_TYPE_STR_CONV.match(d) is not None

    def normalize_address_val(self, a):
        """Converts a single address part (possibly hex) to decimal string."""
        if a is None:
            return ""
        a = str(a).strip().replace(',', '')
        if not a:
            return ""

        # Handle hex formats
        if a.lower().startswith('0x'):
            try:
                return str(int(a, 16))
            except ValueError:
                return a
        elif a.lower().endswith('h'):
            try:
                return str(int(a[:-1], 16))
            except ValueError:
                return a

        # Heuristic: if it contains A-F and is not valid decimal
        if any(c in a.upper() for c in 'ABCDEF'):
            try:
                return str(int(a, 16))
            except ValueError:
                return a

        return a

    def validate_address(self, a, t):
        """Validates address format based on data type."""
        t = t.upper()
        if t == 'STRING':
            return RE_ADDR_STRING.match(a) is not None
        elif t == 'BITS':
            return RE_ADDR_BITS.match(a) is not None
        return RE_ADDR_INT.match(a) is not None

    def get_register_count(self, t, a):
        """Calculates the number of 16-bit registers used by the type."""
        t = t.upper()
        if RE_COUNT_16_8.match(t):
            return 1
        elif RE_COUNT_32.match(t):
            return 2
        elif RE_COUNT_64.match(t):
            return 4
        elif t == 'MAC':
            return 3
        elif t == 'IPV6':
            return 8
        elif t == 'STRING':
            try:
                p = a.split('_')
                length = int(p[1])
                return math.ceil(length / 2)
            except (IndexError, ValueError):
                return 0
        return 1

    def normalize_action(self, a):
        """Normalizes action strings to WebdynSunPM numeric codes."""
        if not a or not str(a).strip():
            return '1' # Default: Read on change

        s = str(a).strip().upper()
        if s in ['R', 'READ', 'READ-ONLY', 'READ ONLY', '4']:
            return '4' # Read on request
        elif s in ['RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', '1']:
            return '1' # Read on change (standard for RW/W in this tool)

        if s in self.allowed_actions:
            return s
        return '1'

    def process_rows(self, rows):
        """Processes simplified CSV rows into WebdynSunPM format."""
        processed = []
        seen_names = {}
        seen_tags = {}
        used_addresses = {} # info1 -> list of (start, end, line, name, type)

        for line_num, row in enumerate(rows, start=2):
            # Skip empty rows
            if not any(v for v in row.values() if v):
                continue

            def get_val(k):
                val = row.get(k)
                if val is not None:
                    return str(val).strip()
                # Case-insensitive fallback
                for rk, rv in row.items():
                    if rk.lower() == k.lower():
                        return str(rv).strip()
                return ''

            # Extract fields
            name = get_val('Name')
            tag = get_val('Tag')
            reg_type_str = get_val('RegisterType')
            address = get_val('Address')
            dtype_raw = get_val('Type')
            factor = get_val('Factor')
            offset = get_val('Offset')
            unit = get_val('Unit')
            action_raw = get_val('Action')
            scale_factor_str = get_val('ScaleFactor')

            if not name and not address:
                continue

            # Normalize and validate Type
            dtype = self.normalize_type(dtype_raw)
            if not self.validate_type(dtype):
                logging.warning(f"Line {line_num}: Invalid data type '{dtype_raw}' (normalized to '{dtype}'). Skipping row.")
                continue

            # Special handling for STR<n> and STRING
            dtype_upper = dtype.upper()
            match_str = RE_TYPE_STR_CONV.match(dtype_upper)
            if match_str:
                dtype = 'STRING'
                length = match_str.group(1)
                if address and '_' not in address:
                    address = f"{address}_{length}"

            # Normalize Address
            if address:
                parts = address.split('_')
                norm_parts = [self.normalize_address_val(p) for p in parts]
                address = '_'.join(norm_parts)

            # Validate Address
            if not self.validate_address(address, dtype):
                logging.warning(f"Line {line_num}: Invalid address format '{address}' for type '{dtype}'. Skipping row.")
                continue

            # Duplicate Name Detection
            if name:
                if name in seen_names:
                    logging.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous occurrence at line {seen_names[name]}.")
                else:
                    seen_names[name] = line_num

            # Automatic Tag Generation
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
                    logging.warning(f"Line {line_num}: Duplicate tag '{tag}' detected. Previous at line {seen_tags[tag]}.")
                else:
                    seen_tags[tag] = line_num

            # Map RegisterType to Info1
            info1 = '3' # Default: Holding Register
            if reg_type_str:
                lt = reg_type_str.lower()
                if lt in self.register_type_map:
                    info1 = self.register_type_map[lt]
                elif reg_type_str in ['1', '2', '3', '4']:
                    info1 = reg_type_str
                else:
                    logging.warning(f"Line {line_num}: Unknown RegisterType '{reg_type_str}'. Defaulting to Holding (3).")

            # Overlap Detection
            try:
                start_addr = int(address.split('_')[0])
                reg_count = self.get_register_count(dtype, address)
                end_addr = start_addr + reg_count - 1

                if info1 not in used_addresses:
                    used_addresses[info1] = []

                is_bits = (dtype.upper() == 'BITS')
                for u_start, u_end, u_line, u_name, u_type in used_addresses[info1]:
                    if max(start_addr, u_start) <= min(end_addr, u_end):
                        # Allow overlap only if both are BITS and same address
                        if not (is_bits and u_type == 'BITS'):
                            logging.warning(f"Line {line_num}: Address overlap detected for '{name}' (Addr: {start_addr}-{end_addr}). Overlaps with '{u_name}' at line {u_line}.")

                used_addresses[info1].append((start_addr, end_addr, line_num, name, dtype.upper()))
            except (ValueError, IndexError):
                logging.debug(f"Line {line_num}: Skipping overlap check due to address format.")

            # Calculate Coefficients
            try:
                f_val = float(factor) if factor else 1.0
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid factor '{factor}'. Using 1.0.")
                f_val = 1.0

            try:
                s_val = int(float(scale_factor_str)) if scale_factor_str else 0
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid scale factor '{scale_factor_str}'. Using 0.")
                s_val = 0

            coef_a = "{:.6f}".format(f_val * (10 ** s_val))

            try:
                o_val = float(offset) if offset else 0.0
            except ValueError:
                logging.warning(f"Line {line_num}: Invalid offset '{offset}'. Using 0.0.")
                o_val = 0.0
            coef_b = "{:.6f}".format(o_val)

            # Normalize Action
            action = self.normalize_action(action_raw)

            processed.append({
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

        return processed

def generate_template(output_file):
    """Generates a template CSV input file."""
    headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
    sample_data = [
        ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
        ['String Variable', 'string_tag', 'Holding Register', '30010_10', 'String', '', '', '', '4', ''],
        ['Bit Variable', 'bit_tag', 'Holding Register', '30020_0_1', 'Bits', '', '', '', '4', ''],
        ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
    ]
    try:
        f = open(output_file, 'w', newline='', encoding='utf-8') if output_file else sys.stdout
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(sample_data)
        if output_file:
            f.close()
            logging.info(f"Template generated at {output_file}")
    except Exception as e:
        logging.error(f"Error generating template: {e}")

def run_generator(input_file, output=None, manufacturer=None, model=None,
                 protocol='modbusRTU', category='Inverter', forced_write='',
                 template=False):
    """Main execution entry point for generation."""
    if template:
        generate_template(output)
        return

    if not input_file:
        logging.error("Input file is required.")
        return
    if not manufacturer or not model:
        logging.error("--manufacturer and --model are required.")
        return

    generator = Generator()
    try:
        with open(input_file, mode='r', encoding='utf-8-sig') as f:
            try:
                dialect = csv.Sniffer().sniff(f.read(1024), delimiters=";,")
                f.seek(0)
            except Exception:
                f.seek(0)
                dialect = csv.excel
                dialect.delimiter = ','

            reader = csv.DictReader(f, dialect=dialect)
            if reader.fieldnames:
                reader.fieldnames = [n.strip() for n in reader.fieldnames]
            else:
                logging.error("Input CSV is empty or malformed.")
                return

            processed = generator.process_rows(reader)

        out = open(output, 'w', newline='', encoding='utf-8') if output else sys.stdout
        writer = csv.writer(out, delimiter=';', lineterminator='\n')

        # Write Webdyn Header
        writer.writerow([protocol, category, manufacturer, model, forced_write, '', '', '', '', '', ''])

        # Write Data Rows
        for i, r in enumerate(processed, start=1):
            writer.writerow([
                str(i), r['Info1'], r['Info2'], r['Info3'], r['Info4'],
                r['Name'], r['Tag'], r['CoefA'], r['CoefB'], r['Unit'], r['Action']
            ])

        if output:
            out.close()
            logging.info(f"Successfully generated definition at {output}")

    except FileNotFoundError:
        logging.error(f"Input file '{input_file}' not found.")
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='WebdynSunPM Definition Generator')
    parser.add_argument('input_file', nargs='?', help='Simplified input CSV')
    parser.add_argument('-o', '--output', help='Output definition file')
    parser.add_argument('--manufacturer', help='Device manufacturer')
    parser.add_argument('--model', help='Device model')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--forced-write', default='')
    parser.add_argument('--template', action='store_true', help='Generate template input file')

    args = parser.parse_args()
    run_generator(args.input_file, args.output, args.manufacturer, args.model,
                  args.protocol, args.category, args.forced_write, args.template)
