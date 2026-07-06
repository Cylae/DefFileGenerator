#!/usr/bin/env python3
import argparse
import csv
import sys
import logging
import re
import math
import itertools
from typing import Dict, List, Optional, Any, Union, Tuple, Set, Iterator, Iterable
import os
from dataclasses import dataclass

# Named logger for the module
logger = logging.getLogger('DefFileGenerator.def_gen')

def peek_generator(iterable: Optional[Iterable]) -> Tuple[bool, Iterator]:
    """
    Checks if an iterable is non-empty without fully consuming it.
    Returns (has_data, original_iterator).
    """
    if iterable is None:
        return False, iter([])
    it = iter(iterable)
    try:
        first = next(it)
    except StopIteration:
        return False, iter([])
    return True, itertools.chain([first], it)

# Pre-compiled regex patterns for optimization
RE_TYPE_NUMERIC = re.compile(r'^([UI](8|16|32|64)|F(32|64))(_(W|B|WB))?$', re.IGNORECASE)
RE_TYPE_STR_CONV = re.compile(r'^STR(\d+)$', re.IGNORECASE)
RE_ADDR_STRING = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_BITS = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)_(\d+)_(\d+)$', re.IGNORECASE)
RE_ADDR_INT = re.compile(r'^([0-9A-F]+|0x[0-9A-F]+|[0-9A-F]+h|-?\d+)$', re.IGNORECASE)
RE_COUNT_16_8 = re.compile(r'^([UI](16|8)(_(W|B|WB))?|BITS)$', re.IGNORECASE)
RE_COUNT_32 = re.compile(r'^([UI]32(_(W|B|WB))?|F32(_(W|B|WB))?|IP)$', re.IGNORECASE)
RE_COUNT_64 = re.compile(r'^([UI]64(_(W|B|WB))?|F64(_(W|B|WB))?)$', re.IGNORECASE)

_CLEAN_TYPE_RE = re.compile(r'[^a-z0-9_]+')

@dataclass
class GeneratorConfig:
    input_file: Optional[str] = None
    output: Optional[str] = None
    manufacturer: str = 'Manufacturer'
    model: str = 'Model'
    protocol: str = 'modbusRTU'
    category: str = 'Inverter'
    forced_write: str = ''
    template: bool = False
    template_mode: str = 'input'
    address_offset: int = 0

class Generator:
    def __init__(self) -> None:
        self.register_type_map = {
            'coil': '1',
            'coils': '1',
            'discrete input': '2',
            'discrete': '2',
            'discrete register': '2',
            'discrete registers': '2',
            'holding register': '3',
            'holding': '3',
            'input register': '4',
            'input': '4'
        }
        self.allowed_actions = ['0', '1', '2', '4', '6', '7', '8', '9']

    @staticmethod
    def sanitize_csv_field(field: Any) -> str:
        """Prevents CSV formula injection by escaping leading dangerous characters."""
        s = str(field) if field is not None else ""
        if s.startswith(('=', '+', '-', '@', '\t', '\r')):
            return "'" + s
        return s

    @staticmethod
    def normalize_type(dtype: Any) -> str:
        if not dtype: return 'U16'
        t = str(dtype).lower().strip()
        suffix = ''
        if any(x in t for x in ['_wb', 'swap', 'big endian']): suffix = '_WB'
        elif any(x in t for x in ['_b', 'big']): suffix = '_B'
        elif any(x in t for x in ['_w', 'word']): suffix = '_W'

        # Handle "string 20" -> "STR20"
        str_match = re.search(r'string\s*(\d+)', t)
        if str_match:
            return f"STR{str_match.group(1)}{suffix}"

        # Mapping ordered by specificity
        synonyms = [
            (r'unsigned integer 64|unsigned int 64|uint64', 'U64'),
            (r'signed integer 64|signed int 64|int64', 'I64'),
            (r'unsigned integer 32|unsigned int 32|uint32', 'U32'),
            (r'signed integer 32|signed int 32|int32', 'I32'),
            (r'unsigned integer 16|unsigned int 16|uint16', 'U16'),
            (r'signed integer 16|signed int 16|int16', 'I16'),
            (r'unsigned integer 8|unsigned int 8|uint8', 'U8'),
            (r'signed integer 8|signed int 8|int8', 'I8'),
            (r'float64|double', 'F64'),
            (r'float32|float', 'F32'),
            (r'string', 'STRING'),
        ]
        for pattern, replacement in synonyms:
            if re.search(pattern, t):
                return f"{replacement}{suffix}"

        if t.startswith('str') and t[3:].isdigit():
            return t.upper()

        t = _CLEAN_TYPE_RE.sub('', t)
        return t.upper() if t else 'U16'

    @staticmethod
    def validate_type(dtype: str) -> bool:
        dtype_upper = str(dtype).upper()
        if dtype_upper in ['STRING', 'BITS', 'IP', 'IPV6', 'MAC']:
            return True
        if RE_TYPE_NUMERIC.match(dtype_upper) or RE_TYPE_STR_CONV.match(dtype_upper):
            return True
        return False

    @staticmethod
    def normalize_address_val(addr_part: Any) -> str:
        addr_part = str(addr_part).strip()
        addr_part = re.sub(r'(?<=\d),(?=\d{3}(?!\d))', '', addr_part)
        if not addr_part: return ""
        if addr_part.lower().startswith('0x'):
            try: return str(int(addr_part, 16))
            except ValueError: return addr_part
        elif addr_part.lower().endswith('h'):
            try: return str(int(addr_part[:-1], 16))
            except ValueError: return addr_part
        try: return str(int(addr_part, 0))
        except ValueError: pass
        if re.match(r'^[0-9A-Fa-f]+$', addr_part):
            try: return str(int(addr_part, 16))
            except ValueError: pass
        return addr_part

    @staticmethod
    def validate_address(address: str, dtype: str) -> bool:
        dtype_upper = dtype.upper()
        if RE_TYPE_STR_CONV.match(dtype_upper):
            dtype_upper = 'STRING'

        if dtype_upper == 'STRING':
            if not RE_ADDR_STRING.match(address): return False
        elif dtype_upper == 'BITS':
            if not RE_ADDR_BITS.match(address): return False
        else:
            if not RE_ADDR_INT.match(address): return False

        try:
            base_addr_str = Generator.normalize_address_val(address.split('_')[0])
            base_addr = int(base_addr_str)
            if not (0 <= base_addr <= 65535):
                logger.warning(f"Address {base_addr} is out of standard Modbus range (0-65535).")
                return False
            return True
        except (ValueError, IndexError):
            return False

    @staticmethod
    def apply_address_offset(address: Any, offset: int, line_num: Optional[int] = None, name: Optional[str] = None) -> str:
        if not address: return ""
        parts = str(address).split('_')
        norm_parts = [Generator.normalize_address_val(p) for p in parts]
        try:
            base_addr = int(norm_parts[0]) + offset
            if base_addr < 0:
                msg = f"Address offset {offset} results in negative address {base_addr}"
                if name: msg += f" for '{name}'"
                if line_num: logger.warning(f"Line {line_num}: {msg}")
                else: logger.warning(msg)
            norm_parts[0] = str(base_addr)
        except (ValueError, IndexError): pass
        return '_'.join(norm_parts)

    @staticmethod
    def get_register_count(dtype: str, address: str) -> int:
        dtype_upper = dtype.upper()
        if RE_COUNT_16_8.match(dtype_upper): return 1
        elif RE_COUNT_32.match(dtype_upper): return 2
        elif RE_COUNT_64.match(dtype_upper): return 4
        elif dtype_upper == 'MAC': return 3
        elif dtype_upper == 'IPV6': return 8
        elif dtype_upper == 'STRING' or RE_TYPE_STR_CONV.match(dtype_upper):
            try:
                parts = address.split('_')
                length = int(parts[1]) if len(parts) > 1 else 0
                return math.ceil(length / 2)
            except (IndexError, ValueError): return 0
        return 1

    def _check_address_overlap(self, info1: str, address: str, dtype: str, name: str, line_num: int, address_usage: Dict[str, Any], warned_lines: Set[Tuple[int, int]], treat_as_error: bool = False) -> bool:
        try:
            base_addr_str = Generator.normalize_address_val(address.split('_')[0])
            start_addr = int(base_addr_str)
            reg_count = self.get_register_count(dtype, address)
            end_addr = start_addr + reg_count - 1

            if info1 not in address_usage:
                address_usage[info1] = {'intervals': [], 'max_len': 0}

            usage = address_usage[info1]
            intervals = usage['intervals']
            max_len = usage['max_len']

            import bisect
            idx = bisect.bisect_left(intervals, (start_addr, -1, -1, '', ''))
            overlap_detected = False

            # Check right
            for j in range(idx, len(intervals)):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if u_start > end_addr: break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if dtype.upper() == 'BITS' and u_type == 'BITS' and start_addr == u_start: continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        msg = f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line})."
                        if treat_as_error: logger.error(msg)
                        else: logger.warning(msg)
                        warned_lines.add(warn_key)

            # Check left
            for j in range(idx - 1, -1, -1):
                u_start, u_end, u_line, u_name, u_type = intervals[j]
                if start_addr - u_start > max_len: break
                if max(start_addr, u_start) <= min(end_addr, u_end):
                    if dtype.upper() == 'BITS' and u_type == 'BITS' and start_addr == u_start: continue
                    overlap_detected = True
                    warn_key = tuple(sorted((line_num, u_line)))
                    if warn_key not in warned_lines:
                        msg = f"Line {line_num}: Address overlap detected for '{name}' at {max(start_addr, u_start)}. Overlaps with '{u_name}' (Line {u_line})."
                        if treat_as_error: logger.error(msg)
                        else: logger.warning(msg)
                        warned_lines.add(warn_key)

            bisect.insort(intervals, (start_addr, end_addr, line_num, name, dtype.upper()))
            if reg_count > max_len: usage['max_len'] = reg_count
            return overlap_detected
        except (ValueError, IndexError):
            return False

    def validate_csv(self, filepath: str, strict_overlap: bool = False) -> bool:
        if not os.path.exists(filepath):
            logger.error(f"File not found: {filepath}")
            return False

        valid = True
        seen_tags = {}
        address_usage = {}
        warned_lines = set()

        try:
            with open(filepath, 'rb') as f:
                header_bytes = f.read(4)
                encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

            with open(filepath, 'r', encoding=encoding) as f:
                snippet = f.read(2048)
                f.seek(0)
                is_webdyn = ';' in snippet and ('modbusRTU' in snippet or 'modbusTCP' in snippet)

                if is_webdyn:
                    reader = csv.reader(f, delimiter=';')
                    next(reader, None) # Skip global header
                    for line_num, row in enumerate(reader, start=2):
                        if not row or all(not c.strip() for c in row): continue
                        if row[0].startswith('#'): continue
                        if len(row) < 11:
                            logger.warning(f"Line {line_num}: Insufficient columns. Skipping.")
                            continue
                        info1, addr, dtype, name, tag = row[1], row[2], row[3], row[5], row[6]
                        if tag:
                            if tag in seen_tags:
                                logger.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (Previously at line {seen_tags[tag]}).")
                                valid = False
                            else: seen_tags[tag] = line_num
                        if not self.validate_address(addr, dtype): valid = False
                        if self._check_address_overlap(info1, addr, dtype, name, line_num, address_usage, warned_lines, treat_as_error=strict_overlap):
                            if strict_overlap: valid = False
                else:
                    reader = csv.DictReader(f)
                    for line_num, row in enumerate(reader, start=2):
                        if not any(row.values()): continue
                        name, addr, dtype, tag = row.get('Name'), row.get('Address'), row.get('Type'), row.get('Tag')
                        if not name and not addr: continue
                        if tag:
                            if tag in seen_tags:
                                logger.error(f"Line {line_num}: Fatal Error - Duplicate Tag '{tag}' (Previously at line {seen_tags[tag]}).")
                                valid = False
                            else: seen_tags[tag] = line_num
                        if addr and dtype:
                            if not self.validate_address(addr, dtype): valid = False
                            if self._check_address_overlap('3', addr, dtype, name or 'Unknown', line_num, address_usage, warned_lines, treat_as_error=strict_overlap):
                                if strict_overlap: valid = False
            return valid
        except Exception as e:
            logger.error(f"Error during validation: {e}")
            return False

    @staticmethod
    def _parse_numeric(val: Any, default: float = 0.0) -> float:
        if val is None or str(val).strip() == '': return default
        s = str(val).strip()
        if '/' in s:
            try:
                p = s.split('/')
                return float(p[0]) / float(p[1])
            except (ValueError, ZeroDivisionError, IndexError): return default
        s = s.replace(',', '.') if ',' in s and '.' not in s else s.replace(',', '') if ',' in s else s
        try: return float(s)
        except ValueError: return default

    @staticmethod
    def _calculate_coefficients(factor_str: Any, offset_str: Any, scale_factor_str: Any) -> Tuple[str, str]:
        f = Generator._parse_numeric(factor_str, 1.0)
        o = Generator._parse_numeric(offset_str, 0.0)
        try: sv = int(float(scale_factor_str)) if scale_factor_str else 0
        except ValueError: sv = 0
        return "{:.6f}".format(f * (10 ** sv)), "{:.6f}".format(o)

    def _process_name_and_tag(self, name: str, tag: str, line_num: int, seen_names: Dict[str, int], seen_tags: Dict[str, int]) -> str:
        if name:
            if name in seen_names: logger.warning(f"Line {line_num}: Duplicate Name '{name}' detected. Previous at line {seen_names[name]}.")
            else: seen_names[name] = line_num
        if not tag and name:
            bt = re.sub(r'[^a-z0-9_]', '', name.lower().replace(' ', '_'))
            bt = re.sub(r'_+', '_', bt).strip('_')
            if not bt or not bt[0].isalpha(): bt = f"v_{bt}" if bt else "var"
            tag = bt
            c = 1
            while tag in seen_tags:
                tag = f"{bt}_{c}"
                c += 1
        if tag:
            if tag in seen_tags:
                if seen_tags[tag] != line_num: # Avoid warning about itself
                     logger.warning(f"Line {line_num}: Duplicate Tag '{tag}' detected. Previous at line {seen_tags[tag]}.")
            else: seen_tags[tag] = line_num
        return tag

    def _determine_info1(self, rts: str, line_num: int = 0) -> str:
        if rts is None: return '3'
        lt = str(rts).lower().strip()
        if not lt: return '3'
        if lt in self.register_type_map: return self.register_type_map[lt]
        if lt in ['1', '2', '3', '4']: return lt
        if line_num: logger.warning(f"Line {line_num}: Unknown RegisterType '{rts}'. Defaulting to 3.")
        return '3'

    def process_rows(self, rows: Iterable[Dict[str, Any]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        sn, st, au, wl = {}, {}, {}, set()
        for ln, r in enumerate(rows, start=2):
            if not any(v for v in r.values() if v): continue
            nr = {k.lower().strip(): (str(v).strip() if v is not None else '') for k, v in r.items()}
            name, tag, rts, addr = nr.get('name', ''), nr.get('tag', ''), nr.get('registertype', ''), nr.get('address', '')
            dtr, fct, off, unt = nr.get('type', ''), nr.get('factor', ''), nr.get('offset', ''), nr.get('unit', '')
            act, sf = nr.get('action', ''), nr.get('scalefactor', '')

            if not name and not addr:
                logger.warning(f"Line {ln}: Skipping row with missing Name and Address.")
                continue
            dtype = self.normalize_type(dtr)
            if not self.validate_type(dtype):
                logger.warning(f"Line {ln}: Invalid Type '{dtr}' (normalized to '{dtype}'). Skipping.")
                continue
            if RE_TYPE_STR_CONV.match(dtype) and '_' not in addr:
                addr = f"{addr}_{RE_TYPE_STR_CONV.match(dtype).group(1)}"
            addr = Generator.apply_address_offset(addr, address_offset, ln, name)
            if not self.validate_address(addr, dtype):
                logger.warning(f"Line {ln}: Invalid Address '{addr}' for Type '{dtype}'. Skipping.")
                continue
            tag = self._process_name_and_tag(name, tag, ln, sn, st)
            info1 = self._determine_info1(rts, ln)
            self._check_address_overlap(info1, addr, dtype, name, ln, au, wl)
            ca, cb = self._calculate_coefficients(fct, off, sf)

            act_s = str(act).strip().upper()
            if not act_s: norm_act = '4' if info1 in ['2', '4'] else '1'
            elif act_s in ['R', 'READ', 'RO', 'READ-ONLY', 'READ ONLY', '4']: norm_act = '4'
            elif act_s in ['RW', 'W', 'WRITE', 'READ/WRITE', 'READ-WRITE', 'R/W', 'WO', 'WRITE-ONLY', 'WRITE ONLY', '1']: norm_act = '1'
            elif act_s in self.allowed_actions: norm_act = act_s
            else: norm_act = '4' if info1 in ['2', '4'] else '1'

            yield {'Info1': info1, 'Info2': addr, 'Info3': dtype.upper(), 'Info4': '', 'Name': name, 'Tag': tag, 'CoefA': ca, 'CoefB': cb, 'Unit': unt, 'Action': norm_act}

    @staticmethod
    def write_output_csv(output: Union[str, Any, None], processed_rows: Iterable[Dict[str, Any]], manufacturer: str, model: str,
                        protocol: str = 'modbusRTU', category: str = 'Inverter', forced_write: str = '') -> None:
        type_counts = {'1': 0, '2': 0, '3': 0, '4': 0}
        type_labels = {'1': 'Coils', '2': 'Discrete', '3': 'Holding', '4': 'Input'}
        outfile = None
        try:
            if isinstance(output, str): outfile = open(output, 'w', newline='', encoding='utf-8-sig')
            elif output is None: outfile = sys.stdout
            else: outfile = output

            writer = csv.writer(outfile, delimiter=';', lineterminator='\n')
            writer.writerow([Generator.sanitize_csv_field(protocol), Generator.sanitize_csv_field(category),
                            Generator.sanitize_csv_field(manufacturer), Generator.sanitize_csv_field(model),
                            Generator.sanitize_csv_field(forced_write), '', '', '', '', '', ''])

            last_index = 0
            for index, row in enumerate(processed_rows, start=1):
                last_index = index
                type_counts[row['Info1']] = type_counts.get(row['Info1'], 0) + 1
                writer.writerow([str(index), row['Info1'], row['Info2'], row['Info3'], row['Info4'],
                                row['Name'], row['Tag'], row['CoefA'], row['CoefB'], row['Unit'], row['Action']])

            summary = ", ".join([f"{type_labels.get(k, k)}: {v}" for k, v in type_counts.items() if v > 0])
            if summary: logger.info(f"Generated {last_index} registers ({summary})")
        except Exception as e:
            logger.error(f"Error writing output CSV: {e}")
        finally:
            if isinstance(output, str) and outfile: outfile.close()

def generate_template(output_file: Optional[str], mode: str = 'input') -> None:
    if mode == 'definition':
        headers = ['#Index', 'Info1', 'Info2', 'Info3', 'Info4', 'Name', 'Tag', 'CoefA', 'CoefB', 'Unit', 'Action']
        rows = [
            ['1', '3', '40001', 'U16', '', 'Active Power', 'active_power', '1.000000', '0.000000', 'W', '4'],
            ['2', '3', '40002', 'U16', '', 'Voltage', 'voltage', '0.100000', '0.000000', 'V', '4']
        ]
    else:
        headers = ['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor']
        rows = [
            ['Example Variable', 'example_tag', 'Holding Register', '30001', 'U16', '1', '0', 'V', '4', '0'],
            ['Convenience String', 'str_tag', 'Holding Register', '30030', 'STR20', '', '', '', '4', '']
        ]
    try:
        f = open(output_file, 'w', newline='', encoding='utf-8') if output_file else sys.stdout
        writer = csv.writer(f, delimiter=';' if mode == 'definition' else ',')
        if mode == 'definition':
            # Prepend Webdyn header
            writer.writerow(['modbusRTU', 'Inverter', 'SampleManufacturer', 'SampleModel', '', '', '', '', '', '', ''])

        writer.writerow(headers)
        writer.writerows(rows)
        if output_file: f.close()
    except Exception as e:
        logger.error(f"Error generating template: {e}")

def run_generator(config: GeneratorConfig, input_data: Optional[Iterable[Dict[str, Any]]] = None) -> None:
    generator = Generator()
    if config.template:
        generate_template(config.output, mode=config.template_mode)
        return

    if input_data is None:
        if not config.input_file or not os.path.exists(config.input_file):
            logger.error(f"Input file not found: {config.input_file}")
            return
        try:
            with open(config.input_file, mode='rb') as f:
                hb = f.read(4)
                enc = 'utf-16' if hb.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
            with open(config.input_file, mode='r', encoding=enc) as cf:
                snip = cf.read(2048); cf.seek(0)
                try: dial = csv.Sniffer().sniff(snip, delimiters=";,")
                except csv.Error: dial = csv.excel
                input_data = list(csv.DictReader(cf, dialect=dial))
        except Exception as e:
            logger.error(f"Error reading input file: {e}")
            return

    processed = generator.process_rows(input_data, config.address_offset)
    Generator.write_output_csv(config.output, processed, config.manufacturer, config.model,
                               config.protocol, config.category, config.forced_write)

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='WebdynSunPM Modbus definition generator.')
    parser.add_argument('input_file', nargs='?')
    parser.add_argument('-o', '--output')
    parser.add_argument('--manufacturer', default='Manufacturer')
    parser.add_argument('--model', default='Model')
    parser.add_argument('--template', action='store_true')
    parser.add_argument('--template-mode', choices=['input', 'definition'], default='input')
    args = parser.parse_args()
    run_generator(GeneratorConfig(input_file=args.input_file, output=args.output, manufacturer=args.manufacturer,
                                  model=args.model, template=args.template, template_mode=args.template_mode))
