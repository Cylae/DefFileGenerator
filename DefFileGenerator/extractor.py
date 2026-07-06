#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import itertools
import zipfile
from typing import Dict, List, Any, Iterator, Optional, Iterable, Union, Tuple

# Named logger
logger = logging.getLogger('DefFileGenerator.extractor')

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from defusedxml import ElementTree as ET
    HAS_DEFUSEDXML = True
except ImportError:
    HAS_DEFUSEDXML = False

# Import Generator and peek_generator from def_gen
try:
    from DefFileGenerator.def_gen import Generator, peek_generator
except ImportError:
    try:
        from def_gen import Generator, peek_generator
    except ImportError:
        Generator = None
        def peek_generator(iterable: Optional[Iterable]) -> Tuple[bool, Iterator]:
            if iterable is None: return False, iter([])
            it = iter(iterable)
            try:
                first = next(it)
                return True, itertools.chain([first], it)
            except StopIteration:
                return False, iter([])

class Extractor:
    COLUMN_MAPPING: Dict[str, List[str]] = {
        'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
        'Address': ['address', 'addr', 'offset', 'register', 'reg'],
        'Name': ['name', 'description', 'parameter', 'variable', 'signal', 'signal name'],
        'Type': ['data type', 'datatype', 'type', 'format'],
        'Unit': ['unit', 'units'],
        'Tag': ['tag'],
        'Action': ['action', 'access'],
        'Factor': ['scale', 'factor', 'multiplier', 'ratio'],
        'Offset': ['offset', 'bias', 'coefficient b'],
        'ScaleFactor': ['scalefactor', 'scale factor'],
        'Length': ['length', 'len', 'size', 'count', 'quantity'],
        'StartBit': ['startbit', 'bit offset', 'bit', 'start']
    }

    def __init__(self, mapping: Optional[Dict[str, str]] = None) -> None:
        self.mapping = mapping or {}

    @staticmethod
    def normalize_type(t: Any) -> str:
        if Generator:
            return Generator.normalize_type(t)
        return str(t).upper() if t else 'U16'

    def extract_from_excel(self, filepath: str, sheet_name: Optional[str] = None) -> Iterator[Iterator[Dict[str, Any]]]:
        if not HAS_OPENPYXL:
            logger.error("openpyxl is required for Excel extraction.")
            return

        wb = None
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            sheets = [wb[sheet_name]] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets

            for ws in sheets:
                def sheet_generator(ws_obj=ws) -> Iterator[Dict[str, Any]]:
                    rows = ws_obj.iter_rows(values_only=True)
                    try:
                        header_row = next(rows)
                    except StopIteration:
                        return

                    headers = [str(h).strip() if h is not None else "" for h in header_row]
                    for row in rows:
                        if any(cell is not None and str(cell).strip() for cell in row):
                            yield {headers[i]: cell for i, cell in enumerate(row) if i < len(headers)}

                yield sheet_generator()
        except Exception as e:
            logger.error(f"Error extracting from Excel {filepath}: {e}")
        finally:
            if wb: wb.close()

    def extract_from_pdf(self, filepath: str, pages: Optional[Union[List[int], str]] = None) -> Iterator[Iterator[Dict[str, Any]]]:
        if not HAS_PDFPLUMBER:
            logger.error("pdfplumber is required for PDF extraction.")
            return

        try:
            with pdfplumber.open(filepath) as pdf:
                target_pages = []
                if pages:
                    requested = [int(p.strip()) for p in pages.split(',')] if isinstance(pages, str) else pages
                    for p in requested:
                        idx = p - 1
                        if 0 <= idx < len(pdf.pages):
                            target_pages.append(pdf.pages[idx])
                        else:
                            logger.warning(f"Page {p} out of range.")
                else:
                    target_pages = pdf.pages

                for page in target_pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or len(table) < 2: continue

                        def table_generator(t=table) -> Iterator[Dict[str, Any]]:
                            headers = [str(c).replace('\n', ' ').strip() if c else "" for c in t[0]]
                            for row in t[1:]:
                                rd = {headers[i]: str(cell).replace('\n', ' ').strip() if cell else ""
                                      for i, cell in enumerate(row) if i < len(headers)}
                                if any(v.strip() for v in rd.values()):
                                    yield rd

                        # Materialize to ensure data is read while PDF is open
                        yield iter(list(table_generator()))
        except Exception as e:
            logger.error(f"Error extracting from PDF {filepath}: {e}")

    def extract_from_csv(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:
        def csv_table_generator() -> Iterator[Dict[str, Any]]:
            try:
                with open(filepath, 'rb') as f:
                    hb = f.read(4)
                    enc = 'utf-16' if hb.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

                with open(filepath, 'r', encoding=enc) as f:
                    snip = f.read(2048); f.seek(0)
                    try:
                        dial = csv.Sniffer().sniff(snip, delimiters=";,")
                        delimiter = dial.delimiter
                    except csv.Error:
                        delimiter = ';' if ';' in snip else ','

                    reader = csv.DictReader(f, delimiter=delimiter)
                    for row in reader:
                        if any(v.strip() for v in row.values() if v):
                            yield dict(row)
            except Exception as e:
                logger.error(f"Error extracting from CSV {filepath}: {e}")

        yield csv_table_generator()

    def extract_from_xml(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:
        if not HAS_DEFUSEDXML:
            logger.error("defusedxml is required for secure XML parsing.")
            return

        def xml_generator() -> Iterator[Dict[str, Any]]:
            try:
                tree = ET.parse(filepath)
                root = tree.getroot()
                seen = set()
                for elem in root.iter():
                    row = {child.tag: child.text.strip() for child in elem if len(child) == 0 and child.text}
                    if len(row) >= 2:
                        js = json.dumps(row, sort_keys=True)
                        if js not in seen:
                            seen.add(js)
                            yield row
            except Exception as e:
                logger.error(f"Error extracting from XML {filepath}: {e}")

        yield xml_generator()

    def map_and_clean(self, tables: Optional[Iterable[Iterable[Dict[str, Any]]]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        if not tables: return

        for table in tables:
            has_rows, table_iter = peek_generator(table)
            if not has_rows: continue

            # Peek first 50 rows to detect columns
            buffer = []
            it = iter(table_iter)
            try:
                for _ in range(50):
                    buffer.append(next(it))
            except StopIteration:
                pass

            if not buffer: continue
            all_keys = set().union(*(r.keys() for r in buffer))
            col_map = {}
            used_src = set()

            # 1. Manual mapping
            for target, src in self.mapping.items():
                if src in all_keys:
                    col_map[target] = src
                    used_src.add(src)

            # 2. Heuristic detection
            order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'Offset', 'ScaleFactor', 'Length', 'StartBit']
            for target in order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src in all_keys:
                    if src in used_src: continue
                    s_low = str(src).lower().strip()
                    if s_low in patterns or any(p in s_low for p in patterns):
                        col_map[target] = src
                        used_src.add(src)
                        break

            def process_row(r: Dict[str, Any]) -> Optional[Dict[str, Any]]:
                nr = {target: r.get(src) for target, src in col_map.items()}
                if not nr.get('Name') and not nr.get('Address'): return None

                dtype = self.normalize_type(nr.get('Type', 'U16'))
                nr['Type'] = dtype
                addr = str(nr.get('Address', '')).strip()
                sbit = str(r.get(col_map.get('StartBit', ''))).strip() if col_map.get('StartBit') else ''
                slen = str(r.get(col_map.get('Length', ''))).strip() if col_map.get('Length') else ''

                if dtype == 'BITS' and sbit and '_' not in addr:
                    addr = f"{addr}_{sbit}_{slen or '1'}"
                elif (dtype == 'STRING' or dtype.startswith('STR')) and slen and '_' not in addr:
                    addr = f"{addr}_{slen}"

                if Generator:
                    nr['Address'] = Generator.apply_address_offset(addr, address_offset)
                else:
                    nr['Address'] = addr

                if nr.get('Factor') is not None and Generator:
                    nr['Factor'] = str(Generator._parse_numeric(nr['Factor'], 1.0))

                if not nr.get('RegisterType'):
                    nr['RegisterType'] = 'Holding Register'
                return nr

            for r in buffer:
                p = process_row(r)
                if p: yield p
            for r in it:
                p = process_row(r)
                if p: yield p

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information.')
    parser.add_argument('input_file'); parser.add_argument('-o', '--output')
    parser.add_argument('--mapping'); parser.add_argument('--sheet'); parser.add_argument('--pages')
    parser.add_argument('--address-offset', type=int, default=0)
    args = parser.parse_args()

    mapping = {}
    if args.mapping and os.path.exists(args.mapping):
        with open(args.mapping, 'r') as f: mapping = json.load(f)

    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()

    if ext in ['.xlsx', '.xlsm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, args.pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logger.error(f"Unsupported: {ext}"); sys.exit(1)

    mapped = list(extractor.map_and_clean(raw, args.address_offset))
    out = open(args.output, 'w', newline='', encoding='utf-8') if args.output else sys.stdout
    writer = csv.DictWriter(out, fieldnames=['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor'], extrasaction='ignore')
    writer.writeheader(); writer.writerows(mapped)
    if args.output: out.close()

if __name__ == "__main__":
    main()
