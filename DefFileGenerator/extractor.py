#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import io
import itertools
import zipfile
import itertools
from typing import Dict, List, Any, Iterator, Optional, Iterable, Union, Tuple

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
    try:
        from pdfminer.pdfparser import PDFSyntaxError
        from pdfplumber.utils.exceptions import PdfminerException
        PDF_ERRORS = (PDFSyntaxError, PdfminerException)
    except ImportError:
        PDF_ERRORS = ()
except ImportError:
    HAS_PDFPLUMBER = False
    PDF_ERRORS = ()

try:
    from defusedxml import ElementTree as ET
    from defusedxml.common import EntitiesForbidden, DTDForbidden, ExternalReferenceForbidden
    HAS_DEFUSEDXML = True
    SECURITY_EXCEPTIONS = (EntitiesForbidden, DTDForbidden, ExternalReferenceForbidden)
except ImportError:
    HAS_DEFUSEDXML = False
    SECURITY_EXCEPTIONS = ()

try:
    import xml.etree.ElementTree as ET_STD
    XML_PARSE_ERRORS = (ET_STD.ParseError,)
except ImportError:
    XML_PARSE_ERRORS = ()

def peek_generator(iterable: Iterable[Any]) -> Tuple[bool, Iterator[Any]]:
    """Checks if an iterable is empty without exhausting it."""
    it = iter(iterable)
    try:
        first = next(it)
    except StopIteration:
        return False, iter([])
    return True, itertools.chain([first], it)

try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    # Support local import if running from within the directory
    try:
        from def_gen import Generator
    except ImportError:
        Generator = None

def peek_generator(iterable: Iterable) -> Tuple[bool, Iterator]:
    """Checks if an iterable is empty without exhausting it. Returns (has_data, iterator)."""
    if iterable is None:
        return False, iter([])
    try:
        it = iter(iterable)
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
            logging.error("openpyxl is required for Excel extraction.")
            return iter([])

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

            def gen_all_sheets():
                try:
                    for ws in sheets:
                        def sheet_generator() -> Iterator[Dict[str, Any]]:
                            rows = ws.iter_rows(values_only=True)
                            try:
                                header_row = next(rows)
                            except StopIteration:
                                return

                            headers = [str(h).strip() if h is not None else "" for h in header_row]

                            for row in rows:
                                if any(cell is not None and str(cell).strip() for cell in row):
                                    yield {headers[i]: cell for i, cell in enumerate(row) if i < len(headers)}
                        yield sheet_generator()
                finally:
                    wb.close()
            return gen_all_sheets()

        except (OSError, zipfile.BadZipFile) as e:
            logging.error(f"File IO Error extracting from Excel {filepath}: {e}")
        except (ValueError, TypeError, KeyError) as e:
            logging.error(f"Error extracting from Excel {filepath}: {e}")
        return iter([])

    def extract_from_pdf(self, filepath: str, pages: Optional[Union[int, List[int], str]] = None) -> Iterator[Iterator[Dict[str, Any]]]:
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return iter([])

        # Handle pages as comma-separated string
        if isinstance(pages, str):
            try:
                with pdfplumber.open(filepath) as pdf:
                    target_page_indices = []
                    if pages is None:
                        target_page_indices = range(len(pdf.pages))
                    elif isinstance(pages, int):
                        target_page_indices = [pages - 1]
                    elif isinstance(pages, list):
                        target_page_indices = [int(p) - 1 for p in pages]
                    elif isinstance(pages, str):
                        target_page_indices = [int(p.strip()) - 1 for p in pages.split(',')]

                    for idx in target_page_indices:
                        if idx < 0 or idx >= len(pdf.pages): continue
                        page = pdf.pages[idx]
                        tables = page.extract_tables()
                        for table in tables:
                            if not table or len(table) < 2: continue

                            def table_generator(current_table: List[List[Any]]) -> Iterator[Dict[str, Any]]:
                                headers = [str(c).replace('\n', ' ').strip() if c else "" for c in current_table[0]]
                                for row in current_table[1:]:
                                    row_dict = {}
                                    for i, cell in enumerate(row):
                                        if i < len(headers):
                                            row_dict[headers[i]] = str(cell).replace('\n', ' ').strip() if cell else ""
                                    if any(row_dict.values()):
                                        yield row_dict
                            yield table_generator(table)
            except (OSError,) + PDF_ERRORS as e:
                logging.error(f"File IO Error or PDF Syntax Error extracting from PDF {filepath}: {e}")
            except (ValueError, TypeError, IndexError) as e:
                logging.error(f"Error extracting from PDF {filepath}: {e}")

        except (OSError,) + PDF_ERRORS as e:
            logging.error(f"File IO Error or PDF Syntax Error extracting from PDF {filepath}: {e}")
        except (ValueError, TypeError, IndexError) as e:
            logging.error(f"Error extracting from PDF {filepath}: {e}")

    def extract_from_csv(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:
        def csv_table_generator() -> Iterator[Dict[str, Any]]:
            try:
                with open(filepath, 'rb') as f:
                    header_bytes = f.read(4)
                    encoding = 'utf-16' if header_bytes.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'

                with open(filepath, 'r', encoding=encoding) as f:
                    snippet = f.read(2048)
                    f.seek(0)
                    try:
                        dialect = csv.Sniffer().sniff(snippet, delimiters=";,")
                        delimiter = dialect.delimiter
                    except csv.Error:
                        delimiter = ','
                        for d in [',', ';', '\t']:
                            if d in snippet:
                                delimiter = d
                                break

                    reader = csv.DictReader(f, delimiter=delimiter)
                    for row in reader:
                        if any(val.strip() for val in row.values() if val is not None):
                            yield dict(row)
            except Exception as e:
                logging.error(f"Error extracting from CSV {filepath}: {e}")

        yield csv_table_generator()

    def extract_from_xml(self, filepath: str) -> Iterator[Iterator[Dict[str, Any]]]:
        if not HAS_DEFUSEDXML:
            logging.error("defusedxml is required for secure XML parsing.")
            return iter([])
        try:
            with open(filepath, 'rb') as f:
                tree = ET.parse(f)
                root = tree.getroot()

            def xml_generator() -> Iterator[Dict[str, Any]]:
                seen = set()
                for elem in root.iter():
                    row = {}
                    for child in elem:
                        if len(child) == 0 and child.text:
                            row[child.tag] = child.text.strip()
                    if len(row) >= 2:
                        js = json.dumps(row, sort_keys=True)
                        if js not in seen:
                            seen.add(js)
                            yield row
            yield xml_generator()
        except Exception as e:
            logging.error(f"Error extracting from XML {filepath}: {e}")
            return iter([])

    def map_and_clean(self, tables: Iterable[Iterable[Dict[str, Any]]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        if tables is None:
            return

        for table in tables:
            if table is None:
                continue

            # Since table could be a generator, we need to extract the first few rows
            # to determine column mapping, then process the rest.
            has_data, iterator = peek_generator(table)
            if not has_data:
                continue

            buffer = []
            try:
                for _ in range(50):
                    buffer.append(next(iterator))
            except StopIteration:
                pass

            if not buffer:
                continue

            all_keys = set()
            for row in buffer:
                all_keys.update(row.keys())

            col_map = {}
            used_src_cols = set()
            for target, source in self.mapping.items():
                if source in all_keys:
                    col_map[target] = source
                    used_src_cols.add(source)

            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'Offset', 'ScaleFactor', 'Length', 'StartBit']
            for target in detection_order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in all_keys:
                    if src_col in used_src_cols: continue
                    s_low = str(src_col).lower().strip()
                    if s_low in patterns:
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            for target in detection_order:
                if target in col_map: continue
                patterns = self.COLUMN_MAPPING.get(target, [target.lower()])
                for src_col in all_keys:
                    if src_col in used_src_cols: continue
                    if any(p in str(src_col).lower() for p in patterns):
                        col_map[target] = src_col
                        used_src_cols.add(src_col)
                        break

            def process_row(r: Dict[str, Any]) -> Optional[Dict[str, Any]]:
                new_row = {target: r.get(src_col) for target, src_col in col_map.items()}
                if not new_row.get('Name') and not new_row.get('Address'): return None
                sbit = r.get(col_map.get('StartBit'))
                slen = r.get(col_map.get('Length'))
                sbit = str(sbit).strip() if sbit is not None else ''
                slen = str(slen).strip() if slen is not None else ''
                dtype = self.normalize_type(new_row.get('Type', 'U16'))
                new_row['Type'] = dtype
                addr = str(new_row.get('Address', '')).strip()
                if dtype == 'BITS' and sbit != '' and '_' not in addr:
                    if slen == '': slen = '1'
                    addr = f"{addr}_{sbit}_{slen}"
                elif (dtype == 'STRING' or dtype.startswith('STR')) and slen != '' and '_' not in addr:
                    addr = f"{addr}_{slen}"
                if Generator:
                    new_row['Address'] = Generator.apply_address_offset(addr, address_offset)
                else:
                    new_row['Address'] = addr
                if new_row.get('Factor') is not None:
                    if Generator:
                        new_row['Factor'] = str(Generator._parse_numeric(new_row['Factor'], 1.0))
                if 'RegisterType' not in new_row or not new_row['RegisterType']:
                    new_row['RegisterType'] = 'Holding Register'
                return new_row

            for row in buffer:
                processed = process_row(row)
                if processed: yield processed
            for row in iterator:
                processed = process_row(row)
                if processed: yield processed

def main():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    parser = argparse.ArgumentParser(description='Extract register information.')
    parser.add_argument('input_file'); parser.add_argument('-o', '--output')
    parser.add_argument('--mapping'); parser.add_argument('--sheet'); parser.add_argument('--pages')
    parser.add_argument('--address-offset', type=int, default=0)
    args = parser.parse_args()
    mapping = {}
    if args.mapping:
        with open(args.mapping, 'r') as f: mapping = json.load(f)
    extractor = Extractor(mapping)
    ext = os.path.splitext(args.input_file)[1].lower()
    pages = args.pages
    if ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']: raw = extractor.extract_from_excel(args.input_file, args.sheet)
    elif ext == '.pdf': raw = extractor.extract_from_pdf(args.input_file, pages)
    elif ext == '.csv': raw = extractor.extract_from_csv(args.input_file)
    elif ext == '.xml': raw = extractor.extract_from_xml(args.input_file)
    else: logging.error(f"Unsupported extension: {ext}"); sys.exit(1)
    mapped = list(extractor.map_and_clean(raw, args.address_offset))
    out = open(args.output, 'w', newline='', encoding='utf-8') if args.output else sys.stdout
    writer = csv.DictWriter(out, fieldnames=['Name', 'Tag', 'RegisterType', 'Address', 'Type', 'Factor', 'Offset', 'Unit', 'Action', 'ScaleFactor'], extrasaction='ignore')
    writer.writeheader(); writer.writerows(mapped)
    if args.output: out.close()

if __name__ == "__main__":
    main()
