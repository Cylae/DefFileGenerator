#!/usr/bin/env python3
import argparse
import csv
import json
import logging
import os
import re
import sys
import io
import zipfile
from typing import Dict, List, Any, Iterator, Optional, Iterable, Union

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

try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    # Support local import if running from within the directory
    try:
        from def_gen import Generator
    except ImportError:
        Generator = None

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

        wb = None
        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
            sheets = [wb[sheet_name]] if sheet_name else wb.worksheets

            for ws in sheets:
                def sheet_generator(ws_obj=ws) -> Iterator[Dict[str, Any]]:
                    rows = ws_obj.iter_rows(values_only=True)
                    try:
                        header_row = next(rows)
                    except StopIteration:
                        return

                    headers = [str(h).strip() if h is not None else "" for h in header_row]

                    for row in rows:
                        # Only yield if row has actual data (not all None/empty)
                        if any(cell is not None and str(cell).strip() for cell in row):
                            yield {headers[i]: cell for i, cell in enumerate(row) if i < len(headers)}

                # We yield a generator for each sheet. We check if it has rows by creating it.
                # To see if it's empty, we would need to peek, but let's yield it.
                # If there are no data rows, it simply yields nothing when iterated.
                yield sheet_generator()

        except (OSError, zipfile.BadZipFile) as e:
            logging.error(f"File IO Error extracting from Excel {filepath}: {e}")
        except (ValueError, TypeError, KeyError) as e:
            logging.error(f"Error extracting from Excel {filepath}: {e}")
        finally:
            if wb:
                wb.close()

    def extract_from_pdf(self, filepath: str, pages: Optional[Union[int, List[int]]] = None) -> Iterator[Iterator[Dict[str, Any]]]:
        if not HAS_PDFPLUMBER:
            logging.error("pdfplumber is required for PDF extraction.")
            return iter([])

        def pdf_tables_generator():
            try:
                with pdfplumber.open(filepath) as pdf:
                    target_pages = pdf.pages if pages is None else [pdf.pages[i-1] for i in (pages if isinstance(pages, list) else [pages])]
                    for page in target_pages:
                        tables = page.extract_tables()
                        logging.debug(f"Found {len(tables)} tables on page {page.page_number}")
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

                            # Since we must keep pdfplumber context open while iterating,
                            # and pdfplumber pages/tables are held in memory anyway,
                            # yielding generators is safe but the entire PDF is open.
                            # A better approach: evaluate the generator immediately if we're yielding it,
                            # but here we're conforming to `Iterator[Iterator[Dict]]`
                            yield table_generator(table)

            except (OSError,) + PDF_ERRORS as e:
                logging.error(f"File IO Error or PDF Syntax Error extracting from PDF {filepath}: {e}")
            except (ValueError, TypeError, IndexError) as e:
                logging.error(f"Error extracting from PDF {filepath}: {e}")

        return pdf_tables_generator()

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

                    for line_num, row in enumerate(reader, start=2):
                        if any(val.strip() for val in row.values() if val is not None):
                            yield dict(row)

            except OSError as e:
                logging.error(f"File IO Error extracting from CSV {filepath}: {e}")
            except csv.Error as e:
                logging.error(f"CSV Parsing Error in {filepath}: {e}")
            except UnicodeError as e:
                logging.error(f"Encoding Error extracting from CSV {filepath}: {e}")
            except (ValueError, TypeError) as e:
                logging.error(f"Unexpected error extracting from CSV {filepath}: {e}")

        # Yield a single table that contains all rows
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

            # Return as a single table
            yield xml_generator()

        except SECURITY_EXCEPTIONS:
            raise
        except (OSError,) + XML_PARSE_ERRORS as e:
            logging.error(f"File IO Error or Parsing Error extracting from XML {filepath}: {e}")
        except (ValueError, TypeError) as e:
            logging.error(f"Error extracting from XML {filepath}: {e}")

    def map_and_clean(self, tables: Iterable[Iterable[Dict[str, Any]]], address_offset: int = 0) -> Iterator[Dict[str, Any]]:
        if not tables:
            return iter([])

        for table in tables:
            if not table: continue

            # Since table could be a generator, we need to extract the first few rows
            # to determine column mapping, then process the rest.
            iterator = iter(table)
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

            # 1. Explicit mapping
            for target, source in self.mapping.items():
                if source in all_keys:
                    col_map[target] = source
                    used_src_cols.add(source)

            # 2. Priority fuzzy matching (Multi-pass)
            detection_order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Action', 'Tag', 'Factor', 'Offset', 'ScaleFactor', 'Length', 'StartBit']

            # Pass 1: Exact or Normalized Exact Match
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

            # Pass 2: Fuzzy Substring Match
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

                # Extract StartBit and Length
                sbit = r.get(col_map.get('StartBit'))
                slen = r.get(col_map.get('Length'))
                sbit = str(sbit).strip() if sbit is not None else ''
                slen = str(slen).strip() if slen is not None else ''

                # Normalize Type
                dtype = self.normalize_type(new_row.get('Type', 'U16'))
                new_row['Type'] = dtype

                # Address normalization/construction
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

                # Factor
                if new_row.get('Factor') is not None:
                    if Generator:
                        new_row['Factor'] = str(Generator._parse_numeric(new_row['Factor'], 1.0))

                if 'RegisterType' not in new_row or not new_row['RegisterType']:
                    new_row['RegisterType'] = 'Holding Register'

                return new_row

            # Process buffer first
            for row in buffer:
                processed = process_row(row)
                if processed:
                    yield processed

            # Process remaining iterator
            for row in iterator:
                processed = process_row(row)
                if processed:
                    yield processed

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

    pages = None
    if args.pages:
        try:
            pages = [int(p.strip()) for p in args.pages.split(',')]
        except ValueError:
            logging.error("Invalid format for --pages. Expected comma-separated integers.")
            sys.exit(1)

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
