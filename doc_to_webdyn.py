#!/usr/bin/env python3
import argparse
import sys
import os
import logging
import re
import csv
import math

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

# Import Generator from the package
try:
    from DefFileGenerator.def_gen import Generator
except ImportError:
    # Fallback for when running from root without proper package setup
    sys.path.append(os.path.join(os.path.dirname(__file__), 'DefFileGenerator'))
    from def_gen import Generator

COLUMN_MAPPING = {
    'RegisterType': ['register type', 'reg type', 'modbus type', 'registertype'],
    'Address': ['address', 'addr', 'offset', 'register', 'reg'],
    'Name': ['name', 'description', 'parameter', 'variable', 'signal'],
    'Type': ['data type', 'datatype', 'type', 'format'],
    'Unit': ['unit', 'units'],
    'Scale': ['scale', 'factor', 'multiplier', 'ratio'],
    'Action': ['action', 'access']
}

def is_na(val):
    if HAS_PANDAS:
        return pd.isna(val)
    return val is None or val == '' or (isinstance(val, float) and math.isnan(val))

class MockDF:
    def __init__(self, rows, columns):
        self.rows = rows
        self.columns = columns
    def iterrows(self):
        for i, row in enumerate(self.rows):
            yield i, row

def main():
    parser = argparse.ArgumentParser(description='WebdynSunPM Documentation Parser')
    parser.add_argument('input_file', help='Path to documentation (PDF, Excel, CSV, XML)')
    parser.add_argument('--manufacturer', required=True)
    parser.add_argument('--model', required=True)
    parser.add_argument('-o', '--output', help='Output definition file')
    parser.add_argument('--protocol', default='modbusRTU')
    parser.add_argument('--category', default='Inverter')
    parser.add_argument('--sheet', help='Excel sheet name')
    parser.add_argument('-v', '--verbose', action='store_true')

    args = parser.parse_args()
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format='%(levelname)s: %(message)s')

    logging.info(f"Processing {args.input_file}")

    ext = os.path.splitext(args.input_file)[1].lower()
    dataframes = []

    if ext == '.csv':
        dataframes = load_csv(args.input_file)
    elif ext in ['.xlsx', '.xls']:
        dataframes = load_excel(args.input_file, args.sheet)
    elif ext == '.xml':
        dataframes = load_xml(args.input_file)
    elif ext == '.pdf':
        dataframes = load_pdf(args.input_file)
    else:
        logging.error(f"Unsupported extension: {ext}")
        sys.exit(1)

    if not dataframes:
        logging.error("No data extracted.")
        sys.exit(1)

    generator = Generator()
    all_extracted = []

    for df in dataframes:
        df.columns = [str(c).strip() for c in df.columns]
        col_map = {}
        assigned = set()

        # Priority mapping
        order = ['RegisterType', 'Address', 'Name', 'Type', 'Unit', 'Scale', 'Action']
        for key in order:
            for col in df.columns:
                if col in assigned: continue
                if any(p in str(col).lower() for p in COLUMN_MAPPING[key]):
                    col_map[key] = col
                    assigned.add(col)
                    break

        if 'Address' not in col_map and 'Name' not in col_map: continue

        for _, row in df.iterrows():
            addr_raw = row.get(col_map.get('Address')) if 'Address' in col_map else None
            name_raw = row.get(col_map.get('Name')) if 'Name' in col_map else None

            if is_na(addr_raw) and is_na(name_raw): continue

            # Simple header skip
            if str(addr_raw).lower() in COLUMN_MAPPING['Address']: continue

            addr = str(addr_raw).strip() if not is_na(addr_raw) else ''
            name = str(name_raw) if not is_na(name_raw) else f"Reg {addr}"

            dtype_raw = row.get(col_map.get('Type')) if 'Type' in col_map else 'U16'
            unit_raw = row.get(col_map.get('Unit')) if 'Unit' in col_map else ''
            scale_raw = row.get(col_map.get('Scale')) if 'Scale' in col_map else '1'
            reg_type_raw = row.get(col_map.get('RegisterType')) if 'RegisterType' in col_map else 'Holding'
            action_raw = row.get(col_map.get('Action')) if 'Action' in col_map else '1'

            # Scale factor cleanup
            scale = str(scale_raw)
            if '/' in scale:
                try:
                    p = scale.split('/')
                    scale = str(float(p[0]) / float(p[1]))
                except Exception: scale = '1'

            all_extracted.append({
                'Name': name,
                'RegisterType': str(reg_type_raw),
                'Address': addr,
                'Type': generator.normalize_type(dtype_raw),
                'Factor': scale if scale != 'nan' else '1',
                'Unit': str(unit_raw) if not is_na(unit_raw) else '',
                'Action': generator.normalize_action(action_raw)
            })

    if not all_extracted:
        logging.error("No valid registers found.")
        sys.exit(1)

    # Final generation
    processed = generator.process_rows(all_extracted)
    out_file = args.output or f"{args.manufacturer}_{args.model}_def.csv".lower().replace(' ', '_')

    try:
        with open(out_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';', lineterminator='\n')
            writer.writerow([args.protocol, args.category, args.manufacturer, args.model, '', '', '', '', '', '', ''])
            for i, r in enumerate(processed, 1):
                writer.writerow([str(i), r['Info1'], r['Info2'], r['Info3'], r['Info4'], r['Name'], r['Tag'], r['CoefA'], r['CoefB'], r['Unit'], r['Action']])
        logging.info(f"Generated {out_file}")
    except Exception as e:
        logging.error(f"Failed to write output: {e}")
        sys.exit(1)

def load_csv(filepath):
    if HAS_PANDAS:
        for d in [',', ';', '\t']:
            try:
                df = pd.read_csv(filepath, sep=d, encoding='utf-8-sig')
                if len(df.columns) > 1: return [df]
            except: pass
    else:
        for d in [',', ';', '\t']:
            try:
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    r = csv.DictReader(f, delimiter=d)
                    rows = list(r)
                    if rows: return [MockDF(rows, r.fieldnames)]
            except: pass
    return []

def load_excel(filepath, sheet=None):
    if not HAS_PANDAS: return []
    try:
        if sheet: return [pd.read_excel(filepath, sheet_name=sheet)]
        xl = pd.ExcelFile(filepath)
        return [xl.parse(s) for s in xl.sheet_names]
    except: return []

def load_xml(filepath):
    if not HAS_PANDAS: return []
    try: return [pd.read_xml(filepath)]
    except:
        try: return [pd.read_xml(filepath, parser='etree')]
        except: return []

def load_pdf(filepath):
    if not HAS_PDFPLUMBER: return []
    dfs = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables() or []:
                    if len(table) > 1:
                        headers = [str(h).replace('\n', ' ') if h else f"Col{i}" for i, h in enumerate(table[0])]
                        if HAS_PANDAS: dfs.append(pd.DataFrame(table[1:], columns=headers))
                        else: dfs.append(MockDF([dict(zip(headers, r)) for r in table[1:]], headers))
    except: pass
    return dfs

if __name__ == "__main__":
    main()
