import csv
import os
import json
import logging
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor

def run_torture_battery():
    logging.basicConfig(level=logging.INFO)
    os.makedirs('torture_test', exist_ok=True)

    # 1. Test ambiguous column detection
    fieldnames = ['Signal Name', 'Register Offset', 'Data Format', 'Measurement Unit', 'Ratio']
    rows = [
        {'Signal Name': 'V_L1', 'Register Offset': '0x7530', 'Data Format': 'float32 swap', 'Measurement Unit': 'V', 'Ratio': '0.1'},
        {'Signal Name': 'Status', 'Register Offset': '30001', 'Data Format': 'BITS', 'Measurement Unit': 'READ', 'Ratio': ''},
        {'Signal Name': 'Serial', 'Register Offset': '30002', 'Data Format': 'STR50', 'Measurement Unit': '4', 'Ratio': ''},
        {'Signal Name': 'Freq', 'Register Offset': '30,050', 'Data Format': 'unsigned int 32 big endian', 'Measurement Unit': '0.01', 'Ratio': 'Hz'}
    ]

    csv_path = 'torture_test/ambiguous.csv'
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    extractor = Extractor()
    raw = extractor.extract_from_csv(csv_path)
    mapped = extractor.map_and_clean(raw)

    print(f"Mapped {len(mapped)} rows from ambiguous columns")
    for row in mapped:
        print(f"  {row.get('Name')} -> {row.get('Address')} ({row.get('Type')})")

    if len(mapped) < 4:
        print("Ambiguous column mapping: FAILED")
    else:
        print("Ambiguous column mapping: SUCCESS")

    # 2. Test large address offset
    config = GeneratorConfig(
        input_file=csv_path,
        output='torture_test/output.csv',
        manufacturer='TortureMfg',
        model='Extreme',
        address_offset=1000000
    )
    run_generator(config)

    # 3. Test STR<n> with offset
    str_data = [{'Name': 'StringTest', 'Address': '100', 'Type': 'STR20'}]
    str_csv = 'torture_test/str_test.csv'
    with open(str_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['Name', 'Address', 'Type'])
        writer.writeheader()
        writer.writerows(str_data)

    config.input_file = str_csv
    config.address_offset = 50
    config.output = 'torture_test/output_str.csv'
    run_generator(config)

    with open('torture_test/output_str.csv', 'r') as f:
        content = f.read()
        if '150_20' in content:
            print("STR<n> expansion with offset: SUCCESS")
        else:
            print("STR<n> expansion with offset: FAILED")

if __name__ == "__main__":
    run_torture_battery()
