import time
import logging
import os
import subprocess
from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor

def run_gigantic_battery():
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    results = {}

    # 1. Performance test with 5000 rows CSV
    start = time.time()
    config = GeneratorConfig(
        input_file='stress_test_data/stress.csv',
        output='stress_test_data/output.csv',
        manufacturer='StressCorp',
        model='Gigantic3000',
        address_offset=0
    )
    run_generator(config)
    results['csv_performance_5000_rows'] = time.time() - start

    # 2. Address offset with negative result
    start = time.time()
    config.address_offset = -100
    config.output = 'stress_test_data/output_offset.csv'
    run_generator(config)
    results['address_offset_stress'] = time.time() - start

    # 3. Extraction stress tests
    extractor = Extractor()

    # Excel
    if os.path.exists('stress_test_data/stress.xlsx'):
        start = time.time()
        raw = extractor.extract_from_excel('stress_test_data/stress.xlsx')
        mapped = extractor.map_and_clean(raw)
        results['excel_extraction_5000_rows'] = time.time() - start
        results['excel_mapped_count'] = len(mapped)

    # XML
    if os.path.exists('stress_test_data/stress.xml'):
        start = time.time()
        raw = extractor.extract_from_xml('stress_test_data/stress.xml')
        mapped = extractor.map_and_clean(raw)
        results['xml_extraction_5000_rows'] = time.time() - start
        results['xml_mapped_count'] = len(mapped)

    # PDF
    if os.path.exists('stress_test_data/stress.pdf'):
        start = time.time()
        raw = extractor.extract_from_pdf('stress_test_data/stress.pdf')
        mapped = extractor.map_and_clean(raw)
        results['pdf_extraction_100_rows'] = time.time() - start
        results['pdf_mapped_count'] = len(mapped)

    # 4. Security test (XXE)
    try:
        raw = extractor.extract_from_xml('stress_test_data/xxe.xml')
        # If it didn't crash or log error correctly, it's a concern
        results['xxe_security_test'] = "Executed (Check logs for security blocks)"
    except Exception as e:
        results['xxe_security_test'] = f"Caught Exception: {e}"

    # 5. Type normalization stress
    generator = Generator()
    test_types = [
        'unsigned int 16 big endian', 'float32 swap', 'signed integer 64 word swap',
        'uint32 big', 'U16_WB', 'STR50', 'IP', 'IPV6', 'MAC', 'bits', 'uint8', 'double'
    ]
    start = time.time()
    for _ in range(1000): # Repeat to see if it's slow
        for t in test_types:
            generator.normalize_type(t)
    results['type_normalization_12000_calls'] = time.time() - start

    return results

if __name__ == "__main__":
    res = run_gigantic_battery()
    print("\n--- GIGANTIC BATTERY RESULTS ---")
    for k, v in res.items():
        print(f"{k}: {v}")
