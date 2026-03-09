import csv
import random
import os
import json
import io
from DefFileGenerator.def_gen import Generator
from DefFileGenerator.extractor import Extractor

try:
    import pandas as pd
    from openpyxl import Workbook
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

def generate_stress_data(num_rows=1000):
    types = ['U16', 'I32_WB', 'F32', 'STR20', 'BITS', 'U64_B', 'IP', 'IPV6', 'MAC', 'unsigned int 16', 'float32 swap']
    reg_types = ['Holding Register', 'Input Register', 'Coil', 'Discrete Input']
    data = []

    for i in range(num_rows):
        addr = 40000 + i * 2
        # Introduce some hex and thousands separators
        if i % 10 == 0:
            addr_str = hex(addr)
        elif i % 15 == 0:
            addr_str = f"{addr // 1000},{addr % 1000:03d}"
        else:
            addr_str = str(addr)

        # Special case for BITS
        row_type = random.choice(types)
        if row_type == 'BITS':
            addr_str = f"{addr}_0_1"

        data.append({
            'Name': f'Stress Var {i}',
            'Tag': f'stress_tag_{i}' if i % 2 == 0 else '',
            'RegisterType': random.choice(reg_types),
            'Address': addr_str,
            'Type': row_type,
            'Factor': random.choice(['1', '0.1', '1/10', '0.001']),
            'Offset': str(random.randint(0, 100)),
            'Unit': random.choice(['V', 'A', 'W', '°C', '%', '']),
            'Action': random.choice(['R', 'RW', '1', '4', 'Write', '']),
            'ScaleFactor': str(random.randint(-3, 3))
        })

    # Add some edge cases
    # 1. Overlap (Forbidden)
    data.append({
        'Name': 'Overlap Var', 'Address': '40000', 'Type': 'U32', 'RegisterType': 'Holding Register'
    })
    # 2. Duplicate Name
    data.append({
        'Name': 'Stress Var 0', 'Address': '50000', 'Type': 'U16', 'RegisterType': 'Holding Register'
    })
    # 3. Negative address after offset
    data.append({
        'Name': 'Negative Var', 'Address': '10', 'Type': 'U16', 'RegisterType': 'Holding Register'
    })

    return data

def save_csv(data, path):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)

def save_excel(data, path):
    if not HAS_LIBS: return
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)

def save_xml(data, path):
    if not HAS_LIBS: return
    df = pd.DataFrame(data)
    df.to_xml(path, index=False)

def save_pdf(data, path):
    if not HAS_LIBS: return
    from reportlab.lib import colors
    doc = SimpleDocTemplate(path, pagesize=letter)
    elements = []
    # Take a subset for PDF
    subset = data[:100]
    headers = list(subset[0].keys())
    table_data = [headers] + [[str(row[h]) for h in headers] for row in subset]
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(t)
    doc.build(elements)

def generate_xxe_payload(path):
    # Malformed XML for security testing
    payload = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE root [
  <!ENTITY xxe SYSTEM "file:///etc/passwd">
]>
<root>
  <row>
    <Name>&xxe;</Name>
    <Address>100</Address>
    <Type>U16</Type>
  </row>
</root>"""
    with open(path, 'w') as f:
        f.write(payload)

if __name__ == "__main__":
    os.makedirs('stress_test_data', exist_ok=True)
    data = generate_stress_data(5000)
    save_csv(data, 'stress_test_data/stress.csv')
    if HAS_LIBS:
        save_excel(data, 'stress_test_data/stress.xlsx')
        save_xml(data, 'stress_test_data/stress.xml')
        save_pdf(data, 'stress_test_data/stress.pdf')
    generate_xxe_payload('stress_test_data/xxe.xml')
    print("Stress test data generated in stress_test_data/")
