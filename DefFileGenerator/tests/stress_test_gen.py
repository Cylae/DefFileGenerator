import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def create_stress_data(num_registers=1000):
    os.makedirs('stress_test_data', exist_ok=True)

    data = []
    for i in range(num_registers):
        data.append({
            'Name': f'Register_{i}',
            'Address': str(40001 + i),
            'RegisterType': 'Holding Register',
            'Type': 'U16',
            'Unit': 'V',
            'Factor': '0.1'
        })

    df = pd.DataFrame(data)

    # CSV
    df.to_csv('stress_test_data/stress.csv', index=False)

    # Excel
    df.to_excel('stress_test_data/stress.xlsx', index=False)

    # PDF
    c = canvas.Canvas("stress_test_data/stress.pdf", pagesize=letter)
    width, height = letter
    y = height - 50
    c.drawString(50, y, "Register | Name | Type")
    y -= 20
    for i in range(min(num_registers, 30)): # Only first 30 for PDF sample
        c.drawString(50, y, f"{40001+i} | Register_{i} | uint16")
        y -= 15
    c.save()

if __name__ == "__main__":
    create_stress_data()
    print("Stress test data generated in stress_test_data/")
