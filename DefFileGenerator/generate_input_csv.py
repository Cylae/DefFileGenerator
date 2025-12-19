import csv

def generate_csv():
    rows = []
    header = ["Name", "Tag", "RegisterType", "Address", "Type", "Factor", "Offset", "Unit", "Action"]

    # Helper to add row
    def add_row(name, tag, reg_type, addr, type_str, factor="1", offset="0", unit="", action="4"):
        rows.append([name, tag, reg_type, addr, type_str, factor, offset, unit, action])

    # --- 1. Inputs (30001) ---
    add_row("Digital Inputs", "DigitalInputs", "Input", "30001", "U16")

    # --- 2. Inst Currents (30002-30025) ---
    for i in range(24):
        add_row(f"Inst Curr Str_{i+1}", f"InstCurrStr{i+1}", "Input", str(30002 + i), "U16", "0.001", "0", "A")

    # --- 3. Analog (30040-30048) ---
    add_row("Inst V_1", "InstV1", "Input", "30040", "U16", "0.1", "0", "V")
    # 30041 Unused
    add_row("Aux 1", "Aux1", "Input", "30042", "U16", "0.01", "0", "V")
    add_row("Aux 2", "Aux2", "Input", "30043", "U16", "0.02", "0", "mA")
    add_row("Inst T_1", "InstT1", "Input", "30044", "U16", "1", "0", "°C")
    add_row("Inst T_2", "InstT2", "Input", "30045", "U16", "0.1", "0", "°C")
    add_row("Inst T_3", "InstT3", "Input", "30046", "U16", "0.1", "0", "°C")
    add_row("Sum Currents", "SumCurrents", "Input", "30047", "U16", "0.001", "0", "A")
    add_row("Power", "Power", "Input", "30048", "U16", "1", "0", "W") # Assuming U16 as per 3.py logic (1 register)

    # --- 4. RMS Currents (30052-30075) ---
    for i in range(24):
        add_row(f"RMS Curr Str_{i+1}", f"RMSCurrStr{i+1}", "Input", str(30052 + i), "U16", "0.001", "0", "A")

    # --- 5. Offset Currents (40002-40025) ---
    # Holding registers
    for i in range(24):
        add_row(f"Offset Curr Str_{i+1}", f"OffsetCurrStr{i+1}", "Holding", str(40002 + i), "U16", "1", "0", "")

    # --- 6. Offset Analog (40040-40046) ---
    add_row("Offset V_1", "OffsetV1", "Holding", "40040", "U16", "1", "0", "")
    add_row("Offset Aux_1", "OffsetAux1", "Holding", "40042", "U16", "1", "0", "")
    add_row("Offset Aux_2", "OffsetAux2", "Holding", "40043", "U16", "1", "0", "")
    add_row("Offset T_1", "OffsetT1", "Holding", "40044", "U16", "1", "0", "")
    add_row("Offset T_2", "OffsetT2", "Holding", "40045", "U16", "1", "0", "")
    add_row("Offset T_3", "OffsetT3", "Holding", "40046", "U16", "1", "0", "")

    # --- 7. Gain Currents (40052-40076) ---
    for i in range(24):
        add_row(f"Gain Curr Str_{i+1}", f"GainCurrStr{i+1}", "Holding", str(40052 + i), "U16", "1", "0", "")

    add_row("Gain Check 40076", "GainCheck40076", "Holding", "40076", "U16", "1", "0", "")

    # --- 8. Gain Analog (40090-40096) ---
    add_row("Gain V_1", "GainV1", "Holding", "40090", "U16", "1", "0", "")
    add_row("Gain Aux_1", "GainAux1", "Holding", "40092", "U16", "1", "0", "")
    add_row("Gain Aux_2", "GainAux2", "Holding", "40093", "U16", "1", "0", "")
    add_row("Gain T_1", "GainT1", "Holding", "40094", "U16", "1", "0", "")
    add_row("Gain T_2", "GainT2", "Holding", "40095", "U16", "1", "0", "")
    add_row("Gain T_3", "GainT3", "Holding", "40096", "U16", "1", "0", "")

    with open("ST12422_input.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)

    print("ST12422_input.csv generated.")

if __name__ == "__main__":
    generate_csv()
