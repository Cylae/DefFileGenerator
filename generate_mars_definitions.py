#!/usr/bin/env python3
"""
High-Precision Generator for Mars Renewable / Mars Local Controller V7.0
Strictly aligned with WebdynSunPM reference specifications (INV_HUAWEI_V4.csv benchmark).

Features:
- Complete parsing of all sheets in MarsLocalController_IO List_ModbusTCP_V7.0.xlsx.
- 11-column standard format (Header + data rows).
- 6-decimal coefficient formatting (1.000000, 0.010000, 0.100000, 0.000000).
- Standard canonical data types: FLOAT32, U16, U32, I32, STRING.
- Canonical Webdyn action codes: 4 (Telemetry/Read), 9 (Status/Alarms), 1 (Write commands), 0 (Reserve).
- Standard Webdyn tags (RealPower, EnergyTotal, EnergyDay) with clean empty fields for custom metrics.
- Multi-slave generation:
    1. MarsLocalController_V7.0_webdyn.csv (Slave 1: BESS Concentrator, PCS 1..50, Banks 1..50, Meters 1..100, AC 1..100, DIO 1..50, PV 1..50, LC 1..100, Alarms)
    2. MarsLocalController_V7.0_EMS_Slave247.csv (Slave 247: LC Overview FC4 + EMS Writes FC16/3)
    3. MarsLocalController_V7.0_RackCell_SlaveN.csv (Slave N+1: Rack + 1000 Cells)
    4. MarsLocalController_V7.0_AllInOne.csv (Full system register map)
"""

import csv
import os

import openpyxl

from DefFileGenerator.def_gen import Generator


def normalize_unit(unit_raw: str | None) -> tuple[str, float]:
    """Normalizes engineering units and extracts scaling factors."""
    if not unit_raw:
        return "", 1.0
    u = str(unit_raw).strip()
    if u in ("-", "/", "1"):
        return "", 1.0
    if u == "0.01Hz":
        return "Hz", 0.01
    if u.lower() in ("℃", "°c", "degc", "celsius"):
        return "°C", 1.0
    if u in ("Ω", "ohm", "Ohm"):
        return "Ω", 1.0
    return u, 1.0


def normalize_type(type_raw: str | None) -> str:
    """Maps vendor type strings to canonical WebdynSunPM data types."""
    if not type_raw:
        return "U16"
    t = str(type_raw).strip().lower()
    if "float" in t or t in ("f32", "float32"):
        return "F32"
    if t in ("u16", "uint16"):
        return "U16"
    if t in ("u32", "uint32"):
        return "U32"
    if t in ("i32", "int32"):
        return "I32"
    if t in ("i16", "int16"):
        return "I16"
    if t in ("bit", "bits", "bool", "boolean"):
        return "U16"
    return "U16"


def determine_action(name: str, fc: str, is_reserve: bool) -> str:
    """Determines WebdynSunPM Action code based on signal role."""
    if is_reserve:
        return "0"
    if fc in ("1", "16", "3") and ("write" in name.lower() or "mode" in name.lower() and "set" in name.lower()):
        return "1"
    name_lower = name.lower()
    if any(k in name_lower for k in ("alarm", "status", "fault", "crash-stop", "trip", "state")):
        return "9"
    return "4"


def build_definition_file(output_path: str, protocol: str, category: str, manufacturer: str, model: str, registers: list[dict], strict_overlap: bool = True) -> bool:
    """Writes a strictly compliant WebdynSunPM definition file with UTF-8 BOM encoding."""
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, delimiter=";", lineterminator="\n")
        # 11 columns in header (10 semicolons)
        writer.writerow([protocol, category, manufacturer, model, "", "", "", "", "", "", ""])

        for idx, reg in enumerate(registers, start=1):
            writer.writerow([
                str(idx),
                str(reg["Info1"]),
                str(reg["Info2"]),
                str(reg["Info3"]),
                "",
                str(reg["Name"]),
                str(reg.get("Tag", "")),
                f"{float(reg['CoefA']):.6f}",
                f"{float(reg['CoefB']):.6f}",
                str(reg.get("Unit", "")),
                str(reg["Action"]),
            ])

    # Validate output file
    gen = Generator()
    report = gen.validate_csv_detailed(output_path, strict=True, strict_overlap=strict_overlap)
    if not report.is_valid:
        print(f"Validation FAILED for {output_path}:")
        for issue in report.issues:
            print(f"  Line {issue.line}: [{issue.severity}] {issue.code} - {issue.message}")
        return False

    print(f"Validated successfully: {output_path} ({len(registers)} registers, 0 errors)")
    return True


def parse_mars_workbook(xlsx_path: str):
    """Parses all sheets of Mars Local Controller V7.0 workbook."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # 1. LC Information (Overview, EMS Write, Alarms)
    ws_lc = wb["LC Information"]
    rows_lc = list(ws_lc.iter_rows(values_only=True))

    lc_overview = []
    # Lines 6-10: Overview FC 0x04, Addr 247
    for r in rows_lc[5:10]:
        if r and r[0] is not None and str(r[0]).isdigit():
            addr = int(r[0])
            name = f"LC Overview - {str(r[2]).strip()}"
            dtype = normalize_type(r[4])
            lc_overview.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": dtype,
                "Name": name,
                "Tag": "",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "",
                "Action": "4",
            })

    # Lines 18-24: EMS Write FC 0x16 (or FC 3), Addr 247
    lc_writes = []
    for r in rows_lc[17:24]:
        if r and r[0] is not None and str(r[0]).isdigit():
            addr = int(r[0])
            name = f"EMS Consigne - {str(r[2]).strip()}"
            unit, factor = normalize_unit(r[3])
            dtype = normalize_type(r[4])
            lc_writes.append({
                "Info1": "3",
                "Info2": str(addr),
                "Info3": dtype,
                "Name": name,
                "Tag": "",
                "CoefA": factor,
                "CoefB": 0.0,
                "Unit": unit,
                "Action": "1",
            })

    # Lines 30-33: LC Alarm Details FC 0x02, Addr 1
    lc_alarms = []
    for r in rows_lc[29:33]:
        if r and r[0] is not None and str(r[0]).isdigit():
            addr = int(r[0])
            name = f"LC Alarm - {str(r[2]).strip()}"
            lc_alarms.append({
                "Info1": "2",
                "Info2": str(addr),
                "Info3": "U16",
                "Name": name,
                "Tag": "",
                "CoefA": 1.0,
                "CoefB": 0.0,
                "Unit": "",
                "Action": "9",
            })

    # Helper function to extract equipment template rows
    def extract_template(sheet_name: str, skip_header_rows: int = 6):
        ws = wb[sheet_name]
        template = []
        rows = list(ws.iter_rows(values_only=True))
        for r in rows[skip_header_rows:]:
            if not r or r[0] is None:
                continue
            addr_str = str(r[0]).strip()
            if not addr_str.isdigit():
                if "..." in addr_str:
                    break
                continue
            addr = int(addr_str)
            name = str(r[2]).strip() if r[2] else ""
            acc = str(r[3]).strip() if len(r) > 3 and r[3] else ""
            unit_val = str(r[4]).strip() if len(r) > 4 and r[4] else ""
            type_val = str(r[5]).strip() if len(r) > 5 and r[5] else ""

            unit, unit_factor = normalize_unit(unit_val)
            acc_factor = 1.0
            if acc:
                try:
                    acc_factor = float(acc)
                except ValueError:
                    pass
            coef_a = acc_factor if acc_factor != 1.0 else unit_factor
            dtype = normalize_type(type_val)
            # Only genuine reserve registers get Action 0. Informational notes like 'Empty point' do not deactivate real physical metrics.
            is_reserve = "reserve" in name.lower()
            action = determine_action(name, "4", is_reserve)

            # Fix known vendor documentation typo: Branch 2~n DC Current listed at 71 instead of 72 (overlaps with 70 F32)
            if addr == 71 and "branch 2~n dc current" in name.lower():
                addr = 72

            template.append({
                "rel_addr": addr,
                "name": name,
                "dtype": dtype,
                "unit": unit,
                "coef_a": coef_a,
                "action": action,
            })
        return template

    pcs_tpl = extract_template("PCS Information")
    bank_tpl = extract_template("Bank Information")
    meter_tpl = extract_template("Meter Information")
    ac_tpl = extract_template("Air Conditioning Information")
    dio_tpl = extract_template("DIO Information")
    pv_tpl = extract_template("PV Information")
    lc_cool_tpl = extract_template("Liquid Cooling Information")

    # Parse Rack & Cell template
    ws_rack = wb["Rack and Cell Information"]
    rack_rows = list(ws_rack.iter_rows(values_only=True))
    rack_tpl = []
    cell_tpl = []
    in_cell = False

    for r in rack_rows[6:]:
        if not r or r[0] is None:
            continue
        addr_str = str(r[0]).strip()
        if not addr_str.isdigit():
            if "..." in addr_str:
                in_cell = True
            continue
        addr = int(addr_str)
        name = str(r[2]).strip() if r[2] else ""
        unit_val = str(r[4]).strip() if len(r) > 4 and r[4] else ""
        type_val = str(r[5]).strip() if len(r) > 5 and r[5] else ""
        unit, factor = normalize_unit(unit_val)
        dtype = normalize_type(type_val)
        action = "9" if "status" in name.lower() or "alarm" in name.lower() else "4"

        if addr < 100 and not in_cell:
            rack_tpl.append({
                "rel_addr": addr,
                "name": name,
                "dtype": dtype,
                "unit": unit,
                "coef_a": factor,
                "action": action,
            })
        elif 100 <= addr <= 109:
            cell_tpl.append({
                "rel_addr": addr - 100,
                "name": name,
                "dtype": dtype,
                "unit": unit,
                "coef_a": factor,
                "action": action,
            })

    return {
        "lc_overview": lc_overview,
        "lc_writes": lc_writes,
        "lc_alarms": lc_alarms,
        "pcs_tpl": pcs_tpl,
        "bank_tpl": bank_tpl,
        "meter_tpl": meter_tpl,
        "ac_tpl": ac_tpl,
        "dio_tpl": dio_tpl,
        "pv_tpl": pv_tpl,
        "lc_cool_tpl": lc_cool_tpl,
        "rack_tpl": rack_tpl,
        "cell_tpl": cell_tpl,
    }


def generate_all_mars_files(xlsx_path: str, workspace_dir: str):
    """Generates the full suite of production-ready Mars Local Controller V7.0 definition files."""
    data = parse_mars_workbook(xlsx_path)

    # -------------------------------------------------------------------------
    # File 1: Slave 1 - BESS Concentrator (MarsLocalController_V7.0_webdyn.csv)
    # -------------------------------------------------------------------------
    # Contains: LC Alarms (FC 2, 0..3)
    # + 50 PCS (FC 4, 0..4999, offset 100)
    # + 50 Banks (FC 4, 5000..9999, offset 100)
    # + 100 Meters (FC 4, 10000..14999, offset 50)
    # + 100 AC (FC 4, 15000..19999, offset 50)
    # + 50 DIO (FC 4, 20000..24999, offset 100)
    # + 50 PV (FC 4, 25000..29999, offset 100)
    # + 100 LC (FC 4, 30000..34999, offset 50)
    slave1_regs = []

    # LC Alarms (FC 2)
    for reg in data["lc_alarms"]:
        slave1_regs.append(reg)

    # 50 PCS (FC 4)
    pcs_base0 = data["pcs_tpl"][0]["rel_addr"] if data["pcs_tpl"] else 0
    for i in range(1, 51):
        offset = (i - 1) * 100
        for item in data["pcs_tpl"]:
            addr = item["rel_addr"] - pcs_base0 + offset
            tag = ""
            if i == 1:
                if "Total AC Active Power" in item["name"]:
                    tag = "RealPower"
                elif "Total AC Discharging Energy" in item["name"]:
                    tag = "EnergyTotal"
                elif "Daily AC Discharging Energy" in item["name"]:
                    tag = "EnergyDay"
                elif "AC Frequency" in item["name"]:
                    tag = "GridFrequency"
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#PCS - {item['name']}",
                "Tag": tag,
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    # 50 Banks (FC 4)
    bank_base0 = data["bank_tpl"][0]["rel_addr"] if data["bank_tpl"] else 5000
    for i in range(1, 51):
        offset = 5000 + (i - 1) * 100
        for item in data["bank_tpl"]:
            addr = item["rel_addr"] - bank_base0 + offset
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#battery bank - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    # 100 Meters (FC 4)
    meter_base0 = data["meter_tpl"][0]["rel_addr"] if data["meter_tpl"] else 10000
    for i in range(1, 101):
        offset = 10000 + (i - 1) * 50
        for item in data["meter_tpl"]:
            addr = item["rel_addr"] - meter_base0 + offset
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#Bidirectional Meter - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    # 100 AC (FC 4)
    ac_base0 = data["ac_tpl"][0]["rel_addr"] if data["ac_tpl"] else 15000
    for i in range(1, 101):
        offset = 15000 + (i - 1) * 50
        for item in data["ac_tpl"]:
            addr = item["rel_addr"] - ac_base0 + offset
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#AC - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    # 50 DIO (FC 4)
    dio_base0 = data["dio_tpl"][0]["rel_addr"] if data["dio_tpl"] else 20000
    for i in range(1, 51):
        offset = 20000 + (i - 1) * 100
        for item in data["dio_tpl"]:
            addr = item["rel_addr"] - dio_base0 + offset
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#DIO - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    # 50 PV (FC 4)
    pv_base0 = data["pv_tpl"][0]["rel_addr"] if data["pv_tpl"] else 25000
    for i in range(1, 51):
        offset = 25000 + (i - 1) * 100
        for item in data["pv_tpl"]:
            addr = item["rel_addr"] - pv_base0 + offset
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#PV - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    # 100 LC (FC 4)
    lc_base0 = data["lc_cool_tpl"][0]["rel_addr"] if data["lc_cool_tpl"] else 30000
    for i in range(1, 101):
        offset = 30000 + (i - 1) * 50
        for item in data["lc_cool_tpl"]:
            addr = item["rel_addr"] - lc_base0 + offset
            slave1_regs.append({
                "Info1": "4",
                "Info2": str(addr),
                "Info3": item["dtype"],
                "Name": f"{i}#Liquid Cooling - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    slave1_path = os.path.join(workspace_dir, "MarsLocalController_V7.0_webdyn.csv")
    build_definition_file(slave1_path, "modbusTCP", "BESS", "Mars Energy", "MarsLocalController V7.0", slave1_regs)

    # -------------------------------------------------------------------------
    # File 2: Slave 247 - EMS Management (MarsLocalController_V7.0_EMS_Slave247.csv)
    # -------------------------------------------------------------------------
    slave247_regs = []
    slave247_regs.extend(data["lc_overview"])
    slave247_regs.extend(data["lc_writes"])
    slave247_path = os.path.join(workspace_dir, "MarsLocalController_V7.0_EMS_Slave247.csv")
    build_definition_file(slave247_path, "modbusTCP", "BESS", "Mars Energy", "MarsLocalController EMS V7.0", slave247_regs)

    # -------------------------------------------------------------------------
    # File 3: Slave N+1 - Rack & 1000 Cells (MarsLocalController_V7.0_RackCell_SlaveN.csv)
    # -------------------------------------------------------------------------
    rack_regs = []
    for item in data["rack_tpl"]:
        rack_regs.append({
            "Info1": "4",
            "Info2": str(item["rel_addr"]),
            "Info3": item["dtype"],
            "Name": f"Rack - {item['name']}",
            "Tag": "",
            "CoefA": item["coef_a"],
            "CoefB": 0.0,
            "Unit": item["unit"],
            "Action": item["action"],
        })

    for cell_idx in range(1, 1001):
        cell_offset = 100 + (cell_idx - 1) * 10
        for item in data["cell_tpl"]:
            cell_addr = cell_offset + item["rel_addr"]
            rack_regs.append({
                "Info1": "4",
                "Info2": str(cell_addr),
                "Info3": item["dtype"],
                "Name": f"Cell {cell_idx} - {item['name']}",
                "Tag": "",
                "CoefA": item["coef_a"],
                "CoefB": 0.0,
                "Unit": item["unit"],
                "Action": item["action"],
            })

    rack_path = os.path.join(workspace_dir, "MarsLocalController_V7.0_RackCell_SlaveN.csv")
    build_definition_file(rack_path, "modbusTCP", "BatteryRack", "Mars Energy", "MarsLocalController Rack V7.0", rack_regs)

    # -------------------------------------------------------------------------
    # File 4: All-In-One Unified Architecture (MarsLocalController_V7.0_AllInOne.csv)
    # -------------------------------------------------------------------------
    all_regs = []
    all_regs.extend(slave247_regs)
    all_regs.extend(slave1_regs)
    all_regs.extend(rack_regs)
    all_path = os.path.join(workspace_dir, "MarsLocalController_V7.0_AllInOne.csv")
    build_definition_file(all_path, "modbusTCP", "BESS", "Mars Energy", "MarsLocalController All-In-One V7.0", all_regs, strict_overlap=False)


if __name__ == "__main__":
    src_xlsx = r"C:\Users\Cylae\Downloads\MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"
    wk_dir = r"c:\Users\Cylae\Documents\GitHub\DefFileGenerator"
    generate_all_mars_files(src_xlsx, wk_dir)
