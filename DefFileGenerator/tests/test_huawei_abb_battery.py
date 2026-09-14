#!/usr/bin/env python3
"""
Diagnostic runner for Huawei and ABB/PowerOne documentation battery.
"""

import logging
import os
import sys
import tempfile

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator

EQUIPEMENTIERS_DIR = r"G:\My Drive\08092026\Equipementiers"

TEST_FILES = [
    # --- Huawei CSVs ---
    r"Huawei\Inverter\HUAWEI_V4.csv",
    r"Huawei\Inverter\HUAWEI_V4 (METER).csv",
    r"Huawei\Inverter\HUAWEI_V4 (BATTERY).csv",
    r"Huawei\Inverter\SunSpec\WPM015E12_modbusRTU_Inverter_Huawei_SUN2000-30KTL-M3_DV3LZV3H4Q.csv",
    r"Huawei\SmartLogger\SL3000\HUAWEI_SmartLogger3000.csv",
    r"Huawei\Meter\HUAWEI_PowerMeter.csv",
    r"Huawei\Inverter\OLD\1\Huawei_364233-A.csv",
    r"Huawei\Inverter\OLD\2\HUAWEI_V2.csv",
    r"Huawei\Inverter\OLD\3\HUAWEI_V3.csv",
    # --- ABB / PowerOne CSVs ---
    r"PowerOne\Modbus\PVS800.csv",
    r"PowerOne\Modbus\INV_ABB_PVI50-60.csv",
    r"PowerOne\Modbus\ABB_TRIO_200_276_FR.csv",
    r"PowerOne\Sunspec\WPM_MODBUS_ABB_TRIO_200_276.csv",
    r"PowerOne\M2M_Ethernet-Meter\ABB_M2M_Meter.csv",
    r"PowerOne\Sunspec\Modbus_ABB_SunSpec.csv",
    r"PowerOne\Sunspec\ABB_PVS50_MODBUS_light.csv",
    # --- ABB / PowerOne XLSX ---
    r"PowerOne\Sunspec\Sources\SUNSPEC_PVS1xx_PICS_V002_GL20181011.xlsx",
    r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-175_PICS_Rev_003.xlsx",
    r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-10_12.5_15_20_30_33_PICS_Rev_001.xlsx",
    r"PowerOne\Sunspec\Sources\ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.xlsx",
    r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-50_60_PICS_Rev_3.xlsx",
    r"PowerOne\Modbus\sources\Meter\Modbus mapping Meter ABB B21.xlsx",
    r"PowerOne\Modbus\sources\analyseur reseau M4M\M4M Modbus map - v.1.3N.xlsx",
    r"Huawei\Inverter\OLD\1\Huawei_Work.xlsx",
    # --- Huawei PDFs ---
    r"Huawei\Inverter\Sources\SUN2000MA_V100R001C00_ModbusInterfaceDefinitions_issue08_2024.pdf",
    r"Huawei\Inverter\Sources\SUN2000-12~25K-MB0 ModbusInterfaceDefinitions_issue01_2023.pdf",
    r"Huawei\Inverter\OLD\1\Sources\SUN2000 364233-A MODBUS Interface Definitions-20170731 (002).pdf",
    r"Huawei\Inverter\OLD\1\Sources\SUN2000 33-40 KTL MODBUS Interface Definitions.pdf",
    r"Huawei\Inverter\OLD\1\Sources\SUN2000-(50KTL-M060KTL-M065KTL-M070KTL-INM070KTL-C175KTL-C1) MODBUS Interface Definitions (003).pdf",
    r"Huawei\Inverter\OLD\1\Sources\SUN2000 8 10121517202328KTL MODBUS Interface Definitions.pdf",
    r"Huawei\SmartLogger\SmartLogger ModBus Interface Definitions 2018.pdf",
    # --- ABB PDFs ---
    r"PowerOne\ABB Terra Heavy charger\ABB_Terra_AC_Charger_ModbusCommunication_v1.7.pdf",
    r"PowerOne\Modbus\sources\Modbus_RTU_register_map_for_TRIO-20.0(27.6)-TL-OUTD_Revision_1.6.pdf",
    r"PowerOne\DocOnduleurs\TRIO 50 60\Modbus_RTU_register_map_for_TRIO-50 0_60 0_Revision_3.4.pdf",
    r"PowerOne\Modbus\sources\analyseur reseau M4M\Modbus manual v. 1.05.pdf",
]


def run_battery():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logger = logging.getLogger("HuaweiAbbBattery")

    if not os.path.exists(EQUIPEMENTIERS_DIR):
        logger.warning(f"Equipementiers dir not found: {EQUIPEMENTIERS_DIR}")
        return

    extractor = Extractor()
    generator = Generator()
    temp_dir = tempfile.TemporaryDirectory()

    results = []
    print(f"\nRunning battery on {len(TEST_FILES)} documentation files from Huawei and ABB...")

    for idx, rel in enumerate(TEST_FILES, start=1):
        full = os.path.join(EQUIPEMENTIERS_DIR, rel)
        ext = os.path.splitext(full)[1].lower()
        res = {
            "idx": idx,
            "rel": rel,
            "ext": ext,
            "exists": os.path.exists(full),
            "regs": 0,
            "def_valid": False,
            "error": None,
        }

        if not res["exists"]:
            res["error"] = "File not found"
            results.append(res)
            print(f"[{idx:02d}/35] FAIL {rel} -> File not found")
            continue

        try:
            if ext == ".csv":
                raw = extractor.extract_from_csv(full)
            elif ext in [".xlsx", ".xlsm"]:
                raw = extractor.extract_from_excel(full)
            elif ext == ".pdf":
                raw = extractor.extract_from_pdf(full)
            else:
                res["error"] = f"Unsupported ext {ext}"
                results.append(res)
                continue

            has_data, it = peek_generator(raw)
            if not has_data:
                res["error"] = "No data yielded by extractor"
                results.append(res)
                print(f"[{idx:02d}/35] FAIL {rel} -> No data yielded by extractor")
                continue

            mapped = list(extractor.map_and_clean(it))
            res["regs"] = len(mapped)
            if len(mapped) == 0:
                res["error"] = "0 registers mapped"
                results.append(res)
                print(f"[{idx:02d}/35] FAIL {rel} -> 0 registers mapped")
                continue

            # Generate Webdyn definition file
            mfg = "Huawei" if "Huawei" in rel else "ABB"
            out_path = os.path.join(temp_dir.name, f"def_{idx}_{ext.strip('.')}.csv")
            cfg = GeneratorConfig(
                input_file=full,
                output=out_path,
                manufacturer=mfg,
                model="ModelTest",
                protocol="modbusRTU",
                category="Inverter",
            )
            run_generator(cfg, input_data=mapped)

            if os.path.exists(out_path):
                # Validate definition file
                valid = generator.validate_csv(out_path, strict=False)
                res["def_valid"] = valid
                if valid:
                    print(f"[{idx:02d}/35] PASS {rel} -> {res['regs']} registers, Valid Def")
                else:
                    res["error"] = "CSV definition validation failed"
                    print(f"[{idx:02d}/35] FAIL {rel} -> {res['regs']} registers, Def Invalid")
                os.remove(out_path)
            else:
                res["error"] = "Definition file was not created"
                print(f"[{idx:02d}/35] FAIL {rel} -> Def file not created")

        except Exception as e:
            res["error"] = str(e)
            print(f"[{idx:02d}/35] ERROR {rel} -> {e}")

        results.append(res)

    temp_dir.cleanup()

    print("\n" + "=" * 70)
    print("SUMMARY OF HUAWEI & ABB REAL-WORLD TESTS:")
    total = len(results)
    passed = sum(1 for r in results if r["def_valid"] and r["regs"] > 0)
    print(f"Passed: {passed} / {total}")
    for r in results:
        status = "PASS" if (r["def_valid"] and r["regs"] > 0) else "FAIL"
        err = f" ({r['error']})" if r["error"] else ""
        print(f"  [{status}] {r['rel']} -> regs: {r['regs']}{err}")
    print("=" * 70)


if __name__ == "__main__":
    run_battery()
