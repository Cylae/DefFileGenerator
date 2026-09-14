#!/usr/bin/env python3
"""
30 Equipementiers Benchmark:
Tests 30 real-world manufacturer documentation files across 10 manufacturers:
Huawei, ABB, Fox ESS, GoodWe, Growatt, Hukseflux, Janitza, Kaco, Lettel, Siebert.
"""

import logging
import os
import sys
import tempfile
import time

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator

EQUIPEMENTIERS_DIR = r"G:\My Drive\08092026\Equipementiers"

DOCS_30 = [
    # 1. Huawei (3)
    (r"Huawei\Inverter\HUAWEI_V4.csv", "Huawei"),
    (
        r"Huawei\Inverter\Sources\SUN2000MA_V100R001C00_ModbusInterfaceDefinitions_issue08_2024.pdf",
        "Huawei",
    ),
    (r"Huawei\SmartLogger\SL3000\HUAWEI_SmartLogger3000.csv", "Huawei"),
    # 2. ABB / PowerOne (3)
    (r"PowerOne\Sunspec\Sources\SUNSPEC_PVS1xx_PICS_V002_GL20181011.xlsx", "ABB"),
    (r"PowerOne\Modbus\sources\analyseur reseau M4M\M4M Modbus map - v.1.3N.xlsx", "ABB"),
    (r"PowerOne\ABB Terra Heavy charger\ABB_Terra_AC_Charger_ModbusCommunication_v1.7.pdf", "ABB"),
    # 3. Fox ESS (3)
    (r"Fox ESS\R&V&H3plus&GmaxPlus.csv", "Fox ESS"),
    (r"Fox ESS\INV_FOXESS_RSERIESINVERTER.csv", "Fox ESS"),
    (r"Fox ESS\G-MAX Communication Protocol modbusTcp-EN-20250425(1).xlsx", "Fox ESS"),
    # 4. Goodwe (3)
    (r"Goodwe\GoodWe_SMT_Series.csv", "GoodWe"),
    (
        r"Goodwe\Sources\Sunspec for Inverters\GOODWE_three_phase_grid-tied_inverter_PICS.xlsx",
        "GoodWe",
    ),
    (r"Goodwe\GoodWe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv", "GoodWe"),
    # 5. Growatt (3)
    (r"Growatt\W_Growatt_Generic.csv", "Growatt"),
    (r"Growatt\Sources\Growatt Inverter Modbus RTU Protocol_II V1 07 （20180918）.pdf", "Growatt"),
    (r"Growatt\Sources\Growatt PV Inverter Modbus RS485 RTU Protocol V3 14.pdf", "Growatt"),
    # 6. Hukseflux (3)
    (r"Hukseflux\Hukseflux_SR30.csv", "Hukseflux"),
    (r"Hukseflux\Hukseflux_SR05-D1A3.csv", "Hukseflux"),
    (r"Hukseflux\Sources\Hukseflux_SR20-D2_Doc_FR.pdf", "Hukseflux"),
    # 7. Janitza (3)
    (r"Janitza\JANITZA_GENERIC_RTU.csv", "Janitza"),
    (r"Janitza\JANITZA_GENERIC_TCP.csv", "Janitza"),
    (r"Janitza\Fichier de def inconnu\METER_JANITZA_UMG604-PRO.csv", "Janitza"),
    # 8. Kaco (3)
    (r"Kaco\DOC from Kaco\SUNSPEC_NX3_G2_G3_1xx_7xx.xlsx", "Kaco"),
    (r"Kaco\Modbus\Kaco_NX1-NX3_ModbusRTU.csv", "Kaco"),
    (r"Kaco\Modbus\KACO_NX1.csv", "Kaco"),
    # 9. Lettel (3)
    (r"Lettel\Sources\map4-345m_table_modbus_fr.xlsx", "Lettel"),
    (r"Lettel\LETTEL_MAP4-34RJ.csv", "Lettel"),
    (r"Lettel\Sources\mcx4-34v_table_modbus_fr.pdf", "Lettel"),
    # 10. Siebert (3)
    (r"Siebert\AFFICHEUR_SIEBERT_XC4XX_TCP.csv", "Siebert"),
    (r"Siebert\AFFICHEUR_SIEBERT_XC4XX.csv", "Siebert"),
    (r"Siebert\AFFICHEUR_SIEBERT_XC4XX_TEST.csv", "Siebert"),
]


def test_30_battery():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logger = logging.getLogger("Equipementiers30")

    if not os.path.exists(EQUIPEMENTIERS_DIR):
        logger.warning(f"Equipementiers dir not found: {EQUIPEMENTIERS_DIR}")
        try:
            import pytest

            pytest.skip(f"Equipementiers directory not found at: {EQUIPEMENTIERS_DIR}")
        except ImportError:
            return

    extractor = Extractor()
    generator = Generator()
    temp_dir = tempfile.TemporaryDirectory()

    results = []
    print("\n" + "=" * 75)
    print("STARTING BENCHMARK ON 30 EQUIPEMENTIERS MODBUS DOCUMENTATIONS")
    print("=" * 75)

    start_time = time.time()

    for idx, (rel, mfg) in enumerate(DOCS_30, start=1):
        full = os.path.join(EQUIPEMENTIERS_DIR, rel)
        ext = os.path.splitext(full)[1].lower()
        res = {
            "idx": idx,
            "mfg": mfg,
            "rel": rel,
            "ext": ext,
            "regs": 0,
            "def_valid": False,
            "error": None,
        }

        if not os.path.exists(full):
            res["error"] = "File not found"
            results.append(res)
            print(f"[{idx:02d}/30] [{mfg:10s}] FAIL {rel} -> File not found")
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
                res["error"] = "No table data yielded by extractor"
                results.append(res)
                print(f"[{idx:02d}/30] [{mfg:10s}] FAIL {rel} -> No table data")
                continue

            mapped = list(extractor.map_and_clean(it))
            res["regs"] = len(mapped)
            if len(mapped) == 0:
                res["error"] = "0 registers mapped from table"
                results.append(res)
                print(f"[{idx:02d}/30] [{mfg:10s}] FAIL {rel} -> 0 registers mapped")
                continue

            out_path = os.path.join(temp_dir.name, f"def_{idx}_{ext.strip('.')}.csv")
            cfg = GeneratorConfig(
                input_file=full,
                output=out_path,
                manufacturer=mfg,
                model=f"Model_{idx}",
                protocol="modbusRTU",
                category="Inverter",
            )
            run_generator(cfg, input_data=mapped)

            if os.path.exists(out_path):
                valid = generator.validate_csv(out_path, strict=False)
                res["def_valid"] = valid
                if valid:
                    print(
                        f"[{idx:02d}/30] [{mfg:10s}] PASS {rel} -> {res['regs']} registers, Valid Def"
                    )
                else:
                    res["error"] = "CSV definition validation failed"
                    print(
                        f"[{idx:02d}/30] [{mfg:10s}] FAIL {rel} -> {res['regs']} registers, Invalid Def"
                    )
                os.remove(out_path)
            else:
                res["error"] = "Definition file was not created"
                print(f"[{idx:02d}/30] [{mfg:10s}] FAIL {rel} -> Output def missing")

        except Exception as e:
            res["error"] = str(e)
            print(f"[{idx:02d}/30] [{mfg:10s}] ERROR {rel} -> {e}")

        results.append(res)

    elapsed = time.time() - start_time
    temp_dir.cleanup()

    print("\n" + "=" * 75)
    print("FINAL 30 EQUIPEMENTIERS BENCHMARK SUMMARY")
    print(f"Elapsed: {elapsed:.2f}s")
    passed = sum(1 for r in results if r["def_valid"] and r["regs"] > 0)
    print(f"TOTAL PASSED: {passed} / 30")
    print("=" * 75)
    for r in results:
        status = "PASS" if (r["def_valid"] and r["regs"] > 0) else "FAIL"
        err = f" -- ERROR: {r['error']}" if r["error"] else ""
        print(f"  [{status}] #{r['idx']:02d} [{r['mfg']:10s}] {r['rel']} ({r['regs']} regs){err}")
    print("=" * 75)

    assert passed == 30


if __name__ == "__main__":
    test_30_battery()
