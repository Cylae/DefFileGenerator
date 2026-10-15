#!/usr/bin/env python3
"""
50+ Equipementiers Industrial Benchmark:
Tests 52 real-world manufacturer documentation files across 20 manufacturers:
Huawei, ABB/PowerOne, Fox ESS, GoodWe, Growatt, Hukseflux, Janitza, Kaco, Lettel, Siebert,
SMA, Acrel, Solax, Carlo Gavazzi, Legrand, Victron, SofarSolar, Socomec, Kipp & Zonen, Chint Power.
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

DEFAULT_EQUIPEMENTIERS_DIR = r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers"
EQUIPEMENTIERS_DIR = os.environ.get(
    "EQUIPEMENTIERS_DIR",
    DEFAULT_EQUIPEMENTIERS_DIR
    if os.path.exists(DEFAULT_EQUIPEMENTIERS_DIR)
    else r"G:\My Drive\08092026\Equipementiers",
)

DOCS_52 = [
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
    # 11. SMA (3)
    (r"SMA\SMA_Modbus.csv", "SMA"),
    (r"SMA\Sunspec\Sources\CORE2_STP_110-60_Sunspec_Modbus_Model_PICS_Overview_V12.xlsx", "SMA"),
    (r"SMA\WPM00C715_SMA_Inverter_SMA_WR21TL09.csv", "SMA"),
    # 12. Acrel (2)
    (r"Acrel\METER_ACREL_DTSD1352.csv", "Acrel"),
    (r"Acrel\Sources\DTSD1352-Manual.pdf", "Acrel"),
    # 13. Solax (3)
    (r"Solax\SOLAX-X1-X3-GENERIC.csv", "Solax"),
    (r"Solax\INV_SOLAX_MICG2-PROG2.csv", "Solax"),
    (r"Solax\SOLAX-X1-X3-GENERIC(TCP).csv", "Solax"),
    # 14. Carlo Gavazzi (2)
    (r"CarloGavazzi\Meter\CarloGavazzi_EM210.csv", "Carlo Gavazzi"),
    (r"CarloGavazzi\Meter\CarloGavazzi_EM24.csv", "Carlo Gavazzi"),
    # 15. Legrand (2)
    (r"Legrand\Meter\METER_LEGRAND_LG412041.csv", "Legrand"),
    (r"Legrand\Meter\METER_LEGRAND_CF-GENERIC.csv", "Legrand"),
    # 16. Victron (2)
    (r"Victron\onduleur\MODBUS_Victron.csv", "Victron"),
    (r"Victron\CCGX-Modbus-TCP-register-list-2.90 (2).xlsx", "Victron"),
    # 17. SofarSolar (3)
    (r"SofarSolar\Sources\SOFAR_Modbus_Protocol_English_G3_V1.17.xlsx", "SofarSolar"),
    (r"SofarSolar\INV_SOFARSOLAR_1-40KTL.csv", "SofarSolar"),
    (r"SofarSolar\INV_SOFARSOLAR_50-70KTL.csv", "SofarSolar"),
    # 18. Socomec (2)
    (r"Socomec\compteurs\ElectricMotionWebdyn_modbusRTU_Meter_SOCOMEC_DIRIS_B10.csv", "Socomec"),
    (
        r"Socomec\compteurs\Diris\Sources\DIRIS M70\DIRIS-DIGIWARE-M-&-D_COMMUNICATION-TABLE_2025-01_CMT-SNMP-OID-LIST_EN.xlsx",
        "Socomec",
    ),
    # 19. Kipp & Zonen (2)
    (r"KippZonen\KippZonen_DustIQ.csv", "Kipp & Zonen"),
    (r"KippZonen\SENSOR_KIPPZONEN_SMP12.csv", "Kipp & Zonen"),
    # 20. Chint Power (1)
    (r"Chint Power\CHINT_DTSU.csv", "Chint Power"),
]


def test_52_battery():
    """Benchmarks 52 real-world manufacturer files across 20 industrial vendors."""
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logger = logging.getLogger("Equipementiers52")

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
    print("STARTING BENCHMARK ON 52 EQUIPEMENTIERS MODBUS DOCUMENTATIONS")
    print("=" * 75)

    start_time = time.time()

    for idx, (rel, mfg) in enumerate(DOCS_52, start=1):
        full = os.path.join(EQUIPEMENTIERS_DIR, rel)
        ext = os.path.splitext(full)[1].lower()
        res = {
            "idx": idx,
            "rel": rel,
            "mfg": mfg,
            "ext": ext,
            "status": "FAIL",
            "regs": 0,
            "valid_def": False,
            "error": "",
        }

        if not os.path.exists(full):
            res["error"] = "File missing from disk"
            results.append(res)
            print(f"[{idx:02d}/52] [{mfg:<12}] FAIL {rel} -> File missing")
            continue

        try:
            # 1. Extraction
            if ext == ".xlsx":
                raw = extractor.extract_from_excel(full)
            elif ext == ".csv":
                raw = extractor.extract_from_csv(full)
            elif ext == ".pdf":
                raw = extractor.extract_from_pdf(full)
            else:
                res["error"] = f"Unsupported extension {ext}"
                results.append(res)
                continue

            has_data, raw_iter = peek_generator(raw)
            if not has_data:
                res["error"] = "No table data yielded by extractor"
                results.append(res)
                print(f"[{idx:02d}/52] [{mfg:<12}] FAIL {rel} -> No table data")
                continue

            mapped = list(extractor.map_and_clean(raw_iter))
            reg_count = len(mapped)
            res["regs"] = reg_count

            if reg_count == 0:
                res["error"] = "0 registers after mapping/cleaning"
                results.append(res)
                print(f"[{idx:02d}/52] [{mfg:<12}] FAIL {rel} -> 0 registers")
                continue

            # 2. Webdyn CSV Generation
            out_csv = os.path.join(temp_dir.name, f"benchmark_{idx}_{mfg}.csv")
            cfg = GeneratorConfig(
                input_file=full,
                output=out_csv,
                manufacturer=mfg,
                model="BenchmarkModel",
                protocol="modbusRTU",
                category="Inverter",
            )
            run_generator(cfg, input_data=mapped)

            # 3. Validation
            if os.path.exists(out_csv):
                is_valid = generator.validate_csv(out_csv, strict=False)
                res["valid_def"] = is_valid
                if is_valid:
                    res["status"] = "PASS"
                else:
                    res["error"] = "Generated CSV failed structural validation"
                try:
                    os.remove(out_csv)
                except OSError:
                    pass
            else:
                res["error"] = "Output CSV was not produced by generator"

        except Exception as e:
            res["error"] = str(e)

        results.append(res)
        status_flag = "PASS" if res["status"] == "PASS" else "FAIL"
        print(
            f"[{idx:02d}/52] [{mfg:<12}] {status_flag} {rel} -> {res['regs']} registers, "
            f"{'Valid Def' if res['valid_def'] else 'INVALID'}"
        )

    temp_dir.cleanup()
    elapsed = time.time() - start_time

    passed = sum(1 for r in results if r["status"] == "PASS")
    total = len(results)

    print("\n" + "=" * 75)
    print("FINAL 52 EQUIPEMENTIERS BENCHMARK SUMMARY")
    print(f"Elapsed: {elapsed:.2f}s")
    print(f"TOTAL PASSED: {passed} / {total}")
    print("=" * 75)
    for r in results:
        status_tag = f"[{r['status']}]"
        err_msg = f" -- ERROR: {r['error']}" if r["error"] else ""
        print(
            f"  {status_tag} #{r['idx']:02d} [{r['mfg']:<12}] {r['rel']} ({r['regs']} regs){err_msg}"
        )
    print("=" * 75 + "\n")

    assert passed >= 45, f"Expected at least 45 passes out of 52, got {passed}"


if __name__ == "__main__":
    test_52_battery()
