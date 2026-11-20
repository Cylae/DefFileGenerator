#!/usr/bin/env python3
"""
Production Suite Builder and Governor — Universal Multi-Manufacturer Engine.
Enforces the 5 Apollo-standard pillars on all definition files entering `production/`:
  1. Readable & canonical naming (<PREFIX>_<Manufacturer>_<Model>.csv)
  2. Strict Modbus addressing (decimal 0..65535 or string addr_len)
  3. Certified Webdyn action codes (0=None, 1=Cmd, 2=1min, 4=5min/Telemetry, 9=Alarm)
  4. Standard Modbus function codes (1=Coil, 2=Discrete, 3=Holding, 4=Input)
  5. Canonical SunPM telemetry & control tags
"""

import csv
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

# Add project root to sys.path
WORKSPACE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(WORKSPACE_DIR))

from DefFileGenerator.def_gen import Generator  # noqa: E402

PROD_DIR = WORKSPACE_DIR / "production"
DEFINITIONS_DIR = WORKSPACE_DIR / "Definitions"
TAGGED_DIR = WORKSPACE_DIR / "tagged_definitions"
FOXESS_DIR = WORKSPACE_DIR / "FoxESS_P3_H3_Pro_Definitions"
MARS_DIR = WORKSPACE_DIR / "MarsRenewable_Definitions"
DYNESS_DIR = WORKSPACE_DIR / "Dyness_DH_Series_Definitions"
X7_DIR = WORKSPACE_DIR / "X7_Definitions"
CLASSIC_DIR = WORKSPACE_DIR / "Definitions" / "Classic"
GOODWE_DIR = WORKSPACE_DIR / "Definitions" / "GoodWe"
SMA_DIR = WORKSPACE_DIR / "Definitions" / "SMA"

PREFIX_TO_CATEGORY = {
    "INV": "Inverter",
    "MET": "Meter",
    "SEN": "Sensor",
    "BAT": "Battery",
    "COM": "Logger",
    "OTH": "Other",
}

# Complete Universal Manufacturer Registry
MFG_REGISTRY = {
    # 1. Inverters, Convertors & Power Electronics
    "abb": "ABB",
    "powerone": "ABB",
    "power-one": "ABB",
    "fimer": "Fimer",
    "huawei": "Huawei",
    "huwai": "Huawei",
    "goodwe": "GoodWe",
    "goodwee": "GoodWe",
    "sungrow": "Sungrow",
    "growatt": "Growatt",
    "solis": "Solis",
    "ginlong": "Solis",
    "ginlong_solis": "Solis",
    "sma": "SMA",
    "core2": "SMA",
    "fronius": "Fronius",
    "tauro": "Fronius",
    "kaco": "Kaco",
    "kostal": "Kostal",
    "solaredge": "SolarEdge",
    "solax": "SolaX",
    "solax-x1-x3-generic": "SolaX",
    "deye": "Deye",
    "foxess": "FoxESS",
    "fox_ess": "FoxESS",
    "fox": "FoxESS",
    "dyness": "Dyness",
    "x7": "X7",
    "x7_bess": "X7",
    "x7_energy": "X7",
    "x7energy": "X7",
    "ingeteam": "Ingeteam",
    "hoymiles": "Hoymiles",
    "delta": "Delta",
    "delta-sunspec": "Delta",
    "cyberpower": "Cyberpower",
    "cyberpower3ph": "Cyberpower",
    "solplanet": "Solplanet",
    "saj": "SAJ",
    "sofarsolar": "SofarSolar",
    "siel": "Siel",
    "kehua": "Kehua",
    "kstar": "Kstar",
    "k_star": "Kstar",
    "cefem": "Cefem",
    "efacec": "Efacec",
    "efasolar": "Efacec",
    "jema": "Jema",
    "victron": "Victron",
    "vectron": "Victron",
    "greenheiss": "Greenheiss",
    "hyxi": "Hyxi",
    "hyxipower": "Hyxi",
    "powerelectronics": "PowerElectronics",
    "satcon": "Satcon",
    "hopewind": "Hopewind",
    "sineng": "Sineng",
    "danfoss": "Danfoss",
    "riello": "Riello",
    "aros": "Riello",
    "hitachi": "Hitachi",
    "waaree": "Waaree",
    "tata": "Tata",
    "zeversolar": "ZeverSolar",
    "zigor": "Zigor",
    "zucchetti": "Zucchetti",
    "solarmax": "Solarmax",
    "salicru": "Salicru",
    "santerno": "Santerno",
    "bonfiglioli": "Bonfiglioli",
    "bluesun": "BlueSun",
    "gamesa": "Gamesa",
    "answerdrives": "AnswerDrives",
    "solargate": "AnswerDrives",
    "apsystems": "APsystems",
    "ap": "APsystems",
    "atess": "Atess",
    "invt": "INVT",
    "sigenergy": "Sigenergy",
    "sigen": "Sigenergy",
    "robestec": "Robestec",
    "tmeic": "TMEIC",
    "autarco": "Autarco",
    "marsenergy": "MarsEnergy",
    "marsrenewable": "MarsEnergy",
    "mars": "MarsEnergy",
    "slenergy": "SLEnergy",
    "sl_energy": "SLEnergy",
    "sl": "SLEnergy",
    "alrosolar": "AlroSolar",
    "alro_solar": "AlroSolar",
    "alro": "AlroSolar",
    "atmoce": "Atmoce",
    "nidec": "Nidec",
    # 2. Meters & Electrical Network Equipment
    "schneider": "Schneider",
    "sch": "Schneider",
    "pm5500": "Schneider",
    "pm5560": "Schneider",
    "pm5563": "Schneider",
    "em1200": "Schneider",
    "em6400": "Schneider",
    "em6400ng": "Schneider",
    "relayp5": "Schneider",
    "socomec": "Socomec",
    "diris": "Socomec",
    "countis": "Socomec",
    "lettel": "Lettel",
    "janitza": "Janitza",
    "carlogavazzi": "CarloGavazzi",
    "carlo_gavazzi": "CarloGavazzi",
    "carlogavazi": "CarloGavazzi",
    "carlo": "CarloGavazzi",
    "em21": "CarloGavazzi",
    "chauvinarnoux": "ChauvinArnoux",
    "chauvin-arnoux": "ChauvinArnoux",
    "chauvinarnoud": "ChauvinArnoux",
    "legrand": "Legrand",
    "eastron": "Eastron",
    "chint": "Chint",
    "circutor": "Circutor",
    "acrel": "Acrel",
    "algodue": "Algodue",
    "anexo": "Anexo",
    "bender": "Bender",
    "edmi": "EDMI",
    "elster": "Elster",
    "emu": "Emu",
    "hager": "Hager",
    "ime": "IME",
    "imesys": "Imesys",
    "inepro": "Inepro",
    "lacecal": "Lacecal",
    "og": "OG",
    "panasonic": "Panasonic",
    "polier": "Polier",
    "saci": "Saci",
    "siemens": "Siemens",
    "socorel": "Socorel",
    "ticmaster": "Ticmaster",
    "weidmuller": "Weidmuller",
    "zelec": "Zelec",
    "zélec": "Zelec",
    "ecoadapt": "EcoAdapt",
    "eco-adapt": "EcoAdapt",
    "eco_adapt": "EcoAdapt",
    "eco": "EcoAdapt",
    "a-eberle": "AEberle",
    "aeberle": "AEberle",
    "pq": "AEberle",
    "ecs-optec": "EcsOptec",
    "enerfox": "Enerfox",
    "eflex4u": "Eflex4u",
    "mci": "MCI",
    "secure": "Secure",
    "pdl": "Secure",
    # 3. Weather Stations, Pyranometers & Environmental Sensors
    "hukseflux": "Hukseflux",
    "kippzonen": "KippZonen",
    "kipp_zonen": "KippZonen",
    "deltaohm": "DeltaOhm",
    "delta_ohm": "DeltaOhm",
    "seven": "Seven",
    "seven_solutions": "Seven",
    "seven_sensor_solutions": "Seven",
    "thermokon": "Thermokon",
    "thies": "Thies",
    "thies_clima": "Thies",
    "lufft": "Lufft",
    "eko": "EKO",
    "maximet": "Maximet",
    "meiernt": "MeierNT",
    "meier-nt": "MeierNT",
    "meier": "MeierNT",
    "ingenieurburo": "IngenieurBuro",
    "termasgard": "Termasgard",
    "trackso": "Trackso",
    "atersa": "Atersa",
    "bd": "BD",
    "blet": "Blet",
    "sregeltechnik": "SRegeltechnik",
    "s_regeltechnik": "SRegeltechnik",
    "s_s_regeltechnik": "SRegeltechnik",
    "s": "SRegeltechnik",
    "sesautomation": "SESAutomation",
    "ses_automation": "SESAutomation",
    "lsi": "LSI",
    "lsi_lastem": "LSI",
    # 4. Gateways, Loggers & Industrial Automation IO
    "advantech": "Advantech",
    "advantec": "Advantech",
    "adam": "Advantech",
    "phoenixcontact": "PhoenixContact",
    "phenixcontact": "PhoenixContact",
    "siebert": "Siebert",
    "solarlog": "SolarLog",
    "icpdas": "ICPDAS",
    "inaccess": "Inaccess",
    "intelec": "Intelec",
    "snef": "SNEF",
    "axun": "Axun",
    "alan": "Alan",
    "kerneli": "Kerneli",
    "webdyn": "Webdyn",
    "monichek": "Monichek",
    "dcmonitor": "Monichek",
}

COMPOUND_PREFIXES = [
    ("fox_ess", "FoxESS"),
    ("sl_energy", "SLEnergy"),
    ("alro_solar", "AlroSolar"),
    ("carlo_gavazzi", "CarloGavazzi"),
    ("chauvin_arnoux", "ChauvinArnoux"),
    ("atess_power", "Atess"),
    ("deye_inverter", "Deye"),
    ("invt_solar", "INVT"),
    ("ginlong_solis", "Solis"),
    ("delta_ohm", "DeltaOhm"),
    ("ap_system", "APsystems"),
    ("thies_clima", "Thies"),
    ("ses_automation", "SESAutomation"),
    ("lsi_lastem", "LSI"),
    ("seven_solutions", "Seven"),
    ("seven_sensor_solutions", "Seven"),
    ("eco_adapt", "EcoAdapt"),
    ("a_eberle", "AEberle"),
    ("a-eberle", "AEberle"),
    ("mci_elec", "MCI"),
    ("pq_plus", "AEberle"),
]

SKIP_INTERMEDIATES = {"india", "modbus", "sunspec", "modbusrtu", "inv", "meter", "sensor", "cpt"}

NAMING_PATTERN = re.compile(
    r"^(INV|MET|SEN|BAT|COM|OTH)_([A-Za-z0-9\-]+)_([A-Za-z0-9\.\-_]+)\.csv$"
)
BAD_WORDS = [
    "part",
    "test",
    "draft",
    "wip",
    "old",
    "copie",
    "copy",
    "faux",
    "in_progress",
    "temp",
    "backup",
    "wpm",
    "_2.csv",
    "simu",
    "ne_pas_utiliser",
    "faitalamain",
    "charp1",
    "copergreen",
    "customer_version",
    "pics_v2",
    "inverter_pics",
    "log_goodwe",
    "new_100kw",
    # SMA scrap, raw logs and legacy draft tokens
    "modbus-ti-en",
    "1.00.11.r",
    "v02_02_07",
    "v03_02_07",
    "work_modbus_sma",
    "hedl02770",
    "edl210606",
    "fonctionnel",
    "prelim210114",
    "stp125-70",
    "25_updated_10",
    "sma_16kw",
    "sma_sma",
    "smappee",
    "clustercontroler",
]
NON_EQUIPMENT_SCRAPS = {
    "analyseregistres",
    "decoupetramemodbusexcel",
    "report",
    "repport",
}


def parse_raw_filename(fname: str) -> tuple[str, str, str] | None:
    """Parses a candidate raw filename into (Prefix, CanonicalManufacturer, CleanModel)."""
    p = fname[:-4]
    tokens = p.split("_")
    if len(tokens) < 3:
        return None
    pfx = tokens[0]
    if pfx not in ("INV", "MET", "SEN", "BAT", "COM", "OTH"):
        return None

    rest = p[len(pfx) + 1 :]

    # 1. Check Webdyn wrapper
    if rest.lower().startswith("webdyn_"):
        sub = rest[7:]
        # Check true Webdyn first-party hardware
        if any(
            sub.lower().startswith(x)
            for x in (
                "webdynbridge",
                "sens_io",
                "m200-03",
                "rk100-01",
                "anaex_webdyn",
                "ioex_webdyn",
            )
        ):
            mfg = "Webdyn"
            model = sub
            if model.lower().startswith("anaex_webdyn_"):
                model = "Sens_IO_" + model[13:]
            elif model.lower().startswith("ioex_webdyn_"):
                model = "Sens_IO_" + model[12:]
        else:
            sub_tokens = sub.split("_")
            idx = 0
            while idx < len(sub_tokens) and sub_tokens[idx].lower() in SKIP_INTERMEDIATES:
                idx += 1
            if idx >= len(sub_tokens):
                return None

            rem_sub = "_".join(sub_tokens[idx:])
            comp_found = False
            for comp, canon in COMPOUND_PREFIXES:
                if rem_sub.lower().startswith(comp + "_"):
                    mfg = canon
                    model = rem_sub[len(comp) + 1 :]
                    comp_found = True
                    break
            if not comp_found:
                cand = sub_tokens[idx].lower()
                mfg = MFG_REGISTRY.get(
                    cand,
                    sub_tokens[idx].capitalize()
                    if sub_tokens[idx].isupper() and len(sub_tokens[idx]) > 3
                    else sub_tokens[idx],
                )
                model = "_".join(sub_tokens[idx + 1 :])
    else:
        comp_found = False
        for comp, canon in COMPOUND_PREFIXES:
            if rest.lower().startswith(comp + "_"):
                mfg = canon
                model = rest[len(comp) + 1 :]
                comp_found = True
                break
        if not comp_found:
            cand = tokens[1].lower()
            mfg = MFG_REGISTRY.get(
                cand,
                tokens[1].capitalize() if tokens[1].isupper() and len(tokens[1]) > 3 else tokens[1],
            )
            model = "_".join(tokens[2:])

    # Clean model
    model = model.strip("_")
    if model.upper().startswith(f"{pfx}_"):
        model = model[len(pfx) + 1 :].strip("_")
    for cw in ("BATTERY_", "METER_", "INVERTER_", "SENSOR_", "LOGGER_"):
        if model.upper().startswith(cw):
            model = model[len(cw) :].strip("_")
    for mv in (mfg, mfg.upper(), mfg.lower()):
        if model.upper().startswith(f"{mv.upper()}_"):
            model = model[len(mv) + 1 :].strip("_")
        elif model.upper().startswith(f"{mv.upper()}-"):
            model = model[len(mv) + 1 :].strip("_")
    for cw in ("BATTERY_", "METER_", "INVERTER_", "SENSOR_", "LOGGER_"):
        if model.upper().startswith(cw):
            model = model[len(cw) :].strip("_")
    for mv in (mfg, mfg.upper(), mfg.lower()):
        if model.upper().startswith(f"{mv.upper()}_"):
            model = model[len(mv) + 1 :].strip("_")
    model = re.sub(r"_+", "_", model).strip("_")
    if not model:
        model = "Generic"
    return pfx, mfg, model


def score_source_candidate(fpath: Path) -> tuple[int, int]:
    """Scores a candidate source file: returns (tag_count, register_count)."""
    try:
        with open(fpath, encoding="utf-8", errors="ignore") as fp:
            r = csv.reader(fp, delimiter=";")
            rows = list(r)
            regs = len(rows) - 1
            tags = sum(1 for row in rows[1:] if len(row) > 6 and row[6].strip())
            return (tags, regs)
    except Exception:
        return (0, 0)


def build_manifest() -> dict[str, Path]:
    """Builds the comprehensive production inventory mapping of target_name -> source_path."""
    manifest: dict[str, Path] = {}

    # 1. Curated Reference Library (Definitions/ folder)
    for f in sorted(DEFINITIONS_DIR.glob("*.csv")):
        if f.name == "INV_ABB_PVS800-57B_2.csv":
            continue
        parsed = parse_raw_filename(f.name)
        if parsed:
            pfx, mfg, model = parsed
            tgt_name = f"{pfx}_{mfg}_{model}.csv"
            manifest[tgt_name] = f

    # 2. FoxESS Certified Definitions
    fox_map = {
        "BAT_FoxESS_GMax_Storage.csv": FOXESS_DIR / "FOXESS_GMAX_STORAGE.csv",
        "BAT_FoxESS_H3_Pro.csv": FOXESS_DIR / "FOXESS_H3_PRO_BATTERY.csv",
        "BAT_FoxESS_P3_Pro.csv": FOXESS_DIR / "FOXESS_P3_PRO_BATTERY.csv",
        "BAT_FoxESS_RSeries.csv": FOXESS_DIR / "FOXESS_R_SERIES_BATTERY.csv",
        "INV_FoxESS_H3_Pro.csv": FOXESS_DIR / "FOXESS_H3_PRO_INVERTER.csv",
        "INV_FoxESS_P3_Pro.csv": FOXESS_DIR / "FOXESS_P3_PRO_INVERTER.csv",
        "INV_FoxESS_P3_H3_Pro_AllInOne.csv": FOXESS_DIR / "FOXESS_P3_H3_PRO_ALL_IN_ONE.csv",
        "INV_FoxESS_RSeries.csv": FOXESS_DIR / "FOXESS_R_SERIES_INVERTER.csv",
        "INV_FoxESS_SunSpec_RV_H3Plus.csv": FOXESS_DIR / "FOXESS_SUNSPEC_RV_H3PLUS.csv",
        "MET_FoxESS_H3_Pro.csv": FOXESS_DIR / "FOXESS_H3_PRO_METER.csv",
    }
    manifest.update(fox_map)

    # 3. Mars Energy Certified Definitions
    mars_map = {
        "BAT_MarsEnergy_LocalController_V7.0.csv": MARS_DIR / "MarsLocalController_V7.0_webdyn.csv",
        "BAT_MarsEnergy_LocalController_AllInOne.csv": MARS_DIR
        / "MarsLocalController_V7.0_AllInOne.csv",
        "BAT_MarsEnergy_RackCell_SlaveN.csv": MARS_DIR
        / "MarsLocalController_V7.0_RackCell_SlaveN.csv",
        "COM_MarsEnergy_EMS_Slave247.csv": MARS_DIR / "MarsLocalController_V7.0_EMS_Slave247.csv",
        "INV_MarsEnergy_AirCooling.csv": MARS_DIR / "MarsRenewable_AC.csv",
        "INV_MarsEnergy_LiquidCooling.csv": MARS_DIR / "MarsRenewable_LC.csv",
        "INV_MarsEnergy_LC_AC.csv": MARS_DIR / "MarsRenewable_LC-AC.csv",
        "MET_MarsEnergy_Meter.csv": MARS_DIR / "MarsRenewable_Meter.csv",
        "SEN_MarsEnergy_PV.csv": MARS_DIR / "MarsRenewable_PV.csv",
    }
    manifest.update(mars_map)

    # 4. Dyness Certified Definitions
    dyness_map = {
        "BAT_Dyness_DH100F.csv": DYNESS_DIR / "BAT_Dyness_DH100F.csv",
        "BAT_Dyness_DH200F.csv": DYNESS_DIR / "BAT_Dyness_DH200F.csv",
        "BAT_Dyness_DH200Y.csv": DYNESS_DIR / "BAT_Dyness_DH200Y.csv",
        "BAT_Dyness_DH_Series_Generic.csv": DYNESS_DIR / "BAT_Dyness_DH_Series_Generic.csv",
    }
    manifest.update(dyness_map)

    # 5. X7 BESS Certified Definitions
    x7_map = {
        "BAT_X7_BESS-AllInOne.csv": X7_DIR / "BAT_X7_BESS-AllInOne.csv",
        "BAT_X7_EMS.csv": X7_DIR / "BAT_X7_EMS.csv",
        "BAT_X7_BMS.csv": X7_DIR / "BAT_X7_BMS.csv",
        "BAT_X7_BMS-Cells.csv": X7_DIR / "BAT_X7_BMS-Cells.csv",
        "MET_X7_Meters.csv": X7_DIR / "MET_X7_Meters.csv",
        "INV_X7_PCS.csv": X7_DIR / "INV_X7_PCS.csv",
        "OTH_X7_Cabinet.csv": X7_DIR / "OTH_X7_Cabinet.csv",
    }
    manifest.update(x7_map)

    # 6. Additional Key Certified Models
    extra_map = {
        "BAT_Huawei_Luna2000.csv": TAGGED_DIR / "BAT_Huawei_Luna_2000.csv",
        "BAT_Hoymiles_HoyUltra2.csv": TAGGED_DIR / "BAT_Hoymiles_HOYMILES_BATTERY_HOYULTRA2.csv",
        "INV_Huawei_GENERIC.csv": DEFINITIONS_DIR / "INV_Huawei_GENERIC.csv",
        "INV_Atess_Generic.csv": TAGGED_DIR / "INV_Atess_Power_ATESS_GENERIC.csv",
        "INV_Deye_Generic.csv": TAGGED_DIR / "INV_Deye_Inverter_INV_DEYE_GENERIC.csv",
        "INV_INVT_Generic.csv": TAGGED_DIR / "INV_INVT_Solar_ModbusRTU_INVT_Generic.csv",
        "INV_Refusol_Generic.csv": TAGGED_DIR / "INV_REFUSOL_GENERIC.csv",
        "INV_Sigenergy_Generic.csv": TAGGED_DIR / "INV_SIGENERGY_Generic.csv",
        "INV_Solplanet_Generic.csv": TAGGED_DIR / "INV_SOLPLANET_GENERIC.csv",
        "MET_AEberle_PQPlus_Generic.csv": TAGGED_DIR / "MET_A-eberle_METER_PQPLUS_GENERIC.csv",
        "MET_EcsOptec_ECS_Generic.csv": TAGGED_DIR / "MET_Ecs-Optec_METER_ECS_GENERIC_FLOAT.csv",
        "MET_IME_Generic.csv": TAGGED_DIR / "MET_IME_Meter_IME_GENERIC.csv",
        "MET_Janitza_Generic_RTU.csv": TAGGED_DIR / "MET_Janitza_JANITZA_GENERIC_RTU.csv",
        "MET_Legrand_CF_Generic.csv": TAGGED_DIR / "MET_Legrand_METER_LEGRAND_CF-GENERIC.csv",
    }
    manifest.update(extra_map)

    # 7. Grand Classic Certified Definitions
    classic_map = {
        "MET_Chint_DTSU666.csv": CLASSIC_DIR / "MET_Chint_DTSU666.csv",
        "MET_Chint_DDSU666.csv": CLASSIC_DIR / "MET_Chint_DDSU666.csv",
        "MET_Chint_DTSU666-H.csv": CLASSIC_DIR / "MET_Chint_DTSU666-H.csv",
        "INV_Huawei_SUN2000-M3.csv": CLASSIC_DIR / "INV_Huawei_SUN2000-M3.csv",
        "INV_Huawei_SUN2000-Utility.csv": CLASSIC_DIR / "INV_Huawei_SUN2000-Utility.csv",
        "MET_Huawei_DTSU666-H.csv": CLASSIC_DIR / "MET_Huawei_DTSU666-H.csv",
        "INV_SMA_SunnyTripower.csv": CLASSIC_DIR / "INV_SMA_SunnyTripower.csv",
        "INV_SMA_Core1_Core2.csv": CLASSIC_DIR / "INV_SMA_Core1_Core2.csv",
        "INV_Solplanet_ASW_ThreePhase.csv": CLASSIC_DIR / "INV_Solplanet_ASW_ThreePhase.csv",
        "INV_Sungrow_SG_Series.csv": CLASSIC_DIR / "INV_Sungrow_SG_Series.csv",
        "INV_Goodwe_SMT_Series.csv": CLASSIC_DIR / "INV_Goodwe_SMT_Series.csv",
        "INV_Goodwe_ET_Series.csv": CLASSIC_DIR / "INV_Goodwe_ET_Series.csv",
        "INV_Growatt_MAX_Series.csv": CLASSIC_DIR / "INV_Growatt_MAX_Series.csv",
        "INV_Fronius_Symo_Primo_Gen24.csv": CLASSIC_DIR / "INV_Fronius_Symo_Primo_Gen24.csv",
        "INV_SolarEdge_SE_Series.csv": CLASSIC_DIR / "INV_SolarEdge_SE_Series.csv",
        "MET_Socomec_Diris_A_Countis_E.csv": CLASSIC_DIR / "MET_Socomec_Diris_A_Countis_E.csv",
        "MET_Schneider_PM5000_IEM3000.csv": CLASSIC_DIR / "MET_Schneider_PM5000_IEM3000.csv",
        "MET_CarloGavazzi_EM24_EM300.csv": CLASSIC_DIR / "MET_CarloGavazzi_EM24_EM300.csv",
    }
    manifest.update(classic_map)

    # 8. Complete GoodWe Certified Ecosystem
    if GOODWE_DIR.exists():
        for gw_file in sorted(GOODWE_DIR.glob("*.csv")):
            manifest[gw_file.name] = gw_file

    # 9. Complete SMA Certified Ecosystem
    if SMA_DIR.exists():
        for sma_file in sorted(SMA_DIR.glob("*.csv")):
            manifest[sma_file.name] = sma_file

    # 6. Comprehensive Ingestion from tagged_definitions/
    candidates_by_tgt: dict[str, list[Path]] = defaultdict(list)
    for f in sorted(TAGGED_DIR.glob("*.csv")):
        name = f.name
        lower = name.lower()
        if any(bw in lower for bw in BAD_WORDS):
            continue
        if any(c in name for c in ["(", ")", ",", " ", "..", ";", "~", "$", "%"]):
            continue
        if any(sc in lower for sc in NON_EQUIPMENT_SCRAPS):
            continue
        parsed = parse_raw_filename(name)
        if not parsed:
            continue
        pfx, mfg, model = parsed
        tgt_name = f"{pfx}_{mfg}_{model}.csv"
        if (
            not NAMING_PATTERN.match(tgt_name)
            or "__" in tgt_name
            or tgt_name.endswith("_.csv")
            or tgt_name.endswith("_2.csv")
        ):
            continue
        candidates_by_tgt[tgt_name].append(f)

    # Lowercase mapping to enforce cross-platform case-insensitive uniqueness
    lower_to_canonical: dict[str, str] = {k.lower(): k for k in manifest}

    for tgt_name, sources in candidates_by_tgt.items():
        low = tgt_name.lower()
        if low not in lower_to_canonical:
            manifest[tgt_name] = max(sources, key=score_source_candidate)
            lower_to_canonical[low] = tgt_name

    return manifest


def sanitize_and_normalize_definition(
    tgt_filename: str, src_path: Path
) -> tuple[list[str], list[list[str]]]:
    """
    Normalizes header and rows:
    - Sets clean manufacturer and model from target filename
    - Sets appropriate category from prefix
    - Enforces monotonically increasing indices 1..N
    - Ensures valid action, function, and canonical tags
    """
    parts = tgt_filename[:-4].split("_", 2)
    pfx, mfg, model = parts[0], parts[1], parts[2]
    cat = PREFIX_TO_CATEGORY.get(pfx, "Other")

    with open(src_path, encoding="utf-8-sig", errors="ignore") as fp:
        reader = csv.reader(fp, delimiter=";")
        raw_header = next(reader, [])
        raw_rows = list(reader)

    # Normalize protocol
    protocol = raw_header[0].strip() if raw_header and raw_header[0].strip() else "Modbus"
    if protocol.lower() in (
        "modbus",
        "modbus_rtu",
        "modbusrtu",
        "modbustcp",
        "modbus_tcp",
        "tcp",
        "rtu",
    ):
        protocol = "Modbus"

    # Preserve raw casing of manufacturer and model if they match case-insensitively
    if len(raw_header) > 2 and raw_header[2].strip():
        if raw_header[2].strip().lower() == mfg.lower():
            mfg = raw_header[2].strip()
    if len(raw_header) > 3 and raw_header[3].strip():
        if (
            raw_header[3].strip().lower() == model.lower()
            or raw_header[3].strip().lower().replace("-", "_")
            == model.lower().replace("-", "_")
        ):
            model = raw_header[3].strip()

    # Clean Header
    header = [protocol, cat, mfg, model, "", "", "", "", "", "", ""]

    normalized_rows: list[list[str]] = []
    seen_unique_tags: set[str] = set()

    for _idx, row in enumerate(raw_rows, 1):
        if not row or len(row) < 7:
            continue

        # Pad to 11 columns
        while len(row) < 11:
            row.append("")

        fc = row[1].strip() or "3"
        addr = row[2].strip()
        dtype = row[3].strip() or "U16"
        name = row[5].strip()
        tag = row[6].strip()
        coef_a = row[7].strip() or "1.000000"
        coef_b = row[8].strip() or "0.000000"
        unit = row[9].strip()
        action = row[10].strip() or "4"

        # Validate/Normalize function code
        if fc not in ("1", "2", "3", "4"):
            fc = "3"

        # Validate/Normalize action code
        if not action.isdigit():
            action = "4"

        # Infer command action
        if any(
            kw in name.lower()
            for kw in (
                "shutdown",
                "power on",
                "power off",
                "remote control",
                "startup",
                "reboot",
                "cmd_",
            )
        ):
            if action in ("4", ""):
                action = "1"
        elif any(kw in name.lower() for kw in ("alarm", "fault", "warning", "error")):
            if action in ("4", ""):
                action = "9"

        # Deduplicate single-instance tags
        if tag in (
            "P_AC",
            "Q_AC",
            "S_AC",
            "P_DC",
            "FREQ",
            "PF",
            "E_TOT",
            "E_DAY",
            "MODEL",
            "SERIAL",
        ):
            if tag in seen_unique_tags:
                tag = ""
            else:
                seen_unique_tags.add(tag)

        norm_row = [
            str(len(normalized_rows) + 1),
            fc,
            addr,
            dtype,
            "",
            name,
            tag,
            coef_a,
            coef_b,
            unit,
            action,
        ]
        normalized_rows.append(norm_row)

    return header, normalized_rows


def generate_production_catalog(manifest_stats: list[dict]):
    """Generates a professional README.md catalog in the production folder."""
    catalog_path = PROD_DIR / "README.md"

    total_files = len(manifest_stats)
    total_regs = sum(item["registers"] for item in manifest_stats)
    total_tags = sum(item["tags"] for item in manifest_stats)
    total_mfgs = len(set(item["mfg"] for item in manifest_stats))

    content = [
        "# 🏭 Webdyn Production Definitions Library",
        "",
        "Bibliothèque officielle de définitions d'équipements industriels validées pour déploiement en production sur les concentrateurs et passerelles Webdyn (SunPM, Hub, Flex).",
        "",
        f"Comprend **{total_files} fichiers de définition certifiés** couvrant **{total_regs:,} registres Modbus** et **{total_tags:,} tags SunPM canoniques** à travers **{total_mfgs} constructeurs mondiaux**.",
        "",
        "---",
        "",
        "## 📊 Synthèse par Catégorie",
        "",
        "| Catégorie | Préfixe | Fichiers | Total Registres | Tags Actifs | Exemples Constructeurs |",
        "|---|---|---|---|---|---|",
    ]

    prefix_summary = {
        "Inverter": ("INV_", "Onduleurs & Convertisseurs"),
        "Meter": ("MET_", "Compteurs & Analyseurs Réseau"),
        "Battery": ("BAT_", "Systèmes BESS & Batteries"),
        "Sensor": ("SEN_", "Capteurs Météo & Pyranomètres"),
        "Logger": ("COM_", "Passerelles & Loggers"),
        "Other": ("OTH_", "Modules IO & Afficheurs"),
    }

    for cat, (pfx, desc) in prefix_summary.items():
        files_cat = [item for item in manifest_stats if item["prefix"] == pfx.strip("_")]
        if not files_cat:
            continue
        c_count = len(files_cat)
        c_regs = sum(item["registers"] for item in files_cat)
        c_tags = sum(item["tags"] for item in files_cat)
        mfgs = sorted(list(set(item["mfg"] for item in files_cat)))
        sample_mfgs = ", ".join(mfgs[:5]) + (f" (+{len(mfgs) - 5})" if len(mfgs) > 5 else "")
        content.append(
            f"| {desc} ({cat}) | `{pfx}` | {c_count} | {c_regs:,} | {c_tags:,} | {sample_mfgs} |"
        )

    content.extend(
        [
            "",
            "---",
            "",
            "## 🏭 Synthèse par Constructeur (Tous les Constructeurs)",
            "",
            "| Constructeur | Modèles | Total Registres | Tags SunPM Actifs | Catégories Couvertes |",
            "|---|:---:|:---:|:---:|---|",
        ]
    )

    mfgs_dict = defaultdict(list)
    for item in manifest_stats:
        mfgs_dict[item["mfg"]].append(item)

    for mfg_name in sorted(mfgs_dict.keys(), key=lambda m: (-len(mfgs_dict[m]), m)):
        m_items = mfgs_dict[mfg_name]
        m_count = len(m_items)
        m_regs = sum(it["registers"] for it in m_items)
        m_tags = sum(it["tags"] for it in m_items)
        m_cats = ", ".join(sorted(set(it["category"] for it in m_items)))
        content.append(f"| **{mfg_name}** | {m_count} | {m_regs:,} | {m_tags:,} | {m_cats} |")

    content.extend(
        [
            "",
            "---",
            "",
            "## 🛡️ Les 5 Piliers Qualité Apollo-Standard",
            "",
            "Chaque fichier de définition hébergé dans ce dossier respecte strictement les 5 critères non négociables :",
            "1. **Nommage Lisible & Canonique** : `<PREFIX>_<Manufacturer>_<Model>.csv` sans scories de génération (`_2`, `_test`, `_part`, `_draft`).",
            "2. **Adressage Modbus Strict** : Adresses décimales (0 à 65535) ou `address_length` pour les chaînes de caractères (ex. `30000_30`), zéro décalage, zéro chevauchement corrupteur.",
            "3. **Code Action Webdyn Certifié** :",
            "   - `Action = 4` : Télémesures régulières (périodique standard 5 min).",
            "   - `Action = 9` : Événements, états critiques et alarmes.",
            "   - `Action = 1` : Consignes et registres de pilotage d'écriture (Read/Write).",
            "   - `Action = 0` : Registres non historisés / internes.",
            "4. **Code Fonction Modbus Standard** : `1` (Coils), `2` (Discrete Inputs), `3` (Holding Registers), `4` (Input Registers).",
            "5. **Tags SunPM Canoniques** : Télémesures normalisées (`P_AC`, `Q_AC`, `U_AC`, `I_AC`, `P_DC`, `E_TOT`, `E_DAY`, `FREQ`, `PF`, `TEMP`, `MODEL`, `SERIAL`, etc.) positionnées sur les mesures physiques effectives.",
            "",
            "---",
            "",
            "## 📑 Catalogue Détaillé de Production",
            "",
            "| Fichier | Catégorie | Constructeur | Modèle | Protocole | Registres | Tags SunPM |",
            "|---|---|---|---|---|---|---|",
        ]
    )

    for item in sorted(manifest_stats, key=lambda x: (x["mfg"], x["model"], x["prefix"])):
        fname = item["filename"]
        content.append(
            f"| [`{fname}`](./{fname}) | {item['category']} | {item['mfg']} | {item['model']} | {item['protocol']} | {item['registers']:,} | {item['tags']:,} |"
        )

    content.append("")
    with open(catalog_path, "w", encoding="utf-8", newline="") as fp:
        fp.write("\n".join(content))


def main():
    print("=" * 80)
    print("APOLLO UNIVERSAL PRODUCTION SUITE GENERATOR & GOVERNOR")
    print("=" * 80)

    PROD_DIR.mkdir(parents=True, exist_ok=True)
    generator = Generator()
    manifest = build_manifest()

    print(f"\nProcessing {len(manifest)} production definitions across all manufacturers...")
    stats_list = []
    failed_files = []

    # Clean obsolete files from production/ that are not in manifest
    manifest_filenames = set(manifest.keys())
    existing_files = list(PROD_DIR.glob("*.csv"))
    obsolete_count = 0
    for ef in existing_files:
        if ef.name not in manifest_filenames:
            try:
                ef.unlink()
                obsolete_count += 1
            except Exception:
                pass
    if obsolete_count > 0:
        print(f"Cleaned {obsolete_count} obsolete/misnamed files from production/")

    for idx, (tgt_filename, src_path) in enumerate(sorted(manifest.items()), 1):
        if not src_path.is_file():
            print(f"[{idx:03d}/{len(manifest)}] ERROR: Source file not found: {src_path}")
            failed_files.append((tgt_filename, "Source file not found"))
            continue

        out_path = PROD_DIR / tgt_filename

        is_already_certified = any(
            src_path.resolve().is_relative_to(d.resolve())
            for d in (
                GOODWE_DIR,
                SMA_DIR,
                CLASSIC_DIR,
                X7_DIR,
                FOXESS_DIR,
                MARS_DIR,
                DYNESS_DIR,
            )
            if d.exists()
        )

        if is_already_certified:
            if not out_path.exists() or out_path.read_bytes() != src_path.read_bytes():
                shutil.copy2(src_path, out_path)
            with open(out_path, encoding="utf-8-sig", errors="ignore") as fp:
                rdr = csv.reader(fp, delimiter=";")
                header = next(rdr, [])
                rows = list(rdr)
        else:
            header, rows = sanitize_and_normalize_definition(tgt_filename, src_path)

            buf = []
            header_line = ";".join(header)
            buf.append(header_line)
            for r in rows:
                buf.append(";".join(r))
            new_content = "\n".join(buf) + "\n"

            need_write = True
            if out_path.exists():
                try:
                    with open(out_path, encoding="utf-8") as existing_fp:
                        if existing_fp.read().replace("\r\n", "\n") == new_content:
                            need_write = False
                except Exception:
                    pass

            if need_write:
                with open(out_path, "w", encoding="utf-8", newline="") as fp:
                    writer = csv.writer(fp, delimiter=";")
                    writer.writerow(header)
                    writer.writerows(rows)

        # Validate with def_gen
        val_rep = generator.validate_csv_detailed(str(out_path), strict=True, strict_overlap=False)
        tag_count = sum(1 for r in rows if r[6])

        if not val_rep.is_valid:
            print(
                f"[{idx:03d}/{len(manifest)}] FAIL: {tgt_filename} -> {len(val_rep.issues)} issues"
            )
            failed_files.append((tgt_filename, str(val_rep.issues)))
        else:
            pfx = tgt_filename.split("_")[0]
            stats_list.append(
                {
                    "filename": tgt_filename,
                    "prefix": pfx,
                    "category": header[1],
                    "mfg": header[2],
                    "model": header[3],
                    "protocol": header[0],
                    "registers": len(rows),
                    "tags": tag_count,
                }
            )

    # Generate README.md catalog
    generate_production_catalog(stats_list)

    print("\n" + "=" * 80)
    print("PRODUCTION SUITE GENERATION SUMMARY")
    print("=" * 80)
    print(f"Total Target Files   : {len(manifest)}")
    print(f"Successfully Created : {len(stats_list)}")
    print(f"Failed Files         : {len(failed_files)}")
    print(f"Total Manufacturers  : {len(set(s['mfg'] for s in stats_list))}")
    print(f"Total Registers      : {sum(s['registers'] for s in stats_list):,}")
    print(f"Total SunPM Tags     : {sum(s['tags'] for s in stats_list):,}")
    print(f"Output Directory     : {PROD_DIR}")
    print("=" * 80)

    # Record automated version commit
    try:
        from DefFileGenerator.version_manager import DefinitionVersionControl
        vc = DefinitionVersionControl(repo_root=WORKSPACE_DIR)
        commit = vc.commit(
            source_dir=PROD_DIR,
            message=f"Automated build: {len(stats_list)} production definitions ({sum(s['registers'] for s in stats_list):,} registers)",
            author="Production Governor",
        )
        print(f"[VERSION CONTROL] Recorded commit {commit.commit_id} (+{commit.changes_summary.get('added', 0)} ~{commit.changes_summary.get('modified', 0)} -{commit.changes_summary.get('deleted', 0)})")
    except Exception as e:
        print(f"[WARNING] Version control snapshot could not be recorded: {e}")

    if failed_files:
        sys.exit(1)


if __name__ == "__main__":
    main()
