#!/usr/bin/env python3
"""
Full Technical Documentation Test Runner & Comprehensive Report Generator.
Executes extraction and strict Webdyn validation across all technical documents
in the dataset (Mars Renewable, 30+ Manufacturers, and R&D productions).
Outputs structured results for audit reporting.
"""

import json
import logging
import os
import sys
import tempfile
import time
from typing import Any

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../..")))

from DefFileGenerator.def_gen import Generator, GeneratorConfig, run_generator
from DefFileGenerator.extractor import Extractor, peek_generator

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger("FullDocBattery")

WORKSPACE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
EQUIPEMENTIERS_DIR = r"G:\My Drive\08092026\Equipementiers"
R_AND_D_DIR = r"G:\My Drive\08092026\DEF_PM_R&D"
MARS_SOURCE = r"C:\Users\Cylae\Downloads\MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"

# Technical documents catalog to test thoroughly
CATALOG = [
    # -------------------------------------------------------------------------
    # 0. User's Main Target: Mars Renewable Local Controller V7.0
    # -------------------------------------------------------------------------
    {
        "id": "MARS_01",
        "category": "Battery / BESS",
        "mfg": "Mars Energy",
        "model": "MarsLocalController_V7.0",
        "protocol": "modbusTCP",
        "path": MARS_SOURCE,
        "sheet": None,  # All sheets merged
        "output_name": "MarsLocalController_V7.0_webdyn.csv",
        "description": "Mars Local Controller V7.0 BESS Architecture (All 12 Subsystems Merged)",
        "difficulty": "Classeur Excel industriel complexe (12 onglets, 2 593 registres), mélange de Holding FC03 et Coils FC01, blocs de cellules fusionnées, adresses 0-based et 1-based, sous-tables de réservation.",
    },
    {
        "id": "MARS_02",
        "category": "Battery / EMS",
        "mfg": "Mars Energy",
        "model": "MarsLocalController_EMS_Slave247",
        "protocol": "modbusTCP",
        "path": MARS_SOURCE,
        "sheet": "LC Information",
        "output_name": "MarsLocalController_V7.0_EMS_Slave247.csv",
        "description": "Mars Local Controller Slave 247 (EMS Energy Management Commands & Setpoints)",
        "difficulty": "Feuille LC Information avec registres de commande en écriture (Action=1 pour Set Active Power Ratio, Reset, Switch On/Off) nécessitant détection automatique de droit d'accès R/W.",
    },
    {
        "id": "MARS_03",
        "category": "Battery / Cell",
        "mfg": "Mars Energy",
        "model": "MarsLocalController_RackCell_SlaveN",
        "protocol": "modbusTCP",
        "path": MARS_SOURCE,
        "sheet": "Rack and Cell Information",
        "output_name": "MarsLocalController_V7.0_RackCell_SlaveN.csv",
        "description": "Mars Local Controller Slave N (Multi-Rack & Cell Subsystems)",
        "difficulty": "Table massive de télémesures de tension/température cellule, normalisation des facteurs de division 100/1000 pour les millivolts et dixièmes de degrés.",
    },
    # -------------------------------------------------------------------------
    # 1. Huawei (Benchmark Gold Standard)
    # -------------------------------------------------------------------------
    {
        "id": "HUA_01",
        "category": "Inverter",
        "mfg": "Huawei",
        "model": "SUN2000MA_V100R001C00",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"Huawei\Inverter\Sources\SUN2000MA_V100R001C00_ModbusInterfaceDefinitions_issue08_2024.pdf",
        ),
        "sheet": None,
        "output_name": "INV_Huawei_SUN2000MA_V100R001C00.csv",
        "description": "Manuel PDF officiel d'interface Modbus Huawei SUN2000 (Issue 08, 2024)",
        "difficulty": "PDF vectoriel technique multi-pages, tableaux avec en-têtes imbriqués, adresses 30000/40000, formats U16/U32_WB, filtres stricts des avertissements de sécurité.",
    },
    {
        "id": "HUA_02",
        "category": "Inverter",
        "mfg": "Huawei",
        "model": "SUN2000",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Huawei\Inverter\HUAWEI_V4.csv"),
        "sheet": None,
        "output_name": "INV_Huawei_SUN2000_V4.csv",
        "description": "Table de référence de production R&D Huawei V4 (Golden Benchmark)",
        "difficulty": "Fichier de référence absolu pour WebdynSunPM (322 registres, 11 colonnes strictes, 6 décimales sur CoefA/CoefB).",
    },
    {
        "id": "HUA_03",
        "category": "Communication",
        "mfg": "Huawei",
        "model": "SmartLogger3000",
        "protocol": "modbusTCP",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR, r"Huawei\SmartLogger\SL3000\HUAWEI_SmartLogger3000.csv"
        ),
        "sheet": None,
        "output_name": "COM_Huawei_SmartLogger3000.csv",
        "description": "Passerelle de communication Huawei SmartLogger 3000",
        "difficulty": "Multiplexage d'alarmes bit-field, registres de statut système, normalisation des tags de puissance totale.",
    },
    # -------------------------------------------------------------------------
    # 2. ABB / PowerOne
    # -------------------------------------------------------------------------
    {
        "id": "ABB_01",
        "category": "Inverter",
        "mfg": "ABB",
        "model": "PVS1xx_SunSpec",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR, r"PowerOne\Sunspec\Sources\SUNSPEC_PVS1xx_PICS_V002_GL20181011.xlsx"
        ),
        "sheet": None,
        "output_name": "INV_ABB_PVS1xx_SunSpec.csv",
        "description": "Spécification PICS SunSpec Excel pour onduleurs triphasés ABB PVS1xx",
        "difficulty": "Modèle SunSpec avec Scale Factors dynamiques (sunssf / sunssf-relative), types de données acc32, acc64, enum16 et pad.",
    },
    {
        "id": "ABB_02",
        "category": "Meter",
        "mfg": "ABB",
        "model": "M4M_NetworkAnalyzer",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"PowerOne\Modbus\sources\analyseur reseau M4M\M4M Modbus map - v.1.3N.xlsx",
        ),
        "sheet": None,
        "output_name": "MET_ABB_M4M_NetworkAnalyzer.csv",
        "description": "Analyseur de réseau haute précision ABB M4M (v1.3N)",
        "difficulty": "Adresses hexadécimales et décimales alternées, formats virgule flottante IEEE 754 32-bit Big Endian, sous-pages de mesures harmoniques.",
    },
    {
        "id": "ABB_03",
        "category": "Other / EV",
        "mfg": "ABB",
        "model": "Terra_AC_Charger",
        "protocol": "modbusTCP",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"PowerOne\ABB Terra Heavy charger\ABB_Terra_AC_Charger_ModbusCommunication_v1.7.pdf",
        ),
        "sheet": None,
        "output_name": "OTH_ABB_Terra_AC_Charger.csv",
        "description": "Manuel PDF Borne de recharge pour véhicules électriques ABB Terra AC (v1.7)",
        "difficulty": "Documentation PDF avec structures de données composites, registres d'état de charge et gestion de puissance dynamique.",
    },
    # -------------------------------------------------------------------------
    # 3. Fox ESS
    # -------------------------------------------------------------------------
    {
        "id": "FOX_01",
        "category": "Inverter",
        "mfg": "Fox ESS",
        "model": "G-MAX_ModbusTCP",
        "protocol": "modbusTCP",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"Fox ESS\G-MAX Communication Protocol modbusTcp-EN-20250425(1).xlsx",
        ),
        "sheet": None,
        "output_name": "INV_FoxESS_G_MAX_ModbusTCP.csv",
        "description": "Protocole de communication Modbus TCP Fox ESS G-MAX (Excel multi-feuilles)",
        "difficulty": "Feuilles multiples contenant des descriptions textuelles et des tables Modbus, détection automatique des lignes d'en-tête Modbus.",
    },
    {
        "id": "FOX_02",
        "category": "Inverter",
        "mfg": "Fox ESS",
        "model": "RSeries_Inverter",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Fox ESS\INV_FOXESS_RSERIESINVERTER.csv"),
        "sheet": None,
        "output_name": "INV_FoxESS_RSeries_Inverter.csv",
        "description": "Définition de production Fox ESS R-Series",
        "difficulty": "Mapping de tags de production, gestion des coefficients de conversion puissance/tension.",
    },
    {
        "id": "FOX_03",
        "category": "Inverter",
        "mfg": "Fox ESS",
        "model": "R_V_H3plus_GmaxPlus",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Fox ESS\R&V&H3plus&GmaxPlus.csv"),
        "sheet": None,
        "output_name": "INV_FoxESS_R_V_H3plus_GmaxPlus.csv",
        "description": "Table Modbus unifiée gammes R, V, H3 Plus et G-Max Plus",
        "difficulty": "Groupement de modèles multi-gammes, prévention des doublons d'adresses Modbus.",
    },
    # -------------------------------------------------------------------------
    # 4. GoodWe
    # -------------------------------------------------------------------------
    {
        "id": "GDW_01",
        "category": "Inverter",
        "mfg": "GoodWe",
        "model": "ThreePhase_SunSpec",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"Goodwe\Sources\Sunspec for Inverters\GOODWE_three_phase_grid-tied_inverter_PICS.xlsx",
        ),
        "sheet": None,
        "output_name": "INV_GoodWe_ThreePhase_SunSpec.csv",
        "description": "Onduleur réseau triphasé GoodWe PICS SunSpec (Excel)",
        "difficulty": "Champs SunSpec normalisés, gestion des offsets d'adresses relatifs au modèle SunSpec de base (Model 1, Model 103, Model 160).",
    },
    {
        "id": "GDW_02",
        "category": "Inverter",
        "mfg": "GoodWe",
        "model": "SMT_Series",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Goodwe\GoodWe_SMT_Series.csv"),
        "sheet": None,
        "output_name": "INV_GoodWe_SMT_Series.csv",
        "description": "Onduleur GoodWe SMT Series",
        "difficulty": "Extraction de paramètres multi-MPPT, tags normalisés de puissance et fréquence réseau.",
    },
    {
        "id": "GDW_03",
        "category": "Inverter",
        "mfg": "GoodWe",
        "model": "ET_EH_Hybrid_Series",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR, r"Goodwe\GoodWe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv"
        ),
        "sheet": None,
        "output_name": "INV_GoodWe_ET_EH_Hybrid_Series.csv",
        "description": "Gamme hybride avec stockage batterie GoodWe ET/EH/BT/BH Modbus RTU",
        "difficulty": "Données hybrides combinant production solaire photovoltaïque et charge/décharge batterie.",
    },
    # -------------------------------------------------------------------------
    # 5. Growatt
    # -------------------------------------------------------------------------
    {
        "id": "GRW_01",
        "category": "Inverter",
        "mfg": "Growatt",
        "model": "Growatt_RTU_Protocol_II",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"Growatt\Sources\Growatt Inverter Modbus RTU Protocol_II V1 07 （20180918）.pdf",
        ),
        "sheet": None,
        "output_name": "INV_Growatt_RTU_Protocol_II.csv",
        "description": "Documentation PDF Protocole Modbus RTU II Growatt (v1.07)",
        "difficulty": "Tableau PDF technique dense avec adresses hexadécimales et décimales, gestion de l'encodage de caractères asiatiques.",
    },
    {
        "id": "GRW_02",
        "category": "Inverter",
        "mfg": "Growatt",
        "model": "Growatt_PV_Inverter_V3.14",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR,
            r"Growatt\Sources\Growatt PV Inverter Modbus RS485 RTU Protocol V3 14.pdf",
        ),
        "sheet": None,
        "output_name": "INV_Growatt_PV_Inverter_V3.14.csv",
        "description": "Manuel PDF Growatt RS485 RTU V3.14",
        "difficulty": "Tableaux Modbus séparés par des notes de bas de page et des sections de configuration de câblage.",
    },
    {
        "id": "GRW_03",
        "category": "Inverter",
        "mfg": "Growatt",
        "model": "Growatt_Generic",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Growatt\W_Growatt_Generic.csv"),
        "sheet": None,
        "output_name": "INV_Growatt_Generic.csv",
        "description": "Définition de production Growatt Générique",
        "difficulty": "Table étendue de registres de mesure de rendement et de production journalière/totale.",
    },
    # -------------------------------------------------------------------------
    # 6. Hukseflux
    # -------------------------------------------------------------------------
    {
        "id": "HUK_01",
        "category": "Sensor",
        "mfg": "Hukseflux",
        "model": "SR20-D2_Doc_FR",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Hukseflux\Sources\Hukseflux_SR20-D2_Doc_FR.pdf"),
        "sheet": None,
        "output_name": "SEN_Hukseflux_SR20_D2.csv",
        "description": "Pyranomètre secondaire standard Hukseflux SR20-D2 (Documentation PDF en français)",
        "difficulty": "PDF complexe où la table Modbus n'a pas de bordures vectorielles de cellules, contient un tableau préalable de câblage et un sous-en-tête interne. Résolu par détection d'en-tête secondaire et fallback par regex textuelle.",
    },
    {
        "id": "HUK_02",
        "category": "Sensor",
        "mfg": "Hukseflux",
        "model": "SR30_Pyranometer",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Hukseflux\Hukseflux_SR30.csv"),
        "sheet": None,
        "output_name": "SEN_Hukseflux_SR30.csv",
        "description": "Pyranomètre Hukseflux SR30 avec ventilation et chauffage intégrés",
        "difficulty": "Capteur météo haute précision avec mesures de température corporelle du capteur et compensation de dérive.",
    },
    {
        "id": "HUK_03",
        "category": "Sensor",
        "mfg": "Hukseflux",
        "model": "SR05_Pyranometer",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Hukseflux\Hukseflux_SR05-D1A3.csv"),
        "sheet": None,
        "output_name": "SEN_Hukseflux_SR05_D1A3.csv",
        "description": "Pyranomètre compact Hukseflux SR05-D1A3",
        "difficulty": "Registres d'irradiance solaire (W/m²), contrôle de qualité des données et tension d'alimentation.",
    },
    # -------------------------------------------------------------------------
    # 7. Janitza
    # -------------------------------------------------------------------------
    {
        "id": "JAN_01",
        "category": "Meter",
        "mfg": "Janitza",
        "model": "UMG604-PRO",
        "protocol": "modbusTCP",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR, r"Janitza\Fichier de def inconnu\METER_JANITZA_UMG604-PRO.csv"
        ),
        "sheet": None,
        "output_name": "MET_Janitza_UMG604_PRO.csv",
        "description": "Analyseur de puissance réseau et qualité d'énergie Janitza UMG 604-PRO",
        "difficulty": "Comptage quadri-quadrant haute vitesse (consommation active, réactive, harmonique, cos phi).",
    },
    {
        "id": "JAN_02",
        "category": "Meter",
        "mfg": "Janitza",
        "model": "Janitza_Generic_RTU",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Janitza\JANITZA_GENERIC_RTU.csv"),
        "sheet": None,
        "output_name": "MET_Janitza_Generic_RTU.csv",
        "description": "Compteur Janitza Modbus RTU Générique",
        "difficulty": "Gestion des formats Float32 et entiers 32 bits non signés (U32_WB / F32_WB).",
    },
    {
        "id": "JAN_03",
        "category": "Meter",
        "mfg": "Janitza",
        "model": "Janitza_Generic_TCP",
        "protocol": "modbusTCP",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Janitza\JANITZA_GENERIC_TCP.csv"),
        "sheet": None,
        "output_name": "MET_Janitza_Generic_TCP.csv",
        "description": "Compteur Janitza Modbus TCP Générique",
        "difficulty": "Passerelle Ethernet avec adresses 19000+, vérification de l'intégrité sans duplication de mot d'adresse.",
    },
    # -------------------------------------------------------------------------
    # 8. Kaco
    # -------------------------------------------------------------------------
    {
        "id": "KAC_01",
        "category": "Inverter",
        "mfg": "Kaco",
        "model": "NX3_G2_G3_SunSpec",
        "protocol": "modbusRTU",
        "path": os.path.join(
            EQUIPEMENTIERS_DIR, r"Kaco\DOC from Kaco\SUNSPEC_NX3_G2_G3_1xx_7xx.xlsx"
        ),
        "sheet": None,
        "output_name": "INV_Kaco_NX3_G2_G3_SunSpec.csv",
        "description": "Tableur Excel SunSpec officiel Kaco new energy (NX3 G2/G3 Models 1xx to 7xx)",
        "difficulty": "Classeur Excel SunSpec dense avec onglets de modèles modulaires (101, 103, 120, 160).",
    },
    {
        "id": "KAC_02",
        "category": "Inverter",
        "mfg": "Kaco",
        "model": "NX1_NX3_ModbusRTU",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Kaco\Modbus\Kaco_NX1-NX3_ModbusRTU.csv"),
        "sheet": None,
        "output_name": "INV_Kaco_NX1_NX3_ModbusRTU.csv",
        "description": "Onduleurs Kaco blueplanet NX1 et NX3 Modbus RTU",
        "difficulty": "Normalisation des tags de puissance réactive et facteur de puissance.",
    },
    {
        "id": "KAC_03",
        "category": "Inverter",
        "mfg": "Kaco",
        "model": "KACO_NX1",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Kaco\Modbus\KACO_NX1.csv"),
        "sheet": None,
        "output_name": "INV_Kaco_NX1.csv",
        "description": "Onduleur monophasé Kaco NX1",
        "difficulty": "Télémesures monophasées (tension, courant, puissance active DC/AC).",
    },
    # -------------------------------------------------------------------------
    # 9. Lettel
    # -------------------------------------------------------------------------
    {
        "id": "LET_01",
        "category": "Meter",
        "mfg": "Lettel",
        "model": "MAP4-345M_Table",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Lettel\Sources\map4-345m_table_modbus_fr.xlsx"),
        "sheet": None,
        "output_name": "MET_Lettel_MAP4_345M.csv",
        "description": "Compteur d'énergie communicant Lettel MAP4-345M (Excel en français)",
        "difficulty": "Feuille avec en-têtes multilingues français ('Adresse', 'Grandeur', 'Format', 'Unité'), facteurs d'échelle 0.01 / 0.1.",
    },
    {
        "id": "LET_02",
        "category": "Meter",
        "mfg": "Lettel",
        "model": "MCX4-34V_Modbus_FR",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Lettel\Sources\mcx4-34v_table_modbus_fr.pdf"),
        "sheet": None,
        "output_name": "MET_Lettel_MCX4_34V.csv",
        "description": "Compteur Modbus Lettel MCX4-34V (Documentation PDF en français)",
        "difficulty": "PDF avec adresses Modbus hexadécimales scindées sur deux colonnes distinctes ('High Byte' et 'Low Byte'), en-têtes à répétitions dégénérées. Résolu par détection des colonnes split et recomposition en '0x{hi}{lo}'.",
    },
    {
        "id": "LET_03",
        "category": "Meter",
        "mfg": "Lettel",
        "model": "MAP4-34RJ",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Lettel\LETTEL_MAP4-34RJ.csv"),
        "sheet": None,
        "output_name": "MET_Lettel_MAP4_34RJ.csv",
        "description": "Compteur divisionnaire Lettel MAP4-34RJ",
        "difficulty": "Mapping de compteurs d'énergie active totale et d'énergie injectée.",
    },
    # -------------------------------------------------------------------------
    # 10. Siebert
    # -------------------------------------------------------------------------
    {
        "id": "SIE_01",
        "category": "Other / Display",
        "mfg": "Siebert",
        "model": "XC4XX_TCP",
        "protocol": "modbusTCP",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Siebert\AFFICHEUR_SIEBERT_XC4XX_TCP.csv"),
        "sheet": None,
        "output_name": "OTH_Siebert_XC4XX_TCP.csv",
        "description": "Afficheur dynamique industriel pour centrales solaires Siebert XC4xx TCP",
        "difficulty": "Afficheur grand format recevant des valeurs de puissance instantanée, énergie cumulée et CO2 économisé.",
    },
    {
        "id": "SIE_02",
        "category": "Other / Display",
        "mfg": "Siebert",
        "model": "XC4XX_RTU",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Siebert\AFFICHEUR_SIEBERT_XC4XX.csv"),
        "sheet": None,
        "output_name": "OTH_Siebert_XC4XX_RTU.csv",
        "description": "Afficheur Siebert XC4xx Série Modbus RTU",
        "difficulty": "Commandes d'affichage de texte et de variables numériques.",
    },
    {
        "id": "SIE_03",
        "category": "Other / Display",
        "mfg": "Siebert",
        "model": "XC4XX_Test",
        "protocol": "modbusRTU",
        "path": os.path.join(EQUIPEMENTIERS_DIR, r"Siebert\AFFICHEUR_SIEBERT_XC4XX_TEST.csv"),
        "sheet": None,
        "output_name": "OTH_Siebert_XC4XX_TEST.csv",
        "description": "Table de banc de test Siebert XC4xx",
        "difficulty": "Conformité stricte de formatage CSV WebdynSunPM.",
    },
]


def run_battery():
    extractor = Extractor()
    generator = Generator()
    tempfile.TemporaryDirectory()

    output_dir = os.path.join(WORKSPACE_DIR, "generated_definitions")
    os.makedirs(output_dir, exist_ok=True)

    summary_results = []
    print("\n" + "=" * 80)
    print("DEMARRAGE DU GRAND AUDIT : EXTRACTION & CONVERSION DE TOUS LES DOCUMENTS")
    print("=" * 80)

    start_global = time.time()
    for idx, item in enumerate(CATALOG, start=1):
        doc_id = item["id"]
        doc_path = item["path"]
        mfg = item["mfg"]
        model = item["model"]
        category = item["category"]
        protocol = item["protocol"]
        sheet = item.get("sheet")
        output_name = item["output_name"]
        final_csv_path = os.path.join(output_dir, output_name)

        print(f"\n[{idx:02d}/{len(CATALOG):02d}] Traitement de : {os.path.basename(doc_path)}")
        print(f"     Constructeur : {mfg} | Modèle : {model} | Catégorie : {category}")

        res: dict[str, Any] = {
            "id": doc_id,
            "source_path": doc_path,
            "source_filename": os.path.basename(doc_path),
            "file_ext": os.path.splitext(doc_path)[1].upper(),
            "mfg": mfg,
            "model": model,
            "category": category,
            "protocol": protocol,
            "sheet": sheet,
            "output_csv_filename": output_name,
            "output_csv_path": final_csv_path,
            "description": item["description"],
            "difficulty": item["difficulty"],
            "success": False,
            "register_count": 0,
            "holding_count": 0,
            "coils_count": 0,
            "error_count": 0,
            "warning_count": 0,
            "duration_ms": 0,
            "error_msg": None,
        }

        t0 = time.time()
        try:
            if not os.path.exists(doc_path):
                raise FileNotFoundError(f"Fichier introuvable: {doc_path}")

            ext = os.path.splitext(doc_path)[1].lower()
            if ext in (".xlsx", ".xlsm", ".xltx"):
                raw_data = extractor.extract_from_excel(doc_path, sheet_name=sheet)
            elif ext == ".pdf":
                raw_data = extractor.extract_from_pdf(doc_path)
            elif ext == ".csv":
                raw_data = extractor.extract_from_csv(doc_path)
            elif ext == ".xml":
                raw_data = extractor.extract_from_xml(doc_path)
            else:
                raise ValueError(f"Extension non supportée: {ext}")

            has_data, peeked_raw = peek_generator(raw_data)
            if not has_data:
                raise RuntimeError("Aucune table Modbus brute détectée.")

            mapped_gen = extractor.map_and_clean(peeked_raw, address_offset=0)
            has_mapped, peeked_mapped = peek_generator(mapped_gen)
            if not has_mapped:
                raise RuntimeError("Aucun registre n'a pu être cartographié.")

            clean_rows = list(peeked_mapped)

            # Deduplicate alternative representations on the same address
            seen_addrs = set()
            dedup_rows = []
            for r in clean_rows:
                k = (str(r.get("RegisterType", "Holding")), str(r.get("Address", "")).strip())
                if k not in seen_addrs:
                    seen_addrs.add(k)
                    dedup_rows.append(r)

            # Generate target Webdyn CSV definition
            cfg = GeneratorConfig(
                input_file=doc_path,
                output=final_csv_path,
                manufacturer=mfg,
                model=model,
                protocol=protocol,
                category=category.split("/")[0].strip(),
                address_offset=0,
                strict_validation=False,
            )

            gen_ok = run_generator(cfg, input_data=dedup_rows)
            if not gen_ok or not os.path.exists(final_csv_path):
                raise RuntimeError("Échec de la génération du fichier CSV.")

            # Strict validation using WebdynSunPM rules
            report = generator.validate_csv_detailed(
                final_csv_path, strict=True, strict_overlap=False
            )
            res["register_count"] = report.register_count
            res["error_count"] = report.stats.get("errors", 0)
            res["warning_count"] = report.stats.get("warnings", 0)

            # Count Modbus types
            types_dict = report.stats.get("types", {})
            res["holding_count"] = types_dict.get("3", 0) + types_dict.get("4", 0)
            res["coils_count"] = types_dict.get("1", 0) + types_dict.get("2", 0)

            res["success"] = report.is_valid and report.register_count > 0
            if not report.is_valid:
                res["error_msg"] = "; ".join([iss.message for iss in report.issues[:3]])

            status_str = "✅ SUCCES (100% CONFORME)" if res["success"] else "❌ ECHEC"
            print(
                f"     --> {status_str} : {res['register_count']} registres extraits en {(time.time() - t0) * 1000:.1f} ms"
            )

        except Exception as exc:
            res["success"] = False
            res["error_msg"] = str(exc)
            print(f"     --> ❌ ECHEC : {exc}")

        res["duration_ms"] = round((time.time() - t0) * 1000, 1)
        summary_results.append(res)

    total_time = time.time() - start_global
    print("\n" + "=" * 80)
    print(f"AUDIT TERMINE EN {total_time:.2f} s")
    print(
        f"Taux de succès : {sum(1 for r in summary_results if r['success'])} / {len(summary_results)} ({sum(1 for r in summary_results if r['success']) / len(summary_results) * 100:.1f}%)"
    )
    print("=" * 80)

    # Save structured audit results as JSON for report compilation
    audit_json_path = os.path.join(WORKSPACE_DIR, "scratch", "doc_audit_results.json")
    os.makedirs(os.path.dirname(audit_json_path), exist_ok=True)
    with open(audit_json_path, "w", encoding="utf-8") as f:
        json.dump(summary_results, f, indent=2, ensure_ascii=False)
    print(f"Résultats JSON enregistrés dans : {audit_json_path}")


if __name__ == "__main__":
    run_battery()
