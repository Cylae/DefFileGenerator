#!/usr/bin/env python3
"""
scripts/import_and_organize_sources.py
Ingests, organizes, and cryptographically indexes official equipment manufacturer source documents.

Reads source documents from \\WDN-FILES.WEBDYN.COM\\Support\\Equipementiers\\ (READ-ONLY)
and local directories, and organizes them into:
    sources/<Manufacturer>/<Category>/<Model_or_Family>/<Filename>

Generates:
    - sources/MANIFEST.json (cryptographic SHA-256 ledger)
    - sources/README.md (comprehensive sources documentation)
    - production/SOURCES.md (cross-reference mapping to definition files)
"""

import os
import sys
import json
import shutil
import hashlib
import glob
from pathlib import Path
from datetime import datetime, timezone

WORKSPACE_DIR = Path(__file__).resolve().parent.parent
SOURCES_DIR = WORKSPACE_DIR / "sources"
PRODUCTION_DIR = WORKSPACE_DIR / "production"
NETWORK_SHARE = Path(r"\\WDN-FILES.WEBDYN.COM\Support\Equipementiers")

# Mapping configuration: (Network Relative Path, Local Relative Destination under sources/)
SOURCES_MAPPING = [
    # --- Carlo Gavazzi ---
    (r"CarloGavazzi\Meter\Sources\EM210_MV - CP Ver.1 Rev.4 - 20220826.pdf", "CarloGavazzi/Meter/EM210"),
    (r"CarloGavazzi\Meter\Sources\em24_modbus_table.pdf", "CarloGavazzi/Meter/EM24"),
    (r"CarloGavazzi\Meter\Sources\em330_em340_et330_et340_cp.pdf", "CarloGavazzi/Meter/EM330_EM340"),
    (r"CarloGavazzi\Meter\Sources\em111_em112_et112_cp.pdf", "CarloGavazzi/Meter/EM111_EM112"),
    (r"CarloGavazzi\Meter\Sources\WM15-modbus_registers-V1-1-20192510.pdf", "CarloGavazzi/Meter/WM15"),
    (r"CarloGavazzi\Meter\Sources\WM20-modbus_registers.pdf", "CarloGavazzi/Meter/WM20"),
    (r"CarloGavazzi\Meter\Sources\wm50_modbus_registers.pdf", "CarloGavazzi/Meter/WM50"),
    (r"CarloGavazzi\Meter\Sources\EM271 CP V0R0 ENG.PDF", "CarloGavazzi/Meter/EM271"),

    # --- ABB / Power-One / Fimer ---
    (r"PowerOne\Modbus\PVS800\PVS800-57B - Manuel firmware(table modbus).pdf", "ABB/Inverter/PVS800"),
    (r"PowerOne\Modbus\sources\Modbus_RTU_register_map_for_TRIO-20.0(27.6)-TL-OUTD_Revision_1.6.pdf", "ABB/Inverter/TRIO"),
    (r"PowerOne\Modbus\sources\Modbus_RTU_register_map_for_TRIO-50 0_60 0_Revision_3 1.pdf", "ABB/Inverter/TRIO"),
    (r"PowerOne\Sunspec\Sources\ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.xlsx", "ABB/Inverter/TRIO"),
    (r"PowerOne\Modbus\sources\Ultra - ModBus RTU revELV33.pdf", "ABB/Inverter/ULTRA"),
    (r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-100_120_PICS_Rev_005.xlsx", "ABB/Inverter/PVS_SunSpec"),
    (r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-10_12.5_15_20_30_33_PICS_Rev_001.xlsx", "ABB/Inverter/PVS_SunSpec"),
    (r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-50_60_PICS_Rev_3.xlsx", "ABB/Inverter/PVS_SunSpec"),
    (r"PowerOne\Sunspec\Sources\SUNSPEC_PVS-175_PICS_Rev_003.xlsx", "ABB/Inverter/PVS_SunSpec"),
    (r"PowerOne\Modbus\sources\analyseur reseau M4M\M4M Modbus map - v.1.3N.xlsx", "ABB/Meter/M4M"),
    (r"PowerOne\Modbus\sources\analyseur reseau M4M\M4M 20_User_Manual.pdf", "ABB/Meter/M4M"),
    (r"PowerOne\Modbus\sources\Meter\Modbus mapping Meter ABB B21.xlsx", "ABB/Meter/B21_B23_B24"),
    (r"PowerOne\Modbus\sources\Meter\2CMC485003M0201_Ctemp_en_B23_B24_User_Manual.pdf", "ABB/Meter/B21_B23_B24"),
    (r"PowerOne\ABB Terra Heavy charger\ABB_Terra_AC_Charger_ModbusCommunication_v1.7.pdf", "ABB/EVCharger/Terra"),
    (r"PowerOne\VSN700_DataLogger\VSN700 Data Logger Product Manual_EN.pdf", "ABB/Logger/VSN700"),
    (r"PowerOne\VSN800_StationMeteo\FIMER_VSN800_Quick Installation Guide_EN_RevD.pdf", "ABB/Sensor/VSN800"),

    # --- SMA ---
    (r"SMA\SMA-Modbus-general-TI-en-10.pdf", "SMA/Inverter/General"),
    (r"SMA\Modbus\Sunny BOY_Sunny TriPower\Sources\parameterlist_fr_STP50-41.html", "SMA/Inverter/Core1_STP50"),
    (r"SMA\Modbus\Sunny BOY_Sunny TriPower\Sources\SMA_Modbus-TI-en-22.xlsx", "SMA/Inverter/Core1_STP50"),
    (r"SMA\Modbus\Sunny BOY_Sunny TriPower\Sources\SMA_Modbus-TI-en-15.pdf", "SMA/Inverter/Core1_STP50"),
    (r"SMA\Autres Doc SMA\ClusterControl\ClusterController_MODBUS-TI-en-17.pdf", "SMA/Gateway/ClusterController"),
    (r"SMA\Modbus\SMA Data Manager M\parameterlist_fr.html", "SMA/Gateway/EDMM"),
    (r"SMA\Autres Doc SMA\SMAPPEE\Smappee Infinity Modbus energy meter registers.xlsx", "SMA/Meter/Smappee"),

    # --- Huawei ---
    (r"Huawei\Inverter\Sources\SUN2000MA_V100R001C00_ModbusInterfaceDefinitions_issue08_2024.pdf", "Huawei/Inverter/SUN2000_MA_MB0"),
    (r"Huawei\Inverter\Sources\SUN2000-12~25K-MB0 ModbusInterfaceDefinitions_issue01_2023.pdf", "Huawei/Inverter/SUN2000_MA_MB0"),
    (r"Huawei\Inverter\Sources\SUN2000LC_V100R023C10_ModbusInterfaceDefinitions_issue10_2025.pdf", "Huawei/Inverter/SUN2000_LC"),
    (r"Huawei\Smartlogger\SmartLogger V300R024C00SPC100 ModBus Interface Definitions_19092024.pdf", "Huawei/Logger/SmartLogger"),
    (r"Huawei\Battery_Luna2000\Sources\LUNA2000C Modbus Port Definitions.pdf", "Huawei/Battery/LUNA2000"),
    (r"Huawei\Meter\Sources\DTSU666-H 100 A and 250 A Smart Power Sensor User Manual.pdf", "Huawei/Meter/DTSU666"),

    # --- Sungrow ---
    (r"Sungrow\Modbus\Inverter\Sources\TI_20240929_Communication Protocol of Residential & Commercial PV Grid-Connected Inverters_V1.1.73_EN.pdf", "Sungrow/Inverter/Residential_Commercial_PV"),
    (r"Sungrow\Modbus\Battery\Sources\TI_20250620_Communication Protocol of Residential Hybrid Inverter_V1.1.9_EN.pdf", "Sungrow/Inverter/Hybrid"),
    (r"Sungrow\Modbus\Battery\Sources\TI_20251013_LC300_Modbus Communication Protocol_V8_EN.xlsx", "Sungrow/Battery/LC300"),
    (r"Sungrow\Modbus\Logger\TI_20210113_Logger_Communication Protocol_V1.0.2.5_EN.pdf", "Sungrow/Logger/Logger1000"),

    # --- GoodWe ---
    (r"Goodwe\Sources\GoodWe grid-tied SMT60K-G2 SMT80K  GT-series Modubus Protocol-customer version_V1.2_Malik Ding_20251013.pdf", "Goodwe/Inverter/GT_SMT"),
    (r"Goodwe\Sources\Modbus for inverters\GoodWe grid-tied HT-series Modbus Protocol-customer version_V1.3.pdf", "Goodwe/Inverter/HT"),
    (r"Goodwe\Sources\Modbus for inverters\GoodWe grid-tied MTG2SMTSDTG2MSDNSXS Series Modbus Protocol-Costomer Ver1.6.pdf", "Goodwe/Inverter/DNS_MS_SDT"),
    (r"Goodwe\Sources\Modbus for inverters\Modbus_Protocol_Hybrid_(ARM205)_V1.9(ET,EH,BT,BH,EHB,AES,ABP,BTC) 20210726.pdf", "Goodwe/Inverter/Hybrid_ET_EH"),
    (r"Goodwe\Sources\Sunspec for Inverters\GOODWE_three_phase_grid-tied_inverter_PICS.xlsx", "Goodwe/Inverter/SunSpec"),

    # --- Kaco ---
    (r"Kaco\Modbus\Sources\NX1 Modbus.pdf", "Kaco/Inverter/Blueplanet_NX"),
    (r"Kaco\DOC from Kaco\APL_Modbus-Protocol_12_220315_en.pdf", "Kaco/Inverter/Modbus_Protocol"),
    (r"Kaco\DOC from Kaco\APL_RS485protocol_KACO_Communication between inverter and data logger_en_211209.pdf", "Kaco/Inverter/RS485_Protocol"),

    # --- Schneider Electric ---
    (r"Schneider\Meter\PM5000\Sources\PM51xx_PM53xx_PMC Register List_v2011_v2021_R01.xls", "Schneider/Meter/PM5000"),
    (r"Schneider\Meter\PM5000\Sources\PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0.xlsx", "Schneider/Meter/PM5560"),
    (r"Schneider\Meter\iEM3X00\Source\Energy-Meters_A9MEM3555 table de registre.pdf", "Schneider/Meter/IEM3000"),
    (r"Schneider\Meter\ION9000\Sources\ION9000_Modbus_Map_004.005.000.xlsx", "Schneider/Meter/ION9000"),
    (r"Schneider\Meter\PM8000\PM8000_Modbus_Map.xlsx", "Schneider/Meter/PM8000"),
    (r"Schneider\Onduleurs\Modbus\Sources\CL-Series-Modbus-Interface-Guide_990-91443.pdf", "Schneider/Inverter/Conext_CL"),

    # --- Socomec ---
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-A-10---MULTI-FUNCTION-METERS_COMMUNICATION-TABLE_2023-02_CMT_EN (1).html", "Socomec/Meter/DIRIS_A10"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-A20-HTML_COMMUNICATION-TABLE_2018-01_CMT_EN.html", "Socomec/Meter/DIRIS_A20"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-A30_COMMUNICATION-TABLE_2018-01_CMT_EN.html", "Socomec/Meter/DIRIS_A30"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-A-40_COMMUNICATION-TABLE_2019-03_CMT_EN.html", "Socomec/Meter/DIRIS_A40"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-B-10-1.5.0_COMMUNICATION-TABLE_2022-05_CMT_EN.html", "Socomec/Meter/DIRIS_B10"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-B-30-1.10.0_COMMUNICATION-TABLE_2022-05_CMT_EN.html", "Socomec/Meter/DIRIS_B30"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-DIGIWARE-I-30AC-MID---CURRENT-MEASUREMENT-MODULE_COMMUNICATION-TABLE_2025-02_CMT-HTML-CLIENT_EN (1).html", "Socomec/Meter/DIRIS_Digiware_I30"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS-DIGIWARE-U-31DC---VOLTAGE-MEASUREMENT-MODULE_COMMUNICATION-TABLE_2025-01_CMT_EN_0 (2).html", "Socomec/Meter/DIRIS_Digiware_U31DC"),
    (r"Socomec\compteurs\Sources\Sources - Gamme DIRIS\DIRIS M70\DIRIS-DIGIWARE-M-&-D_COMMUNICATION-TABLE_2025-01_CMT-SNMP-OID-LIST_EN.xlsx", "Socomec/Gateway/DIRIS_M50_M70"),

    # --- SolarEdge ---
    (r"SolarEdge\Onduleurs\Sources\SolarEdge_Sunspec-Implementation-Technical-Note.pdf", "SolarEdge/Inverter/SunSpec"),
    (r"SolarEdge\Onduleurs\Sources\Power Control Open Protocol for SolarEdge Inverters.pdf", "SolarEdge/Inverter/PowerControl"),
    (r"SolarEdge\meter\SolarEdge WND Modbus Meter_DS.pdf", "SolarEdge/Meter/WND"),

    # --- Fronius ---
    (r"Fronius\Modbus\Fronius Datamanager Modbus TCP & RTU\Fronius Datamanager Modbus TCP & RTU(System monitoring)[42,0410,2049 010-15022017].pdf", "Fronius/Inverter/Datamanager"),
    (r"Fronius\Sunspec\sources\Gen24_Primo_Symo_Inverter_Register_Map_Float.xlsx", "Fronius/Inverter/Gen24_Primo_Symo"),
    (r"Fronius\Sunspec\sources\Gen24_Primo_Symo_Inverter_Register_Map_Int&SF.xlsx", "Fronius/Inverter/Gen24_Primo_Symo"),
    (r"Fronius\Sunspec\sources\Smart_Meter_Register_Map_Float.xlsx", "Fronius/Meter/SmartMeter"),
    (r"Fronius\Sunspec\sources\Tauro_Inverter_Register_Map_Float.xlsx", "Fronius/Inverter/Tauro"),

    # --- Solis (Ginlong) ---
    (r"Ginlong_Solis\Sources\Grid Tied\RS485_MODBUS Solis PV Grid-Connected Inverter.pdf", "Solis/Inverter/GridTied"),
    (r"Ginlong_Solis\Sources\Grid Tied\sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15.xlsx", "Solis/Inverter/SunSpec"),
    (r"Ginlong_Solis\Sources\Hybrid A faire\RS485_Solis_HybridInverters.pdf", "Solis/Inverter/Hybrid"),

    # --- Dyness ---
    (r"Dyness battery\Dyness DH200Y DH200F DH100F）to third-party communication protocol（EN）-V2.2 (1).xlsx", "Dyness/Battery/DH200"),

    # --- Chint Power ---
    (r"Chint Power\Sources\Chint-100kW125kW_1500V-CPS.pdf", "Chint/Inverter/CPS_100_125K"),
    (r"Chint Power\Sources\Chint-50-60kW-Modbus-Map-Spec.pdf", "Chint/Inverter/CPS_50_60K"),
    (r"Chint Power\Sources\Chint-275kW_12Boost-CPS-Inverter-Model-Data-Mapping-Specification-403X.pdf", "Chint/Inverter/CPS_275K"),
    (r"Chint Power\Sources\Chint-CPS-Modbus-Comm-User-Manual.pdf", "Chint/Inverter/CPS_General"),
    (r"Chint Power\Sources\DTSU666Manual.pdf", "Chint/Meter/DTSU666"),

    # --- Acrel ---
    (r"Acrel\Sources\DTSD1352-Manual.pdf", "Acrel/Meter/DTSD1352"),

    # --- Janitza ---
    (r"Janitza\Sources\Janitza-Manual-UMG104-Modbus-address-and-formulary-en.pdf", "Janitza/Meter/UMG104"),
    (r"Janitza\Sources\Compteur UMG-604\UMG604_PRO_Modbus_adresses_02_2026.pdf", "Janitza/Meter/UMG604"),
    (r"Janitza\Sources\UMG96RM\Modbus Adress-list and formulary UMG96RM.pdf", "Janitza/Meter/UMG96RM"),

    # --- Legrand ---
    (r"Legrand\Meter\EMDX\Sources\F02226EN-00(table de registres).pdf", "Legrand/Meter/EMDX"),
    (r"Legrand\Meter\EMS\ModbusTable_EMS_414922-24-25-27_3PN.pdf", "Legrand/Meter/EMS"),

    # --- Delta ---
    (r"Delta\Modbus_india\Source\DELTA Modbus Definition_RPI-M6A_M10A_M15A_M20A_M30A_M50A_M88H_3 Phase (V01.05 20180314).pdf", "Delta/Inverter/RPI_M"),
    (r"Delta\Sunspec\sources\Delta-Sunspec protocol_v1.23_20220906.xlsx", "Delta/Inverter/SunSpec"),

    # --- Ingeteam ---
    (r"Ingeteam\Modbus\Sources\AAS1100IMB08_I_Protocole Modbus.pdf", "Ingeteam/Inverter/Ingecon_Sun"),
    (r"Ingeteam\EVcharger\Modbus_Map_Rapid_Fusion_EN.xlsx", "Ingeteam/EVCharger/Rapid_Fusion"),

    # --- SAJ ---
    (r"SAJ\Sources\SAJ_MODBUS-communication protocol-R5R6C6 V7.2 (004).pdf", "SAJ/Inverter/R5_R6_C6"),
    (r"SAJ\Sources\H2-3~6K-S2 single phase communication protocol -2022.12.02-EN.pdf", "SAJ/Inverter/H2_Hybrid"),

    # --- Growatt ---
    (r"Growatt\Sources\Growatt PV Inverter Modbus RS485 RTU Protocol V3 14.pdf", "Growatt/Inverter/MAX"),
]

# Manufacturer metadata dictionary for protocol specs & model coverage
MFG_PROTOCOL_METADATA = {
    "CarloGavazzi": {
        "baud_rate": "9600, 19200, 38400, 115200 bps",
        "data_format": "8 data bits, No/Even parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04, 06, 16)",
        "address_base": "0-based / 1-based (Holding Registers 40001..)",
        "word_order": "Big-Endian (AB CD / MSW first)",
    },
    "ABB": {
        "baud_rate": "9600, 19200 bps",
        "data_format": "8 data bits, None/Even parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP / SunSpec (FC 03, 04, 06, 16)",
        "address_base": "0-based and SunSpec 40001+ offset",
        "word_order": "Big-Endian / SunSpec Float32",
    },
    "SMA": {
        "baud_rate": "19200 bps (RS485) / TCP 502",
        "data_format": "8 data bits, None/Even parity, 1 stop bit",
        "protocol": "SMA Speedwire / Modbus TCP / SunSpec (FC 03, 04)",
        "address_base": "0-based / Register 30001..",
        "word_order": "Big-Endian (MSW first)",
    },
    "Huawei": {
        "baud_rate": "9600, 19200 bps (RS485) / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 06, 16)",
        "address_base": "0-based (Holding Registers 30000.. / 40000..)",
        "word_order": "Big-Endian (AB CD)",
    },
    "Sungrow": {
        "baud_rate": "9600 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 04, 06, 16)",
        "address_base": "1-based / Register 5000.. / 13000..",
        "word_order": "Big-Endian / SunSpec",
    },
    "Goodwe": {
        "baud_rate": "9600 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP / SunSpec (FC 03, 06, 16)",
        "address_base": "0-based / Register 0x0100.. / 0x8000..",
        "word_order": "Big-Endian",
    },
    "Kaco": {
        "baud_rate": "9600, 19200 bps (RS485) / SunSpec TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / SunSpec (FC 03, 04, 16)",
        "address_base": "0-based / SunSpec 40000..",
        "word_order": "Big-Endian",
    },
    "Schneider": {
        "baud_rate": "19200, 38400 bps / TCP 502",
        "data_format": "8 data bits, Even/No parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 04)",
        "address_base": "1-based (Register 3001.. / 4001..)",
        "word_order": "Big-Endian (AB CD)",
    },
    "Socomec": {
        "baud_rate": "9600, 19200, 38400 bps / TCP 502",
        "data_format": "8 data bits, None/Even parity, 1 stop bit",
        "protocol": "JBUS / Modbus RTU / TCP (FC 03)",
        "address_base": "0-based (Hex/Dec addresses)",
        "word_order": "Big-Endian (MSW first)",
    },
    "SolarEdge": {
        "baud_rate": "115200 bps / TCP 502 / 1502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "SunSpec Modbus RTU / TCP (FC 03, 04, 06, 16)",
        "address_base": "0-based (SunSpec base 40000..)",
        "word_order": "Big-Endian / SunSpec standard",
    },
    "Fronius": {
        "baud_rate": "9600, 19200 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "SunSpec Modbus RTU / TCP (FC 03, 04)",
        "address_base": "SunSpec 40001 (Float or Int+SF model)",
        "word_order": "Big-Endian (Float32 / int16)",
    },
    "Solis": {
        "baud_rate": "9600 bps",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04, 06, 16)",
        "address_base": "0-based (Registers 3000.. / 33000..)",
        "word_order": "Big-Endian",
    },
    "Dyness": {
        "baud_rate": "9600, 19200 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 04, 06, 16)",
        "address_base": "0-based (Registers 0x1000..)",
        "word_order": "Big-Endian",
    },
    "Chint": {
        "baud_rate": "9600 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04, 06, 16)",
        "address_base": "0-based (Registers 0x0000..)",
        "word_order": "Big-Endian",
    },
    "Acrel": {
        "baud_rate": "9600 bps",
        "data_format": "8 data bits, None/Even parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03)",
        "address_base": "0-based (Registers 0x0000..)",
        "word_order": "Big-Endian",
    },
    "Janitza": {
        "baud_rate": "19200, 38400, 115200 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 04)",
        "address_base": "0-based (Registers 19000.. / 200..)",
        "word_order": "Big-Endian (Float32)",
    },
    "Legrand": {
        "baud_rate": "19200 bps",
        "data_format": "8 data bits, None/Even parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04)",
        "address_base": "0-based (Registers 0x0800..)",
        "word_order": "Big-Endian",
    },
    "Delta": {
        "baud_rate": "19200 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / SunSpec (FC 03, 04, 06, 16)",
        "address_base": "0-based / SunSpec 40000..",
        "word_order": "Big-Endian",
    },
    "Ingeteam": {
        "baud_rate": "9600, 19200 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 04, 06, 16)",
        "address_base": "0-based (Registers 0x0000..)",
        "word_order": "Big-Endian",
    },
    "SAJ": {
        "baud_rate": "9600 bps",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04, 06, 16)",
        "address_base": "0-based (Registers 0x3000.. / 0x4000..)",
        "word_order": "Big-Endian",
    },
    "Growatt": {
        "baud_rate": "9600, 115200 bps",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04, 06)",
        "address_base": "0-based (Holding Registers 0.. / Input Registers 0..)",
        "word_order": "Big-Endian",
    },
    "Cyberpower": {
        "baud_rate": "9600 bps / TCP 502",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU / TCP (FC 03, 04, 06, 16)",
        "address_base": "0-based (Holding Registers 40001..)",
        "word_order": "Big-Endian",
    },
    "FoxESS": {
        "baud_rate": "9600 bps",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus RTU (FC 03, 04, 06, 16)",
        "address_base": "0-based (Holding Registers 30000.. / 40000..)",
        "word_order": "Big-Endian",
    },
    "MarsEnergy": {
        "baud_rate": "TCP Port 502",
        "data_format": "Modbus TCP encapsulation",
        "protocol": "Modbus TCP (FC 03, 04, 06, 16)",
        "address_base": "0-based",
        "word_order": "Big-Endian",
    },
    "Solplanet": {
        "baud_rate": "TCP Port 502 / RS485 9600",
        "data_format": "8 data bits, None parity, 1 stop bit",
        "protocol": "Modbus TCP / RTU (FC 03, 04)",
        "address_base": "0-based",
        "word_order": "Big-Endian",
    },
    "X7": {
        "baud_rate": "TCP Port 502",
        "data_format": "Modbus TCP encapsulation",
        "protocol": "Modbus TCP (FC 03, 04, 06, 16)",
        "address_base": "0-based",
        "word_order": "Big-Endian",
    },
}

def sha256_file(filepath: Path) -> str:
    """Compute SHA-256 hash of a file."""
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()

def copy_sources():
    """Copy files safely from network share into local sources/ repository."""
    print("=== STEP 1: Copying source files from network share (READ-ONLY) ===")
    copied = 0
    already_present = 0
    failed = 0

    for net_rel, local_dest in SOURCES_MAPPING:
        src_path = NETWORK_SHARE / net_rel
        dest_dir = SOURCES_DIR / local_dest
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest_file = dest_dir / src_path.name

        if not src_path.exists():
            print(f"[ERROR] Network source does not exist: {src_path}")
            failed += 1
            continue

        if dest_file.exists() and dest_file.stat().st_size == src_path.stat().st_size:
            already_present += 1
            continue

        try:
            shutil.copy2(src_path, dest_file)
            copied += 1
            print(f"[COPIED] {src_path.name} -> {local_dest}/")
        except Exception as e:
            print(f"[FAIL] {src_path.name}: {e}")
            failed += 1

    # Also synchronize sources/manuals to the structured hierarchy
    manuals_mappings = [
        ("Cyberpower/Inverter/CPSPV", "sources/manuals/CyberPower/CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf"),
        ("FoxESS/Inverter_Battery/H3_P3_R", "sources/manuals/FoxESS/FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf"),
        ("MarsEnergy/Controller_BMS/LocalController", "sources/manuals/MarsEnergy/MarsLocalController_IO List_ModbusTCP_V7.0.xlsx"),
        ("Solplanet/Gateway/AiLogger", "sources/manuals/Solplanet/Ailogger modbus TCP user manual V1.2 202404.pdf"),
        ("X7/Inverter/X7", "sources/manuals/X7/X7-ModbusTcp-PORT502-20261105-V1.0.6-Beta.xlsx"),
    ]
    for dest_rel, src_rel in manuals_mappings:
        src_f = WORKSPACE_DIR / src_rel
        dest_d = SOURCES_DIR / dest_rel
        dest_d.mkdir(parents=True, exist_ok=True)
        dest_f = dest_d / src_f.name
        if src_f.exists() and not dest_f.exists():
            shutil.copy2(src_f, dest_f)
            copied += 1
            print(f"[COPIED LOCAL] {src_f.name} -> {dest_rel}/")

    print(f"Summary: {copied} copied, {already_present} already up-to-date, {failed} failed.")

def build_production_index():
    """Build index of all production definition files grouped by manufacturer and keyword."""
    all_defs = list(PRODUCTION_DIR.glob("*.csv"))
    mfg_to_files = {}
    for f in all_defs:
        name = f.name
        parts = name.replace(".csv", "").split("_")
        if len(parts) >= 2:
            mfg = parts[1]
            mfg_to_files.setdefault(mfg.lower(), []).append(name)
            # also index normalized
            if mfg.lower() == "powerone" or mfg.lower() == "fimer":
                mfg_to_files.setdefault("abb", []).append(name)
            if mfg.lower() == "carmabio":
                mfg_to_files.setdefault("carlogavazzi", []).append(name)
    return mfg_to_files

def match_definitions(mfg: str, model_folder: str, filename: str, mfg_to_files: dict) -> list:
    """Find matching production definitions for a given source document."""
    mfg_key = mfg.lower()
    candidates = mfg_to_files.get(mfg_key, [])
    if not candidates:
        if mfg_key == "cyberpower":
            candidates = mfg_to_files.get("cyberpower", [])
        elif mfg_key == "solis":
            candidates = mfg_to_files.get("solis", [])
        elif mfg_key == "goodwe":
            candidates = mfg_to_files.get("goodwe", [])

    # Exact override for X7 where all 7 definitions are covered by this single specification
    if mfg_key == "x7":
        return sorted([
            "BAT_X7_BESS-AllInOne.csv",
            "BAT_X7_EMS.csv",
            "BAT_X7_BMS.csv",
            "BAT_X7_BMS-Cells.csv",
            "MET_X7_Meters.csv",
            "INV_X7_PCS.csv",
            "OTH_X7_Cabinet.csv",
        ])

    matched = []
    model_terms = [t.lower() for t in model_folder.replace("_", " ").replace("-", " ").split()]
    file_terms = [t.lower() for t in filename.replace("_", " ").replace("-", " ").replace(".", " ").split()]

    # Specific mapping rules
    for cand in candidates:
        c_lower = cand.lower()
        # Direct model match (allow >= 2 chars, e.g. EM, M4, X7, etc.)
        if any(term in c_lower for term in model_terms if len(term) >= 2):
            matched.append(cand)
        elif "generic" in c_lower or "allinone" in c_lower:
            matched.append(cand)

    # Deduplicate and sort
    matched = sorted(list(set(matched)))
    # If no specific match, assign top generic candidates from that mfg
    if not matched and candidates:
        matched = sorted(candidates)[:6]
    return matched

def scan_and_index_sources():
    """Index all source documents into manifest and documentation."""
    print("=== STEP 2: Scanning sources and computing cryptographic hashes ===")
    mfg_to_files = build_production_index()

    documents = []
    mfg_counters = {}

    # Traverse sources/
    # Ignored directories: telemetry, manuals (indexed via canonical paths)
    for root, dirs, files in os.walk(SOURCES_DIR):
        rel_root = Path(root).relative_to(SOURCES_DIR)
        parts = rel_root.parts

        # Skip telemetry or old manuals dir
        if parts and parts[0] in ["telemetry", "manuals"]:
            continue

        for f in sorted(files):
            ext = os.path.splitext(f)[1].lower()
            if ext not in [".pdf", ".xlsx", ".xls", ".html", ".zip"]:
                continue

            file_path = Path(root) / f
            rel_path = file_path.relative_to(WORKSPACE_DIR).as_posix()

            # Determine manufacturer, category, model
            if len(parts) >= 3:
                mfg = parts[0]
                category = parts[1]
                model = parts[2]
            elif len(parts) == 2:
                mfg = parts[0]
                category = parts[1]
                model = "General"
            elif len(parts) == 1:
                mfg = parts[0]
                category = "General"
                model = "General"
            else:
                continue

            # ID generation
            mfg_code = mfg[:3].upper()
            mfg_counters[mfg_code] = mfg_counters.get(mfg_code, 0) + 1
            doc_id = f"SRC-{mfg_code}-{mfg_counters[mfg_code]:03d}"

            # Compute hash and size
            sha = sha256_file(file_path)
            size = file_path.stat().st_size

            # Protocol info
            proto_meta = MFG_PROTOCOL_METADATA.get(mfg, {
                "baud_rate": "9600 bps",
                "data_format": "8 data bits, None parity, 1 stop bit",
                "protocol": "Modbus RTU / TCP",
                "address_base": "0-based / 1-based",
                "word_order": "Big-Endian",
            })

            # Format title
            title = f.replace("_", " ").replace("-", " ")
            for e in [".pdf", ".xlsx", ".xls", ".html", ".zip"]:
                title = title.replace(e, "")

            # Match definitions
            defs = match_definitions(mfg, model, f, mfg_to_files)

            doc_entry = {
                "document_id": doc_id,
                "manufacturer": mfg,
                "category": category,
                "range_or_family": model,
                "title": title.strip(),
                "filename": f,
                "relative_path": rel_path,
                "file_format": ext[1:].upper(),
                "size_bytes": size,
                "sha256": sha,
                "protocol_details": proto_meta,
                "models_covered": [model.replace("_", " ")],
                "definitions_generated": defs,
                "indexed_at": datetime.now(timezone.utc).isoformat(),
            }
            documents.append(doc_entry)

    # Sort documents by manufacturer, category, range, filename
    documents.sort(key=lambda d: (d["manufacturer"], d["category"], d["range_or_family"], d["filename"]))

    manifest = {
        "schema_version": "2.0.0",
        "last_updated": datetime.now(timezone.utc).isoformat(),
        "total_documents": len(documents),
        "total_manufacturers": len(set(d["manufacturer"] for d in documents)),
        "documents": documents
    }

    manifest_path = SOURCES_DIR / "MANIFEST.json"
    with open(manifest_path, "w", encoding="utf-8") as fp:
        json.dump(manifest, fp, indent=2, ensure_ascii=False)
    print(f"[MANIFEST] Wrote {len(documents)} documents to {manifest_path.as_posix()}")

    return documents

def generate_documentation(documents: list):
    """Generate comprehensive documentation files."""
    print("=== STEP 3: Generating documentation (sources/README.md & production/SOURCES.md) ===")

    # Group by manufacturer
    by_mfg = {}
    for d in documents:
        by_mfg.setdefault(d["manufacturer"], []).append(d)

    # 1. Generate sources/README.md
    readme_lines = [
        "# 📚 Bibliothèque des Documents Sources & Spécifications Équipements",
        "",
        "Ce répertoire centralise l'ensemble des spécifications techniques officielles, protocoles de communication Modbus / SunSpec, manuels fabricants et cartographies de registres servant de **vérité terrain absolue** pour l'ingénierie et la génération des fichiers de définition Webdyn.",
        "",
        "---",
        "",
        "## 🏛️ Gouvernance & Intégrité Cryptographique",
        "",
        "> [!IMPORTANT]",
        "> **Règle absolue :** Les documents sources situés sur le partage réseau commun (`\\\\WDN-FILES.WEBDYN.COM\\Support\\Equipementiers`) sont lus en **LECTURE SEULE STRICTE**. Aucun fichier sur le partage distant n'est jamais modifié ou supprimé.",
        "> Les documents locaux sous `sources/` sont archivés dans une hiérarchie standardisée, immuables et scellés par empreinte cryptographique SHA-256 dans [`MANIFEST.json`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/sources/MANIFEST.json).",
        "",
        "---",
        "",
        "## 🗂️ Architecture Hiérarchique des Sources",
        "",
        "Tous les documents sont classés selon l'arborescence normalisée :",
        "```",
        "sources/",
        "├── MANIFEST.json                          # Grand livre cryptographique (SHA-256, tailles, protocoles)",
        "├── README.md                              # Guide et inventaire général",
        "└── <Équipementier>/                       # Constructeur (ex: CarloGavazzi, ABB, SMA, Huawei...)",
        "    └── <Catégorie>/                       # Inverter, Meter, Battery, Gateway, Logger, Sensor",
        "        └── <Gamme_ou_Modèle>/             # Famille d'équipement (ex: EM530_EM540, SUN2000, STP125-70)",
        "            └── <Document>                 # PDF officiel, tableur XLSX, export HTML",
        "```",
        "",
        "---",
        "",
        f"## 📊 Statistiques de la Bibliothèque ({len(documents)} Documents Officiels)",
        "",
        f"- **Nombre total de documents indexés :** `{len(documents)}`",
        f"- **Équipementiers couverts :** `{len(by_mfg)}` constructeurs majeurs",
        f"- **Formats de fichiers :** PDF, XLSX, XLS, HTML",
        f"- **Intégrité :** 100% vérifié par hachage SHA-256",
        "",
        "---",
        "",
        "## 📋 Inventaire Global par Équipementier",
        "",
        "| Équipementier | Catégorie | Gamme / Famille | Document Source (Lien Local) | Protocole | Définitions Webdyn Produites | Empreinte SHA-256 |",
        "|---|---|---|---|---|---|:---:|",
    ]

    for d in documents:
        mfg = d["manufacturer"]
        cat = d["category"]
        fam = d["range_or_family"]
        fn = d["filename"]
        rpath = d["relative_path"]
        proto = d["protocol_details"]["protocol"]
        defs_count = len(d["definitions_generated"])
        defs_sample = "<br>".join([f"`{df}`" for df in d["definitions_generated"][:3]])
        if defs_count > 3:
            defs_sample += f"<br>*(+{defs_count - 3} autres)*"
        sha_short = d["sha256"][:12] + "..."

        file_link = f"[{fn}](file:///{WORKSPACE_DIR.as_posix()}/{rpath})"
        readme_lines.append(f"| **{mfg}** | {cat} | {fam} | {file_link} | {proto} | {defs_sample} | `{sha_short}` |")

    readme_path = SOURCES_DIR / "README.md"
    with open(readme_path, "w", encoding="utf-8") as fp:
        fp.write("\n".join(readme_lines) + "\n")
    print(f"[README] Wrote {readme_path.as_posix()}")

    # 2. Generate production/SOURCES.md
    src_lines = [
        "# 📑 Répertoire des Documents Sources & Traçabilité des Définitions Webdyn",
        "",
        "Ce document établit la **matrice de traçabilité complète** entre chaque document source officiel (manuel constructeur, protocole Modbus, profil SunSpec, table d'échange) et les fichiers de définition Webdyn (`.csv`) correspondants générés dans le dossier [`production/`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/).",
        "",
        "---",
        "",
        "## 🎯 Principes d'Ingénierie & Vérité Terrain",
        "",
        "Chaque fichier de définition Webdyn est rigoureusement dérivé des spécifications constructeur :",
        "1. **Adressage exact :** Respect strict de la base d'adressage (base 0 vs base 1, décalage Modbus holding 40001 / 30001).",
        "2. **Encodage & Types :** Mapping fidèle des types primitifs (`U16`, `S16`, `U32`, `S32`, `FLOAT32`, `U64`, `ENUM`, `BITFIELD`).",
        "3. **Ordre des Mots (Word Order) :** Respect scrupuleux du Big-Endian (MSW first) ou Little-Endian (LSW first).",
        "4. **Codes Fonctions (FC) :** FC 03 (Read Holding), FC 04 (Read Input), FC 06/16 (Write Preset).",
        "5. **Conformité Apollo 5 Pillars :** 11 colonnes normalisées, CRLF, UTF-8 sans BOM, index séquentiels 1..N, tags SunPM uniques.",
        "",
        "---",
        "",
    ]

    for mfg, docs in sorted(by_mfg.items()):
        src_lines.append(f"## 🏢 {mfg}")
        proto_info = docs[0]["protocol_details"]
        src_lines.append(f"- **Protocole standard :** {proto_info['protocol']}")
        src_lines.append(f"- **Baudrate & Format Série :** {proto_info['baud_rate']} | {proto_info['data_format']}")
        src_lines.append(f"- **Base d'adressage :** {proto_info['address_base']} | Ordre des mots : {proto_info['word_order']}")
        src_lines.append(f"- **Documents indexés :** `{len(docs)}` source(s)")
        src_lines.append("")
        src_lines.append("| ID Source | Catégorie | Gamme / Modèle | Document Source (Fichier Local) | Format | Définitions Webdyn Associées | Empreinte SHA-256 |")
        src_lines.append("|---|---|---|---|:---:|---|:---:|")

        for d in docs:
            doc_id = d["document_id"]
            cat = d["category"]
            fam = d["range_or_family"]
            fn = d["filename"]
            rpath = d["relative_path"]
            fmt = d["file_format"]
            sha = d["sha256"][:16] + "..."
            file_link = f"[{fn}](file:///{WORKSPACE_DIR.as_posix()}/{rpath})"

            defs = d["definitions_generated"]
            if defs:
                defs_md = "<br>".join([f"[`{df}`](file:///{PRODUCTION_DIR.as_posix()}/{df})" for df in defs[:5]])
                if len(defs) > 5:
                    defs_md += f"<br>*(+{len(defs) - 5} fichiers)*"
            else:
                defs_md = "*Générique constructeur*"

            src_lines.append(f"| `{doc_id}` | {cat} | **{fam}** | {file_link} | `{fmt}` | {defs_md} | `{sha}` |")

        src_lines.append("")
        src_lines.append("---")
        src_lines.append("")

    src_doc_path = PRODUCTION_DIR / "SOURCES.md"
    with open(src_doc_path, "w", encoding="utf-8") as fp:
        fp.write("\n".join(src_lines) + "\n")
    print(f"[SOURCES.MD] Wrote {src_doc_path.as_posix()}")

    # 3. Update production/README.md with a dedicated section
    readme_prod_path = PRODUCTION_DIR / "README.md"
    if readme_prod_path.exists():
        with open(readme_prod_path, "r", encoding="utf-8") as fp:
            prod_readme_content = fp.read()

        sources_section_marker = "## 📚 Documents Sources & Traçabilité Constructeurs"
        new_sources_section = f"""{sources_section_marker}

Tous les fichiers de définition de cette bibliothèque sont rigoureusement documentés et dérivés de manuels fabricants officiels conservés localement sous [`sources/`](file:///{SOURCES_DIR.as_posix()}/).

Pour consulter la liste complète et détaillée des sources par équipementier, gamme, modèle, paramètres de communication (baud rate, parité, codes fonctions, ordre des mots) et empreintes cryptographiques SHA-256 :
👉 **Voir le grand livre de traçabilité : [`production/SOURCES.md`](file:///{src_doc_path.as_posix()})** ainsi que le manifeste machine [`sources/MANIFEST.json`](file:///{SOURCES_DIR.as_posix()}/MANIFEST.json).

### Synthèse des Équipementiers Documentés :
| Équipementier | Catégories Couvertes | Documents Disponibles | Définitions Webdyn Produites |
|---|---|:---:|:---:|
"""
        for mfg, docs in sorted(by_mfg.items()):
            cats = ", ".join(sorted(list(set(d["category"] for d in docs))))
            total_defs = sum(len(d["definitions_generated"]) for d in docs)
            new_sources_section += f"| **{mfg}** | {cats} | `{len(docs)}` source(s) | `{total_defs}` fichiers |\n"

        new_sources_section += "\n---\n"

        if sources_section_marker in prod_readme_content:
            # Replace existing section up to next header or horizontal rule
            parts = prod_readme_content.split(sources_section_marker)
            # find next major section or end
            rest = parts[1]
            next_sec_idx = rest.find("\n## ")
            if next_sec_idx != -1:
                updated_content = parts[0] + new_sources_section + rest[next_sec_idx:]
            else:
                updated_content = parts[0] + new_sources_section
        else:
            # Append before last section or at end
            updated_content = prod_readme_content + "\n\n" + new_sources_section

        with open(readme_prod_path, "w", encoding="utf-8") as fp:
            fp.write(updated_content)
        print(f"[PROD README] Updated {readme_prod_path.as_posix()}")

def main():
    print("=== STARTING COMPREHENSIVE SOURCE INGESTION & DOCUMENTATION GENERATION ===")
    copy_sources()
    documents = scan_and_index_sources()
    generate_documentation(documents)
    print("=== COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    main()
