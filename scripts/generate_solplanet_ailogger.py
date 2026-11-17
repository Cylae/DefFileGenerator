#!/usr/bin/env python3
"""
Generator for Solplanet Ai-Logger Modbus TCP Communication Gateway.
Based on official document: 'Ailogger modbus TCP user manual V1.2 202404.pdf' (UM0052_ASW600-800S-S2_EN_V01_0424).
Pillars enforced:
  1. Standard naming: COM_Solplanet_Ai-Logger.csv
  2. Decimal addressing (0..65535 or STRING_len)
  3. Action codes (4=Telemetry, 1=RW/Command, 9=Alarm/Fault)
  4. Standard Function codes (4=Input Registers, 3=Holding Registers)
  5. Canonical SunPM tags without duplicates
  6. Strict UTF-8 without BOM
"""

from pathlib import Path

RAW_REGISTERS = [
    # --- 1. Weather Station Data (FC 4 - Read Input Registers) ---
    (4, "986", "U16", "Wind Speed Sensor 1", "WIND_SPEED", 1.0, 0.0, "m/s", 4),
    (4, "987", "U16", "Irradiance Sensor 1", "IRR", 0.1, 0.0, "W/m2", 4),
    (4, "988", "I16", "Cell Temperature Sensor 1", "T_MOD", 0.1, 0.0, "°C", 4),
    (4, "989", "I16", "External Temperature 1 Sensor 1", "T_AMB", 0.1, 0.0, "°C", 4),
    (4, "990", "I16", "External Temperature 2 Sensor 1", "", 0.1, 0.0, "°C", 4),
    (4, "991", "U16", "Relative Humidity Sensor 1", "HUMIDITY", 1.0, 0.0, "%", 4),
    (4, "992", "U16", "Wind Direction Sensor 1", "WIND_DIR", 1.0, 0.0, "°", 4),
    (4, "993", "U16", "Wind Speed Sensor 2", "", 1.0, 0.0, "m/s", 4),
    (4, "994", "U16", "Irradiance Sensor 2", "", 0.1, 0.0, "W/m2", 4),
    (4, "995", "I16", "Cell Temperature Sensor 2", "", 0.1, 0.0, "°C", 4),
    (4, "996", "I16", "External Temperature 1 Sensor 2", "", 0.1, 0.0, "°C", 4),
    (4, "997", "I16", "External Temperature 2 Sensor 2", "", 0.1, 0.0, "°C", 4),
    (4, "998", "U16", "Relative Humidity Sensor 2", "", 1.0, 0.0, "%", 4),
    (4, "999", "U16", "Wind Direction Sensor 2", "", 1.0, 0.0, "°", 4),

    # --- 2. Inverter Telemetry Mapping (FC 4 - Read Input Registers) ---
    (4, "1000", "U16", "Inverter Device Type (1=Single-phase, 3=Three-phase)", "", 1.0, 0.0, "", 4),
    (4, "1001", "U16", "Inverter Modbus RTU Address", "", 1.0, 0.0, "", 4),
    (4, "1002_16", "STRING", "Inverter Serial Number", "INV_SN", 1.0, 0.0, "", 4),
    (4, "1018_8", "STRING", "Inverter Machine Type", "", 1.0, 0.0, "", 4),
    (4, "1026", "U16", "Current Grid Code", "", 1.0, 0.0, "", 4),
    (4, "1027", "U32", "Rated Nominal Power", "P_NOM", 1.0, 0.0, "W", 4),
    (4, "1029_7", "STRING", "Master Software Version", "", 1.0, 0.0, "", 4),
    (4, "1036_7", "STRING", "Slave Software Version", "", 1.0, 0.0, "", 4),
    (4, "1043_7", "STRING", "Safety Version", "", 1.0, 0.0, "", 4),
    (4, "1050_6", "STRING", "Communication Protocol Version", "", 1.0, 0.0, "", 4),
    (4, "1056_8", "STRING", "Manufacturer Name", "", 1.0, 0.0, "", 4),
    (4, "1064_8", "STRING", "Brand Name", "", 1.0, 0.0, "", 4),
    (4, "1072", "U16", "Inverter Model Code", "", 1.0, 0.0, "", 4),
    (4, "1073", "U16", "MPPT Number", "", 1.0, 0.0, "", 4),
    (4, "1074", "U16", "String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1075_2", "STRING", "Hardware Version", "", 1.0, 0.0, "", 4),
    (4, "1077_3", "STRING", "Master CPU Sub-Version", "", 1.0, 0.0, "", 4),
    (4, "1096_7", "STRING", "Slave CPU Version", "", 1.0, 0.0, "", 4),
    (4, "1103_3", "STRING", "Slave CPU Sub-Version", "", 1.0, 0.0, "", 4),
    (4, "1116", "U16", "PV1-PV4 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1117", "U16", "PV5-PV8 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1118", "U16", "PV9-PV12 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1119", "U16", "PV13-PV16 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1120", "U16", "PV17-PV20 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1121", "U16", "PV21-PV24 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1122", "U16", "PV25-PV28 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1123", "U16", "PV29-PV32 String Current Number", "", 1.0, 0.0, "", 4),
    (4, "1300", "U16", "Grid Rated Voltage", "", 0.1, 0.0, "V", 4),
    (4, "1301", "U16", "Grid Rated Frequency", "", 0.01, 0.0, "Hz", 4),
    (4, "1302", "U32", "E-Today of Inverter", "E_DAY", 0.1, 0.0, "kWh", 4),
    (4, "1304", "U32", "E-Total of Inverter", "E_TOT", 0.1, 0.0, "kWh", 4),
    (4, "1306", "U32", "Hours Total Operating Time", "HOURS_TOT", 1.0, 0.0, "h", 4),
    (4, "1308", "U16", "Inverter Operating State", "STATUS", 1.0, 0.0, "", 9),
    (4, "1309", "U16", "Grid Connect Time", "", 1.0, 0.0, "s", 4),
    (4, "1310", "I16", "Internal Inverter Temperature", "T_INT", 0.1, 0.0, "°C", 4),
    (4, "1311", "I16", "Inverter U-Phase Temperature", "", 0.1, 0.0, "°C", 4),
    (4, "1312", "I16", "Inverter V-Phase Temperature", "", 0.1, 0.0, "°C", 4),
    (4, "1313", "I16", "Inverter W-Phase Temperature", "", 0.1, 0.0, "°C", 4),
    (4, "1314", "I16", "Boost Stage Temperature", "", 0.1, 0.0, "°C", 4),
    (4, "1315", "I16", "Bidirectional DC-DC Converter Temperature", "", 0.1, 0.0, "°C", 4),
    (4, "1316", "U16", "DC Bus Voltage", "", 0.1, 0.0, "V", 4),
    (4, "1318", "U16", "PV1 DC Voltage", "U_DC_1", 0.1, 0.0, "V", 4),
    (4, "1319", "U16", "PV1 DC Current", "I_DC_1", 0.01, 0.0, "A", 4),
    (4, "1320", "U16", "PV2 DC Voltage", "U_DC_2", 0.1, 0.0, "V", 4),
    (4, "1321", "U16", "PV2 DC Current", "I_DC_2", 0.01, 0.0, "A", 4),
    (4, "1322", "U16", "PV3 DC Voltage", "U_DC_3", 0.1, 0.0, "V", 4),
    (4, "1323", "U16", "PV3 DC Current", "I_DC_3", 0.01, 0.0, "A", 4),
    (4, "1324", "U16", "PV4 DC Voltage", "U_DC_4", 0.1, 0.0, "V", 4),
    (4, "1325", "U16", "PV4 DC Current", "I_DC_4", 0.01, 0.0, "A", 4),
    (4, "1326", "U16", "PV5 DC Voltage", "U_DC_5", 0.1, 0.0, "V", 4),
    (4, "1327", "U16", "PV5 DC Current", "I_DC_5", 0.01, 0.0, "A", 4),
    (4, "1328", "U16", "PV6 DC Voltage", "U_DC_6", 0.1, 0.0, "V", 4),
    (4, "1329", "U16", "PV6 DC Current", "I_DC_6", 0.01, 0.0, "A", 4),
    (4, "1330", "U16", "PV7 DC Voltage", "U_DC_7", 0.1, 0.0, "V", 4),
    (4, "1331", "U16", "PV7 DC Current", "I_DC_7", 0.01, 0.0, "A", 4),
    (4, "1332", "U16", "PV8 DC Voltage", "U_DC_8", 0.1, 0.0, "V", 4),
    (4, "1333", "U16", "PV8 DC Current", "I_DC_8", 0.01, 0.0, "A", 4),
    (4, "1334", "U16", "PV9 DC Voltage", "U_DC_9", 0.1, 0.0, "V", 4),
    (4, "1335", "U16", "PV9 DC Current", "I_DC_9", 0.01, 0.0, "A", 4),
    (4, "1336", "U16", "PV10 DC Voltage", "U_DC_10", 0.1, 0.0, "V", 4),
    (4, "1337", "U16", "PV10 DC Current", "I_DC_10", 0.01, 0.0, "A", 4),
    (4, "1338", "U16", "PV String 1 Current", "I_STR_1", 0.1, 0.0, "A", 4),
    (4, "1339", "U16", "PV String 2 Current", "I_STR_2", 0.1, 0.0, "A", 4),
    (4, "1340", "U16", "PV String 3 Current", "I_STR_3", 0.1, 0.0, "A", 4),
    (4, "1341", "U16", "PV String 4 Current", "I_STR_4", 0.1, 0.0, "A", 4),
    (4, "1342", "U16", "PV String 5 Current", "I_STR_5", 0.1, 0.0, "A", 4),
    (4, "1343", "U16", "PV String 6 Current", "I_STR_6", 0.1, 0.0, "A", 4),
    (4, "1344", "U16", "PV String 7 Current", "I_STR_7", 0.1, 0.0, "A", 4),
    (4, "1345", "U16", "PV String 8 Current", "I_STR_8", 0.1, 0.0, "A", 4),
    (4, "1346", "U16", "PV String 9 Current", "I_STR_9", 0.1, 0.0, "A", 4),
    (4, "1347", "U16", "PV String 10 Current", "I_STR_10", 0.1, 0.0, "A", 4),
    (4, "1348", "U16", "PV String 11 Current", "I_STR_11", 0.1, 0.0, "A", 4),
    (4, "1349", "U16", "PV String 12 Current", "I_STR_12", 0.1, 0.0, "A", 4),
    (4, "1350", "U16", "PV String 13 Current", "I_STR_13", 0.1, 0.0, "A", 4),
    (4, "1351", "U16", "PV String 14 Current", "I_STR_14", 0.1, 0.0, "A", 4),
    (4, "1352", "U16", "PV String 15 Current", "I_STR_15", 0.1, 0.0, "A", 4),
    (4, "1353", "U16", "PV String 16 Current", "I_STR_16", 0.1, 0.0, "A", 4),
    (4, "1354", "U16", "PV String 17 Current", "I_STR_17", 0.1, 0.0, "A", 4),
    (4, "1355", "U16", "PV String 18 Current", "I_STR_18", 0.1, 0.0, "A", 4),
    (4, "1356", "U16", "PV String 19 Current", "I_STR_19", 0.1, 0.0, "A", 4),
    (4, "1357", "U16", "PV String 20 Current", "I_STR_20", 0.1, 0.0, "A", 4),
    (4, "1358", "U16", "Grid L1 Phase Voltage", "U_AC_1", 0.1, 0.0, "V", 4),
    (4, "1359", "U16", "Grid L1 Phase Current", "I_AC_1", 0.1, 0.0, "A", 4),
    (4, "1360", "U16", "Grid L2 Phase Voltage", "U_AC_2", 0.1, 0.0, "V", 4),
    (4, "1361", "U16", "Grid L2 Phase Current", "I_AC_2", 0.1, 0.0, "A", 4),
    (4, "1362", "U16", "Grid L3 Phase Voltage", "U_AC_3", 0.1, 0.0, "V", 4),
    (4, "1363", "U16", "Grid L3 Phase Current", "I_AC_3", 0.1, 0.0, "A", 4),
    (4, "1364", "U16", "Grid Line Voltage RS", "U_AC_12", 0.1, 0.0, "V", 4),
    (4, "1365", "U16", "Grid Line Voltage RT", "U_AC_31", 0.1, 0.0, "V", 4),
    (4, "1366", "U16", "Grid Line Voltage ST", "U_AC_23", 0.1, 0.0, "V", 4),
    (4, "1367", "U16", "Grid Frequency", "FREQ", 0.01, 0.0, "Hz", 4),
    (4, "1368", "U32", "Inverter Apparent Power", "S_AC", 1.0, 0.0, "VA", 4),
    (4, "1370", "I32", "Inverter Active Power", "P_AC", 1.0, 0.0, "W", 4),
    (4, "1372", "I32", "Inverter Reactive Power", "Q_AC", 1.0, 0.0, "Var", 4),
    (4, "1374", "I16", "Inverter Power Factor", "PF", 0.01, 0.0, "", 4),
    (4, "1376", "U16", "Inverter Internal Fault State", "", 1.0, 0.0, "", 9),
    (4, "1377", "U16", "Inverter Error Message Code", "", 1.0, 0.0, "", 9),
    (4, "1378", "U16", "Inverter Warning Message Code", "", 1.0, 0.0, "", 9),
    (4, "1388", "U16", "IOS Measure Value", "", 1.0, 0.0, "", 4),

    # --- 3. Total Inverter Power Data (FC 4 - Read Input Registers) ---
    (4, "36100", "I32", "Total Inverters Active Power", "P_AC_TOTAL", 1.0, 0.0, "W", 4),
    (4, "36102", "I32", "Total Inverters Reactive Power", "Q_AC_TOTAL", 1.0, 0.0, "Var", 4),

    # --- 4. SDM630 Meter Data Mapping (FC 4 - Read Input Registers, F32 IEEE-754) ---
    (4, "36104", "F32", "Meter Phase 1 Line to Neutral Voltage", "U_AC_METER_1", 1.0, 0.0, "V", 4),
    (4, "36106", "F32", "Meter Phase 2 Line to Neutral Voltage", "U_AC_METER_2", 1.0, 0.0, "V", 4),
    (4, "36108", "F32", "Meter Phase 3 Line to Neutral Voltage", "U_AC_METER_3", 1.0, 0.0, "V", 4),
    (4, "36110", "F32", "Meter Phase 1 Current", "I_AC_METER_1", 1.0, 0.0, "A", 4),
    (4, "36112", "F32", "Meter Phase 2 Current", "I_AC_METER_2", 1.0, 0.0, "A", 4),
    (4, "36114", "F32", "Meter Phase 3 Current", "I_AC_METER_3", 1.0, 0.0, "A", 4),
    (4, "36116", "F32", "Meter Phase 1 Active Power", "P_AC_METER_1", 1.0, 0.0, "W", 4),
    (4, "36118", "F32", "Meter Phase 2 Active Power", "P_AC_METER_2", 1.0, 0.0, "W", 4),
    (4, "36120", "F32", "Meter Phase 3 Active Power", "P_AC_METER_3", 1.0, 0.0, "W", 4),
    (4, "36122", "F32", "Meter Phase 1 Apparent Power", "S_AC_METER_1", 1.0, 0.0, "VA", 4),
    (4, "36124", "F32", "Meter Phase 2 Apparent Power", "S_AC_METER_2", 1.0, 0.0, "VA", 4),
    (4, "36126", "F32", "Meter Phase 3 Apparent Power", "S_AC_METER_3", 1.0, 0.0, "VA", 4),
    (4, "36128", "F32", "Meter Phase 1 Reactive Power", "Q_AC_METER_1", 1.0, 0.0, "VAr", 4),
    (4, "36130", "F32", "Meter Phase 2 Reactive Power", "Q_AC_METER_2", 1.0, 0.0, "VAr", 4),
    (4, "36132", "F32", "Meter Phase 3 Reactive Power", "Q_AC_METER_3", 1.0, 0.0, "VAr", 4),
    (4, "36134", "F32", "Meter Phase 1 Power Factor", "PF_METER_1", 1.0, 0.0, "", 4),
    (4, "36136", "F32", "Meter Phase 2 Power Factor", "PF_METER_2", 1.0, 0.0, "", 4),
    (4, "36138", "F32", "Meter Phase 3 Power Factor", "PF_METER_3", 1.0, 0.0, "", 4),
    (4, "36140", "F32", "Meter Phase 1 Phase Angle", "", 1.0, 0.0, "°", 4),
    (4, "36142", "F32", "Meter Phase 2 Phase Angle", "", 1.0, 0.0, "°", 4),
    (4, "36144", "F32", "Meter Phase 3 Phase Angle", "", 1.0, 0.0, "°", 4),
    (4, "36146", "F32", "Meter Average Line to Neutral Voltage", "U_AC_AVG", 1.0, 0.0, "V", 4),
    (4, "36150", "F32", "Meter Average Line Current", "I_AC_AVG", 1.0, 0.0, "A", 4),
    (4, "36152", "F32", "Meter Sum of Line Currents", "I_AC_SUM", 1.0, 0.0, "A", 4),
    (4, "36156", "F32", "Meter Total Active Power", "P_AC_METER", 1.0, 0.0, "W", 4),
    (4, "36160", "F32", "Meter Total Apparent Power", "S_AC_METER", 1.0, 0.0, "VA", 4),
    (4, "36164", "F32", "Meter Total Reactive Power", "Q_AC_METER", 1.0, 0.0, "VAr", 4),
    (4, "36166", "F32", "Meter Total Power Factor", "PF_METER", 1.0, 0.0, "", 4),
    (4, "36170", "F32", "Meter Total Phase Angle", "", 1.0, 0.0, "°", 4),
    (4, "36174", "F32", "Meter Frequency of Supply Voltages", "FREQ_METER", 1.0, 0.0, "Hz", 4),
    (4, "36176", "F32", "Meter Import Active Energy Wh Since Reset", "E_IMP", 1.0, 0.0, "kWh", 4),
    (4, "36178", "F32", "Meter Export Active Energy Wh Since Reset", "E_EXP", 1.0, 0.0, "kWh", 4),
    (4, "36180", "F32", "Meter Import Reactive Energy VArh Since Reset", "", 1.0, 0.0, "kVArh", 4),
    (4, "36182", "F32", "Meter Export Reactive Energy VArh Since Reset", "", 1.0, 0.0, "kVArh", 4),
    (4, "36184", "F32", "Meter Apparent Energy VAh Since Reset", "", 1.0, 0.0, "kVAh", 4),
    (4, "36186", "F32", "Meter Line 1 to Line 2 Voltage", "U_AC_METER_12", 1.0, 0.0, "V", 4),
    (4, "36188", "F32", "Meter Line 2 to Line 3 Voltage", "U_AC_METER_23", 1.0, 0.0, "V", 4),
    (4, "36190", "F32", "Meter Line 3 to Line 1 Voltage", "U_AC_METER_31", 1.0, 0.0, "V", 4),
    (4, "36204", "F32", "Meter Average Line to Line Voltage", "U_LL_AVG", 1.0, 0.0, "V", 4),

    # --- 5. Inverter Remote Control (FC 3 - Holding Registers / FC 06/16 Write, Action 1) ---
    (3, "60000", "I16", "Active Power Control Command (-100 to 100% Pn)", "CMD_P_AC_SET_PERCENT", 0.01, 0.0, "%Pn", 1),
    (3, "60001", "I16", "AC Power Factor Control Command (-1.0 to 1.0)", "CMD_PF_SET", 0.0001, 0.0, "", 1),
    (3, "60002", "U16", "Inverter ON/OFF Control Command (1=Turn On, 2=Turn Off)", "CMD_ON_OFF", 1.0, 0.0, "", 1),
    (3, "60003", "U16", "Active Power Limitation Percentage (0 to 110% Pn)", "CMD_P_LIM_PERCENT", 0.01, 0.0, "%Pn", 1),
    (3, "60004", "I16", "Reactive Power Q Value Control Command (-65 to 65% Sn)", "CMD_Q_SET_PERCENT", 0.01, 0.0, "%Sn", 1),
]

def generate_csv_text() -> str:
    lines = ["Modbus;Logger;Solplanet;Ai-Logger;;;;;;;"]
    for idx, (fc, addr, dtype, desc, tag, factor, offset, unit, action) in enumerate(RAW_REGISTERS, 1):
        line = f"{idx};{fc};{addr};{dtype};;{desc};{tag};{factor:.6f};{offset:.6f};{unit};{action}"
        lines.append(line)
    return "\n".join(lines) + "\n"

def main():
    csv_text = generate_csv_text()
    targets = [
        Path("production/COM_Solplanet_Ai-Logger.csv"),
        Path("Definitions/COM_Solplanet_Ai-Logger.csv"),
        Path("tagged_definitions/COM_Solplanet_Ai-Logger.csv"),
        Path("generated_definitions/COM_Solplanet_Ai-Logger.csv"),
    ]
    for target in targets:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(csv_text, encoding="utf-8")
        print(f"Generated {target} with {len(RAW_REGISTERS)} registers.")

if __name__ == "__main__":
    main()
