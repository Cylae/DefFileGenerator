# 🏭 Webdyn Production Definitions Library

Bibliothèque officielle de définitions d'équipements industriels validées pour déploiement en production sur les concentrateurs et passerelles Webdyn (SunPM, Hub, Flex).

Comprend **163 fichiers de définition certifiés** couvrant **70,445 registres Modbus** et **6,152 tags SunPM canoniques** à travers plus de 65 constructeurs mondiaux.

---

## 📊 Synthèse par Catégorie

| Catégorie | Préfixe | Fichiers | Total Registres | Tags Actifs | Exemples Constructeurs |
|---|---|---|---|---|---|
| Onduleurs & Convertisseurs (Inverter) | `INV_` | 82 | 25,654 | 4,206 | ABB, AnswerDrives, Atess, CyberPower (+31) |
| Compteurs & Analyseurs Réseau (Meter) | `MET_` | 38 | 3,727 | 612 | ABB, AEberle, Algodue, CHAUVIN-ARNOUX (+20) |
| Systèmes BESS & Batteries (Battery) | `BAT_` | 13 | 40,206 | 1,172 | Dyness, FoxESS, Hoymiles, Huawei (+1) |
| Capteurs Météo & Pyranomètres (Sensor) | `SEN_` | 22 | 495 | 96 | ATERSA, BD, DELTA, DELTAOHM (+13) |
| Passerelles & Loggers (Logger) | `COM_` | 3 | 154 | 28 | GoodWe, Huawei, MarsEnergy |
| Modules IO & Afficheurs (Other) | `OTH_` | 5 | 209 | 38 | ICPDAS, SIEBERT, Webdyn |

---

## 🛡️ Les 5 Piliers Qualité Apollo-Standard

Chaque fichier de définition hébergé dans ce dossier respecte strictement les 5 critères non négociables :
1. **Nommage Lisible & Canonique** : `<PREFIX>_<Manufacturer>_<Model>.csv` sans scories de génération (`_2`, `_test`, `_part`, `_draft`).
2. **Adressage Modbus Strict** : Adresses décimales (0 à 65535) ou `address_length` pour les chaînes de caractères (ex. `30000_30`), zéro décalage, zéro chevauchement corrupteur.
3. **Code Action Webdyn Certifié** :
   - `Action = 4` : Télémesures régulières (périodique standard 5 min).
   - `Action = 9` : Événements, états critiques et alarmes.
   - `Action = 1` : Consignes et registres de pilotage d'écriture (Read/Write).
   - `Action = 0` : Registres non historisés / internes.
4. **Code Fonction Modbus Standard** : `1` (Coils), `2` (Discrete Inputs), `3` (Holding Registers), `4` (Input Registers).
5. **Tags SunPM Canoniques** : Télémesures normalisées (`P_AC`, `Q_AC`, `U_AC`, `I_AC`, `P_DC`, `E_TOT`, `E_DAY`, `FREQ`, `PF`, `TEMP`, `MODEL`, `SERIAL`, etc.) positionnées sur les mesures physiques effectives.

---

## 📑 Catalogue Détaillé de Production

| Fichier | Catégorie | Constructeur | Modèle | Protocole | Registres | Tags SunPM |
|---|---|---|---|---|---|---|
| [`BAT_Dyness_DH100F.csv`](./BAT_Dyness_DH100F.csv) | Battery | Dyness | DH100F | Modbus | 913 | 51 |
| [`BAT_Dyness_DH200F.csv`](./BAT_Dyness_DH200F.csv) | Battery | Dyness | DH200F | Modbus | 913 | 51 |
| [`BAT_Dyness_DH200Y.csv`](./BAT_Dyness_DH200Y.csv) | Battery | Dyness | DH200Y | Modbus | 912 | 47 |
| [`BAT_Dyness_DH_Series_Generic.csv`](./BAT_Dyness_DH_Series_Generic.csv) | Battery | Dyness | DH_Series_Generic | Modbus | 962 | 51 |
| [`BAT_FoxESS_GMax_Storage.csv`](./BAT_FoxESS_GMax_Storage.csv) | Battery | FoxESS | GMax_Storage | modbusTCP | 894 | 894 |
| [`BAT_FoxESS_H3_Pro.csv`](./BAT_FoxESS_H3_Pro.csv) | Battery | FoxESS | H3_Pro | Modbus | 893 | 18 |
| [`BAT_FoxESS_P3_Pro.csv`](./BAT_FoxESS_P3_Pro.csv) | Battery | FoxESS | P3_Pro | Modbus | 893 | 18 |
| [`BAT_FoxESS_RSeries.csv`](./BAT_FoxESS_RSeries.csv) | Battery | FoxESS | RSeries | Modbus | 49 | 3 |
| [`BAT_Hoymiles_HoyUltra2.csv`](./BAT_Hoymiles_HoyUltra2.csv) | Battery | Hoymiles | HoyUltra2 | Modbus | 127 | 6 |
| [`BAT_Huawei_Luna2000.csv`](./BAT_Huawei_Luna2000.csv) | Battery | Huawei | Luna2000 | Modbus | 170 | 25 |
| [`BAT_MarsEnergy_LocalController_AllInOne.csv`](./BAT_MarsEnergy_LocalController_AllInOne.csv) | Battery | MarsEnergy | LocalController_AllInOne | modbusTCP | 16,746 | 4 |
| [`BAT_MarsEnergy_LocalController_V7.0.csv`](./BAT_MarsEnergy_LocalController_V7.0.csv) | Battery | MarsEnergy | LocalController_V7.0 | modbusTCP | 11,704 | 4 |
| [`BAT_MarsEnergy_RackCell_SlaveN.csv`](./BAT_MarsEnergy_RackCell_SlaveN.csv) | Battery | MarsEnergy | RackCell_SlaveN | modbusTCP | 5,030 | 0 |
| [`COM_GoodWe_Ezlogger3000.csv`](./COM_GoodWe_Ezlogger3000.csv) | Logger | GoodWe | Ezlogger3000 | Modbus | 60 | 10 |
| [`COM_Huawei_SL3000.csv`](./COM_Huawei_SL3000.csv) | Logger | Huawei | SL3000 | Modbus | 82 | 18 |
| [`COM_MarsEnergy_EMS_Slave247.csv`](./COM_MarsEnergy_EMS_Slave247.csv) | Logger | MarsEnergy | EMS_Slave247 | modbusTCP | 12 | 0 |
| [`INV_ABB_10.5.csv`](./INV_ABB_10.5.csv) | Inverter | ABB | 10.5 | Modbus | 19 | 8 |
| [`INV_ABB_PVS800-57B.csv`](./INV_ABB_PVS800-57B.csv) | Inverter | ABB | PVS800-57B | Modbus | 1,053 | 24 |
| [`INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv`](./INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv) | Inverter | ABB | TRIO_20.0-27.6_TL_OUTD | Modbus | 89 | 16 |
| [`INV_ABB_Ultra_Inverter_Module.csv`](./INV_ABB_Ultra_Inverter_Module.csv) | Inverter | ABB | Ultra_Inverter_Module | Modbus | 57 | 11 |
| [`INV_AnswerDrives_Solargate.csv`](./INV_AnswerDrives_Solargate.csv) | Inverter | AnswerDrives | Solargate | Modbus | 288 | 288 |
| [`INV_Atess_Generic.csv`](./INV_Atess_Generic.csv) | Inverter | Atess | Generic | Modbus | 243 | 29 |
| [`INV_CyberPower_1-Phase.csv`](./INV_CyberPower_1-Phase.csv) | Inverter | CyberPower | 1-Phase | Modbus | 85 | 10 |
| [`INV_CyberPower_3-Phases.csv`](./INV_CyberPower_3-Phases.csv) | Inverter | CyberPower | 3-Phases | Modbus | 71 | 13 |
| [`INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv`](./INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv) | Inverter | CyberPower | CPSHB6000ETL48-CPSHB6000ETLFR | Modbus | 103 | 20 |
| [`INV_CyberPower_GENERIC.csv`](./INV_CyberPower_GENERIC.csv) | Inverter | CyberPower | GENERIC | Modbus | 245 | 67 |
| [`INV_Deye_Generic.csv`](./INV_Deye_Generic.csv) | Inverter | Deye | Generic | Modbus | 327 | 23 |
| [`INV_EFACEC_Efasolar.csv`](./INV_EFACEC_Efasolar.csv) | Inverter | EFACEC | Efasolar | Modbus | 58 | 18 |
| [`INV_FoxESS_H3_Pro.csv`](./INV_FoxESS_H3_Pro.csv) | Inverter | FoxESS | H3_Pro | Modbus | 285 | 114 |
| [`INV_FoxESS_P3_H3_Pro_AllInOne.csv`](./INV_FoxESS_P3_H3_Pro_AllInOne.csv) | Inverter | FoxESS | P3_H3_Pro_AllInOne | Modbus | 1,233 | 127 |
| [`INV_FoxESS_P3_Pro.csv`](./INV_FoxESS_P3_Pro.csv) | Inverter | FoxESS | P3_Pro | Modbus | 285 | 114 |
| [`INV_FoxESS_RSeries.csv`](./INV_FoxESS_RSeries.csv) | Inverter | FoxESS | RSeries | Modbus | 448 | 61 |
| [`INV_FoxESS_SunSpec_RV_H3Plus.csv`](./INV_FoxESS_SunSpec_RV_H3Plus.csv) | Inverter | FoxESS | SunSpec_RV_H3Plus | Modbus | 87 | 11 |
| [`INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv`](./INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv) | Inverter | GoodWe | ET_EH_BT_BH_EHB_AES_ABP_BTC_Series | Modbus | 530 | 24 |
| [`INV_Goodwe_ES_EM_S-BP_BP_Series.csv`](./INV_Goodwe_ES_EM_S-BP_BP_Series.csv) | Inverter | Goodwe | ES_EM_S-BP_BP_Series | Modbus | 370 | 39 |
| [`INV_Goodwe_GT_Series.csv`](./INV_Goodwe_GT_Series.csv) | Inverter | Goodwe | GT_Series | Modbus | 305 | 86 |
| [`INV_Goodwe_HT1100.csv`](./INV_Goodwe_HT1100.csv) | Inverter | Goodwe | HT1100 | Modbus | 276 | 76 |
| [`INV_Goodwe_HT1500.csv`](./INV_Goodwe_HT1500.csv) | Inverter | Goodwe | HT1500 | Modbus | 287 | 78 |
| [`INV_Goodwe_HT_Series.csv`](./INV_Goodwe_HT_Series.csv) | Inverter | Goodwe | HT_Series | Modbus | 293 | 64 |
| [`INV_Goodwe_HT_others_Series.csv`](./INV_Goodwe_HT_others_Series.csv) | Inverter | Goodwe | HT_others_Series | Modbus | 290 | 77 |
| [`INV_Goodwe_MS_DNS_XS.csv`](./INV_Goodwe_MS_DNS_XS.csv) | Inverter | Goodwe | MS_DNS_XS | Modbus | 40 | 28 |
| [`INV_Goodwe_MT.csv`](./INV_Goodwe_MT.csv) | Inverter | Goodwe | MT | Modbus | 95 | 49 |
| [`INV_Goodwe_MT_series.csv`](./INV_Goodwe_MT_series.csv) | Inverter | Goodwe | MT_series | Modbus | 123 | 52 |
| [`INV_Goodwe_SDTG2.csv`](./INV_Goodwe_SDTG2.csv) | Inverter | Goodwe | SDTG2 | Modbus | 116 | 54 |
| [`INV_Goodwe_SMT.csv`](./INV_Goodwe_SMT.csv) | Inverter | Goodwe | SMT | Modbus | 115 | 54 |
| [`INV_Goodwe_SMT50K_SMT60K.csv`](./INV_Goodwe_SMT50K_SMT60K.csv) | Inverter | Goodwe | SMT50K_SMT60K | Modbus | 119 | 58 |
| [`INV_Growatt_Generic.csv`](./INV_Growatt_Generic.csv) | Inverter | Growatt | Generic | Modbus | 851 | 851 |
| [`INV_Growatt_MAX1500V_MAX-XLV_Series.csv`](./INV_Growatt_MAX1500V_MAX-XLV_Series.csv) | Inverter | Growatt | MAX1500V_MAX-XLV_Series | Modbus | 468 | 151 |
| [`INV_Growatt_MOD_TL3-XH_Series.csv`](./INV_Growatt_MOD_TL3-XH_Series.csv) | Inverter | Growatt | MOD_TL3-XH_Series | Modbus | 409 | 53 |
| [`INV_Growatt_TL-X_TL-XH_Series.csv`](./INV_Growatt_TL-X_TL-XH_Series.csv) | Inverter | Growatt | TL-X_TL-XH_Series | Modbus | 425 | 53 |
| [`INV_Growatt_TL3-X_Series.csv`](./INV_Growatt_TL3-X_Series.csv) | Inverter | Growatt | TL3-X_Series | Modbus | 371 | 95 |
| [`INV_Huawei_GENERIC.csv`](./INV_Huawei_GENERIC.csv) | Inverter | Huawei | GENERIC | Modbus | 256 | 80 |
| [`INV_Huawei_V2.csv`](./INV_Huawei_V2.csv) | Inverter | Huawei | V2 | Modbus | 214 | 48 |
| [`INV_Huawei_V3.csv`](./INV_Huawei_V3.csv) | Inverter | Huawei | V3 | Modbus | 182 | 80 |
| [`INV_Huawei_V4.csv`](./INV_Huawei_V4.csv) | Inverter | Huawei | V4 | Modbus | 392 | 86 |
| [`INV_INVT_Generic.csv`](./INV_INVT_Generic.csv) | Inverter | INVT | Generic | Modbus | 41 | 29 |
| [`INV_JEMA_Generic.csv`](./INV_JEMA_Generic.csv) | Inverter | JEMA | Generic | Modbus | 60 | 26 |
| [`INV_KEHUA_SPI_100K-250K-B.csv`](./INV_KEHUA_SPI_100K-250K-B.csv) | Inverter | KEHUA | SPI_100K-250K-B | Modbus | 238 | 117 |
| [`INV_KOSTAL_PIKO_CI.csv`](./INV_KOSTAL_PIKO_CI.csv) | Inverter | KOSTAL | PIKO_CI | Modbus | 78 | 30 |
| [`INV_KSTAR_Generic.csv`](./INV_KSTAR_Generic.csv) | Inverter | KSTAR | Generic | Modbus | 52 | 21 |
| [`INV_KSTAR_KIG1-250K.csv`](./INV_KSTAR_KIG1-250K.csv) | Inverter | KSTAR | KIG1-250K | Modbus | 254 | 56 |
| [`INV_Kaco_NX1-NX3_Series.csv`](./INV_Kaco_NX1-NX3_Series.csv) | Inverter | Kaco | NX1-NX3_Series | Modbus | 182 | 36 |
| [`INV_Kehua_SinglePhase.csv`](./INV_Kehua_SinglePhase.csv) | Inverter | Kehua | SinglePhase | Modbus | 159 | 28 |
| [`INV_Kehua_ThreePhase.csv`](./INV_Kehua_ThreePhase.csv) | Inverter | Kehua | ThreePhase | Modbus | 84 | 18 |
| [`INV_Kostal_Modbus.csv`](./INV_Kostal_Modbus.csv) | Inverter | Kostal | Modbus | Modbus | 94 | 16 |
| [`INV_MarsEnergy_AirCooling.csv`](./INV_MarsEnergy_AirCooling.csv) | Inverter | MarsEnergy | AirCooling | modbusTCP | 1,600 | 2 |
| [`INV_MarsEnergy_LC_AC.csv`](./INV_MarsEnergy_LC_AC.csv) | Inverter | MarsEnergy | LC_AC | modbusTCP | 3,400 | 3 |
| [`INV_MarsEnergy_LiquidCooling.csv`](./INV_MarsEnergy_LiquidCooling.csv) | Inverter | MarsEnergy | LiquidCooling | modbusTCP | 1,800 | 1 |
| [`INV_Refusol_Generic.csv`](./INV_Refusol_Generic.csv) | Inverter | Refusol | Generic | Modbus | 536 | 36 |
| [`INV_SAJ_R6.csv`](./INV_SAJ_R6.csv) | Inverter | SAJ | R6 | Modbus | 313 | 11 |
| [`INV_SAJ_Sununo_Plus_Suntrio.csv`](./INV_SAJ_Sununo_Plus_Suntrio.csv) | Inverter | SAJ | Sununo_Plus_Suntrio | Modbus | 201 | 9 |
| [`INV_SATCON_PowerGate.csv`](./INV_SATCON_PowerGate.csv) | Inverter | SATCON | PowerGate | Modbus | 96 | 12 |
| [`INV_SIEL_10_TL_SOLAR_INVERTER.csv`](./INV_SIEL_10_TL_SOLAR_INVERTER.csv) | Inverter | SIEL | 10_TL_SOLAR_INVERTER | Modbus | 17 | 1 |
| [`INV_SIEL_DSPX.csv`](./INV_SIEL_DSPX.csv) | Inverter | SIEL | DSPX | Modbus | 21 | 1 |
| [`INV_SIEL_FUEL_CELL.csv`](./INV_SIEL_FUEL_CELL.csv) | Inverter | SIEL | FUEL_CELL | Modbus | 22 | 1 |
| [`INV_SIEL_Monophase.csv`](./INV_SIEL_Monophase.csv) | Inverter | SIEL | Monophase | Modbus | 11 | 2 |
| [`INV_SIEL_RECTIFIERS.csv`](./INV_SIEL_RECTIFIERS.csv) | Inverter | SIEL | RECTIFIERS | Modbus | 19 | 1 |
| [`INV_SIEL_STS.csv`](./INV_SIEL_STS.csv) | Inverter | SIEL | STS | Modbus | 25 | 1 |
| [`INV_SIEL_Solar_Converters.csv`](./INV_SIEL_Solar_Converters.csv) | Inverter | SIEL | Solar_Converters | Modbus | 19 | 1 |
| [`INV_SIEL_UPS_UNITS.csv`](./INV_SIEL_UPS_UNITS.csv) | Inverter | SIEL | UPS_UNITS | Modbus | 22 | 2 |
| [`INV_SIEL_WINDWAVE.csv`](./INV_SIEL_WINDWAVE.csv) | Inverter | SIEL | WINDWAVE | Modbus | 21 | 1 |
| [`INV_SMA_Generic.csv`](./INV_SMA_Generic.csv) | Inverter | SMA | Generic | Modbus | 857 | 44 |
| [`INV_SMA_SC2xxx.csv`](./INV_SMA_SC2xxx.csv) | Inverter | SMA | SC2xxx | Modbus | 78 | 5 |
| [`INV_SMA_STP_X.csv`](./INV_SMA_STP_X.csv) | Inverter | SMA | STP_X | Modbus | 435 | 27 |
| [`INV_Schneider_CONEXT_CORE_XC.csv`](./INV_Schneider_CONEXT_CORE_XC.csv) | Inverter | Schneider | CONEXT_CORE_XC | Modbus | 207 | 42 |
| [`INV_Schneider_Conext_CL_30-33-50.csv`](./INV_Schneider_Conext_CL_30-33-50.csv) | Inverter | Schneider | Conext_CL_30-33-50 | Modbus | 61 | 44 |
| [`INV_Sigenergy_Generic.csv`](./INV_Sigenergy_Generic.csv) | Inverter | Sigenergy | Generic | Modbus | 194 | 61 |
| [`INV_SofarSolar_1-40KTL.csv`](./INV_SofarSolar_1-40KTL.csv) | Inverter | SofarSolar | 1-40KTL | Modbus | 115 | 31 |
| [`INV_SofarSolar_50-70KTL.csv`](./INV_SofarSolar_50-70KTL.csv) | Inverter | SofarSolar | 50-70KTL | Modbus | 152 | 41 |
| [`INV_SofarSolar_HYDG3.csv`](./INV_SofarSolar_HYDG3.csv) | Inverter | SofarSolar | HYDG3 | Modbus | 538 | 32 |
| [`INV_Solax_Generic.csv`](./INV_Solax_Generic.csv) | Inverter | Solax | Generic | Modbus | 214 | 9 |
| [`INV_Solax_X1-X3_Generic.csv`](./INV_Solax_X1-X3_Generic.csv) | Inverter | Solax | X1-X3_Generic | Modbus | 157 | 11 |
| [`INV_Solis_GENERIC.csv`](./INV_Solis_GENERIC.csv) | Inverter | Solis | GENERIC | Modbus | 454 | 23 |
| [`INV_Solplanet_Generic.csv`](./INV_Solplanet_Generic.csv) | Inverter | Solplanet | Generic | Modbus | 625 | 60 |
| [`INV_Sungrow_Generic.csv`](./INV_Sungrow_Generic.csv) | Inverter | Sungrow | Generic | Modbus | 164 | 80 |
| [`INV_TATA_Generic.csv`](./INV_TATA_Generic.csv) | Inverter | TATA | Generic | Modbus | 20 | 11 |
| [`INV_VICTRON_Generic.csv`](./INV_VICTRON_Generic.csv) | Inverter | VICTRON | Generic | Modbus | 12 | 6 |
| [`INV_Zucchetti_Azzuro-3PH.csv`](./INV_Zucchetti_Azzuro-3PH.csv) | Inverter | Zucchetti | Azzuro-3PH | Modbus | 503 | 27 |
| [`MET_ABB_B21-B23-B24.csv`](./MET_ABB_B21-B23-B24.csv) | Meter | ABB | B21-B23-B24 | Modbus | 137 | 29 |
| [`MET_ABB_M2M_Meter.csv`](./MET_ABB_M2M_Meter.csv) | Meter | ABB | M2M_Meter | Modbus | 126 | 18 |
| [`MET_AEberle_PQPlus_Generic.csv`](./MET_AEberle_PQPlus_Generic.csv) | Meter | AEberle | PQPlus_Generic | Modbus | 60 | 24 |
| [`MET_Algodue_UPM209-309.csv`](./MET_Algodue_UPM209-309.csv) | Meter | Algodue | UPM209-309 | Modbus | 98 | 28 |
| [`MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv`](./MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv) | Meter | CHAUVIN-ARNOUX | Ulys_TT-M | Modbus | 19 | 4 |
| [`MET_CHINT_DTSU666.csv`](./MET_CHINT_DTSU666.csv) | Meter | CHINT | DTSU666 | Modbus | 29 | 10 |
| [`MET_CIRCUTOR_CVME3MINI.csv`](./MET_CIRCUTOR_CVME3MINI.csv) | Meter | CIRCUTOR | CVME3MINI | Modbus | 31 | 15 |
| [`MET_Carlo_Gavazzi_EM210.csv`](./MET_Carlo_Gavazzi_EM210.csv) | Meter | Carlo | Gavazzi_EM210 | Modbus | 115 | 9 |
| [`MET_ECO_ADAPT_POWERELEC_6.csv`](./MET_ECO_ADAPT_POWERELEC_6.csv) | Meter | ECO | ADAPT_POWERELEC_6 | Modbus | 221 | 7 |
| [`MET_EDMI_Mk6.csv`](./MET_EDMI_Mk6.csv) | Meter | EDMI | Mk6 | Modbus | 59 | 6 |
| [`MET_EcsOptec_ECS_Generic.csv`](./MET_EcsOptec_ECS_Generic.csv) | Meter | EcsOptec | ECS_Generic | Modbus | 44 | 8 |
| [`MET_FoxESS_H3_Pro.csv`](./MET_FoxESS_H3_Pro.csv) | Meter | FoxESS | H3_Pro | Modbus | 61 | 27 |
| [`MET_HAGER_Generic.csv`](./MET_HAGER_Generic.csv) | Meter | HAGER | Generic | Modbus | 102 | 10 |
| [`MET_HUAWEI_Power_Meter.csv`](./MET_HUAWEI_Power_Meter.csv) | Meter | HUAWEI | Power_Meter | Modbus | 44 | 13 |
| [`MET_IME_Generic.csv`](./MET_IME_Generic.csv) | Meter | IME | Generic | Modbus | 85 | 16 |
| [`MET_Inepro_PRO380-MOD-CT.csv`](./MET_Inepro_PRO380-MOD-CT.csv) | Meter | Inepro | PRO380-MOD-CT | Modbus | 24 | 9 |
| [`MET_Janitza_Generic.csv`](./MET_Janitza_Generic.csv) | Meter | Janitza | Generic | Modbus | 61 | 30 |
| [`MET_Janitza_Generic_RTU.csv`](./MET_Janitza_Generic_RTU.csv) | Meter | Janitza | Generic_RTU | Modbus | 61 | 14 |
| [`MET_Legrand_CF_Generic.csv`](./MET_Legrand_CF_Generic.csv) | Meter | Legrand | CF_Generic | Modbus | 11 | 1 |
| [`MET_Legrand_LEG412052.csv`](./MET_Legrand_LEG412052.csv) | Meter | Legrand | LEG412052 | Modbus | 489 | 30 |
| [`MET_Lettel_MAP4-34RJ.csv`](./MET_Lettel_MAP4-34RJ.csv) | Meter | Lettel | MAP4-34RJ | Modbus | 405 | 17 |
| [`MET_Lettel_MCT4-45.csv`](./MET_Lettel_MCT4-45.csv) | Meter | Lettel | MCT4-45 | Modbus | 30 | 9 |
| [`MET_Lettel_MCX4-34.csv`](./MET_Lettel_MCX4-34.csv) | Meter | Lettel | MCX4-34 | Modbus | 139 | 32 |
| [`MET_Lettel_MCX4-34V.csv`](./MET_Lettel_MCX4-34V.csv) | Meter | Lettel | MCX4-34V | Modbus | 102 | 32 |
| [`MET_MarsEnergy_Meter.csv`](./MET_MarsEnergy_Meter.csv) | Meter | MarsEnergy | Meter | modbusTCP | 11 | 3 |
| [`MET_POLIER_MTR5LMOD.csv`](./MET_POLIER_MTR5LMOD.csv) | Meter | POLIER | MTR5LMOD | Modbus | 33 | 17 |
| [`MET_SOCOMEC_COUNTIS_E23.csv`](./MET_SOCOMEC_COUNTIS_E23.csv) | Meter | SOCOMEC | COUNTIS_E23 | Modbus | 59 | 18 |
| [`MET_SOCOMEC_DIRIS_B10.csv`](./MET_SOCOMEC_DIRIS_B10.csv) | Meter | SOCOMEC | DIRIS_B10 | Modbus | 23 | 7 |
| [`MET_SOCOMEC_Diris_A40.csv`](./MET_SOCOMEC_Diris_A40.csv) | Meter | SOCOMEC | Diris_A40 | Modbus | 35 | 17 |
| [`MET_SOCOMEC_E33_E43.csv`](./MET_SOCOMEC_E33_E43.csv) | Meter | SOCOMEC | E33_E43 | Modbus | 73 | 18 |
| [`MET_Schneider_EM1200-EM1220.csv`](./MET_Schneider_EM1200-EM1220.csv) | Meter | Schneider | EM1200-EM1220 | Modbus | 96 | 24 |
| [`MET_Schneider_Generic.csv`](./MET_Schneider_Generic.csv) | Meter | Schneider | Generic | Modbus | 517 | 16 |
| [`MET_Schneider_IEM_3xxx.csv`](./MET_Schneider_IEM_3xxx.csv) | Meter | Schneider | IEM_3xxx | Modbus | 55 | 20 |
| [`MET_Schneider_PM5xxx.csv`](./MET_Schneider_PM5xxx.csv) | Meter | Schneider | PM5xxx | Modbus | 35 | 18 |
| [`MET_Socomec_Diris_A10.csv`](./MET_Socomec_Diris_A10.csv) | Meter | Socomec | Diris_A10 | Modbus | 75 | 18 |
| [`MET_Socomec_Diris_B30.csv`](./MET_Socomec_Diris_B30.csv) | Meter | Socomec | Diris_B30 | Modbus | 20 | 7 |
| [`MET_Socomec_Diris_I30.csv`](./MET_Socomec_Diris_I30.csv) | Meter | Socomec | Diris_I30 | Modbus | 20 | 6 |
| [`MET_Weidmuller_610.csv`](./MET_Weidmuller_610.csv) | Meter | Weidmuller | 610 | Modbus | 122 | 25 |
| [`OTH_ICPDAS_M7028.csv`](./OTH_ICPDAS_M7028.csv) | Other | ICPDAS | M7028 | Modbus | 70 | 19 |
| [`OTH_SIEBERT_XC4XX.csv`](./OTH_SIEBERT_XC4XX.csv) | Other | SIEBERT | XC4XX | Modbus | 3 | 3 |
| [`OTH_Webdyn_Sens_IO_16DI.csv`](./OTH_Webdyn_Sens_IO_16DI.csv) | Other | Webdyn | Sens_IO_16DI | Modbus | 60 | 4 |
| [`OTH_Webdyn_Sens_IO_8AI.csv`](./OTH_Webdyn_Sens_IO_8AI.csv) | Other | Webdyn | Sens_IO_8AI | Modbus | 34 | 1 |
| [`OTH_Webdyn_Sens_IO_8DI-6DO.csv`](./OTH_Webdyn_Sens_IO_8DI-6DO.csv) | Other | Webdyn | Sens_IO_8DI-6DO | Modbus | 42 | 11 |
| [`SEN_ATERSA_DATASOLMET.csv`](./SEN_ATERSA_DATASOLMET.csv) | Sensor | ATERSA | DATASOLMET | Modbus | 5 | 1 |
| [`SEN_BD_SENSOR_FUEL_SENSOR.csv`](./SEN_BD_SENSOR_FUEL_SENSOR.csv) | Sensor | BD | SENSOR_FUEL_SENSOR | Modbus | 17 | 3 |
| [`SEN_DELTA_OHM_LP_PYRA_03.csv`](./SEN_DELTA_OHM_LP_PYRA_03.csv) | Sensor | DELTA | OHM_LP_PYRA_03 | Modbus | 6 | 0 |
| [`SEN_DELTAOHM_HD523D.csv`](./SEN_DELTAOHM_HD523D.csv) | Sensor | DELTAOHM | HD523D | Modbus | 10 | 5 |
| [`SEN_EKO_MS40.csv`](./SEN_EKO_MS40.csv) | Sensor | EKO | MS40 | Modbus | 28 | 2 |
| [`SEN_Hukseflux_SR05-D1A3.csv`](./SEN_Hukseflux_SR05-D1A3.csv) | Sensor | Hukseflux | SR05-D1A3 | Modbus | 28 | 4 |
| [`SEN_Hukseflux_SR30.csv`](./SEN_Hukseflux_SR30.csv) | Sensor | Hukseflux | SR30 | Modbus | 28 | 4 |
| [`SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv`](./SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv) | Sensor | IngenieurBuro | Si-RS485TC-2T-v-MB | Modbus | 7 | 3 |
| [`SEN_KIPPZONEN_RT1.csv`](./SEN_KIPPZONEN_RT1.csv) | Sensor | KIPPZONEN | RT1 | Modbus | 15 | 1 |
| [`SEN_KippZonen_DustIQ.csv`](./SEN_KippZonen_DustIQ.csv) | Sensor | KippZonen | DustIQ | Modbus | 19 | 2 |
| [`SEN_KippZonen_SMP12.csv`](./SEN_KippZonen_SMP12.csv) | Sensor | KippZonen | SMP12 | Modbus | 53 | 4 |
| [`SEN_KippZonen_Si-SMP-full.csv`](./SEN_KippZonen_Si-SMP-full.csv) | Sensor | KippZonen | Si-SMP-full | Modbus | 35 | 1 |
| [`SEN_LSI_LASTEM_DPA870_873.csv`](./SEN_LSI_LASTEM_DPA870_873.csv) | Sensor | LSI | LASTEM_DPA870_873 | Modbus | 30 | 1 |
| [`SEN_Lufft_WS10.csv`](./SEN_Lufft_WS10.csv) | Sensor | Lufft | WS10 | Modbus | 23 | 4 |
| [`SEN_MEIER-NT_ADL-SR.csv`](./SEN_MEIER-NT_ADL-SR.csv) | Sensor | MEIER-NT | ADL-SR | Modbus | 18 | 4 |
| [`SEN_MarsEnergy_PV.csv`](./SEN_MarsEnergy_PV.csv) | Sensor | MarsEnergy | PV | modbusTCP | 49 | 45 |
| [`SEN_S_S_REGELTECHNIK_HYDRASGARD.csv`](./SEN_S_S_REGELTECHNIK_HYDRASGARD.csv) | Sensor | S | S_REGELTECHNIK_HYDRASGARD | Modbus | 8 | 1 |
| [`SEN_Seven_Sensor_3S-IS.csv`](./SEN_Seven_Sensor_3S-IS.csv) | Sensor | Seven | Sensor_3S-IS | Modbus | 45 | 6 |
| [`SEN_TRACKSO_SOZ03.csv`](./SEN_TRACKSO_SOZ03.csv) | Sensor | TRACKSO | SOZ03 | Modbus | 2 | 0 |
| [`SEN_Webdyn_M200-03.csv`](./SEN_Webdyn_M200-03.csv) | Sensor | Webdyn | M200-03 | Modbus | 1 | 0 |
| [`SEN_Webdyn_RK100-01.csv`](./SEN_Webdyn_RK100-01.csv) | Sensor | Webdyn | RK100-01 | Modbus | 2 | 1 |
| [`SEN_Webdyn_Thermokon_AGS55.csv`](./SEN_Webdyn_Thermokon_AGS55.csv) | Sensor | Webdyn | Thermokon_AGS55 | Modbus | 66 | 4 |
