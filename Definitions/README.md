# 📚 Webdyn Definitions Library

Bibliothèque de définitions d'équipements industriels au format WebdynSunPM. Comprend **125 fichiers de définition** validés couvrant **18 131 registres Modbus** à travers plus de 60 constructeurs mondiaux.

---

## 📊 Synthèse par Catégorie

| Catégorie | Préfixe | Fichiers | Total Registres | Constructeurs Notables |
|-----------|---------|----------|-----------------|------------------------|
| Onduleurs & Convertisseurs (Inverters) | `INV_` | 66 | 13 940 | ABB, CyberPower, EFACEC, GoodWe, Goodwe (+21) |
| Compteurs & Analyseurs Réseau (Meters) | `MET_` | 31 | 3 394 | ABB, Algodue, CHAUVIN-ARNOUX, CHINT, CIRCUTOR (+14) |
| Capteurs & Pyranomètres (Sensors) | `SEN_` | 21 | 446 | ATERSA, BD_SENSOR, DELTAOHM, DELTA_OHM, EKO (+11) |
| Passerelles & Loggers de Communication | `COM_` | 2 | 142 | GoodWe, Huawei |
| Autres Équipements (IO Modules & Afficheurs) | `OTH_` | 5 | 209 | ICPDAS, SIEBERT, Webdyn |

---

## 🏷️ Convention de Nommage

Tous les fichiers suivent la convention canonique Webdyn :
```text
<CATEGORY_PREFIX>_<Manufacturer>_<Model>.csv
```
- `INV_` : Onduleurs solaires, hybrides et industriels (*Inverters*)
- `MET_` : Compteurs d'énergie et analyseurs de réseau (*Meters*)
- `SEN_` : Capteurs météo, pyranomètres et sondes thermiques (*Sensors*)
- `COM_` : Loggers, concentrateurs et passerelles de données (*Loggers*)
- `OTH_` : Modules d'entrées/sorties (Sens'IO), afficheurs et bornes (*Other*)

---

## 📑 Catalogue Détaillé

### Onduleurs & Convertisseurs (Inverters) (66 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`INV_ABB_10.5.csv`](./INV_ABB_10.5.csv) | ABB | 10.5 | Modbus | 19 |
| [`INV_ABB_PVS800-57B.csv`](./INV_ABB_PVS800-57B.csv) | ABB | PVS800-57B | Modbus | 1053 |
| [`INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv`](./INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv) | ABB | TRIO_20.0-27.6_TL_OUTD | Modbus | 89 |
| [`INV_ABB_Ultra_Inverter_Module.csv`](./INV_ABB_Ultra_Inverter_Module.csv) | ABB | Ultra_Inverter_Module | Modbus | 57 |
| [`INV_CyberPower_1-Phase.csv`](./INV_CyberPower_1-Phase.csv) | CyberPower | 1-Phase | Modbus | 85 |
| [`INV_CyberPower_3-Phases.csv`](./INV_CyberPower_3-Phases.csv) | CyberPower | 3-Phases | Modbus | 71 |
| [`INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv`](./INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv) | CyberPower | CPSHB6000ETL48-CPSHB6000ETLFR | Modbus | 103 |
| [`INV_CyberPower_GENERIC.csv`](./INV_CyberPower_GENERIC.csv) | CyberPower | GENERIC | Modbus | 245 |
| [`INV_EFACEC_Efasolar.csv`](./INV_EFACEC_Efasolar.csv) | EFACEC | Efasolar | Modbus | 58 |
| [`INV_Goodwe_ES_EM_S-BP_BP_Series.csv`](./INV_Goodwe_ES_EM_S-BP_BP_Series.csv) | Goodwe | ES_EM_S-BP_BP_Series | Modbus | 370 |
| [`INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv`](./INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv) | GoodWe | ET_EH_BT_BH_EHB_AES_ABP_BTC_Series | Modbus | 530 |
| [`INV_Goodwe_GT_Series.csv`](./INV_Goodwe_GT_Series.csv) | Goodwe | GT_Series | Modbus | 305 |
| [`INV_Goodwe_HT1100.csv`](./INV_Goodwe_HT1100.csv) | Goodwe | HT1100 | Modbus | 276 |
| [`INV_Goodwe_HT1500.csv`](./INV_Goodwe_HT1500.csv) | Goodwe | HT1500 | Modbus | 287 |
| [`INV_Goodwe_HT_others_Series.csv`](./INV_Goodwe_HT_others_Series.csv) | Goodwe | HT_others_Series | Modbus | 290 |
| [`INV_Goodwe_HT_Series.csv`](./INV_Goodwe_HT_Series.csv) | Goodwe | HT_Series | Modbus | 293 |
| [`INV_Goodwe_MS_DNS_XS.csv`](./INV_Goodwe_MS_DNS_XS.csv) | Goodwe | MS_DNS_XS | Modbus | 40 |
| [`INV_Goodwe_MT.csv`](./INV_Goodwe_MT.csv) | Goodwe | MT | Modbus | 95 |
| [`INV_Goodwe_MT_series.csv`](./INV_Goodwe_MT_series.csv) | Goodwe | MT_series | Modbus | 123 |
| [`INV_Goodwe_SDTG2.csv`](./INV_Goodwe_SDTG2.csv) | Goodwe | SDTG2 | Modbus | 116 |
| [`INV_Goodwe_SMT.csv`](./INV_Goodwe_SMT.csv) | Goodwe | SMT | Modbus | 115 |
| [`INV_Goodwe_SMT50K_SMT60K.csv`](./INV_Goodwe_SMT50K_SMT60K.csv) | Goodwe | SMT50K_SMT60K | Modbus | 119 |
| [`INV_Growatt_Generic.csv`](./INV_Growatt_Generic.csv) | Growatt | Growatt_Generic | ﻿modbusRTU | 851 |
| [`INV_Growatt_MAX1500V_MAX-XLV_Series.csv`](./INV_Growatt_MAX1500V_MAX-XLV_Series.csv) | Growatt | MAX1500V_MAX-XLV_Series | Modbus | 468 |
| [`INV_Growatt_MOD_TL3-XH_Series.csv`](./INV_Growatt_MOD_TL3-XH_Series.csv) | Growatt | MOD_TL3-XH_Series | Modbus | 410 |
| [`INV_Growatt_TL-X_TL-XH_Series.csv`](./INV_Growatt_TL-X_TL-XH_Series.csv) | Growatt | TL-X_TL-XH_Series | Modbus | 426 |
| [`INV_Growatt_TL3-X_Series.csv`](./INV_Growatt_TL3-X_Series.csv) | Growatt | TL3-X_Series | Modbus | 371 |
| [`INV_Huawei_V2.csv`](./INV_Huawei_V2.csv) | Huawei | V2 | Modbus | 214 |
| [`INV_Huawei_V3.csv`](./INV_Huawei_V3.csv) | Huawei | V3 | Modbus | 182 |
| [`INV_Huawei_V4.csv`](./INV_Huawei_V4.csv) | Huawei | V4 | Modbus | 322 |
| [`INV_JEMA_Generic.csv`](./INV_JEMA_Generic.csv) | JEMA | Generic | Modbus | 60 |
| [`INV_Kaco_NX1-NX3_Series.csv`](./INV_Kaco_NX1-NX3_Series.csv) | Kaco | NX1-NX3_Series | Modbus | 182 |
| [`INV_Kehua_SinglePhase.csv`](./INV_Kehua_SinglePhase.csv) | Kehua | SinglePhase | Modbus | 159 |
| [`INV_KEHUA_SPI_100K-250K-B.csv`](./INV_KEHUA_SPI_100K-250K-B.csv) | KEHUA | SPI_100K-250K-B | Modbus | 238 |
| [`INV_Kehua_ThreePhase.csv`](./INV_Kehua_ThreePhase.csv) | Kehua | ThreePhase | Modbus | 84 |
| [`INV_Kostal_Modbus.csv`](./INV_Kostal_Modbus.csv) | Kostal | Modbus | Modbus | 94 |
| [`INV_KOSTAL_PIKO_CI.csv`](./INV_KOSTAL_PIKO_CI.csv) | KOSTAL | PIKO_CI | Modbus | 78 |
| [`INV_KSTAR_Generic.csv`](./INV_KSTAR_Generic.csv) | KSTAR | Generic | Modbus | 52 |
| [`INV_KSTAR_KIG1-250K.csv`](./INV_KSTAR_KIG1-250K.csv) | KSTAR | KIG1-250K | Modbus | 254 |
| [`INV_SAJ_R6.csv`](./INV_SAJ_R6.csv) | SAJ | R6 | Modbus | 313 |
| [`INV_SAJ_Sununo_Plus_Suntrio.csv`](./INV_SAJ_Sununo_Plus_Suntrio.csv) | SAJ | Sununo_Plus_Suntrio | Modbus | 201 |
| [`INV_SATCON_PowerGate.csv`](./INV_SATCON_PowerGate.csv) | SATCON | PowerGate | Modbus | 96 |
| [`INV_Schneider_Conext_CL_30-33-50.csv`](./INV_Schneider_Conext_CL_30-33-50.csv) | Schneider | Conext_CL_30-33-50 | Modbus | 61 |
| [`INV_Schneider_CONEXT_CORE_XC.csv`](./INV_Schneider_CONEXT_CORE_XC.csv) | Schneider | CONEXT_CORE_XC | Modbus | 207 |
| [`INV_SIEL_10_TL_SOLAR_INVERTER.csv`](./INV_SIEL_10_TL_SOLAR_INVERTER.csv) | SIEL | 10_TL_SOLAR_INVERTER | Modbus | 17 |
| [`INV_SIEL_DSPX.csv`](./INV_SIEL_DSPX.csv) | SIEL | DSPX | Modbus | 21 |
| [`INV_SIEL_FUEL_CELL.csv`](./INV_SIEL_FUEL_CELL.csv) | SIEL | FUEL_CELL | Modbus | 22 |
| [`INV_SIEL_Monophase.csv`](./INV_SIEL_Monophase.csv) | SIEL | Monophase | Modbus | 11 |
| [`INV_SIEL_RECTIFIERS.csv`](./INV_SIEL_RECTIFIERS.csv) | SIEL | RECTIFIERS | Modbus | 19 |
| [`INV_SIEL_Solar_Converters.csv`](./INV_SIEL_Solar_Converters.csv) | SIEL | Solar_Converters | Modbus | 19 |
| [`INV_SIEL_STS.csv`](./INV_SIEL_STS.csv) | SIEL | STS | Modbus | 25 |
| [`INV_SIEL_UPS_UNITS.csv`](./INV_SIEL_UPS_UNITS.csv) | SIEL | UPS_UNITS | Modbus | 22 |
| [`INV_SIEL_WINDWAVE.csv`](./INV_SIEL_WINDWAVE.csv) | SIEL | WINDWAVE | Modbus | 21 |
| [`INV_SMA_Generic.csv`](./INV_SMA_Generic.csv) | SMA | Generic | Modbus | 857 |
| [`INV_SMA_SC2xxx.csv`](./INV_SMA_SC2xxx.csv) | SMA | SC2xxx | Modbus | 78 |
| [`INV_SMA_STP_X.csv`](./INV_SMA_STP_X.csv) | SMA | STP_X | Modbus | 437 |
| [`INV_SofarSolar_1-40KTL.csv`](./INV_SofarSolar_1-40KTL.csv) | SofarSolar | 1-40KTL | Modbus | 115 |
| [`INV_SofarSolar_50-70KTL.csv`](./INV_SofarSolar_50-70KTL.csv) | SofarSolar | 50-70KTL | Modbus | 152 |
| [`INV_SofarSolar_HYDG3.csv`](./INV_SofarSolar_HYDG3.csv) | SofarSolar | HYDG3 | Modbus | 538 |
| [`INV_Solax_Generic.csv`](./INV_Solax_Generic.csv) | Solax | Generic | Modbus | 214 |
| [`INV_Solax_X1-X3_Generic.csv`](./INV_Solax_X1-X3_Generic.csv) | Solax | X1-X3_Generic | Modbus | 157 |
| [`INV_Solis_GENERIC.csv`](./INV_Solis_GENERIC.csv) | Solis | GENERIC | Modbus | 454 |
| [`INV_Sungrow_Generic.csv`](./INV_Sungrow_Generic.csv) | Sungrow | Generic | Modbus | 164 |
| [`INV_TATA_Generic.csv`](./INV_TATA_Generic.csv) | TATA | Generic | Modbus | 20 |
| [`INV_VICTRON_Generic.csv`](./INV_VICTRON_Generic.csv) | VICTRON | Generic | Modbus | 12 |
| [`INV_Zucchetti_Azzuro-3PH.csv`](./INV_Zucchetti_Azzuro-3PH.csv) | Zucchetti | Azzuro-3PH | Modbus | 503 |

### Compteurs & Analyseurs Réseau (Meters) (31 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`MET_ABB_B21-B23-B24.csv`](./MET_ABB_B21-B23-B24.csv) | ABB | B21-B23-B24 | Modbus | 137 |
| [`MET_ABB_M2M_Meter.csv`](./MET_ABB_M2M_Meter.csv) | ABB | M2M_Meter | Modbus | 126 |
| [`MET_Algodue_UPM209-309.csv`](./MET_Algodue_UPM209-309.csv) | Algodue | UPM209-309 | Modbus | 98 |
| [`MET_Carlo_Gavazzi_EM210.csv`](./MET_Carlo_Gavazzi_EM210.csv) | Carlo_Gavazzi | EM210 | Modbus | 115 |
| [`MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv`](./MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv) | CHAUVIN-ARNOUX | Ulys_TT-M | Modbus | 19 |
| [`MET_CHINT_DTSU666.csv`](./MET_CHINT_DTSU666.csv) | CHINT | DTSU666 | Modbus | 29 |
| [`MET_CIRCUTOR_CVME3MINI.csv`](./MET_CIRCUTOR_CVME3MINI.csv) | CIRCUTOR | CVME3MINI | Modbus | 31 |
| [`MET_ECO_ADAPT_POWERELEC_6.csv`](./MET_ECO_ADAPT_POWERELEC_6.csv) | ECO_ADAPT | POWERELEC_6 | Modbus | 221 |
| [`MET_EDMI_Mk6.csv`](./MET_EDMI_Mk6.csv) | EDMI | Mk6 | Modbus | 59 |
| [`MET_HAGER_Generic.csv`](./MET_HAGER_Generic.csv) | HAGER | Generic | Modbus | 102 |
| [`MET_HUAWEI_Power_Meter.csv`](./MET_HUAWEI_Power_Meter.csv) | HUAWEI | Power_Meter | Modbus | 44 |
| [`MET_Inepro_PRO380-MOD-CT.csv`](./MET_Inepro_PRO380-MOD-CT.csv) | Inepro | PRO380-MOD-CT | Modbus | 24 |
| [`MET_Janitza_Generic.csv`](./MET_Janitza_Generic.csv) | Janitza | Generic | Modbus | 61 |
| [`MET_Legrand_LEG412052.csv`](./MET_Legrand_LEG412052.csv) | Legrand | LEG412052 | Modbus | 489 |
| [`MET_Lettel_MAP4-34RJ.csv`](./MET_Lettel_MAP4-34RJ.csv) | Lettel | MAP4-34RJ | Modbus | 405 |
| [`MET_Lettel_MCT4-45.csv`](./MET_Lettel_MCT4-45.csv) | Lettel | MCT4-45 | Modbus | 30 |
| [`MET_Lettel_MCX4-34.csv`](./MET_Lettel_MCX4-34.csv) | Lettel | MCX4-34 | Modbus | 139 |
| [`MET_Lettel_MCX4-34V.csv`](./MET_Lettel_MCX4-34V.csv) | Lettel | MCX4-34V | Modbus | 102 |
| [`MET_POLIER_MTR5LMOD.csv`](./MET_POLIER_MTR5LMOD.csv) | POLIER | MTR5LMOD | Modbus | 33 |
| [`MET_Schneider_EM1200-EM1220.csv`](./MET_Schneider_EM1200-EM1220.csv) | Schneider | EM1200-EM1220 | Modbus | 96 |
| [`MET_Schneider_Generic.csv`](./MET_Schneider_Generic.csv) | Schneider | Generic | Modbus | 517 |
| [`MET_Schneider_IEM_3xxx.csv`](./MET_Schneider_IEM_3xxx.csv) | Schneider | IEM_3xxx | Modbus | 55 |
| [`MET_Schneider_PM5xxx.csv`](./MET_Schneider_PM5xxx.csv) | Schneider | PM5xxx | Modbus | 35 |
| [`MET_SOCOMEC_COUNTIS_E23.csv`](./MET_SOCOMEC_COUNTIS_E23.csv) | SOCOMEC | COUNTIS_E23 | Modbus | 59 |
| [`MET_Socomec_Diris_A10.csv`](./MET_Socomec_Diris_A10.csv) | Socomec | Diris_A10 | Modbus | 75 |
| [`MET_SOCOMEC_Diris_A40.csv`](./MET_SOCOMEC_Diris_A40.csv) | SOCOMEC | Diris_A40 | Modbus | 35 |
| [`MET_SOCOMEC_DIRIS_B10.csv`](./MET_SOCOMEC_DIRIS_B10.csv) | SOCOMEC | DIRIS_B10 | Modbus | 23 |
| [`MET_Socomec_Diris_B30.csv`](./MET_Socomec_Diris_B30.csv) | Socomec | Diris_B30 | Modbus | 20 |
| [`MET_Socomec_Diris_I30.csv`](./MET_Socomec_Diris_I30.csv) | Socomec | Diris_I30 | Modbus | 20 |
| [`MET_SOCOMEC_E33_E43.csv`](./MET_SOCOMEC_E33_E43.csv) | SOCOMEC | E33_E43 | Modbus | 73 |
| [`MET_Weidmuller_610.csv`](./MET_Weidmuller_610.csv) | Weidmuller | 610 | Modbus | 122 |

### Capteurs & Pyranomètres (Sensors) (21 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`SEN_ATERSA_DATASOLMET.csv`](./SEN_ATERSA_DATASOLMET.csv) | ATERSA | DATASOLMET | Modbus | 5 |
| [`SEN_BD_SENSOR_FUEL_SENSOR.csv`](./SEN_BD_SENSOR_FUEL_SENSOR.csv) | BD_SENSOR | FUEL_SENSOR | Modbus | 17 |
| [`SEN_DELTA_OHM_LP_PYRA_03_.csv`](./SEN_DELTA_OHM_LP_PYRA_03_.csv) | DELTA_OHM | LP_PYRA_03_ | Modbus | 6 |
| [`SEN_DELTAOHM_HD523D.csv`](./SEN_DELTAOHM_HD523D.csv) | DELTAOHM | HD523D | Modbus | 10 |
| [`SEN_EKO_MS40.csv`](./SEN_EKO_MS40.csv) | EKO | MS40 | Modbus | 28 |
| [`SEN_Hukseflux_SR05-D1A3.csv`](./SEN_Hukseflux_SR05-D1A3.csv) | Hukseflux | SR05-D1A3 | Modbus | 28 |
| [`SEN_Hukseflux_SR30.csv`](./SEN_Hukseflux_SR30.csv) | Hukseflux | SR30 | Modbus | 28 |
| [`SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv`](./SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv) | IngenieurBuro | Si-RS485TC-2T-v-MB | Modbus | 7 |
| [`SEN_KippZonen_DustIQ.csv`](./SEN_KippZonen_DustIQ.csv) | KippZonen | DustIQ | Modbus | 19 |
| [`SEN_KIPPZONEN_RT1.csv`](./SEN_KIPPZONEN_RT1.csv) | KIPPZONEN | RT1 | Modbus | 15 |
| [`SEN_KippZonen_Si-SMP-full.csv`](./SEN_KippZonen_Si-SMP-full.csv) | KippZonen | Si-SMP-full | Modbus | 35 |
| [`SEN_KippZonen_SMP12.csv`](./SEN_KippZonen_SMP12.csv) | KippZonen | SMP12 | Modbus | 53 |
| [`SEN_LSI_LASTEM_DPA870_873.csv`](./SEN_LSI_LASTEM_DPA870_873.csv) | LSI_LASTEM | DPA870_873 | Modbus | 30 |
| [`SEN_Lufft_WS10.csv`](./SEN_Lufft_WS10.csv) | Lufft | WS10 | Modbus | 23 |
| [`SEN_MEIER-NT_ADL-SR.csv`](./SEN_MEIER-NT_ADL-SR.csv) | MEIER-NT | ADL-SR | Modbus | 18 |
| [`SEN_S_S_REGELTECHNIK_HYDRASGARD.csv`](./SEN_S_S_REGELTECHNIK_HYDRASGARD.csv) | S_S_REGELTECHNIK | HYDRASGARD | Modbus | 8 |
| [`SEN_Seven_Sensor_3S-IS.csv`](./SEN_Seven_Sensor_3S-IS.csv) | Seven_Sensor | 3S-IS | Modbus | 45 |
| [`SEN_TRACKSO_SOZ03.csv`](./SEN_TRACKSO_SOZ03.csv) | TRACKSO | SOZ03 | Modbus | 2 |
| [`SEN_Webdyn_M200-03.csv`](./SEN_Webdyn_M200-03.csv) | Webdyn | M200-03 | Modbus | 1 |
| [`SEN_Webdyn_RK100-01.csv`](./SEN_Webdyn_RK100-01.csv) | Webdyn | RK100-01 | Modbus | 2 |
| [`SEN_Webdyn_Thermokon_AGS55.csv`](./SEN_Webdyn_Thermokon_AGS55.csv) | Webdyn | Thermokon_AGS55 | Modbus | 66 |

### Passerelles & Loggers de Communication (2 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`COM_GoodWe_Ezlogger3000.csv`](./COM_GoodWe_Ezlogger3000.csv) | GoodWe | Ezlogger3000 | Modbus | 60 |
| [`COM_Huawei_SL3000.csv`](./COM_Huawei_SL3000.csv) | Huawei | SL3000 | Modbus | 82 |

### Autres Équipements (IO Modules & Afficheurs) (5 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`OTH_ICPDAS_M7028.csv`](./OTH_ICPDAS_M7028.csv) | ICPDAS | M7028 | Modbus | 70 |
| [`OTH_SIEBERT_XC4XX.csv`](./OTH_SIEBERT_XC4XX.csv) | SIEBERT | XC4XX | Modbus | 3 |
| [`OTH_Webdyn_Sens_IO_16DI.csv`](./OTH_Webdyn_Sens_IO_16DI.csv) | Webdyn | Sens_IO_16DI | Modbus | 60 |
| [`OTH_Webdyn_Sens_IO_8AI.csv`](./OTH_Webdyn_Sens_IO_8AI.csv) | Webdyn | Sens_IO_8AI | Modbus | 34 |
| [`OTH_Webdyn_Sens_IO_8DI-6DO.csv`](./OTH_Webdyn_Sens_IO_8DI-6DO.csv) | Webdyn | Sens_IO_8DI-6DO | Modbus | 42 |

