# 📚 Webdyn Definitions Library

Bibliothèque de définitions d'équipements industriels au format WebdynSunPM. Comprend **291 fichiers de définition** validés couvrant **61 002 registres Modbus** à travers plus de 60 constructeurs mondiaux.

---

## 📊 Synthèse par Catégorie

| Catégorie | Préfixe | Fichiers | Total Registres | Constructeurs Notables |
|-----------|---------|----------|-----------------|------------------------|
| Onduleurs & Convertisseurs (Inverters) | `INV_` | 165 | 42 385 | ABB, Answer Drives, CyberPower, Cyberpower, EFACEC (+25) |
| Compteurs & Analyseurs Réseau (Meters) | `MET_` | 73 | 8 121 | ABB, Algodue, CHAUVIN-ARNOUX, CHINT, CIRCUTOR (+20) |
| Systèmes de Stockage BESS (Batteries) | `BAT_` | 11 | 7 491 | Dyness, GoodWe, Huawei, Sungrow |
| Capteurs & Pyranomètres (Sensors) | `SEN_` | 22 | 500 | ATERSA, BD_SENSOR, DELTAOHM, DELTA_OHM, EKO (+12) |
| Passerelles & Loggers de Communication | `COM_` | 13 | 2 212 | GoodWe, Huawei, SMA, Solplanet, Sungrow |
| Autres Équipements (IO Modules & Afficheurs) | `OTH_` | 7 | 293 | ABB, ICPDAS, SIEBERT, Sungrow, Webdyn |

---

## 📁 Arborescence Propre & Structurée par Constructeur

Pour une lisibilité et une maintenabilité optimales, les fichiers de définition sont classés dans des sous-dossiers dédiés par grand équipementier tout en maintenant une racine plate pour la rétro-compatibilité :

| Dossier Dédié | Équipementier / Gamme | Fichiers Certifiés | Description |
|---------------|----------------------|--------------------|-------------|
| [`Definitions/SMA/`](./SMA/) | SMA Solar Technology | 27 | Gammes Sunny Boy, Sunny Tripower (Core1, Core2, STP-X, SE), Sunny Island, Sunny Central, Data Manager M |
| [`Definitions/GoodWe/`](./GoodWe/) | GoodWe Power Supply | 32 | Gammes résidentielles & tertiaires SMT, ET/EH, HT, MT, XS/DNS/MS, SDT, Lynx Home HV, EzLogger |
| [`Definitions/Huawei/`](./Huawei/) | Huawei FusionSolar | 65 | Gammes SUN2000 M0/M1/M2/M3/Utility, SmartLogger 1000/3000, Luna2000, Power Meters DTSU666-H |
| [`Definitions/Chint/`](./Chint/) | Chint Instrument | 5 | Compteurs communicants monophasés et triphasés DDSU666, DTSU666, DTSU666-H |
| [`Definitions/Solplanet/`](./Solplanet/) | Solplanet (AISWEI) | 13 | Gammes ASW monophasés, triphasés, hybrides et passerelles Ai-Logger |
| [`Definitions/Schneider/`](./Schneider/) | Schneider Electric | 79 | Gammes Conext CL, Core XC, compteurs réseau PM5000, iEM3000, EM1200 |
| [`Definitions/Socomec/`](./Socomec/) | Socomec Power | 36 | Analyseurs de réseau Diris A10/A20/A40/B10/B30/I30, compteurs d'énergie Countis E |
| [`Definitions/Sungrow/`](./Sungrow/) | Sungrow Power | 24 | Onduleurs SG tertiaires & utility, batteries BESS, coffrets string PV |
| [`Definitions/Growatt/`](./Growatt/) | Growatt New Energy | 16 | Onduleurs string MAX, MOD, MIN, TL-X, TL3-X, hybrides et stockage |
| [`Definitions/Fronius/`](./Fronius/) | Fronius International | 34 | Gammes Symo, Primo, Eco, Gen24 Plus, cartes d'entrées/capteurs IG |
| [`Definitions/SolarEdge/`](./SolarEdge/) | SolarEdge Technologies | 20 | Onduleurs string SE monophasés/triphasés, blocs de contrôle de puissance active/réactive |
| [`Definitions/CarloGavazzi/`](./CarloGavazzi/) | Carlo Gavazzi Automation | 27 | Compteurs d'énergie et analyseurs EM24, EM210, EM210_MV, EM530, EM540, EM100, EM300, WM |
| [`Definitions/ABB/`](./ABB/) | ABB / Fimer | 71 | Onduleurs PVS (10..175, 800), Trio, Pro33, Ultra, compteurs B21/23/24, M2M, M4M, Terra AC |
| [`Definitions/Dyness/`](./Dyness/) | Dyness Renewable | 4 | Systèmes de stockage batteries haute tension BESS DH100F, DH200F, DH200Y |
| [`Definitions/Classic/`](./Classic/) | Grands Classiques Multi-Constructeurs | 18 | Profils unifiés et standardisés des 18 équipements les plus déployés |
| [`Definitions/X7/`](./X7/) | X7 C&I BESS | 7 | Définitions ultra-détaillées du conteneur de stockage X7 100kW/215kWh |

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

### Onduleurs & Convertisseurs (Inverters) (165 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`INV_ABB_10.5.csv`](./INV_ABB_10.5.csv) | ABB | 10.5 | Modbus | 19 |
| [`INV_ABB_AURORA_CENTRAL.csv`](./INV_ABB_AURORA_CENTRAL.csv) | ABB | AURORA_Central | Modbus | 1522 |
| [`INV_ABB_PRO-33.0-TL.csv`](./INV_ABB_PRO-33.0-TL.csv) | ABB | PRO-33.0-TL | Modbus | 62 |
| [`INV_ABB_PVI-10.0_12.5.csv`](./INV_ABB_PVI-10.0_12.5.csv) | ABB | PVI-10.0_12.5 | Modbus | 19 |
| [`INV_ABB_PVI-110_275.csv`](./INV_ABB_PVI-110_275.csv) | ABB | PVI-110_275 | Modbus | 441 |
| [`INV_ABB_PVI-50_60.csv`](./INV_ABB_PVI-50_60.csv) | ABB | PVI-50_60 | Modbus | 104 |
| [`INV_ABB_PVS-100_120.csv`](./INV_ABB_PVS-100_120.csv) | ABB | PVS-100_120 | Modbus | 986 |
| [`INV_ABB_PVS-10_12.5_15_20_30_33.csv`](./INV_ABB_PVS-10_12.5_15_20_30_33.csv) | ABB | PVS-10_12.5_15_20_30_33 | Modbus | 963 |
| [`INV_ABB_PVS-175.csv`](./INV_ABB_PVS-175.csv) | ABB | PVS-175 | Modbus | 1068 |
| [`INV_ABB_PVS-30-TL.csv`](./INV_ABB_PVS-30-TL.csv) | ABB | PVS-30-TL | Modbus | 956 |
| [`INV_ABB_PVS-50_60.csv`](./INV_ABB_PVS-50_60.csv) | ABB | PVS-50_60 | Modbus | 956 |
| [`INV_ABB_PVS800-57B.csv`](./INV_ABB_PVS800-57B.csv) | ABB | PVS800-57B | Modbus | 1053 |
| [`INV_ABB_PVS800-57B_2.csv`](./INV_ABB_PVS800-57B_2.csv) | ABB | PVS800-57B | Modbus | 1053 |
| [`INV_ABB_TRIO-50.0-TM-OUTD.csv`](./INV_ABB_TRIO-50.0-TM-OUTD.csv) | ABB | TRIO-50.0-TM-OUTD | Modbus | 956 |
| [`INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv`](./INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv) | ABB | TRIO_20.0-27.6_TL_OUTD | Modbus | 89 |
| [`INV_ABB_Ultra_Inverter_Module.csv`](./INV_ABB_Ultra_Inverter_Module.csv) | ABB | ULTRA_Inverter_Module | Modbus | 57 |
| [`INV_AnswerDrives_Solargate.csv`](./INV_AnswerDrives_Solargate.csv) | Answer Drives | Solargate | Modbus | 288 |
| [`INV_CyberPower_1-Phase.csv`](./INV_CyberPower_1-Phase.csv) | CyberPower | 1-Phase | Modbus | 85 |
| [`INV_CyberPower_3-Phases.csv`](./INV_CyberPower_3-Phases.csv) | CyberPower | 3-Phases | Modbus | 71 |
| [`INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv`](./INV_CyberPower_CPSHB6000ETL48-CPSHB6000ETLFR.csv) | CyberPower | CPSHB6000ETL48-CPSHB6000ETLFR | Modbus | 103 |
| [`INV_Cyberpower_CPSPV120K-140K.csv`](./INV_Cyberpower_CPSPV120K-140K.csv) | Cyberpower | CPSPV120K-140K | Modbus | 244 |
| [`INV_CyberPower_GENERIC.csv`](./INV_CyberPower_GENERIC.csv) | Cyberpower | GENERIC | Modbus | 244 |
| [`INV_EFACEC_Efasolar.csv`](./INV_EFACEC_Efasolar.csv) | EFACEC | Efasolar | Modbus | 58 |
| [`INV_Fronius_ECO_Series.csv`](./INV_Fronius_ECO_Series.csv) | Fronius | ECO_Series | Modbus | 257 |
| [`INV_Fronius_Gen24_Primo_Symo_Float.csv`](./INV_Fronius_Gen24_Primo_Symo_Float.csv) | Fronius | Gen24_Primo_Symo_Float | Modbus | 262 |
| [`INV_Fronius_Gen24_Primo_Symo_IntSF.csv`](./INV_Fronius_Gen24_Primo_Symo_IntSF.csv) | Fronius | Gen24_Primo_Symo_IntSF | Modbus | 269 |
| [`INV_Fronius_Symo_10_20.csv`](./INV_Fronius_Symo_10_20.csv) | Fronius | Symo_10_20 | Modbus | 177 |
| [`INV_Fronius_Symo_GEN24_Plus.csv`](./INV_Fronius_Symo_GEN24_Plus.csv) | Fronius | Symo_GEN24_Plus | Modbus | 223 |
| [`INV_Fronius_Tauro_50-3-D.csv`](./INV_Fronius_Tauro_50-3-D.csv) | Fronius | Tauro_50-3-D | Modbus | 193 |
| [`INV_Fronius_Tauro_ECO_100-3-D.csv`](./INV_Fronius_Tauro_ECO_100-3-D.csv) | Fronius | Tauro_ECO_100-3-D | Modbus | 167 |
| [`INV_Fronius_Verto_3MPPT.csv`](./INV_Fronius_Verto_3MPPT.csv) | Fronius | Verto_3MPPT | Modbus | 279 |
| [`INV_Fronius_Verto_4MPPT.csv`](./INV_Fronius_Verto_4MPPT.csv) | Fronius | Verto_4MPPT | Modbus | 299 |
| [`INV_Goodwe_ES_EM_S-BP_BP_Series.csv`](./INV_Goodwe_ES_EM_S-BP_BP_Series.csv) | GoodWe | ES-EM-S-BP-BP-Series | Modbus | 370 |
| [`INV_Goodwe_ES-EM-SBP-BP.csv`](./INV_Goodwe_ES-EM-SBP-BP.csv) | GoodWe | ES-EM-SBP-BP | Modbus | 49 |
| [`INV_Goodwe_ES_EM_SBP.csv`](./INV_Goodwe_ES_EM_SBP.csv) | GoodWe | ES-EM-SBP-Series | Modbus | 49 |
| [`INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv`](./INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv) | GoodWe | ET-EH-BT-BH-EHB-Series | Modbus | 530 |
| [`INV_Goodwe_ET_EH_Hybrid_Series.csv`](./INV_Goodwe_ET_EH_Hybrid_Series.csv) | GoodWe | ET-EH-Hybrid-Series | Modbus | 79 |
| [`INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv`](./INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv) | GoodWe | ET-EH-ModbusRTU | Modbus | 79 |
| [`INV_Goodwe_ET_EH_Series.csv`](./INV_Goodwe_ET_EH_Series.csv) | GoodWe | ET-EH-Series | Modbus | 79 |
| [`INV_Goodwe_ET_Series.csv`](./INV_Goodwe_ET_Series.csv) | GoodWe | ET-Series | Modbus | 79 |
| [`INV_Goodwe_GT-SERIES.csv`](./INV_Goodwe_GT-SERIES.csv) | GoodWe | GT-SERIES | Modbus | 305 |
| [`INV_Goodwe_GT_Series.csv`](./INV_Goodwe_GT_Series.csv) | GoodWe | GT-Series | Modbus | 305 |
| [`INV_Goodwe_HT_1500V.csv`](./INV_Goodwe_HT_1500V.csv) | GoodWe | HT-1500V-Series | Modbus | 288 |
| [`INV_Goodwe_HT_others_Series.csv`](./INV_Goodwe_HT_others_Series.csv) | GoodWe | HT-Others-Series | Modbus | 91 |
| [`INV_Goodwe_HT_Series.csv`](./INV_Goodwe_HT_Series.csv) | GoodWe | HT-Series | Modbus | 294 |
| [`INV_Goodwe_HT1100.csv`](./INV_Goodwe_HT1100.csv) | GoodWe | HT1100 | Modbus | 277 |
| [`INV_Goodwe_HT1500.csv`](./INV_Goodwe_HT1500.csv) | GoodWe | HT1500 | Modbus | 288 |
| [`INV_Goodwe_MS_DNS_XS.csv`](./INV_Goodwe_MS_DNS_XS.csv) | GoodWe | MS-DNS-XS | Modbus | 50 |
| [`INV_Goodwe_MT.csv`](./INV_Goodwe_MT.csv) | GoodWe | MT | Modbus | 95 |
| [`INV_Goodwe_MT_series.csv`](./INV_Goodwe_MT_series.csv) | GoodWe | MT-Series | Modbus | 68 |
| [`INV_Goodwe_MT-Series_ModbusRTU.csv`](./INV_Goodwe_MT-Series_ModbusRTU.csv) | GoodWe | MT-Series-ModbusRTU | Modbus | 68 |
| [`INV_Goodwe_SDT_G2_G3.csv`](./INV_Goodwe_SDT_G2_G3.csv) | GoodWe | SDT-G2-G3-Series | Modbus | 60 |
| [`INV_Goodwe_SDTG2.csv`](./INV_Goodwe_SDTG2.csv) | GoodWe | SDTG2 | Modbus | 116 |
| [`INV_Goodwe_SMT.csv`](./INV_Goodwe_SMT.csv) | GoodWe | SMT | Modbus | 115 |
| [`INV_Goodwe_SMT_Series.csv`](./INV_Goodwe_SMT_Series.csv) | GoodWe | SMT-Series | Modbus | 68 |
| [`INV_Goodwe_SMT50K_SMT60K.csv`](./INV_Goodwe_SMT50K_SMT60K.csv) | GoodWe | SMT50K-SMT60K | Modbus | 119 |
| [`INV_Goodwe_ThreePhase_SunSpec.csv`](./INV_Goodwe_ThreePhase_SunSpec.csv) | GoodWe | ThreePhase-SunSpec | Modbus | 29 |
| [`INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv`](./INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv) | GoodWe | XS-DNS-MS-SDT-G3-Series | Modbus | 50 |
| [`INV_Goodwe_XS_DNS_MS.csv`](./INV_Goodwe_XS_DNS_MS.csv) | GoodWe | XS-DNS-MS-Series | Modbus | 50 |
| [`INV_Growatt_125kw.csv`](./INV_Growatt_125kw.csv) | Growatt | 125kw | Modbus | 156 |
| [`INV_Growatt_Generic.csv`](./INV_Growatt_Generic.csv) | Growatt | Growatt_Generic | Modbus | 851 |
| [`INV_Growatt_MAX_Series.csv`](./INV_Growatt_MAX_Series.csv) | Growatt | MAX-Series | Modbus | 29 |
| [`INV_Growatt_MAX1500V_MAX-XLV_Series.csv`](./INV_Growatt_MAX1500V_MAX-XLV_Series.csv) | Growatt | MAX1500V_MAX-XLV_Series | Modbus | 468 |
| [`INV_Growatt_MOD_TL3-XH_Series.csv`](./INV_Growatt_MOD_TL3-XH_Series.csv) | Growatt | MOD_TL3-XH_Series | Modbus | 409 |
| [`INV_Growatt_PM_5.csv`](./INV_Growatt_PM_5.csv) | Growatt | PM_5 | Modbus | 203 |
| [`INV_Growatt_PV_Inverter_V3.14.csv`](./INV_Growatt_PV_Inverter_V3.14.csv) | Growatt | PV_Inverter_V3.14 | Modbus | 136 |
| [`INV_Growatt_TL-X_TL-XH_Series.csv`](./INV_Growatt_TL-X_TL-XH_Series.csv) | Growatt | TL-X_TL-XH_Series | Modbus | 425 |
| [`INV_Growatt_TL3-X_Series.csv`](./INV_Growatt_TL3-X_Series.csv) | Growatt | TL3-X_Series | Modbus | 371 |
| [`INV_Huawei_GENERIC.csv`](./INV_Huawei_GENERIC.csv) | Huawei | GENERIC | Modbus | 256 |
| [`INV_Huawei_SUN2000-2_5KTL.csv`](./INV_Huawei_SUN2000-2_5KTL.csv) | Huawei | SUN2000-2_5KTL | Modbus | 47 |
| [`INV_Huawei_SUN2000-30_40KTL-M3.csv`](./INV_Huawei_SUN2000-30_40KTL-M3.csv) | Huawei | SUN2000-30_40KTL-M3 | Modbus | 64 |
| [`INV_Huawei_SUN2000-33A_36_42KTL.csv`](./INV_Huawei_SUN2000-33A_36_42KTL.csv) | Huawei | SUN2000-33A_36_42KTL | Modbus | 192 |
| [`INV_Huawei_SUN2000-50_70KTL-M0.csv`](./INV_Huawei_SUN2000-50_70KTL-M0.csv) | Huawei | SUN2000-50_70KTL-M0 | Modbus | 73 |
| [`INV_Huawei_SUN2000-60KTL-HV.csv`](./INV_Huawei_SUN2000-60KTL-HV.csv) | Huawei | SUN2000-60KTL-HV | Modbus | 347 |
| [`INV_Huawei_SUN2000-8_20KTL.csv`](./INV_Huawei_SUN2000-8_20KTL.csv) | Huawei | SUN2000-8_20KTL | Modbus | 328 |
| [`INV_Huawei_SUN2000-90_105KTL.csv`](./INV_Huawei_SUN2000-90_105KTL.csv) | Huawei | SUN2000-90_105KTL | Modbus | 73 |
| [`INV_Huawei_SUN2000-All_V3.csv`](./INV_Huawei_SUN2000-All_V3.csv) | Huawei | SUN2000-All_V3 | Modbus | 97 |
| [`INV_Huawei_SUN2000-V3.csv`](./INV_Huawei_SUN2000-V3.csv) | Huawei | SUN2000-V3 | Modbus | 97 |
| [`INV_Huawei_SUN2000-V4_Generic.csv`](./INV_Huawei_SUN2000-V4_Generic.csv) | Huawei | SUN2000-V4_Generic | Modbus | 322 |
| [`INV_Huawei_V4.csv`](./INV_Huawei_V4.csv) | Huawei | SUN2000_V4 | Modbus | 392 |
| [`INV_Huawei_V2.csv`](./INV_Huawei_V2.csv) | Huawei | V2 | Modbus | 214 |
| [`INV_Huawei_V3.csv`](./INV_Huawei_V3.csv) | Huawei | V3 | Modbus | 182 |
| [`INV_JEMA_Generic.csv`](./INV_JEMA_Generic.csv) | JEMA | Generic | Modbus | 60 |
| [`INV_Kaco_blueplanet_100_NX3_M8.csv`](./INV_Kaco_blueplanet_100_NX3_M8.csv) | Kaco | blueplanet_100_NX3_M8 | Modbus | 908 |
| [`INV_Kaco_blueplanet_125_NX3_M10.csv`](./INV_Kaco_blueplanet_125_NX3_M10.csv) | Kaco | blueplanet_125_NX3_M10 | Modbus | 928 |
| [`INV_Kaco_blueplanet_125_TL3.csv`](./INV_Kaco_blueplanet_125_TL3.csv) | Kaco | blueplanet_125_TL3 | Modbus | 828 |
| [`INV_Kaco_blueplanet_15.0_TL3.csv`](./INV_Kaco_blueplanet_15.0_TL3.csv) | Kaco | blueplanet_15.0_TL3 | Modbus | 638 |
| [`INV_Kaco_blueplanet_20.0_TL3.csv`](./INV_Kaco_blueplanet_20.0_TL3.csv) | Kaco | blueplanet_20.0_TL3 | Modbus | 697 |
| [`INV_Kaco_blueplanet_50.0_NX3_M5.csv`](./INV_Kaco_blueplanet_50.0_NX3_M5.csv) | Kaco | blueplanet_50.0_NX3_M5 | Modbus | 197 |
| [`INV_Kaco_blueplanet_50.0_TL3.csv`](./INV_Kaco_blueplanet_50.0_TL3.csv) | Kaco | blueplanet_50.0_TL3 | Modbus | 628 |
| [`INV_Kaco_blueplanet_NX1.csv`](./INV_Kaco_blueplanet_NX1.csv) | Kaco | blueplanet_NX1 | Modbus | 182 |
| [`INV_Kaco_blueplanet_NX3.csv`](./INV_Kaco_blueplanet_NX3.csv) | Kaco | blueplanet_NX3 | Modbus | 182 |
| [`INV_Kaco_NX1-NX3_Series.csv`](./INV_Kaco_NX1-NX3_Series.csv) | Kaco | NX1-NX3_Series | Modbus | 182 |
| [`INV_Kaco_Powador_60.0_TL3.csv`](./INV_Kaco_Powador_60.0_TL3.csv) | Kaco | Powador_60.0_TL3 | Modbus | 101 |
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
| [`INV_Schneider_Conext_CL.csv`](./INV_Schneider_Conext_CL.csv) | Schneider | Conext_CL | Modbus | 61 |
| [`INV_Schneider_Conext_CL_30-33-50.csv`](./INV_Schneider_Conext_CL_30-33-50.csv) | Schneider | Conext_CL_30-33-50 | Modbus | 61 |
| [`INV_Schneider_CONEXT_CORE_XC.csv`](./INV_Schneider_CONEXT_CORE_XC.csv) | Schneider | CONEXT_CORE_XC | Modbus | 210 |
| [`INV_Schneider_SunSpec.csv`](./INV_Schneider_SunSpec.csv) | Schneider | SunSpec | Modbus | 55 |
| [`INV_SIEL_10_TL_SOLAR_INVERTER.csv`](./INV_SIEL_10_TL_SOLAR_INVERTER.csv) | SIEL | 10_TL_SOLAR_INVERTER | Modbus | 17 |
| [`INV_SIEL_DSPX.csv`](./INV_SIEL_DSPX.csv) | SIEL | DSPX | Modbus | 21 |
| [`INV_SIEL_FUEL_CELL.csv`](./INV_SIEL_FUEL_CELL.csv) | SIEL | FUEL_CELL | Modbus | 22 |
| [`INV_SIEL_Monophase.csv`](./INV_SIEL_Monophase.csv) | SIEL | Monophase | Modbus | 11 |
| [`INV_SIEL_RECTIFIERS.csv`](./INV_SIEL_RECTIFIERS.csv) | SIEL | RECTIFIERS | Modbus | 19 |
| [`INV_SIEL_Solar_Converters.csv`](./INV_SIEL_Solar_Converters.csv) | SIEL | Solar_Converters | Modbus | 19 |
| [`INV_SIEL_STS.csv`](./INV_SIEL_STS.csv) | SIEL | STS | Modbus | 25 |
| [`INV_SIEL_UPS_UNITS.csv`](./INV_SIEL_UPS_UNITS.csv) | SIEL | UPS_UNITS | Modbus | 22 |
| [`INV_SIEL_WINDWAVE.csv`](./INV_SIEL_WINDWAVE.csv) | SIEL | WINDWAVE | Modbus | 21 |
| [`INV_SMA_Core1.csv`](./INV_SMA_Core1.csv) | SMA | Core1 | Modbus | 105 |
| [`INV_SMA_Core1_Core2.csv`](./INV_SMA_Core1_Core2.csv) | SMA | Core1_Core2 | Modbus | 77 |
| [`INV_SMA_Core2.csv`](./INV_SMA_Core2.csv) | SMA | Core2 | Modbus | 1211 |
| [`INV_SMA_Generic.csv`](./INV_SMA_Generic.csv) | SMA | Generic | Modbus | 857 |
| [`INV_SMA_Modbus.csv`](./INV_SMA_Modbus.csv) | SMA | Modbus | Modbus | 29 |
| [`INV_SMA_SC2xxx.csv`](./INV_SMA_SC2xxx.csv) | SMA | SC2xxx | Modbus | 78 |
| [`INV_SMA_SC_Webbox.csv`](./INV_SMA_SC_Webbox.csv) | SMA | SC_Webbox | Modbus | 350 |
| [`INV_SMA_STP_110-60.csv`](./INV_SMA_STP_110-60.csv) | SMA | STP_110-60 | Modbus | 1211 |
| [`INV_SMA_STP_125-70.csv`](./INV_SMA_STP_125-70.csv) | SMA | STP_125-70 | Modbus | 464 |
| [`INV_SMA_STP_5-10-SE40.csv`](./INV_SMA_STP_5-10-SE40.csv) | SMA | STP_5-10-SE40 | Modbus | 372 |
| [`INV_SMA_STP_X.csv`](./INV_SMA_STP_X.csv) | SMA | STP_X | Modbus | 435 |
| [`INV_SMA_STPCORE1_50-41.csv`](./INV_SMA_STPCORE1_50-41.csv) | SMA | STPCORE1_50-41 | Modbus | 105 |
| [`INV_SMA_SUNNY_CENTRAL.csv`](./INV_SMA_SUNNY_CENTRAL.csv) | SMA | SUNNY_CENTRAL | Modbus | 182 |
| [`INV_SMA_SunnyBoy.csv`](./INV_SMA_SunnyBoy.csv) | SMA | SunnyBoy | Modbus | 36 |
| [`INV_SMA_SunnyBoyStorage.csv`](./INV_SMA_SunnyBoyStorage.csv) | SMA | SunnyBoyStorage | Modbus | 45 |
| [`INV_SMA_SunnyCentral.csv`](./INV_SMA_SunnyCentral.csv) | SMA | SunnyCentral | Modbus | 182 |
| [`INV_SMA_SunnyCentral_Storage.csv`](./INV_SMA_SunnyCentral_Storage.csv) | SMA | SunnyCentral-Storage | Modbus | 193 |
| [`INV_SMA_SunnyIsland.csv`](./INV_SMA_SunnyIsland.csv) | SMA | SunnyIsland | Modbus | 35 |
| [`INV_SMA_SunnyTripower.csv`](./INV_SMA_SunnyTripower.csv) | SMA | SunnyTripower | Modbus | 59 |
| [`INV_SMA_SunnyTripower_SmartEnergy.csv`](./INV_SMA_SunnyTripower_SmartEnergy.csv) | SMA | SunnyTripower-SmartEnergy | Modbus | 372 |
| [`INV_SMA_SunnyTripower_X.csv`](./INV_SMA_SunnyTripower_X.csv) | SMA | SunnyTripower-X | Modbus | 435 |
| [`INV_SMA_Sunspec.csv`](./INV_SMA_Sunspec.csv) | SMA | Sunspec | Modbus | 29 |
| [`INV_SMA_SunSpec_Inverter.csv`](./INV_SMA_SunSpec_Inverter.csv) | SMA | SunSpec-Inverter | Modbus | 29 |
| [`INV_SofarSolar_1-40KTL.csv`](./INV_SofarSolar_1-40KTL.csv) | SofarSolar | 1-40KTL | Modbus | 115 |
| [`INV_SofarSolar_50-70KTL.csv`](./INV_SofarSolar_50-70KTL.csv) | SofarSolar | 50-70KTL | Modbus | 152 |
| [`INV_SofarSolar_HYDG3.csv`](./INV_SofarSolar_HYDG3.csv) | SofarSolar | HYDG3 | Modbus | 538 |
| [`INV_SolarEdge_66KW.csv`](./INV_SolarEdge_66KW.csv) | SolarEdge | 66KW | Modbus | 38 |
| [`INV_SolarEdge_ActivePowerControl.csv`](./INV_SolarEdge_ActivePowerControl.csv) | SolarEdge | ActivePowerControl | Modbus | 16 |
| [`INV_SolarEdge_EnhancedPowerControl.csv`](./INV_SolarEdge_EnhancedPowerControl.csv) | SolarEdge | EnhancedPowerControl | Modbus | 18 |
| [`INV_SolarEdge_SE_Series.csv`](./INV_SolarEdge_SE_Series.csv) | SolarEdge | SE-Series | Modbus | 28 |
| [`INV_SolarEdge_SE25K_SunSpec.csv`](./INV_SolarEdge_SE25K_SunSpec.csv) | SolarEdge | SE25K_SunSpec | Modbus | 53 |
| [`INV_SolarEdge_SE33.3K_SunSpec.csv`](./INV_SolarEdge_SE33.3K_SunSpec.csv) | SolarEdge | SE33.3K_SunSpec | Modbus | 69 |
| [`INV_SolarEdge_SunSpec_Generic.csv`](./INV_SolarEdge_SunSpec_Generic.csv) | SolarEdge | SunSpec_Generic | Modbus | 46 |
| [`INV_Solax_Generic.csv`](./INV_Solax_Generic.csv) | Solax | Generic | Modbus | 214 |
| [`INV_Solax_X1-X3_Generic.csv`](./INV_Solax_X1-X3_Generic.csv) | Solax | X1-X3_Generic | Modbus | 157 |
| [`INV_Solis_GENERIC.csv`](./INV_Solis_GENERIC.csv) | Solis | GENERIC | Modbus | 454 |
| [`INV_Solis_Inverter.csv`](./INV_Solis_Inverter.csv) | Solis | Inverter_Standard | Modbus | 93 |
| [`INV_Solis_50_25_SunPM.csv`](./INV_Solis_50_25_SunPM.csv) | Solis | Solis_50_25_SunPM | Modbus | 47 |
| [`INV_Solis_SOLIS_GENERIC.csv`](./INV_Solis_SOLIS_GENERIC.csv) | Solis | SOLIS_GENERIC | Modbus | 455 |
| [`INV_Solis_SOLIS_GENERIC_GridTied.csv`](./INV_Solis_SOLIS_GENERIC_GridTied.csv) | Solis | SOLIS_GENERIC_GridTied | Modbus | 690 |
| [`INV_Solis_SunSpec_25_66K.csv`](./INV_Solis_SunSpec_25_66K.csv) | Solis | SunSpec_25_66K | Modbus | 308 |
| [`INV_Solplanet_Ai-Logger.csv`](./INV_Solplanet_Ai-Logger.csv) | Solplanet | Ai-Logger | Modbus | 106 |
| [`INV_Sungrow_Generic.csv`](./INV_Sungrow_Generic.csv) | Sungrow | Generic | Modbus | 164 |
| [`INV_Sungrow_SG_Series.csv`](./INV_Sungrow_SG_Series.csv) | Sungrow | SG-Series | Modbus | 29 |
| [`INV_Sungrow_SG_Series_Generic.csv`](./INV_Sungrow_SG_Series_Generic.csv) | Sungrow | SG_Series_Generic | Modbus | 200 |
| [`INV_TATA_Generic.csv`](./INV_TATA_Generic.csv) | TATA | Generic | Modbus | 20 |
| [`INV_VICTRON_Generic.csv`](./INV_VICTRON_Generic.csv) | VICTRON | Generic | Modbus | 12 |
| [`INV_Zucchetti_Azzuro-3PH.csv`](./INV_Zucchetti_Azzuro-3PH.csv) | Zucchetti | Azzuro-3PH | Modbus | 503 |

### Compteurs & Analyseurs Réseau (Meters) (73 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`MET_ABB_B21-B23-B24.csv`](./MET_ABB_B21-B23-B24.csv) | ABB | B21-B23-B24 | Modbus | 137 |
| [`MET_ABB_B21_B23_B24.csv`](./MET_ABB_B21_B23_B24.csv) | ABB | B21_B23_B24 | Modbus | 137 |
| [`MET_ABB_M2M.csv`](./MET_ABB_M2M.csv) | ABB | M2M | Modbus | 126 |
| [`MET_ABB_M2M_Meter.csv`](./MET_ABB_M2M_Meter.csv) | ABB | M2M_Meter | Modbus | 126 |
| [`MET_ABB_M4M.csv`](./MET_ABB_M4M.csv) | ABB | M4M | Modbus | 298 |
| [`MET_Algodue_UPM209-309.csv`](./MET_Algodue_UPM209-309.csv) | Algodue | UPM209-309 | Modbus | 98 |
| [`MET_Carlo_Gavazzi_EM210.csv`](./MET_Carlo_Gavazzi_EM210.csv) | Carlo_Gavazzi | EM210 | Modbus | 115 |
| [`MET_CarloGavazzi_EM111.csv`](./MET_CarloGavazzi_EM111.csv) | CarloGavazzi | EM111 | Modbus | 41 |
| [`MET_CarloGavazzi_EM112.csv`](./MET_CarloGavazzi_EM112.csv) | CarloGavazzi | EM112 | Modbus | 41 |
| [`MET_CarloGavazzi_EM21.csv`](./MET_CarloGavazzi_EM21.csv) | CarloGavazzi | EM21 | Modbus | 35 |
| [`MET_CarloGavazzi_EM210.csv`](./MET_CarloGavazzi_EM210.csv) | CarloGavazzi | EM210 | Modbus | 115 |
| [`MET_CarloGavazzi_EM210_MV.csv`](./MET_CarloGavazzi_EM210_MV.csv) | CarloGavazzi | EM210_MV | Modbus | 115 |
| [`MET_CarloGavazzi_EM24.csv`](./MET_CarloGavazzi_EM24.csv) | CarloGavazzi | EM24 | Modbus | 57 |
| [`MET_CarloGavazzi_EM271.csv`](./MET_CarloGavazzi_EM271.csv) | CarloGavazzi | EM271 | Modbus | 25 |
| [`MET_CarloGavazzi_EM330.csv`](./MET_CarloGavazzi_EM330.csv) | CarloGavazzi | EM330 | Modbus | 65 |
| [`MET_CarloGavazzi_EM340.csv`](./MET_CarloGavazzi_EM340.csv) | CarloGavazzi | EM340 | Modbus | 65 |
| [`MET_CarloGavazzi_EM530.csv`](./MET_CarloGavazzi_EM530.csv) | CarloGavazzi | EM530 | Modbus | 143 |
| [`MET_CarloGavazzi_EM540.csv`](./MET_CarloGavazzi_EM540.csv) | CarloGavazzi | EM540 | Modbus | 143 |
| [`MET_CarloGavazzi_ET112.csv`](./MET_CarloGavazzi_ET112.csv) | CarloGavazzi | ET112 | Modbus | 41 |
| [`MET_CarloGavazzi_ET330.csv`](./MET_CarloGavazzi_ET330.csv) | CarloGavazzi | ET330 | Modbus | 65 |
| [`MET_CarloGavazzi_ET340.csv`](./MET_CarloGavazzi_ET340.csv) | CarloGavazzi | ET340 | Modbus | 65 |
| [`MET_CarloGavazzi_WM15.csv`](./MET_CarloGavazzi_WM15.csv) | CarloGavazzi | WM15 | Modbus | 33 |
| [`MET_CarloGavazzi_WM20.csv`](./MET_CarloGavazzi_WM20.csv) | CarloGavazzi | WM20 | Modbus | 40 |
| [`MET_CarloGavazzi_WM30.csv`](./MET_CarloGavazzi_WM30.csv) | CarloGavazzi | WM30 | Modbus | 40 |
| [`MET_CarloGavazzi_WM40.csv`](./MET_CarloGavazzi_WM40.csv) | CarloGavazzi | WM40 | Modbus | 40 |
| [`MET_CarloGavazzi_WM50.csv`](./MET_CarloGavazzi_WM50.csv) | CarloGavazzi | WM50 | Modbus | 40 |
| [`MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv`](./MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv) | CHAUVIN-ARNOUX | Ulys_TT-M | Modbus | 19 |
| [`MET_CHINT_DTSU666.csv`](./MET_CHINT_DTSU666.csv) | CHINT | DTSU666 | Modbus | 41 |
| [`MET_CIRCUTOR_CVME3MINI.csv`](./MET_CIRCUTOR_CVME3MINI.csv) | CIRCUTOR | CVME3MINI | Modbus | 31 |
| [`MET_ECO_ADAPT_POWERELEC_6.csv`](./MET_ECO_ADAPT_POWERELEC_6.csv) | ECO_ADAPT | POWERELEC_6 | Modbus | 221 |
| [`MET_EDMI_Mk6.csv`](./MET_EDMI_Mk6.csv) | EDMI | Mk6 | Modbus | 59 |
| [`MET_Fronius_Smart_Meter_50kA-3.csv`](./MET_Fronius_Smart_Meter_50kA-3.csv) | Fronius | Smart_Meter_50kA | Modbus | 72 |
| [`MET_Fronius_Smart_Meter_IntSF.csv`](./MET_Fronius_Smart_Meter_IntSF.csv) | Fronius | Smart_Meter_IntSF | Modbus | 165 |
| [`MET_Fronius_Smart_Meter_TS.csv`](./MET_Fronius_Smart_Meter_TS.csv) | Fronius | Smart_Meter_TS | Modbus | 144 |
| [`MET_Goodwe_GM1000_GM3000.csv`](./MET_Goodwe_GM1000_GM3000.csv) | GoodWe | GM1000-GM3000 | Modbus | 18 |
| [`MET_HAGER_Generic.csv`](./MET_HAGER_Generic.csv) | HAGER | Generic | Modbus | 102 |
| [`MET_HUAWEI_Power_Meter.csv`](./MET_HUAWEI_Power_Meter.csv) | HUAWEI | Power_Meter | Modbus | 44 |
| [`MET_Huawei_PowerMeter.csv`](./MET_Huawei_PowerMeter.csv) | Huawei | PowerMeter | Modbus | 44 |
| [`MET_Inepro_PRO380-MOD-CT.csv`](./MET_Inepro_PRO380-MOD-CT.csv) | Inepro | PRO380-MOD-CT | Modbus | 24 |
| [`MET_Janitza_Generic.csv`](./MET_Janitza_Generic.csv) | Janitza | Generic | Modbus | 61 |
| [`MET_Legrand_LEG412052.csv`](./MET_Legrand_LEG412052.csv) | Legrand | LEG412052 | Modbus | 489 |
| [`MET_Lettel_MAP4-34RJ.csv`](./MET_Lettel_MAP4-34RJ.csv) | Lettel | MAP4-34RJ | Modbus | 405 |
| [`MET_Lettel_MCT4-45.csv`](./MET_Lettel_MCT4-45.csv) | Lettel | MCT4-45 | Modbus | 30 |
| [`MET_Lettel_MCX4-34.csv`](./MET_Lettel_MCX4-34.csv) | Lettel | MCX4-34 | Modbus | 139 |
| [`MET_Lettel_MCX4-34V.csv`](./MET_Lettel_MCX4-34V.csv) | Lettel | MCX4-34V | Modbus | 102 |
| [`MET_POLIER_MTR5LMOD.csv`](./MET_POLIER_MTR5LMOD.csv) | POLIER | MTR5LMOD | Modbus | 33 |
| [`MET_Schneider_EM1200.csv`](./MET_Schneider_EM1200.csv) | Schneider | EM1200 | Modbus | 96 |
| [`MET_Schneider_EM1200-EM1220.csv`](./MET_Schneider_EM1200-EM1220.csv) | Schneider | EM1200-EM1220 | Modbus | 96 |
| [`MET_Schneider_Generic.csv`](./MET_Schneider_Generic.csv) | Schneider | Generic | Modbus | 517 |
| [`MET_Schneider_iEM3000.csv`](./MET_Schneider_iEM3000.csv) | Schneider | iEM3000 | Modbus | 55 |
| [`MET_Schneider_iEM3400_3500.csv`](./MET_Schneider_iEM3400_3500.csv) | Schneider | iEM3400_3500 | Modbus | 140 |
| [`MET_Schneider_IEM_3xxx.csv`](./MET_Schneider_IEM_3xxx.csv) | Schneider | IEM_3xxx | Modbus | 55 |
| [`MET_Schneider_PM5xxx.csv`](./MET_Schneider_PM5xxx.csv) | Schneider | PM5xxx | Modbus | 35 |
| [`MET_Schneider_PM5xxx_Series.csv`](./MET_Schneider_PM5xxx_Series.csv) | Schneider | PM5xxx_Series | Modbus | 875 |
| [`MET_Schneider_PowerLogic_PMC.csv`](./MET_Schneider_PowerLogic_PMC.csv) | Schneider | PowerLogic_PMC | Modbus | 517 |
| [`MET_Schneider_RELAYP5.csv`](./MET_Schneider_RELAYP5.csv) | Schneider | RELAYP5 | Modbus | 57 |
| [`MET_SMA_EnergyMeter.csv`](./MET_SMA_EnergyMeter.csv) | SMA | EnergyMeter | Modbus | 26 |
| [`MET_SMA_HomeManager.csv`](./MET_SMA_HomeManager.csv) | SMA | HomeManager | Modbus | 26 |
| [`MET_SOCOMEC_COUNTIS_E23.csv`](./MET_SOCOMEC_COUNTIS_E23.csv) | Socomec | Countis_E23 | Modbus | 59 |
| [`MET_Socomec_Countis_E27.csv`](./MET_Socomec_Countis_E27.csv) | Socomec | Countis_E27 | Modbus | 162 |
| [`MET_Socomec_Countis_E33_E43.csv`](./MET_Socomec_Countis_E33_E43.csv) | Socomec | Countis_E33_E43 | Modbus | 73 |
| [`MET_Socomec_Countis_E50_E53.csv`](./MET_Socomec_Countis_E50_E53.csv) | Socomec | Countis_E50_E53 | Modbus | 141 |
| [`MET_Socomec_Countis_P44.csv`](./MET_Socomec_Countis_P44.csv) | Socomec | Countis_P44 | Modbus | 37 |
| [`MET_Socomec_Diris_A10.csv`](./MET_Socomec_Diris_A10.csv) | Socomec | Diris_A10 | Modbus | 75 |
| [`MET_SOCOMEC_Diris_A40.csv`](./MET_SOCOMEC_Diris_A40.csv) | Socomec | Diris_A40 | Modbus | 35 |
| [`MET_Socomec_Diris_Axx.csv`](./MET_Socomec_Diris_Axx.csv) | Socomec | Diris_Axx | Modbus | 25 |
| [`MET_SOCOMEC_DIRIS_B10.csv`](./MET_SOCOMEC_DIRIS_B10.csv) | Socomec | Diris_B10 | Modbus | 23 |
| [`MET_Socomec_Diris_B30.csv`](./MET_Socomec_Diris_B30.csv) | Socomec | Diris_B30 | Modbus | 20 |
| [`MET_Socomec_Diris_Digiware_I35.csv`](./MET_Socomec_Diris_Digiware_I35.csv) | Socomec | Diris_Digiware_I35 | Modbus | 26 |
| [`MET_Socomec_Diris_I30.csv`](./MET_Socomec_Diris_I30.csv) | Socomec | Diris_I30 | Modbus | 20 |
| [`MET_SOCOMEC_E33_E43.csv`](./MET_SOCOMEC_E33_E43.csv) | SOCOMEC | E33_E43 | Modbus | 73 |
| [`MET_SolarEdge_SE-MTR-3Y-400V-A.csv`](./MET_SolarEdge_SE-MTR-3Y-400V-A.csv) | SolarEdge | SE-MTR-3Y-400V-A | Modbus | 246 |
| [`MET_Weidmuller_610.csv`](./MET_Weidmuller_610.csv) | Weidmuller | 610 | Modbus | 122 |

### Systèmes de Stockage BESS (Batteries) (11 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`BAT_Dyness_DH100F.csv`](./BAT_Dyness_DH100F.csv) | Dyness | DH100F | Modbus | 913 |
| [`BAT_Dyness_DH200F.csv`](./BAT_Dyness_DH200F.csv) | Dyness | DH200F | Modbus | 913 |
| [`BAT_Dyness_DH200Y.csv`](./BAT_Dyness_DH200Y.csv) | Dyness | DH200Y | Modbus | 912 |
| [`BAT_Dyness_DH_Series_Generic.csv`](./BAT_Dyness_DH_Series_Generic.csv) | Dyness | DH_Series_Generic | Modbus | 962 |
| [`BAT_Goodwe_Lynx_Home_HV.csv`](./BAT_Goodwe_Lynx_Home_HV.csv) | GoodWe | Lynx-Home-HV | Modbus | 18 |
| [`BAT_Huawei_LUNA2000.csv`](./BAT_Huawei_LUNA2000.csv) | Huawei | LUNA2000 | Modbus | 170 |
| [`BAT_Huawei_LUNA2000_Storage.csv`](./BAT_Huawei_LUNA2000_Storage.csv) | Huawei | LUNA2000_Storage | Modbus | 272 |
| [`BAT_Sungrow_Battery.csv`](./BAT_Sungrow_Battery.csv) | Sungrow | Battery | Modbus | 172 |
| [`BAT_Sungrow_FullMap_v2.csv`](./BAT_Sungrow_FullMap_v2.csv) | Sungrow | FullMap_v2 | Modbus | 801 |
| [`BAT_Sungrow_LC3000.csv`](./BAT_Sungrow_LC3000.csv) | Sungrow | LC3000 | Modbus | 916 |
| [`BAT_Sungrow_PowerStack_ST255CS.csv`](./BAT_Sungrow_PowerStack_ST255CS.csv) | Sungrow | PowerStack_ST255CS | Modbus | 1442 |

### Capteurs & Pyranomètres (Sensors) (22 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`SEN_ATERSA_DATASOLMET.csv`](./SEN_ATERSA_DATASOLMET.csv) | ATERSA | DATASOLMET | Modbus | 5 |
| [`SEN_BD_SENSOR_FUEL_SENSOR.csv`](./SEN_BD_SENSOR_FUEL_SENSOR.csv) | BD_SENSOR | FUEL_SENSOR | Modbus | 17 |
| [`SEN_DELTA_OHM_LP_PYRA_03_.csv`](./SEN_DELTA_OHM_LP_PYRA_03_.csv) | DELTA_OHM | LP_PYRA_03_ | Modbus | 6 |
| [`SEN_DELTAOHM_HD523D.csv`](./SEN_DELTAOHM_HD523D.csv) | DELTAOHM | HD523D | Modbus | 10 |
| [`SEN_EKO_MS40.csv`](./SEN_EKO_MS40.csv) | EKO | MS40 | Modbus | 28 |
| [`SEN_Fronius_Sensor_Card.csv`](./SEN_Fronius_Sensor_Card.csv) | Fronius | Sensor_Card | Modbus | 54 |
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

### Passerelles & Loggers de Communication (13 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`COM_Goodwe_EzLogger_Pro.csv`](./COM_Goodwe_EzLogger_Pro.csv) | GoodWe | EzLogger-Pro | Modbus | 25 |
| [`COM_GoodWe_Ezlogger3000.csv`](./COM_GoodWe_Ezlogger3000.csv) | GoodWe | Ezlogger3000 | Modbus | 25 |
| [`COM_Goodwe_SCB3000.csv`](./COM_Goodwe_SCB3000.csv) | GoodWe | SCB3000 | Modbus | 29 |
| [`COM_Huawei_SL3000.csv`](./COM_Huawei_SL3000.csv) | Huawei | SL3000 | Modbus | 82 |
| [`COM_Huawei_SmartLogger3000.csv`](./COM_Huawei_SmartLogger3000.csv) | Huawei | SmartLogger3000 | Modbus | 83 |
| [`COM_Huawei_SmartLogger_VSeries.csv`](./COM_Huawei_SmartLogger_VSeries.csv) | Huawei | SmartLogger_VSeries | Modbus | 176 |
| [`COM_SMA_ClusterController.csv`](./COM_SMA_ClusterController.csv) | SMA | ClusterController | Modbus | 24 |
| [`COM_SMA_ComGateway.csv`](./COM_SMA_ComGateway.csv) | SMA | ComGateway | Modbus | 22 |
| [`COM_SMA_DataManager_M.csv`](./COM_SMA_DataManager_M.csv) | SMA | DataManager-M | Modbus | 28 |
| [`COM_Solplanet_Ai-Logger.csv`](./COM_Solplanet_Ai-Logger.csv) | Solplanet | Ai-Logger | Modbus | 161 |
| [`COM_Sungrow_EMS300CP.csv`](./COM_Sungrow_EMS300CP.csv) | Sungrow | EMS300CP | Modbus | 694 |
| [`COM_Sungrow_LC300_Modbus.csv`](./COM_Sungrow_LC300_Modbus.csv) | Sungrow | LC300_Modbus | Modbus | 813 |
| [`COM_Sungrow_Logger1000_3000.csv`](./COM_Sungrow_Logger1000_3000.csv) | Sungrow | Logger1000_3000 | Modbus | 50 |

### Autres Équipements (IO Modules & Afficheurs) (7 fichiers)

| Fichier | Constructeur | Modèle | Protocole | Registres |
|---------|--------------|--------|-----------|-----------|
| [`OTH_ABB_Terra_AC_Charger.csv`](./OTH_ABB_Terra_AC_Charger.csv) | ABB | Terra_AC_Charger | Modbus | 20 |
| [`OTH_ICPDAS_M7028.csv`](./OTH_ICPDAS_M7028.csv) | ICPDAS | M7028 | Modbus | 70 |
| [`OTH_SIEBERT_XC4XX.csv`](./OTH_SIEBERT_XC4XX.csv) | SIEBERT | XC4XX | Modbus | 3 |
| [`OTH_Sungrow_PVCombinerBox.csv`](./OTH_Sungrow_PVCombinerBox.csv) | Sungrow | PVCombinerBox | Modbus | 64 |
| [`OTH_Webdyn_Sens_IO_16DI.csv`](./OTH_Webdyn_Sens_IO_16DI.csv) | Webdyn | Sens_IO_16DI | Modbus | 60 |
| [`OTH_Webdyn_Sens_IO_8AI.csv`](./OTH_Webdyn_Sens_IO_8AI.csv) | Webdyn | Sens_IO_8AI | Modbus | 34 |
| [`OTH_Webdyn_Sens_IO_8DI-6DO.csv`](./OTH_Webdyn_Sens_IO_8DI-6DO.csv) | Webdyn | Sens_IO_8DI-6DO | Modbus | 42 |

