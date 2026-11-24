# 🏭 Webdyn Production Definitions Library

Bibliothèque officielle de définitions d'équipements industriels validées pour déploiement en production sur les concentrateurs et passerelles Webdyn (SunPM, Hub, Flex).

Comprend **1194 fichiers de définition certifiés** couvrant **358,911 registres Modbus** et **36,072 tags SunPM canoniques** à travers **188 constructeurs mondiaux**.

---

## 📊 Synthèse par Catégorie

| Catégorie | Préfixe | Fichiers | Total Registres | Tags Actifs | Exemples Constructeurs |
|---|---|---|---|---|---|
| Onduleurs & Convertisseurs (Inverter) | `INV_` | 483 | 131,658 | 17,659 | 5, ABB, APsystems, ATESS, AnswerDrives (+84) |
| Compteurs & Analyseurs Réseau (Meter) | `MET_` | 232 | 31,485 | 3,606 | ABB, AEberle, ALGODUE, Acrel, Algodue (+57) |
| Systèmes BESS & Batteries (Battery) | `BAT_` | 56 | 55,004 | 3,282 | ATMOCE, AlroSolar, Dyness, FOX ESS, FoxESS (+11) |
| Capteurs Météo & Pyranomètres (Sensor) | `SEN_` | 55 | 1,434 | 181 | ATERSA, BD, Blet, DELTAOHM, DeltaOhm (+20) |
| Passerelles & Loggers (Logger) | `COM_` | 38 | 33,386 | 2,437 | ATMOCE, FoxESS, GoodWe, Greenheiss, HOYMILES (+7) |
| Modules IO & Afficheurs (Other) | `OTH_` | 364 | 105,944 | 8,907 | 5, ABB, AEberle, Advantech, Alan (+97) |

---

## 🏭 Synthèse par Constructeur (Tous les Constructeurs)

| Constructeur | Modèles | Total Registres | Tags SunPM Actifs | Catégories Couvertes |
|---|:---:|:---:|:---:|---|
| **Schneider** | 82 | 13,540 | 1,380 | Inverter, Meter, Other |
| **ABB** | 75 | 39,509 | 2,641 | Inverter, Meter, Other |
| **Huawei** | 69 | 11,630 | 2,259 | Battery, Inverter, Logger, Meter, Other |
| **GoodWe** | 61 | 6,836 | 3,062 | Battery, Inverter, Logger, Meter |
| **MarsEnergy** | 57 | 38,938 | 1,090 | Battery, Inverter, Logger, Meter, Other |
| **Fronius** | 47 | 11,479 | 1,477 | Inverter, Meter, Other, Sensor |
| **Delta** | 35 | 20,387 | 1,095 | Inverter, Other |
| **SMA** | 35 | 9,615 | 1,445 | Inverter, Logger, Meter, Other |
| **Socomec** | 33 | 2,190 | 378 | Inverter, Meter, Other |
| **Sungrow** | 30 | 13,329 | 2,965 | Battery, Inverter, Logger, Other |
| **CarloGavazzi** | 27 | 1,710 | 546 | Meter, Other |
| **FoxESS** | 27 | 11,376 | 2,285 | Battery, Inverter, Logger, Meter, Other, Storage |
| **Kaco** | 27 | 12,185 | 779 | Inverter, Other |
| **SolarEdge** | 22 | 1,334 | 351 | Inverter, Meter, Other |
| **Ingeteam** | 19 | 1,099 | 296 | Inverter, Other |
| **SAJ** | 19 | 3,187 | 391 | Inverter, Other |
| **SLEnergy** | 18 | 6,794 | 1,019 | Battery, Inverter, Meter, Other |
| **Growatt** | 16 | 5,647 | 1,928 | Inverter |
| **Legrand** | 14 | 2,029 | 231 | Inverter, Meter, Other |
| **Lettel** | 14 | 2,722 | 296 | Meter, Other |
| **Janitza** | 13 | 6,582 | 236 | Meter |
| **Solis** | 13 | 4,712 | 385 | Inverter |
| **IngenieurBuro** | 12 | 81 | 21 | Other, Sensor |
| **Kehua** | 12 | 1,500 | 408 | Inverter, Other |
| **Solplanet** | 12 | 4,392 | 784 | Inverter, Logger, Other |
| **SIEL** | 11 | 215 | 3 | Inverter |
| **CyberPower** | 10 | 834 | 128 | Inverter, Other |
| **KippZonen** | 10 | 346 | 17 | Other, Sensor |
| **Kostal** | 9 | 1,061 | 171 | Inverter, Other |
| **Solax** | 9 | 1,429 | 78 | Inverter |
| **Deye** | 8 | 1,601 | 122 | Inverter, Meter, Other |
| **HUAWEI** | 8 | 1,065 | 155 | Battery, Inverter, Meter |
| **Webdyn** | 8 | 225 | 18 | Other, Sensor |
| **AlroSolar** | 7 | 1,904 | 184 | Battery, Other |
| **Efacec** | 7 | 820 | 113 | Inverter, Other |
| **INGETEAM** | 7 | 423 | 108 | Inverter |
| **Kstar** | 7 | 628 | 161 | Inverter, Other |
| **SofarSolar** | 7 | 1,549 | 92 | Inverter |
| **X7** | 7 | 1,588 | 1,492 | Battery, Inverter, Meter, Other |
| **Hukseflux** | 6 | 136 | 18 | Sensor |
| **Hyxi** | 6 | 1,037 | 223 | Inverter |
| **IME** | 6 | 458 | 106 | Meter, Other |
| **Jema** | 6 | 303 | 127 | Inverter, Other |
| **SOCOMEC** | 6 | 282 | 71 | Meter |
| **AnswerDrives** | 5 | 742 | 338 | Inverter |
| **CEFEM** | 5 | 265 | 65 | Inverter, Other |
| **CHINT** | 5 | 141 | 92 | Meter, Other |
| **Lufft** | 5 | 153 | 15 | Meter, Other, Sensor |
| **MarsRenewable** | 5 | 6,860 | 54 | Inverter, Meter |
| **Seven** | 5 | 181 | 30 | Sensor |
| **Victron** | 5 | 1,347 | 139 | Inverter, Other |
| **ATMOCE** | 4 | 188 | 52 | Battery, Logger |
| **ChauvinArnoux** | 4 | 298 | 66 | Meter, Other |
| **Dyness** | 4 | 3,700 | 210 | Battery |
| **EKO** | 4 | 136 | 14 | Other, Sensor |
| **Greenheiss** | 4 | 1,081 | 33 | Logger, Other |
| **Hoymiles** | 4 | 9,645 | 84 | Battery, Inverter, Meter, Other |
| **MCI** | 4 | 242 | 57 | Meter, Other |
| **Mars Energy** | 4 | 33,492 | 8 | BESS, BatteryRack |
| **PowerElectronics** | 4 | 522 | 38 | Inverter, Other |
| **Siel** | 4 | 65 | 4 | Other |
| **Thermokon** | 4 | 264 | 12 | Other, Sensor |
| **Ticmaster** | 4 | 178 | 13 | Meter, Other |
| **Zucchetti** | 4 | 3,291 | 95 | Inverter, Other |
| **AEberle** | 3 | 296 | 42 | Meter, Other |
| **Bonfiglioli** | 3 | 124 | 30 | Inverter, Meter, Other |
| **CIRCUTOR** | 3 | 66 | 10 | Meter |
| **Cyberpower** | 3 | 591 | 392 | Inverter, Other |
| **EcoAdapt** | 3 | 663 | 11 | Meter, Other |
| **HOYMILES** | 3 | 9,518 | 78 | Logger |
| **IMESYS** | 3 | 1,651 | 86 | Meter, Other |
| **INVT** | 3 | 123 | 87 | Inverter, Other |
| **Inepro** | 3 | 72 | 22 | Meter, Other |
| **KOSTAL** | 3 | 310 | 74 | Inverter |
| **KSTAR** | 3 | 482 | 22 | Inverter |
| **Kerneli** | 3 | 441 | 58 | Inverter, Other |
| **LSI** | 3 | 90 | 2 | Other, Sensor |
| **MeierNT** | 3 | 54 | 9 | Other, Sensor |
| **POLIER** | 3 | 429 | 64 | Meter |
| **SIEBERT** | 3 | 9 | 8 | Other |
| **SRegeltechnik** | 3 | 24 | 2 | Other, Sensor |
| **Siemens** | 3 | 7,537 | 59 | Inverter, Other |
| **SolarLog** | 3 | 252 | 66 | Meter, Other |
| **Tata** | 3 | 60 | 33 | Inverter, Other |
| **TicMaster** | 3 | 18 | 12 | Meter |
| **Universal** | 3 | 309 | 84 | Meter, Other |
| **Weidmuller** | 3 | 277 | 61 | Meter, Other |
| **ZeverSolar** | 3 | 255 | 34 | Other |
| **5** | 2 | 98 | 40 | Inverter, Other |
| **APsystems** | 2 | 226 | 66 | Inverter |
| **Advantech** | 2 | 72 | 2 | Other |
| **Algodue** | 2 | 103 | 7 | Meter, Other |
| **Autarco** | 2 | 551 | 46 | Inverter, Other |
| **DELTA** | 2 | 196 | 48 | Inverter |
| **DeltaOhm** | 2 | 11 | 0 | Sensor |
| **EDMI** | 2 | 118 | 7 | Meter, Other |
| **EcsOptec** | 2 | 88 | 16 | Meter |
| **Emu** | 2 | 565 | 42 | Meter |
| **FOX ESS** | 2 | 497 | 64 | Battery, Inverter |
| **GREENHEISS** | 2 | 604 | 18 | Inverter |
| **HAGER** | 2 | 204 | 14 | Meter, Other |
| **Hitachi** | 2 | 220 | 56 | Inverter, Other |
| **ICPDAS** | 2 | 140 | 37 | Other |
| **JEMA** | 2 | 120 | 27 | Inverter |
| **KEHUA** | 2 | 476 | 122 | Inverter |
| **LEGRAND** | 2 | 53 | 11 | Inverter, Meter |
| **Maximet** | 2 | 22 | 22 | Other, Sensor |
| **OG** | 2 | 192 | 30 | Meter, Other |
| **POWERELECTRONICS** | 2 | 30 | 8 | Inverter |
| **Panasonic** | 2 | 28 | 8 | Meter, Other |
| **PhoenixContact** | 2 | 174 | 14 | Meter, Other |
| **Refusol** | 2 | 680 | 53 | Other |
| **Robestec** | 2 | 3,180 | 34 | Battery, Logger |
| **SATCON** | 2 | 191 | 14 | Inverter |
| **SNEF** | 2 | 32 | 26 | Inverter, Other |
| **SOFARSOLAR** | 2 | 1,085 | 64 | Inverter |
| **Siebert** | 2 | 6 | 6 | Other |
| **Socorel** | 2 | 1,384 | 16 | Meter, Other |
| **SolaX** | 2 | 528 | 61 | Inverter |
| **TMEIC** | 2 | 120 | 43 | Inverter, Other |
| **Thies** | 2 | 201 | 11 | Other, Sensor |
| **Waaree** | 2 | 92 | 10 | Inverter, Other |
| **Zelec** | 2 | 268 | 40 | Meter |
| **ALGODUE** | 1 | 5 | 3 | Meter |
| **ATERSA** | 1 | 5 | 0 | Sensor |
| **ATESS** | 1 | 243 | 29 | Inverter |
| **Acrel** | 1 | 257 | 16 | Meter |
| **Alan** | 1 | 75 | 1 | Other |
| **Anexo** | 1 | 14 | 3 | Meter |
| **Atmoce** | 1 | 47 | 12 | Other |
| **Axun** | 1 | 227 | 7 | Other |
| **BD** | 1 | 17 | 0 | Sensor |
| **BONFIGLIOLI** | 1 | 55 | 29 | Inverter |
| **Bender** | 1 | 26 | 2 | Meter |
| **Blet** | 1 | 29 | 6 | Sensor |
| **BlueSun** | 1 | 6 | 0 | Inverter |
| **Cefem** | 1 | 19 | 4 | Meter |
| **DELTAOHM** | 1 | 10 | 0 | Sensor |
| **EASTRON** | 1 | 125 | 28 | Meter |
| **EFACEC** | 1 | 58 | 6 | Inverter |
| **Eastron** | 1 | 125 | 28 | Inverter |
| **Eflex4u** | 1 | 405 | 11 | Meter |
| **Elster** | 1 | 98 | 10 | Meter |
| **Enerfox** | 1 | 64 | 24 | Meter |
| **FOX_ESS** | 1 | 87 | 11 | Inverter |
| **Gamesa** | 1 | 17 | 0 | Inverter |
| **Goodwe** | 1 | 36 | 18 | Inverter |
| **Hopewind** | 1 | 78 | 68 | Other |
| **Hypercharger** | 1 | 80 | 5 | Other |
| **Inaccess** | 1 | 1,098 | 2 | Other |
| **Intelec** | 1 | 10 | 0 | Other |
| **K** | 1 | 36 | 18 | Other |
| **KIPPZONEN** | 1 | 15 | 0 | Sensor |
| **LACECAL** | 1 | 67 | 0 | Meter |
| **Lacecal** | 1 | 67 | 0 | Other |
| **Monichek** | 1 | 524 | 56 | Other |
| **Nidec** | 1 | 73 | 12 | Inverter |
| **Polier** | 1 | 33 | 17 | Other |
| **Power** | 1 | 58 | 12 | Other |
| **REFUSOL** | 1 | 536 | 36 | Inverter |
| **Riello** | 1 | 48 | 17 | Other |
| **SACI** | 1 | 86 | 14 | Meter |
| **SALICRU** | 1 | 249 | 64 | Inverter |
| **SCHNEIDER** | 1 | 73 | 41 | Inverter |
| **SESAutomation** | 1 | 10 | 3 | Sensor |
| **SIGENERGY** | 1 | 194 | 61 | Inverter |
| **SINENG** | 1 | 443 | 24 | Inverter |
| **SOLAREDGE** | 1 | 46 | 10 | Inverter |
| **SOLARLOG** | 1 | 148 | 30 | Meter |
| **SOLAX** | 1 | 347 | 37 | Battery |
| **SOLPLANET** | 1 | 625 | 60 | Inverter |
| **Saci** | 1 | 86 | 14 | Other |
| **Salgueirinha** | 1 | 1,294 | 34 | Other |
| **Satcon** | 1 | 95 | 11 | Other |
| **Secure** | 1 | 198 | 4 | Other |
| **Sigenergy** | 1 | 194 | 61 | Other |
| **Sineng** | 1 | 43 | 10 | Other |
| **Solaredge** | 1 | 38 | 9 | Inverter |
| **TATA** | 1 | 20 | 2 | Inverter |
| **THIES** | 1 | 91 | 4 | Sensor |
| **TRACKSO** | 1 | 2 | 0 | Sensor |
| **Termasgard** | 1 | 5 | 0 | Sensor |
| **VICTRON** | 1 | 12 | 0 | Inverter |
| **ZEVERSOLAR** | 1 | 82 | 16 | Inverter |
| **ZUCCHETTI** | 1 | 1,778 | 39 | Inverter |
| **Zeversolar** | 1 | 75 | 15 | Inverter |
| **fronius** | 1 | 193 | 34 | Inverter |
| **map** | 1 | 7 | 1 | Meter |

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
| [`INV_5_Generic.csv`](./INV_5_Generic.csv) | Inverter | 5 | Generic | Modbus | 49 | 20 |
| [`OTH_5_Generic.csv`](./OTH_5_Generic.csv) | Other | 5 | Generic | Modbus | 49 | 20 |
| [`INV_ABB_10.0_et_12.5_ABB_PVIMODBUS.csv`](./INV_ABB_10.0_et_12.5_ABB_PVIMODBUS.csv) | Inverter | ABB | 10.0_et_12.5_ABB_PVIMODBUS | Modbus | 19 | 8 |
| [`OTH_ABB_10.0_et_12.5_ABB_PVIMODBUS.csv`](./OTH_ABB_10.0_et_12.5_ABB_PVIMODBUS.csv) | Other | ABB | 10.0_et_12.5_ABB_PVIMODBUS | Modbus | 19 | 8 |
| [`INV_ABB_10.5.csv`](./INV_ABB_10.5.csv) | Inverter | ABB | 10.5 | Modbus | 19 | 3 |
| [`INV_ABB_5583-CHEVALIERP2_SunSpec_inverter_ABB_3Q4200_2117B_modbusRTU.csv`](./INV_ABB_5583-CHEVALIERP2_SunSpec_inverter_ABB_3Q4200_2117B_modbusRTU.csv) | Inverter | ABB | 5583-CHEVALIERP2_SunSpec_inverter_ABB_3Q4200_2117B_modbusRTU | Modbus | 1,111 | 37 |
| [`INV_ABB_AURORA_CENTRAL.csv`](./INV_ABB_AURORA_CENTRAL.csv) | Inverter | ABB | AURORA_Central | Modbus | 1,522 | 0 |
| [`MET_ABB_B21-B23-B24.csv`](./MET_ABB_B21-B23-B24.csv) | Meter | ABB | B21-B23-B24 | Modbus | 137 | 5 |
| [`MET_ABB_B21_B23_B24.csv`](./MET_ABB_B21_B23_B24.csv) | Meter | ABB | B21_B23_B24 | Modbus | 137 | 29 |
| [`INV_ABB_LM1_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv`](./INV_ABB_LM1_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv) | Inverter | ABB | LM1_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP | Modbus | 956 | 36 |
| [`MET_ABB_M2M.csv`](./MET_ABB_M2M.csv) | Meter | ABB | M2M | Modbus | 126 | 4 |
| [`MET_ABB_M2M_Meter.csv`](./MET_ABB_M2M_Meter.csv) | Meter | ABB | M2M_Meter | Modbus | 126 | 4 |
| [`MET_ABB_M4M.csv`](./MET_ABB_M4M.csv) | Meter | ABB | M4M | Modbus | 298 | 31 |
| [`OTH_ABB_M4M_Modbus_map_-_v.1.3N.csv`](./OTH_ABB_M4M_Modbus_map_-_v.1.3N.csv) | Other | ABB | M4M_Modbus_map_-_v.1.3N | Modbus | 532 | 31 |
| [`MET_ABB_M4M_NetworkAnalyzer.csv`](./MET_ABB_M4M_NetworkAnalyzer.csv) | Meter | ABB | M4M_NetworkAnalyzer | Modbus | 298 | 31 |
| [`INV_ABB_MODBUS_POWERONE_Ultra_Inverter_Module.csv`](./INV_ABB_MODBUS_POWERONE_Ultra_Inverter_Module.csv) | Inverter | ABB | MODBUS_POWERONE_Ultra_Inverter_Module | Modbus | 56 | 11 |
| [`OTH_ABB_Modbus_ABB_SunSpec.csv`](./OTH_ABB_Modbus_ABB_SunSpec.csv) | Other | ABB | Modbus_ABB_SunSpec | Modbus | 987 | 37 |
| [`INV_ABB_Modbus_ABB_SunSpec_Light.csv`](./INV_ABB_Modbus_ABB_SunSpec_Light.csv) | Inverter | ABB | Modbus_ABB_SunSpec_Light | Modbus | 89 | 33 |
| [`OTH_ABB_Modbus_ABB_SunSpec_Light.csv`](./OTH_ABB_Modbus_ABB_SunSpec_Light.csv) | Other | ABB | Modbus_ABB_SunSpec_Light | Modbus | 89 | 35 |
| [`OTH_ABB_Modbus_ABB_SunSpec_full.csv`](./OTH_ABB_Modbus_ABB_SunSpec_full.csv) | Other | ABB | Modbus_ABB_SunSpec_full | Modbus | 981 | 37 |
| [`INV_ABB_Modbus_PRO_33.csv`](./INV_ABB_Modbus_PRO_33.csv) | Inverter | ABB | Modbus_PRO_33 | Modbus | 62 | 12 |
| [`OTH_ABB_Modbus_PRO_33.csv`](./OTH_ABB_Modbus_PRO_33.csv) | Other | ABB | Modbus_PRO_33 | Modbus | 62 | 12 |
| [`INV_ABB_Modbus_PVI.csv`](./INV_ABB_Modbus_PVI.csv) | Inverter | ABB | Modbus_PVI | Modbus | 19 | 8 |
| [`OTH_ABB_Modbus_PVI.csv`](./OTH_ABB_Modbus_PVI.csv) | Other | ABB | Modbus_PVI | Modbus | 19 | 8 |
| [`INV_ABB_Modbus_TRIO.csv`](./INV_ABB_Modbus_TRIO.csv) | Inverter | ABB | Modbus_TRIO | Modbus | 17 | 6 |
| [`OTH_ABB_Modbus_TRIO.csv`](./OTH_ABB_Modbus_TRIO.csv) | Other | ABB | Modbus_TRIO | Modbus | 17 | 6 |
| [`INV_ABB_Modbus_Ultra_Module.csv`](./INV_ABB_Modbus_Ultra_Module.csv) | Inverter | ABB | Modbus_Ultra_Module | Modbus | 56 | 10 |
| [`OTH_ABB_Modbus_Ultra_Module.csv`](./OTH_ABB_Modbus_Ultra_Module.csv) | Other | ABB | Modbus_Ultra_Module | Modbus | 56 | 10 |
| [`MET_ABB_Modbus_mapping_Meter_ABB_B21.csv`](./MET_ABB_Modbus_mapping_Meter_ABB_B21.csv) | Meter | ABB | Modbus_mapping_Meter_ABB_B21 | Modbus | 41 | 7 |
| [`INV_ABB_PRO-33.0-TL.csv`](./INV_ABB_PRO-33.0-TL.csv) | Inverter | ABB | PRO-33.0-TL | Modbus | 62 | 12 |
| [`INV_ABB_PVI-10.0_12.5.csv`](./INV_ABB_PVI-10.0_12.5.csv) | Inverter | ABB | PVI-10.0_12.5 | Modbus | 19 | 0 |
| [`INV_ABB_PVI-110_275.csv`](./INV_ABB_PVI-110_275.csv) | Inverter | ABB | PVI-110_275 | Modbus | 441 | 9 |
| [`INV_ABB_PVI-50_60.csv`](./INV_ABB_PVI-50_60.csv) | Inverter | ABB | PVI-50_60 | Modbus | 104 | 19 |
| [`INV_ABB_PVI110-275.csv`](./INV_ABB_PVI110-275.csv) | Inverter | ABB | PVI110-275 | Modbus | 441 | 9 |
| [`INV_ABB_PVI50-60.csv`](./INV_ABB_PVI50-60.csv) | Inverter | ABB | PVI50-60 | Modbus | 104 | 19 |
| [`INV_ABB_PVS-100_120.csv`](./INV_ABB_PVS-100_120.csv) | Inverter | ABB | PVS-100_120 | Modbus | 986 | 35 |
| [`INV_ABB_PVS-100_120_PVS1xx_TCP_PICS.csv`](./INV_ABB_PVS-100_120_PVS1xx_TCP_PICS.csv) | Inverter | ABB | PVS-100_120_PVS1xx_TCP_PICS | Modbus | 986 | 35 |
| [`OTH_ABB_PVS-100_120_PVS1xx_TCP_PICS.csv`](./OTH_ABB_PVS-100_120_PVS1xx_TCP_PICS.csv) | Other | ABB | PVS-100_120_PVS1xx_TCP_PICS | Modbus | 986 | 39 |
| [`INV_ABB_PVS-10_12.5_15_20_30_33.csv`](./INV_ABB_PVS-10_12.5_15_20_30_33.csv) | Inverter | ABB | PVS-10_12.5_15_20_30_33 | Modbus | 963 | 33 |
| [`INV_ABB_PVS-175.csv`](./INV_ABB_PVS-175.csv) | Inverter | ABB | PVS-175 | Modbus | 1,068 | 37 |
| [`INV_ABB_PVS-30-TL.csv`](./INV_ABB_PVS-30-TL.csv) | Inverter | ABB | PVS-30-TL | Modbus | 956 | 5 |
| [`INV_ABB_PVS-50_60.csv`](./INV_ABB_PVS-50_60.csv) | Inverter | ABB | PVS-50_60 | Modbus | 956 | 33 |
| [`INV_ABB_PVS1xx_SunSpec.csv`](./INV_ABB_PVS1xx_SunSpec.csv) | Inverter | ABB | PVS1xx_SunSpec | Modbus | 673 | 27 |
| [`INV_ABB_PVS50_MODBUS_light.csv`](./INV_ABB_PVS50_MODBUS_light.csv) | Inverter | ABB | PVS50_MODBUS_light | Modbus | 76 | 24 |
| [`OTH_ABB_PVS50_MODBUS_light.csv`](./OTH_ABB_PVS50_MODBUS_light.csv) | Other | ABB | PVS50_MODBUS_light | Modbus | 76 | 24 |
| [`INV_ABB_PVS800.csv`](./INV_ABB_PVS800.csv) | Inverter | ABB | PVS800 | Modbus | 1,053 | 23 |
| [`OTH_ABB_PVS800.csv`](./OTH_ABB_PVS800.csv) | Other | ABB | PVS800 | Modbus | 1,053 | 23 |
| [`INV_ABB_PVS800-57B.csv`](./INV_ABB_PVS800-57B.csv) | Inverter | ABB | PVS800-57B | Modbus | 1,053 | 17 |
| [`INV_ABB_PVS800_Work.csv`](./INV_ABB_PVS800_Work.csv) | Inverter | ABB | PVS800_Work | Modbus | 1,044 | 22 |
| [`OTH_ABB_PVS800_Work.csv`](./OTH_ABB_PVS800_Work.csv) | Other | ABB | PVS800_Work | Modbus | 1,044 | 1,044 |
| [`INV_ABB_PVS800_en_cours.csv`](./INV_ABB_PVS800_en_cours.csv) | Inverter | ABB | PVS800_en_cours | Modbus | 1,053 | 22 |
| [`OTH_ABB_PVS800_en_cours.csv`](./OTH_ABB_PVS800_en_cours.csv) | Other | ABB | PVS800_en_cours | Modbus | 1,053 | 22 |
| [`INV_ABB_PVS800_en_cours_noBitfield.csv`](./INV_ABB_PVS800_en_cours_noBitfield.csv) | Inverter | ABB | PVS800_en_cours_noBitfield | Modbus | 1,053 | 22 |
| [`OTH_ABB_PVS800_en_cours_noBitfield.csv`](./OTH_ABB_PVS800_en_cours_noBitfield.csv) | Other | ABB | PVS800_en_cours_noBitfield | Modbus | 1,053 | 22 |
| [`INV_ABB_PVS_100_120_SunspecLight.csv`](./INV_ABB_PVS_100_120_SunspecLight.csv) | Inverter | ABB | PVS_100_120_SunspecLight | Modbus | 118 | 26 |
| [`OTH_ABB_PVS_100_120_SunspecLight.csv`](./OTH_ABB_PVS_100_120_SunspecLight.csv) | Other | ABB | PVS_100_120_SunspecLight | Modbus | 118 | 26 |
| [`OTH_ABB_SUNSPEC_PVS-100_120_PICS_Rev_005.csv`](./OTH_ABB_SUNSPEC_PVS-100_120_PICS_Rev_005.csv) | Other | ABB | SUNSPEC_PVS-100_120_PICS_Rev_005 | Modbus | 1,143 | 33 |
| [`OTH_ABB_SUNSPEC_PVS-10_12.5_15_20_30_33_PICS_Rev_001.csv`](./OTH_ABB_SUNSPEC_PVS-10_12.5_15_20_30_33_PICS_Rev_001.csv) | Other | ABB | SUNSPEC_PVS-10_12.5_15_20_30_33_PICS_Rev_001 | Modbus | 963 | 33 |
| [`OTH_ABB_SUNSPEC_PVS-175_PICS_Rev_003.csv`](./OTH_ABB_SUNSPEC_PVS-175_PICS_Rev_003.csv) | Other | ABB | SUNSPEC_PVS-175_PICS_Rev_003 | Modbus | 1,068 | 37 |
| [`OTH_ABB_SUNSPEC_PVS-50_60_PICS_Rev_3.csv`](./OTH_ABB_SUNSPEC_PVS-50_60_PICS_Rev_3.csv) | Other | ABB | SUNSPEC_PVS-50_60_PICS_Rev_3 | Modbus | 956 | 33 |
| [`INV_ABB_SUNSPEC_PVS1xx_PICS_V002_GL20181011.csv`](./INV_ABB_SUNSPEC_PVS1xx_PICS_V002_GL20181011.csv) | Inverter | ABB | SUNSPEC_PVS1xx_PICS_V002_GL20181011 | Modbus | 1,118 | 36 |
| [`OTH_ABB_SUNSPEC_PVS1xx_PICS_V002_GL20181011.csv`](./OTH_ABB_SUNSPEC_PVS1xx_PICS_V002_GL20181011.csv) | Other | ABB | SUNSPEC_PVS1xx_PICS_V002_GL20181011 | Modbus | 1,119 | 40 |
| [`INV_ABB_SunSpec_IM_1_103_160.csv`](./INV_ABB_SunSpec_IM_1_103_160.csv) | Inverter | ABB | SunSpec_IM_1_103_160 | Modbus | 127 | 23 |
| [`OTH_ABB_SunSpec_IM_1_103_160.csv`](./OTH_ABB_SunSpec_IM_1_103_160.csv) | Other | ABB | SunSpec_IM_1_103_160 | Modbus | 127 | 23 |
| [`OTH_ABB_SunSpec_PVS.csv`](./OTH_ABB_SunSpec_PVS.csv) | Other | ABB | SunSpec_PVS | Modbus | 981 | 37 |
| [`INV_ABB_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv`](./INV_ABB_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv) | Inverter | ABB | SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP | Modbus | 956 | 36 |
| [`INV_ABB_TRIO-50.0-TM-OUTD.csv`](./INV_ABB_TRIO-50.0-TM-OUTD.csv) | Inverter | ABB | TRIO-50.0-TM-OUTD | Modbus | 956 | 35 |
| [`INV_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.csv`](./INV_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.csv) | Inverter | ABB | TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS | Modbus | 956 | 35 |
| [`OTH_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.csv`](./OTH_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.csv) | Other | ABB | TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS | Modbus | 956 | 39 |
| [`INV_ABB_TRIO20-27-6.csv`](./INV_ABB_TRIO20-27-6.csv) | Inverter | ABB | TRIO20-27-6 | Modbus | 128 | 18 |
| [`INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv`](./INV_ABB_TRIO_20.0-27.6_TL_OUTD.csv) | Inverter | ABB | TRIO_20.0-27.6_TL_OUTD | Modbus | 89 | 7 |
| [`INV_ABB_TRIO_200_276_FR.csv`](./INV_ABB_TRIO_200_276_FR.csv) | Inverter | ABB | TRIO_200_276_FR | Modbus | 89 | 15 |
| [`OTH_ABB_TRIO_200_276_FR.csv`](./OTH_ABB_TRIO_200_276_FR.csv) | Other | ABB | TRIO_200_276_FR | Modbus | 89 | 15 |
| [`OTH_ABB_Terra_AC_Charger.csv`](./OTH_ABB_Terra_AC_Charger.csv) | Other | ABB | Terra_AC_Charger | Modbus | 20 | 9 |
| [`INV_ABB_ULTRA_Inverter_Module.csv`](./INV_ABB_ULTRA_Inverter_Module.csv) | Inverter | ABB | ULTRA_Inverter_Module | Modbus | 57 | 6 |
| [`INV_ABB_Ultra_Inverter_Module.csv`](./INV_ABB_Ultra_Inverter_Module.csv) | Inverter | ABB | ULTRA_Inverter_Module | Modbus | 57 | 6 |
| [`INV_ABB_Velux_SunSpec_inverter_FIMER_-3Q15-_2201B_modbusTCP.csv`](./INV_ABB_Velux_SunSpec_inverter_FIMER_-3Q15-_2201B_modbusTCP.csv) | Inverter | ABB | Velux_SunSpec_inverter_FIMER_-3Q15-_2201B_modbusTCP | Modbus | 1,111 | 37 |
| [`MET_AEberle_14178-PM_MODBUS_RTU_CPT_PQPLUS_UMD96.csv`](./MET_AEberle_14178-PM_MODBUS_RTU_CPT_PQPLUS_UMD96.csv) | Meter | AEberle | 14178-PM_MODBUS_RTU_CPT_PQPLUS_UMD96 | Modbus | 126 | 9 |
| [`OTH_AEberle_14178-PM_MODBUS_RTU_CPT_PQPLUS_UMD96.csv`](./OTH_AEberle_14178-PM_MODBUS_RTU_CPT_PQPLUS_UMD96.csv) | Other | AEberle | 14178-PM_MODBUS_RTU_CPT_PQPLUS_UMD96 | Modbus | 110 | 9 |
| [`MET_AEberle_PQPlus_Generic.csv`](./MET_AEberle_PQPlus_Generic.csv) | Meter | AEberle | PQPlus_Generic | Modbus | 60 | 24 |
| [`MET_Algodue_UPM209.csv`](./MET_Algodue_UPM209.csv) | Meter | ALGODUE | UPM209 | Modbus | 5 | 3 |
| [`INV_APsystems_Altenergy_Power_System_Inc._single_phase_MicroinverterAPpcs_11-12.csv`](./INV_APsystems_Altenergy_Power_System_Inc._single_phase_MicroinverterAPpcs_11-12.csv) | Inverter | APsystems | Altenergy_Power_System_Inc._single_phase_MicroinverterAPpcs_11-12 | Modbus | 113 | 33 |
| [`INV_APsystems_Altenergy_Power_System_Inc._three_phase_microinverter_11-12.csv`](./INV_APsystems_Altenergy_Power_System_Inc._three_phase_microinverter_11-12.csv) | Inverter | APsystems | Altenergy_Power_System_Inc._three_phase_microinverter_11-12 | Modbus | 113 | 33 |
| [`SEN_Atersa_DATASOLMET.csv`](./SEN_Atersa_DATASOLMET.csv) | Sensor | ATERSA | DATASOLMET | Modbus | 5 | 0 |
| [`INV_Atess_Generic.csv`](./INV_Atess_Generic.csv) | Inverter | ATESS | GENERIC | Modbus | 243 | 29 |
| [`COM_Atmoce_Gateway_V1.2_WebdynSunPM_COMPLET.csv`](./COM_Atmoce_Gateway_V1.2_WebdynSunPM_COMPLET.csv) | Logger | ATMOCE | Gateway_V1.2_WebdynSunPM_COMPLET | Modbus | 47 | 15 |
| [`BAT_Atmoce_MC-MG.csv`](./BAT_Atmoce_MC-MG.csv) | Battery | ATMOCE | MC-MG | Modbus | 47 | 13 |
| [`COM_Atmoce_MC-MG.csv`](./COM_Atmoce_MC-MG.csv) | Logger | ATMOCE | MC-MG | Modbus | 47 | 12 |
| [`COM_Atmoce_MC100_MC100-T_MG100.csv`](./COM_Atmoce_MC100_MC100-T_MG100.csv) | Logger | ATMOCE | MC100_MC100-T_MG100 | Modbus | 47 | 12 |
| [`MET_Acrel_DTSD1352.csv`](./MET_Acrel_DTSD1352.csv) | Meter | Acrel | DTSD1352 | Modbus | 257 | 16 |
| [`OTH_Advantech_6224.csv`](./OTH_Advantech_6224.csv) | Other | Advantech | 6224 | Modbus | 36 | 1 |
| [`OTH_Advantech_MODBUS_ADAM_6224.csv`](./OTH_Advantech_MODBUS_ADAM_6224.csv) | Other | Advantech | MODBUS_ADAM_6224 | Modbus | 36 | 1 |
| [`OTH_Alan_india_Alan_Annunciator.csv`](./OTH_Alan_india_Alan_Annunciator.csv) | Other | Alan | india_Alan_Annunciator | Modbus | 75 | 1 |
| [`OTH_Algodue_UPM209.csv`](./OTH_Algodue_UPM209.csv) | Other | Algodue | UPM209 | Modbus | 5 | 3 |
| [`MET_Algodue_UPM209-309.csv`](./MET_Algodue_UPM209-309.csv) | Meter | Algodue | UPM209-309 | Modbus | 98 | 4 |
| [`BAT_AlroSolar_Alro_BMS.csv`](./BAT_AlroSolar_Alro_BMS.csv) | Battery | AlroSolar | Alro_BMS | Modbus | 114 | 6 |
| [`OTH_AlroSolar_Alro_BMS.csv`](./OTH_AlroSolar_Alro_BMS.csv) | Other | AlroSolar | Alro_BMS | Modbus | 114 | 5 |
| [`OTH_AlroSolar_Alro_PCS.csv`](./OTH_AlroSolar_Alro_PCS.csv) | Other | AlroSolar | Alro_PCS | Modbus | 443 | 16 |
| [`BAT_AlroSolar_BMS_Definition.csv`](./BAT_AlroSolar_BMS_Definition.csv) | Battery | AlroSolar | BMS_Definition | Modbus | 114 | 0 |
| [`OTH_AlroSolar_BMS_Definition.csv`](./OTH_AlroSolar_BMS_Definition.csv) | Other | AlroSolar | BMS_Definition | Modbus | 114 | 114 |
| [`OTH_AlroSolar_EMS_Definition.csv`](./OTH_AlroSolar_EMS_Definition.csv) | Other | AlroSolar | EMS_Definition | Modbus | 167 | 0 |
| [`OTH_AlroSolar_PCS_Definition.csv`](./OTH_AlroSolar_PCS_Definition.csv) | Other | AlroSolar | PCS_Definition | Modbus | 838 | 43 |
| [`MET_Anexo_METER.csv`](./MET_Anexo_METER.csv) | Meter | Anexo | METER | Modbus | 14 | 3 |
| [`INV_AnswerDrives_2009_WebdynSunPM.csv`](./INV_AnswerDrives_2009_WebdynSunPM.csv) | Inverter | AnswerDrives | 2009_WebdynSunPM | Modbus | 76 | 13 |
| [`INV_AnswerDrives_Inverter.csv`](./INV_AnswerDrives_Inverter.csv) | Inverter | AnswerDrives | Inverter | Modbus | 151 | 12 |
| [`INV_AnswerDrives_SOLARGATE_2009_WebdynSunPM.csv`](./INV_AnswerDrives_SOLARGATE_2009_WebdynSunPM.csv) | Inverter | AnswerDrives | SOLARGATE_2009_WebdynSunPM | Modbus | 76 | 13 |
| [`INV_AnswerDrives_SOLARGATE_Inverter.csv`](./INV_AnswerDrives_SOLARGATE_Inverter.csv) | Inverter | AnswerDrives | SOLARGATE_Inverter | Modbus | 151 | 12 |
| [`INV_AnswerDrives_Solargate.csv`](./INV_AnswerDrives_Solargate.csv) | Inverter | AnswerDrives | Solargate | Modbus | 288 | 288 |
| [`OTH_Atmoce_MC100_MC100-T_MG100.csv`](./OTH_Atmoce_MC100_MC100-T_MG100.csv) | Other | Atmoce | MC100_MC100-T_MG100 | Modbus | 47 | 12 |
| [`OTH_Autarco_Modbus_Table.csv`](./OTH_Autarco_Modbus_Table.csv) | Other | Autarco | Modbus_Table | Modbus | 75 | 14 |
| [`INV_Autarco_Solis_table2023.csv`](./INV_Autarco_Solis_table2023.csv) | Inverter | Autarco | Solis_table2023 | Modbus | 476 | 32 |
| [`OTH_Axun_SunSpecInformationModelReference.csv`](./OTH_Axun_SunSpecInformationModelReference.csv) | Other | Axun | SunSpecInformationModelReference | Modbus | 227 | 7 |
| [`SEN_BD_FUEL_SENSOR.csv`](./SEN_BD_FUEL_SENSOR.csv) | Sensor | BD | FUEL_SENSOR | Modbus | 17 | 0 |
| [`INV_Bonfiglioli_Generic.csv`](./INV_Bonfiglioli_Generic.csv) | Inverter | BONFIGLIOLI | Generic | Modbus | 55 | 29 |
| [`MET_Bender_ISOMETER.csv`](./MET_Bender_ISOMETER.csv) | Meter | Bender | ISOMETER | Modbus | 26 | 2 |
| [`SEN_Blet_WINDSENSOR_ANEH5-US4MBP.csv`](./SEN_Blet_WINDSENSOR_ANEH5-US4MBP.csv) | Sensor | Blet | WINDSENSOR_ANEH5-US4MBP | Modbus | 29 | 6 |
| [`INV_BlueSun_Hybrid_Inverter_Modbus_Protocol.csv`](./INV_BlueSun_Hybrid_Inverter_Modbus_Protocol.csv) | Inverter | BlueSun | Hybrid_Inverter_Modbus_Protocol | Modbus | 6 | 0 |
| [`OTH_Bonfiglioli_Generic.csv`](./OTH_Bonfiglioli_Generic.csv) | Other | Bonfiglioli | Generic | Modbus | 55 | 29 |
| [`MET_Bonfiglioli_Scada_Parameters_updated_Bonfiglioli_7_2_1_4_003.csv`](./MET_Bonfiglioli_Scada_Parameters_updated_Bonfiglioli_7_2_1_4_003.csv) | Meter | Bonfiglioli | Scada_Parameters_updated_Bonfiglioli_7_2_1_4_003 | Modbus | 64 | 0 |
| [`INV_Bonfiglioli_inverter.csv`](./INV_Bonfiglioli_inverter.csv) | Inverter | Bonfiglioli | inverter | Modbus | 5 | 1 |
| [`INV_Cefem_DEF_InverterTrioSun.csv`](./INV_Cefem_DEF_InverterTrioSun.csv) | Inverter | CEFEM | DEF_InverterTrioSun | Modbus | 42 | 15 |
| [`INV_Cefem_DEF_InverterTrioSun36KVA_FichierFourniParCEFEM.csv`](./INV_Cefem_DEF_InverterTrioSun36KVA_FichierFourniParCEFEM.csv) | Inverter | CEFEM | DEF_InverterTrioSun36KVA_FichierFourniParCEFEM | Modbus | 55 | 15 |
| [`INV_Cefem_Modbus_TrioSunTrioTop.csv`](./INV_Cefem_Modbus_TrioSunTrioTop.csv) | Inverter | CEFEM | Modbus_TrioSunTrioTop | Modbus | 68 | 14 |
| [`OTH_Cefem_Modbus_TrioSunTrioTop.csv`](./OTH_Cefem_Modbus_TrioSunTrioTop.csv) | Other | CEFEM | Modbus_TrioSunTrioTop | Modbus | 68 | 13 |
| [`INV_Cefem_Table_Modbus_Trio-Sun_Trio-Top_2020_07_01.csv`](./INV_Cefem_Table_Modbus_Trio-Sun_Trio-Top_2020_07_01.csv) | Inverter | CEFEM | Table_Modbus_Trio-Sun_Trio-Top_2020_07_01 | Modbus | 32 | 8 |
| [`MET_Chint_DDSU666.csv`](./MET_Chint_DDSU666.csv) | Meter | CHINT | DDSU666 | Modbus | 14 | 14 |
| [`MET_Chint_DTSU666.csv`](./MET_Chint_DTSU666.csv) | Meter | CHINT | DTSU666 | Modbus | 41 | 41 |
| [`MET_Chint_DTSU666-H.csv`](./MET_Chint_DTSU666-H.csv) | Meter | CHINT | DTSU666-H | Modbus | 22 | 22 |
| [`MET_Chint_Power_CHINT_DTSU.csv`](./MET_Chint_Power_CHINT_DTSU.csv) | Meter | CHINT | Power_CHINT_DTSU | Modbus | 29 | 10 |
| [`OTH_Chint_Power_CHINT_IO_NM8N-MCCB.csv`](./OTH_Chint_Power_CHINT_IO_NM8N-MCCB.csv) | Other | CHINT | Power_CHINT_IO_NM8N-MCCB | Modbus | 35 | 5 |
| [`MET_Circutor_CVME3MINI.csv`](./MET_Circutor_CVME3MINI.csv) | Meter | CIRCUTOR | CVME3MINI | Modbus | 31 | 4 |
| [`MET_Circutor_MODBUS_PM_RTU_CIRCUTOR_CVME3MINI_carte2.csv`](./MET_Circutor_MODBUS_PM_RTU_CIRCUTOR_CVME3MINI_carte2.csv) | Meter | CIRCUTOR | MODBUS_PM_RTU_CIRCUTOR_CVME3MINI_carte2 | Modbus | 19 | 3 |
| [`MET_Circutor_MODBUS_PM_RTU_CIRCUTOR_CVME3MINI_kW.csv`](./MET_Circutor_MODBUS_PM_RTU_CIRCUTOR_CVME3MINI_kW.csv) | Meter | CIRCUTOR | MODBUS_PM_RTU_CIRCUTOR_CVME3MINI_kW | Modbus | 16 | 3 |
| [`MET_CarloGavazzi_72RVV23XOSX.csv`](./MET_CarloGavazzi_72RVV23XOSX.csv) | Meter | CarloGavazzi | 72RVV23XOSX | Modbus | 7 | 2 |
| [`OTH_CarloGavazzi_72RVV23XOSX.csv`](./OTH_CarloGavazzi_72RVV23XOSX.csv) | Other | CarloGavazzi | 72RVV23XOSX | Modbus | 7 | 2 |
| [`MET_CarloGavazzi_EM-ET300.csv`](./MET_CarloGavazzi_EM-ET300.csv) | Meter | CarloGavazzi | EM-ET300 | Modbus | 70 | 1 |
| [`OTH_CarloGavazzi_EM-ET300.csv`](./OTH_CarloGavazzi_EM-ET300.csv) | Other | CarloGavazzi | EM-ET300 | Modbus | 69 | 1 |
| [`MET_CarloGavazzi_EM111.csv`](./MET_CarloGavazzi_EM111.csv) | Meter | CarloGavazzi | EM111 | Modbus | 41 | 10 |
| [`MET_CarloGavazzi_EM112.csv`](./MET_CarloGavazzi_EM112.csv) | Meter | CarloGavazzi | EM112 | Modbus | 41 | 10 |
| [`MET_CarloGavazzi_EM21.csv`](./MET_CarloGavazzi_EM21.csv) | Meter | CarloGavazzi | EM21 | Modbus | 35 | 29 |
| [`MET_CarloGavazzi_EM210.csv`](./MET_CarloGavazzi_EM210.csv) | Meter | CarloGavazzi | EM210 | Modbus | 115 | 29 |
| [`MET_CarloGavazzi_EM210_MV.csv`](./MET_CarloGavazzi_EM210_MV.csv) | Meter | CarloGavazzi | EM210_MV | Modbus | 115 | 29 |
| [`MET_CarloGavazzi_EM24.csv`](./MET_CarloGavazzi_EM24.csv) | Meter | CarloGavazzi | EM24 | Modbus | 57 | 32 |
| [`OTH_CarloGavazzi_EM24.csv`](./OTH_CarloGavazzi_EM24.csv) | Other | CarloGavazzi | EM24 | Modbus | 28 | 8 |
| [`MET_CarloGavazzi_EM24_EM300.csv`](./MET_CarloGavazzi_EM24_EM300.csv) | Meter | CarloGavazzi | EM24-EM300 | Modbus | 27 | 27 |
| [`MET_CarloGavazzi_EM271.csv`](./MET_CarloGavazzi_EM271.csv) | Meter | CarloGavazzi | EM271 | Modbus | 25 | 17 |
| [`MET_CarloGavazzi_EM300-ET300.csv`](./MET_CarloGavazzi_EM300-ET300.csv) | Meter | CarloGavazzi | EM300-ET300 | Modbus | 147 | 1 |
| [`OTH_CarloGavazzi_EM300-ET300.csv`](./OTH_CarloGavazzi_EM300-ET300.csv) | Other | CarloGavazzi | EM300-ET300 | Modbus | 146 | 1 |
| [`MET_CarloGavazzi_EM330.csv`](./MET_CarloGavazzi_EM330.csv) | Meter | CarloGavazzi | EM330 | Modbus | 65 | 30 |
| [`MET_CarloGavazzi_EM340.csv`](./MET_CarloGavazzi_EM340.csv) | Meter | CarloGavazzi | EM340 | Modbus | 65 | 30 |
| [`MET_CarloGavazzi_EM530.csv`](./MET_CarloGavazzi_EM530.csv) | Meter | CarloGavazzi | EM530 | Modbus | 143 | 31 |
| [`MET_CarloGavazzi_EM540.csv`](./MET_CarloGavazzi_EM540.csv) | Meter | CarloGavazzi | EM540 | Modbus | 143 | 31 |
| [`MET_CarloGavazzi_ET112.csv`](./MET_CarloGavazzi_ET112.csv) | Meter | CarloGavazzi | ET112 | Modbus | 41 | 10 |
| [`MET_CarloGavazzi_ET330.csv`](./MET_CarloGavazzi_ET330.csv) | Meter | CarloGavazzi | ET330 | Modbus | 65 | 30 |
| [`MET_CarloGavazzi_ET340.csv`](./MET_CarloGavazzi_ET340.csv) | Meter | CarloGavazzi | ET340 | Modbus | 65 | 30 |
| [`MET_CarloGavazzi_WM15.csv`](./MET_CarloGavazzi_WM15.csv) | Meter | CarloGavazzi | WM15 | Modbus | 33 | 31 |
| [`MET_CarloGavazzi_WM20.csv`](./MET_CarloGavazzi_WM20.csv) | Meter | CarloGavazzi | WM20 | Modbus | 40 | 31 |
| [`MET_CarloGavazzi_WM30.csv`](./MET_CarloGavazzi_WM30.csv) | Meter | CarloGavazzi | WM30 | Modbus | 40 | 31 |
| [`MET_CarloGavazzi_WM40.csv`](./MET_CarloGavazzi_WM40.csv) | Meter | CarloGavazzi | WM40 | Modbus | 40 | 31 |
| [`MET_CarloGavazzi_WM50.csv`](./MET_CarloGavazzi_WM50.csv) | Meter | CarloGavazzi | WM50 | Modbus | 40 | 31 |
| [`MET_Cefem_DEF_GridMeterChauvinArnouxUlysTT_M.csv`](./MET_Cefem_DEF_GridMeterChauvinArnouxUlysTT_M.csv) | Meter | Cefem | DEF_GridMeterChauvinArnouxUlysTT_M | Modbus | 19 | 4 |
| [`MET_ChauvinArnoux_ULYS-TT_M.csv`](./MET_ChauvinArnoux_ULYS-TT_M.csv) | Meter | ChauvinArnoux | ULYS-TT_M | Modbus | 19 | 4 |
| [`OTH_ChauvinArnoux_ULYS_FLEX.csv`](./OTH_ChauvinArnoux_ULYS_FLEX.csv) | Other | ChauvinArnoux | ULYS_FLEX | Modbus | 130 | 29 |
| [`MET_ChauvinArnoux_ULYS_FLEX.csv`](./MET_ChauvinArnoux_ULYS_FLEX.csv) | Meter | ChauvinArnoux | Ulys_Flex | Modbus | 130 | 29 |
| [`MET_ChauvinArnoux_Ulys_TT-M.csv`](./MET_ChauvinArnoux_Ulys_TT-M.csv) | Meter | ChauvinArnoux | Ulys_TT-M | Modbus | 19 | 4 |
| [`INV_Cyberpower_1-Phase.csv`](./INV_Cyberpower_1-Phase.csv) | Inverter | CyberPower | 1-Phase | Modbus | 85 | 7 |
| [`INV_Cyberpower_3-Phases.csv`](./INV_Cyberpower_3-Phases.csv) | Inverter | CyberPower | 3-Phases | Modbus | 71 | 6 |
| [`INV_Cyberpower_50000ETL.csv`](./INV_Cyberpower_50000ETL.csv) | Inverter | CyberPower | 50000ETL | Modbus | 80 | 22 |
| [`INV_Cyberpower_CPSHB6000ETL48-CPSHB6000ETLFR.csv`](./INV_Cyberpower_CPSHB6000ETL48-CPSHB6000ETLFR.csv) | Inverter | CyberPower | CPSHB6000ETL48-CPSHB6000ETLFR | Modbus | 103 | 4 |
| [`INV_Cyberpower_CPSHB6000ETL48_CPSHB6000ETLFR.csv`](./INV_Cyberpower_CPSHB6000ETL48_CPSHB6000ETLFR.csv) | Inverter | CyberPower | CPSHB6000ETL48-CPSHB6000ETLFR | Modbus | 103 | 20 |
| [`OTH_Cyberpower_Cyberpower3Ph_50000ETL.csv`](./OTH_Cyberpower_Cyberpower3Ph_50000ETL.csv) | Other | CyberPower | Cyberpower3Ph_50000ETL | Modbus | 80 | 22 |
| [`INV_Cyberpower_Modbus_1Ph.csv`](./INV_Cyberpower_Modbus_1Ph.csv) | Inverter | CyberPower | Modbus_1Ph | Modbus | 85 | 11 |
| [`OTH_Cyberpower_Modbus_1Ph.csv`](./OTH_Cyberpower_Modbus_1Ph.csv) | Other | CyberPower | Modbus_1Ph | Modbus | 85 | 10 |
| [`INV_Cyberpower_Modbus_3Ph.csv`](./INV_Cyberpower_Modbus_3Ph.csv) | Inverter | CyberPower | Modbus_3Ph | Modbus | 71 | 13 |
| [`OTH_Cyberpower_Modbus_3Ph.csv`](./OTH_Cyberpower_Modbus_3Ph.csv) | Other | CyberPower | Modbus_3Ph | Modbus | 71 | 13 |
| [`OTH_Cyberpower_CPSHB6000ETL48_CPSHB6000ETLFR.csv`](./OTH_Cyberpower_CPSHB6000ETL48_CPSHB6000ETLFR.csv) | Other | Cyberpower | CPSHB6000ETL48_CPSHB6000ETLFR | Modbus | 103 | 20 |
| [`INV_Cyberpower_CPSPV120K-140K.csv`](./INV_Cyberpower_CPSPV120K-140K.csv) | Inverter | Cyberpower | CPSPV120K-140K | Modbus | 244 | 186 |
| [`INV_Cyberpower_GENERIC.csv`](./INV_Cyberpower_GENERIC.csv) | Inverter | Cyberpower | GENERIC | Modbus | 244 | 186 |
| [`INV_Delta_MODBUS_DELTA_M15A_Flex_Sunspec.csv`](./INV_Delta_MODBUS_DELTA_M15A_Flex_Sunspec.csv) | Inverter | DELTA | MODBUS_DELTA_M15A_Flex_Sunspec | Modbus | 112 | 24 |
| [`INV_Delta_Sunspec_Light.csv`](./INV_Delta_Sunspec_Light.csv) | Inverter | DELTA | Sunspec_Light | Modbus | 84 | 24 |
| [`SEN_DeltaOhm_HD523D.csv`](./SEN_DeltaOhm_HD523D.csv) | Sensor | DELTAOHM | HD523D | Modbus | 10 | 0 |
| [`INV_Delta_1043_Ur_PM_SunSpec_inverter_Delta_RPI-M88H_modbusRTU.csv`](./INV_Delta_1043_Ur_PM_SunSpec_inverter_Delta_RPI-M88H_modbusRTU.csv) | Inverter | Delta | 1043_Ur_PM_SunSpec_inverter_Delta_RPI-M88H_modbusRTU | Modbus | 904 | 34 |
| [`INV_Delta_2962CASTELSOLTISNERATPM_Delta_RPI-M88H.csv`](./INV_Delta_2962CASTELSOLTISNERATPM_Delta_RPI-M88H.csv) | Inverter | Delta | 2962CASTELSOLTISNERATPM_Delta_RPI-M88H | Modbus | 1,048 | 43 |
| [`OTH_Delta_2962CASTELSOLTISNERATPM_Delta_RPI-M88H.csv`](./OTH_Delta_2962CASTELSOLTISNERATPM_Delta_RPI-M88H.csv) | Other | Delta | 2962CASTELSOLTISNERATPM_Delta_RPI-M88H | Modbus | 1,048 | 43 |
| [`INV_Delta_6409FORSANS_SunSpec_inverter_Delta_M50A_260_01.14_modbusRTU_M50AFLEX.csv`](./INV_Delta_6409FORSANS_SunSpec_inverter_Delta_M50A_260_01.14_modbusRTU_M50AFLEX.csv) | Inverter | Delta | 6409FORSANS_SunSpec_inverter_Delta_M50A_260_01.14_modbusRTU_M50AFLEX | Modbus | 980 | 36 |
| [`INV_Delta_866LABATJ173PM_Delta_RPI-M88H.csv`](./INV_Delta_866LABATJ173PM_Delta_RPI-M88H.csv) | Inverter | Delta | 866LABATJ173PM_Delta_RPI-M88H | Modbus | 1,048 | 43 |
| [`OTH_Delta_866LABATJ173PM_Delta_RPI-M88H.csv`](./OTH_Delta_866LABATJ173PM_Delta_RPI-M88H.csv) | Other | Delta | 866LABATJ173PM_Delta_RPI-M88H | Modbus | 1,048 | 43 |
| [`INV_Delta_AIMAR11_inverter_Delta_M50A_260.csv`](./INV_Delta_AIMAR11_inverter_Delta_M50A_260.csv) | Inverter | Delta | AIMAR11_inverter_Delta_M50A_260 | Modbus | 301 | 28 |
| [`OTH_Delta_ConfigModel.csv`](./OTH_Delta_ConfigModel.csv) | Other | Delta | ConfigModel | Modbus | 87 | 1 |
| [`OTH_Delta_M100A_280_COM1.15.csv`](./OTH_Delta_M100A_280_COM1.15.csv) | Other | Delta | M100A_280_COM1.15 | Modbus | 1,019 | 36 |
| [`OTH_Delta_M15A-flex_COM115.csv`](./OTH_Delta_M15A-flex_COM115.csv) | Other | Delta | M15A-flex_COM115 | Modbus | 969 | 36 |
| [`OTH_Delta_M30A-flex_COM117.csv`](./OTH_Delta_M30A-flex_COM117.csv) | Other | Delta | M30A-flex_COM117 | Modbus | 969 | 36 |
| [`OTH_Delta_M30A_COM147.csv`](./OTH_Delta_M30A_COM147.csv) | Other | Delta | M30A_COM147 | Modbus | 738 | 35 |
| [`OTH_Delta_M50A-flex_COM114.csv`](./OTH_Delta_M50A-flex_COM114.csv) | Other | Delta | M50A-flex_COM114 | Modbus | 999 | 36 |
| [`OTH_Delta_M50A_COM139.csv`](./OTH_Delta_M50A_COM139.csv) | Other | Delta | M50A_COM139 | Modbus | 726 | 35 |
| [`OTH_Delta_M70AFLEX_COM01.28.csv`](./OTH_Delta_M70AFLEX_COM01.28.csv) | Other | Delta | M70AFLEX_COM01.28 | Modbus | 999 | 36 |
| [`OTH_Delta_M70A_COM01.20.csv`](./OTH_Delta_M70A_COM01.20.csv) | Other | Delta | M70A_COM01.20 | Modbus | 999 | 36 |
| [`INV_Delta_MARKET_SunSpec_inverter_Delta_M70A_260_01.28_modbusRTU.csv`](./INV_Delta_MARKET_SunSpec_inverter_Delta_M70A_260_01.28_modbusRTU.csv) | Inverter | Delta | MARKET_SunSpec_inverter_Delta_M70A_260_01.28_modbusRTU | Modbus | 980 | 37 |
| [`OTH_Delta_MODBUS_DELTA_M15A_Flex_Sunspec.csv`](./OTH_Delta_MODBUS_DELTA_M15A_Flex_Sunspec.csv) | Other | Delta | MODBUS_DELTA_M15A_Flex_Sunspec | Modbus | 112 | 24 |
| [`INV_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv`](./INV_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv) | Inverter | Delta | MODBUS_DELTA_M88H-COM1.20-Sunspec | Modbus | 640 | 34 |
| [`OTH_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv`](./OTH_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv) | Other | Delta | MODBUS_DELTA_M88H-COM1.20-Sunspec | Modbus | 640 | 39 |
| [`OTH_Delta_SunSpec-Delta_ed00_adding_FW_versions.csv`](./OTH_Delta_SunSpec-Delta_ed00_adding_FW_versions.csv) | Other | Delta | SunSpec-Delta_ed00_adding_FW_versions | Modbus | 71 | 23 |
| [`INV_Delta_SunSpec_Inverter_Delta_RPI-M88H_modbusRTU.csv`](./INV_Delta_SunSpec_Inverter_Delta_RPI-M88H_modbusRTU.csv) | Inverter | Delta | SunSpec_Inverter_Delta_RPI-M88H_modbusRTU | Modbus | 904 | 36 |
| [`OTH_Delta_Sunspec_Light.csv`](./OTH_Delta_Sunspec_Light.csv) | Other | Delta | Sunspec_Light | Modbus | 84 | 24 |
| [`INV_Delta_Sunspec_V1.20.csv`](./INV_Delta_Sunspec_V1.20.csv) | Inverter | Delta | Sunspec_V1.20 | Modbus | 640 | 34 |
| [`OTH_Delta_Sunspec_V1.20.csv`](./OTH_Delta_Sunspec_V1.20.csv) | Other | Delta | Sunspec_V1.20 | Modbus | 640 | 39 |
| [`OTH_Delta_Sunspec_forM30-50AandM88H_170120_fromDelta.csv`](./OTH_Delta_Sunspec_forM30-50AandM88H_170120_fromDelta.csv) | Other | Delta | Sunspec_forM30-50AandM88H_170120_fromDelta | Modbus | 45 | 21 |
| [`OTH_Delta_Sunspec_protocol_v1.11_20180725.csv`](./OTH_Delta_Sunspec_protocol_v1.11_20180725.csv) | Other | Delta | Sunspec_protocol_v1.11_20180725 | Modbus | 92 | 24 |
| [`OTH_Delta_Sunspec_protocol_v1.12_20190403.csv`](./OTH_Delta_Sunspec_protocol_v1.12_20190403.csv) | Other | Delta | Sunspec_protocol_v1.12_20190403 | Modbus | 93 | 24 |
| [`OTH_Delta_Sunspec_protocol_v1.16_20210318.csv`](./OTH_Delta_Sunspec_protocol_v1.16_20210318.csv) | Other | Delta | Sunspec_protocol_v1.16_20210318 | Modbus | 122 | 24 |
| [`OTH_Delta_Sunspec_protocol_v1.23_20220906.csv`](./OTH_Delta_Sunspec_protocol_v1.23_20220906.csv) | Other | Delta | Sunspec_protocol_v1.23_20220906 | Modbus | 126 | 25 |
| [`OTH_Delta_Sunspec_protocol_v1_03_20170907.csv`](./OTH_Delta_Sunspec_protocol_v1_03_20170907.csv) | Other | Delta | Sunspec_protocol_v1_03_20170907 | Modbus | 71 | 23 |
| [`OTH_Delta_TechniqueSolaire_INI_M70A.csv`](./OTH_Delta_TechniqueSolaire_INI_M70A.csv) | Other | Delta | TechniqueSolaire_INI_M70A | Modbus | 989 | 36 |
| [`OTH_Delta_protocol_v1.11_20180725.csv`](./OTH_Delta_protocol_v1.11_20180725.csv) | Other | Delta | protocol_v1.11_20180725 | Modbus | 92 | 24 |
| [`OTH_Delta_protocol_v1.12_20190403.csv`](./OTH_Delta_protocol_v1.12_20190403.csv) | Other | Delta | protocol_v1.12_20190403 | Modbus | 93 | 24 |
| [`OTH_Delta_protocol_v1_03_20170907.csv`](./OTH_Delta_protocol_v1_03_20170907.csv) | Other | Delta | protocol_v1_03_20170907 | Modbus | 71 | 23 |
| [`SEN_DeltaOhm_LP_PYRA.csv`](./SEN_DeltaOhm_LP_PYRA.csv) | Sensor | DeltaOhm | LP_PYRA | Modbus | 5 | 0 |
| [`SEN_DeltaOhm_LP_PYRA_03.csv`](./SEN_DeltaOhm_LP_PYRA_03.csv) | Sensor | DeltaOhm | LP_PYRA_03 | Modbus | 6 | 0 |
| [`OTH_Deye_GENERIC.csv`](./OTH_Deye_GENERIC.csv) | Other | Deye | GENERIC | Modbus | 86 | 11 |
| [`INV_Deye_GENERICreg-1.csv`](./INV_Deye_GENERICreg-1.csv) | Inverter | Deye | GENERICreg-1 | Modbus | 86 | 11 |
| [`OTH_Deye_GENERICreg-1.csv`](./OTH_Deye_GENERICreg-1.csv) | Other | Deye | GENERICreg-1 | Modbus | 86 | 11 |
| [`INV_Deye_GENERICreg1.csv`](./INV_Deye_GENERICreg1.csv) | Inverter | Deye | GENERICreg1 | Modbus | 86 | 11 |
| [`OTH_Deye_GENERICreg1.csv`](./OTH_Deye_GENERICreg1.csv) | Other | Deye | GENERICreg1 | Modbus | 86 | 11 |
| [`INV_Deye_Generic.csv`](./INV_Deye_Generic.csv) | Inverter | Deye | Generic | Modbus | 327 | 23 |
| [`MET_Deye_Modbus_-Required_Basic_Parameter_Pour_Kyaw_Thu_Hein_Client_Aaron_Hsu.csv`](./MET_Deye_Modbus_-Required_Basic_Parameter_Pour_Kyaw_Thu_Hein_Client_Aaron_Hsu.csv) | Meter | Deye | Modbus_-Required_Basic_Parameter_Pour_Kyaw_Thu_Hein_Client_Aaron_Hsu | Modbus | 477 | 10 |
| [`OTH_Deye_Sunspec-PICS-DEYE-SUN-0820.csv`](./OTH_Deye_Sunspec-PICS-DEYE-SUN-0820.csv) | Other | Deye | Sunspec-PICS-DEYE-SUN-0820 | Modbus | 367 | 34 |
| [`BAT_Dyness_DH100F.csv`](./BAT_Dyness_DH100F.csv) | Battery | Dyness | DH100F | Modbus | 913 | 53 |
| [`BAT_Dyness_DH200F.csv`](./BAT_Dyness_DH200F.csv) | Battery | Dyness | DH200F | Modbus | 913 | 53 |
| [`BAT_Dyness_DH200Y.csv`](./BAT_Dyness_DH200Y.csv) | Battery | Dyness | DH200Y | Modbus | 912 | 50 |
| [`BAT_Dyness_DH_Series_Generic.csv`](./BAT_Dyness_DH_Series_Generic.csv) | Battery | Dyness | DH_Series_Generic | Modbus | 962 | 54 |
| [`MET_Eastron_SMART-X96-5E.csv`](./MET_Eastron_SMART-X96-5E.csv) | Meter | EASTRON | SMART-X96-5E | Modbus | 125 | 28 |
| [`MET_EDMI_Mk6.csv`](./MET_EDMI_Mk6.csv) | Meter | EDMI | Mk6 | Modbus | 59 | 4 |
| [`OTH_EDMI_Mk6.csv`](./OTH_EDMI_Mk6.csv) | Other | EDMI | Mk6 | Modbus | 59 | 3 |
| [`INV_Efacec_Efasolar.csv`](./INV_Efacec_Efasolar.csv) | Inverter | EFACEC | Efasolar | Modbus | 58 | 6 |
| [`OTH_EKO_MS40.csv`](./OTH_EKO_MS40.csv) | Other | EKO | MS40 | Modbus | 28 | 2 |
| [`SEN_EKO_MS40.csv`](./SEN_EKO_MS40.csv) | Sensor | EKO | MS40 | Modbus | 28 | 0 |
| [`OTH_EKO_MS80.csv`](./OTH_EKO_MS80.csv) | Other | EKO | MS80 | Modbus | 40 | 6 |
| [`SEN_EKO_MS80.csv`](./SEN_EKO_MS80.csv) | Sensor | EKO | MS80 | Modbus | 40 | 6 |
| [`INV_Eastron_SMART-X96-5E.csv`](./INV_Eastron_SMART-X96-5E.csv) | Inverter | Eastron | SMART-X96-5E | Modbus | 125 | 28 |
| [`MET_EcoAdapt_MODBUS_PowerElec6.csv`](./MET_EcoAdapt_MODBUS_PowerElec6.csv) | Meter | EcoAdapt | MODBUS_PowerElec6 | Modbus | 221 | 4 |
| [`OTH_EcoAdapt_MODBUS_PowerElec6.csv`](./OTH_EcoAdapt_MODBUS_PowerElec6.csv) | Other | EcoAdapt | MODBUS_PowerElec6 | Modbus | 221 | 4 |
| [`MET_EcoAdapt_POWERELEC_6.csv`](./MET_EcoAdapt_POWERELEC_6.csv) | Meter | EcoAdapt | POWERELEC_6 | Modbus | 221 | 3 |
| [`MET_EcsOptec_ECS_GENERIC_FLOAT.csv`](./MET_EcsOptec_ECS_GENERIC_FLOAT.csv) | Meter | EcsOptec | ECS_GENERIC_FLOAT | Modbus | 44 | 8 |
| [`MET_EcsOptec_ECS_Generic.csv`](./MET_EcsOptec_ECS_Generic.csv) | Meter | EcsOptec | ECS_Generic | Modbus | 44 | 8 |
| [`OTH_Efacec_Efasolar.csv`](./OTH_Efacec_Efasolar.csv) | Other | Efacec | Efasolar | Modbus | 58 | 18 |
| [`OTH_Efacec_Exemplo_de_Mapeamento_PV500_Reflexos.csv`](./OTH_Efacec_Exemplo_de_Mapeamento_PV500_Reflexos.csv) | Other | Efacec | Exemplo_de_Mapeamento_PV500_Reflexos | Modbus | 27 | 4 |
| [`INV_Efacec_PV250HE_PV375HE_PV500HE.csv`](./INV_Efacec_PV250HE_PV375HE_PV500HE.csv) | Inverter | Efacec | PV250HE_PV375HE_PV500HE | Modbus | 143 | 42 |
| [`OTH_Efacec_PV250HE_PV375HE_PV500HE.csv`](./OTH_Efacec_PV250HE_PV375HE_PV500HE.csv) | Other | Efacec | PV250HE_PV375HE_PV500HE | Modbus | 143 | 34 |
| [`INV_Efacec_PV550_Efacec_MODBUS.csv`](./INV_Efacec_PV550_Efacec_MODBUS.csv) | Inverter | Efacec | PV550_Efacec_MODBUS | Modbus | 145 | 9 |
| [`OTH_Efacec_PV550_Efacec_MODBUS.csv`](./OTH_Efacec_PV550_Efacec_MODBUS.csv) | Other | Efacec | PV550_Efacec_MODBUS | Modbus | 148 | 0 |
| [`OTH_Efacec_ntu-pv500.mbmap-0.2.3.csv`](./OTH_Efacec_ntu-pv500.mbmap-0.2.3.csv) | Other | Efacec | ntu-pv500.mbmap-0.2.3 | Modbus | 156 | 6 |
| [`MET_Eflex4u_OTMETRIC_EFLEX4U.csv`](./MET_Eflex4u_OTMETRIC_EFLEX4U.csv) | Meter | Eflex4u | OTMETRIC_EFLEX4U | Modbus | 405 | 11 |
| [`MET_Elster_A1800.csv`](./MET_Elster_A1800.csv) | Meter | Elster | A1800 | Modbus | 98 | 10 |
| [`MET_Emu_EMUELECTRONIC_EMUPROFESSIONAL.csv`](./MET_Emu_EMUELECTRONIC_EMUPROFESSIONAL.csv) | Meter | Emu | EMUELECTRONIC_EMUPROFESSIONAL | Modbus | 177 | 17 |
| [`MET_Emu_PROFESIONNAL-II.csv`](./MET_Emu_PROFESIONNAL-II.csv) | Meter | Emu | PROFESIONNAL-II | Modbus | 388 | 25 |
| [`MET_Enerfox_FXM-210-XXX.csv`](./MET_Enerfox_FXM-210-XXX.csv) | Meter | Enerfox | FXM-210-XXX | Modbus | 64 | 24 |
| [`INV_FoxESS_RSeries.csv`](./INV_FoxESS_RSeries.csv) | Inverter | FOX ESS | R SERIES | Modbus | 448 | 61 |
| [`BAT_FoxESS_RSeries.csv`](./BAT_FoxESS_RSeries.csv) | Battery | FOX ESS | R SERIES Storage | Modbus | 49 | 3 |
| [`INV_FoxESS_SunSpec_RV_H3Plus.csv`](./INV_FoxESS_SunSpec_RV_H3Plus.csv) | Inverter | FOX_ESS | R_V_H3plus | Modbus | 87 | 11 |
| [`BAT_FoxESS_FOX_ESS_G-MAX_EMS.csv`](./BAT_FoxESS_FOX_ESS_G-MAX_EMS.csv) | Battery | FoxESS | FOX_ESS_G-MAX_EMS | Modbus | 588 | 11 |
| [`INV_FoxESS_FOX_ESS_G-MAX_EMS.csv`](./INV_FoxESS_FOX_ESS_G-MAX_EMS.csv) | Inverter | FoxESS | FOX_ESS_G-MAX_EMS | Modbus | 577 | 11 |
| [`BAT_FoxESS_GMax_Storage.csv`](./BAT_FoxESS_GMax_Storage.csv) | Storage | FoxESS | G-MAX Series | Modbus | 894 | 894 |
| [`COM_FoxESS_G-MAX_Communication_Protocol_modbusTcp-EN-20250425_1.csv`](./COM_FoxESS_G-MAX_Communication_Protocol_modbusTcp-EN-20250425_1.csv) | Logger | FoxESS | G-MAX_Communication_Protocol_modbusTcp-EN-20250425_1 | Modbus | 808 | 37 |
| [`INV_FoxESS_G_MAX_ModbusTCP.csv`](./INV_FoxESS_G_MAX_ModbusTCP.csv) | Inverter | FoxESS | G-MAX_ModbusTCP | Modbus | 246 | 29 |
| [`BAT_FoxESS_G-MAX_OTHER.csv`](./BAT_FoxESS_G-MAX_OTHER.csv) | Battery | FoxESS | G-MAX_OTHER | Modbus | 113 | 24 |
| [`BAT_FoxESS_G-MAX_WebdynSunPM_BMS.csv`](./BAT_FoxESS_G-MAX_WebdynSunPM_BMS.csv) | Battery | FoxESS | G-MAX_WebdynSunPM_BMS | Modbus | 157 | 3 |
| [`INV_FoxESS_G-MAX_WebdynSunPM_BMS.csv`](./INV_FoxESS_G-MAX_WebdynSunPM_BMS.csv) | Inverter | FoxESS | G-MAX_WebdynSunPM_BMS | Modbus | 181 | 3 |
| [`OTH_FoxESS_G-MAX_WebdynSunPM_BMS.csv`](./OTH_FoxESS_G-MAX_WebdynSunPM_BMS.csv) | Other | FoxESS | G-MAX_WebdynSunPM_BMS | Modbus | 181 | 3 |
| [`BAT_FoxESS_H3_Pro.csv`](./BAT_FoxESS_H3_Pro.csv) | Battery | FoxESS | H3 Pro | Modbus | 893 | 18 |
| [`INV_FoxESS_H3_Pro.csv`](./INV_FoxESS_H3_Pro.csv) | Inverter | FoxESS | H3 Pro | Modbus | 285 | 114 |
| [`MET_FoxESS_H3_Pro.csv`](./MET_FoxESS_H3_Pro.csv) | Meter | FoxESS | H3 Pro | Modbus | 61 | 27 |
| [`INV_FoxESS_Modbus_Protocol.csv`](./INV_FoxESS_Modbus_Protocol.csv) | Inverter | FoxESS | Modbus_Protocol | Modbus | 387 | 387 |
| [`BAT_FoxESS_P3_Pro.csv`](./BAT_FoxESS_P3_Pro.csv) | Battery | FoxESS | P3 Pro | Modbus | 893 | 18 |
| [`INV_FoxESS_P3_Pro.csv`](./INV_FoxESS_P3_Pro.csv) | Inverter | FoxESS | P3 Pro | Modbus | 285 | 114 |
| [`INV_FoxESS_P3_H3_Pro_AllInOne.csv`](./INV_FoxESS_P3_H3_Pro_AllInOne.csv) | Inverter | FoxESS | P3 Pro / H3 Pro | Modbus | 1,233 | 127 |
| [`BAT_FoxESS_P3_H3_PRO.csv`](./BAT_FoxESS_P3_H3_PRO.csv) | Battery | FoxESS | P3_H3_PRO | Modbus | 893 | 20 |
| [`INV_FoxESS_P3_H3_PRO.csv`](./INV_FoxESS_P3_H3_PRO.csv) | Inverter | FoxESS | P3_H3_PRO | Modbus | 285 | 115 |
| [`MET_FoxESS_P3_H3_PRO.csv`](./MET_FoxESS_P3_H3_PRO.csv) | Meter | FoxESS | P3_H3_PRO | Modbus | 61 | 26 |
| [`BAT_FoxESS_P3_H3_PRO_BATTERY.csv`](./BAT_FoxESS_P3_H3_PRO_BATTERY.csv) | Battery | FoxESS | P3_H3_PRO_BATTERY | Modbus | 893 | 18 |
| [`INV_FoxESS_P3_H3_PRO_INVERTER.csv`](./INV_FoxESS_P3_H3_PRO_INVERTER.csv) | Inverter | FoxESS | P3_H3_PRO_INVERTER | Modbus | 285 | 114 |
| [`MET_FoxESS_P3_H3_PRO_METER.csv`](./MET_FoxESS_P3_H3_PRO_METER.csv) | Meter | FoxESS | P3_H3_PRO_METER | Modbus | 61 | 27 |
| [`INV_FoxESS_RSERIESINVERTER.csv`](./INV_FoxESS_RSERIESINVERTER.csv) | Inverter | FoxESS | RSERIESINVERTER | Modbus | 448 | 61 |
| [`BAT_FoxESS_RSERIES_BATTERY.csv`](./BAT_FoxESS_RSERIES_BATTERY.csv) | Battery | FoxESS | RSERIES_BATTERY | Modbus | 49 | 3 |
| [`INV_FoxESS_RSeries_Inverter.csv`](./INV_FoxESS_RSeries_Inverter.csv) | Inverter | FoxESS | RSeries_Inverter | Modbus | 445 | 59 |
| [`INV_FoxESS_R_V_H3plus_GmaxPlus.csv`](./INV_FoxESS_R_V_H3plus_GmaxPlus.csv) | Inverter | FoxESS | R_V_H3plus_GmaxPlus | Modbus | 87 | 11 |
| [`OTH_FoxESS_R_V_H3plus_GmaxPlus.csv`](./OTH_FoxESS_R_V_H3plus_GmaxPlus.csv) | Other | FoxESS | R_V_H3plus_GmaxPlus | Modbus | 87 | 11 |
| [`OTH_Fronius_Argeno_Register_Map.csv`](./OTH_Fronius_Argeno_Register_Map.csv) | Other | Fronius | Argeno_Register_Map | Modbus | 1,671 | 35 |
| [`SEN_Fronius_Card.csv`](./SEN_Fronius_Card.csv) | Sensor | Fronius | Card | Modbus | 54 | 6 |
| [`INV_Fronius_ECO_Inverter_Register_Map_Int_SF.csv`](./INV_Fronius_ECO_Inverter_Register_Map_Int_SF.csv) | Inverter | Fronius | ECO_Inverter_Register_Map_Int_SF | Modbus | 257 | 35 |
| [`INV_Fronius_ECO_Series.csv`](./INV_Fronius_ECO_Series.csv) | Inverter | Fronius | ECO_Series | Modbus | 257 | 35 |
| [`INV_Fronius_Gen24_Primo_Symo_Float.csv`](./INV_Fronius_Gen24_Primo_Symo_Float.csv) | Inverter | Fronius | Gen24_Primo_Symo_Float | Modbus | 262 | 35 |
| [`INV_Fronius_Gen24_Primo_Symo_IntSF.csv`](./INV_Fronius_Gen24_Primo_Symo_IntSF.csv) | Inverter | Fronius | Gen24_Primo_Symo_IntSF | Modbus | 269 | 35 |
| [`INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Float.csv`](./INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Float.csv) | Inverter | Fronius | Gen24_Primo_Symo_Inverter_Register_Map_Float | Modbus | 262 | 35 |
| [`INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Float_storage.csv`](./INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Float_storage.csv) | Inverter | Fronius | Gen24_Primo_Symo_Inverter_Register_Map_Float_storage | Modbus | 327 | 35 |
| [`INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Int_SF.csv`](./INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Int_SF.csv) | Inverter | Fronius | Gen24_Primo_Symo_Inverter_Register_Map_Int_SF | Modbus | 269 | 35 |
| [`INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Int_SF_storage.csv`](./INV_Fronius_Gen24_Primo_Symo_Inverter_Register_Map_Int_SF_storage.csv) | Inverter | Fronius | Gen24_Primo_Symo_Inverter_Register_Map_Int_SF_storage | Modbus | 335 | 35 |
| [`SEN_Fronius_IG_Sensor_Card_Register_Map_Float_and_Int_SF_v1.0.csv`](./SEN_Fronius_IG_Sensor_Card_Register_Map_Float_and_Int_SF_v1.0.csv) | Sensor | Fronius | IG_Sensor_Card_Register_Map_Float_and_Int_SF_v1.0 | Modbus | 54 | 6 |
| [`OTH_Fronius_IG_String_Control_Register_Map_Float_and_Int_SF_v1.0.csv`](./OTH_Fronius_IG_String_Control_Register_Map_Float_and_Int_SF_v1.0.csv) | Other | Fronius | IG_String_Control_Register_Map_Float_and_Int_SF_v1.0 | Modbus | 95 | 5 |
| [`INV_Fronius_Modbus_SunS_Fronius_IntSF.csv`](./INV_Fronius_Modbus_SunS_Fronius_IntSF.csv) | Inverter | Fronius | Modbus_SunS_Fronius_IntSF | Modbus | 55 | 24 |
| [`OTH_Fronius_Modbus_SunS_Fronius_IntSF.csv`](./OTH_Fronius_Modbus_SunS_Fronius_IntSF.csv) | Other | Fronius | Modbus_SunS_Fronius_IntSF | Modbus | 55 | 24 |
| [`MET_Fronius_Register_Map_Float_v1.0.csv`](./MET_Fronius_Register_Map_Float_v1.0.csv) | Meter | Fronius | Register_Map_Float_v1.0 | Modbus | 145 | 34 |
| [`INV_Fronius_Register_Map_Float_v1.0_with_SYMOHYBRID_MODEL_124.csv`](./INV_Fronius_Register_Map_Float_v1.0_with_SYMOHYBRID_MODEL_124.csv) | Inverter | Fronius | Register_Map_Float_v1.0_with_SYMOHYBRID_MODEL_124 | Modbus | 309 | 37 |
| [`MET_Fronius_Register_Map_Int_SF_v1.0.csv`](./MET_Fronius_Register_Map_Int_SF_v1.0.csv) | Meter | Fronius | Register_Map_Int_SF_v1.0 | Modbus | 165 | 34 |
| [`INV_Fronius_Register_Map_Int_SF_v1.0_with_SYMOHYBRID_MODEL_124.csv`](./INV_Fronius_Register_Map_Int_SF_v1.0_with_SYMOHYBRID_MODEL_124.csv) | Inverter | Fronius | Register_Map_Int_SF_v1.0_with_SYMOHYBRID_MODEL_124 | Modbus | 316 | 37 |
| [`SEN_Fronius_Sensor_Card.csv`](./SEN_Fronius_Sensor_Card.csv) | Sensor | Fronius | Sensor_Card | Modbus | 54 | 6 |
| [`MET_Fronius_Smart_Meter_50kA-3.csv`](./MET_Fronius_Smart_Meter_50kA-3.csv) | Meter | Fronius | Smart_Meter_50kA | Modbus | 72 | 34 |
| [`MET_Fronius_Smart_Meter_IntSF.csv`](./MET_Fronius_Smart_Meter_IntSF.csv) | Meter | Fronius | Smart_Meter_IntSF | Modbus | 165 | 31 |
| [`MET_Fronius_Smart_Meter_Register_Map_Float.csv`](./MET_Fronius_Smart_Meter_Register_Map_Float.csv) | Meter | Fronius | Smart_Meter_Register_Map_Float | Modbus | 144 | 34 |
| [`MET_Fronius_Smart_Meter_Register_Map_Int_SF.csv`](./MET_Fronius_Smart_Meter_Register_Map_Int_SF.csv) | Meter | Fronius | Smart_Meter_Register_Map_Int_SF | Modbus | 165 | 31 |
| [`MET_Fronius_Smart_Meter_TS.csv`](./MET_Fronius_Smart_Meter_TS.csv) | Meter | Fronius | Smart_Meter_TS | Modbus | 144 | 34 |
| [`INV_Fronius_SunSpec_inverter_Fronius_Symo_GEN24.csv`](./INV_Fronius_SunSpec_inverter_Fronius_Symo_GEN24.csv) | Inverter | Fronius | SunSpec_inverter_Fronius_Symo_GEN24 | Modbus | 223 | 35 |
| [`INV_Fronius_Sunspec_IntSF.csv`](./INV_Fronius_Sunspec_IntSF.csv) | Inverter | Fronius | Sunspec_IntSF | Modbus | 55 | 24 |
| [`OTH_Fronius_Sunspec_IntSF.csv`](./OTH_Fronius_Sunspec_IntSF.csv) | Other | Fronius | Sunspec_IntSF | Modbus | 55 | 24 |
| [`INV_Fronius_Symo_Primo_Gen24.csv`](./INV_Fronius_Symo_Primo_Gen24.csv) | Inverter | Fronius | Symo-Primo-Gen24 | Modbus | 26 | 26 |
| [`INV_Fronius_Symo_10_20.csv`](./INV_Fronius_Symo_10_20.csv) | Inverter | Fronius | Symo_10_20 | Modbus | 177 | 43 |
| [`INV_Fronius_Symo_GEN24_Plus.csv`](./INV_Fronius_Symo_GEN24_Plus.csv) | Inverter | Fronius | Symo_GEN24_Plus | Modbus | 223 | 35 |
| [`INV_Fronius_Tauro_50-3-D.csv`](./INV_Fronius_Tauro_50-3-D.csv) | Inverter | Fronius | Tauro_50-3-D | Modbus | 193 | 34 |
| [`OTH_Fronius_Tauro_50-3-D_1.35.4-1_Float.csv`](./OTH_Fronius_Tauro_50-3-D_1.35.4-1_Float.csv) | Other | Fronius | Tauro_50-3-D_1.35.4-1_Float | Modbus | 193 | 34 |
| [`INV_Fronius_Tauro_ECO_100-3-D.csv`](./INV_Fronius_Tauro_ECO_100-3-D.csv) | Inverter | Fronius | Tauro_ECO_100-3-D | Modbus | 167 | 40 |
| [`INV_Fronius_Tauro_ECO_Inverter_Register_Map_Float.csv`](./INV_Fronius_Tauro_ECO_Inverter_Register_Map_Float.csv) | Inverter | Fronius | Tauro_ECO_Inverter_Register_Map_Float | Modbus | 249 | 35 |
| [`INV_Fronius_Tauro_ECO_Inverter_Register_Map_Int_SF.csv`](./INV_Fronius_Tauro_ECO_Inverter_Register_Map_Int_SF.csv) | Inverter | Fronius | Tauro_ECO_Inverter_Register_Map_Int_SF | Modbus | 257 | 35 |
| [`INV_Fronius_Tauro_Inverter_Register_Map_Float.csv`](./INV_Fronius_Tauro_Inverter_Register_Map_Float.csv) | Inverter | Fronius | Tauro_Inverter_Register_Map_Float | Modbus | 279 | 35 |
| [`INV_Fronius_Tauro_Inverter_Register_Map_Int_SF.csv`](./INV_Fronius_Tauro_Inverter_Register_Map_Int_SF.csv) | Inverter | Fronius | Tauro_Inverter_Register_Map_Int_SF | Modbus | 287 | 35 |
| [`INV_Fronius_Verto_3MPPT.csv`](./INV_Fronius_Verto_3MPPT.csv) | Inverter | Fronius | Verto_3MPPT | Modbus | 279 | 35 |
| [`INV_Fronius_Verto_4MPPT.csv`](./INV_Fronius_Verto_4MPPT.csv) | Inverter | Fronius | Verto_4MPPT | Modbus | 299 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float.csv`](./INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_3MPPT_Float | Modbus | 279 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float_storage.csv`](./INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float_storage.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_3MPPT_Float_storage | Modbus | 347 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Int_SF.csv`](./INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Int_SF.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_3MPPT_Int_SF | Modbus | 287 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Int_SF_storage.csv`](./INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Int_SF_storage.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_3MPPT_Int_SF_storage | Modbus | 355 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Float.csv`](./INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Float.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_4MPPT_Float | Modbus | 299 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Float_storage.csv`](./INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Float_storage.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_4MPPT_Float_storage | Modbus | 367 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Int_SF.csv`](./INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Int_SF.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_4MPPT_Int_SF | Modbus | 307 | 35 |
| [`INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Int_SF_storage.csv`](./INV_Fronius_Verto_Inverter_Register_Map_4MPPT_Int_SF_storage.csv) | Inverter | Fronius | Verto_Inverter_Register_Map_4MPPT_Int_SF_storage | Modbus | 375 | 35 |
| [`INV_Greenheiss_GH-IT_3-4-9M.csv`](./INV_Greenheiss_GH-IT_3-4-9M.csv) | Inverter | GREENHEISS | GH-IT_3-4-9M | Modbus | 385 | 10 |
| [`INV_Greenheiss_GH-I_Style_GH-IT_2M.csv`](./INV_Greenheiss_GH-I_Style_GH-IT_2M.csv) | Inverter | GREENHEISS | GH-I_Style_GH-IT_2M | Modbus | 219 | 8 |
| [`INV_Gamesa_Inverter100KW.csv`](./INV_Gamesa_Inverter100KW.csv) | Inverter | Gamesa | Inverter100KW | Modbus | 17 | 0 |
| [`INV_GoodWe_ES_EM_S-BP_BP_Series.csv`](./INV_GoodWe_ES_EM_S-BP_BP_Series.csv) | Inverter | GoodWe | ES-EM-S-BP-BP-Series | Modbus | 370 | 14 |
| [`INV_Goodwe_ES_EM_S-BP_BP_Series.csv`](./INV_Goodwe_ES_EM_S-BP_BP_Series.csv) | Inverter | GoodWe | ES-EM-S-BP-BP-Series | Modbus | 370 | 14 |
| [`INV_GoodWe_ES-EM-SBP-BP.csv`](./INV_GoodWe_ES-EM-SBP-BP.csv) | Inverter | GoodWe | ES-EM-SBP-BP | Modbus | 49 | 49 |
| [`INV_Goodwe_ES-EM-SBP-BP.csv`](./INV_Goodwe_ES-EM-SBP-BP.csv) | Inverter | GoodWe | ES-EM-SBP-BP | Modbus | 49 | 49 |
| [`INV_Goodwe_ES_EM_SBP.csv`](./INV_Goodwe_ES_EM_SBP.csv) | Inverter | GoodWe | ES-EM-SBP-Series | Modbus | 49 | 49 |
| [`INV_GoodWe_ES_EM_SBP.csv`](./INV_GoodWe_ES_EM_SBP.csv) | Inverter | GoodWe | ES_EM_SBP | Modbus | 49 | 49 |
| [`INV_GoodWe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv`](./INV_GoodWe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv) | Inverter | GoodWe | ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU | Modbus | 79 | 79 |
| [`INV_Goodwe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv`](./INV_Goodwe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv) | Inverter | GoodWe | ET-EH-BT-BH-EHB-Series | Modbus | 530 | 16 |
| [`INV_GoodWe_ET_EH_Hybrid_Series.csv`](./INV_GoodWe_ET_EH_Hybrid_Series.csv) | Inverter | GoodWe | ET-EH-Hybrid-Series | Modbus | 79 | 79 |
| [`INV_Goodwe_ET_EH_Hybrid_Series.csv`](./INV_Goodwe_ET_EH_Hybrid_Series.csv) | Inverter | GoodWe | ET-EH-Hybrid-Series | Modbus | 79 | 79 |
| [`INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv`](./INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv) | Inverter | GoodWe | ET-EH-ModbusRTU | Modbus | 79 | 79 |
| [`INV_GoodWe_ET_EH_Series.csv`](./INV_GoodWe_ET_EH_Series.csv) | Inverter | GoodWe | ET-EH-Series | Modbus | 79 | 79 |
| [`INV_Goodwe_ET_EH_Series.csv`](./INV_Goodwe_ET_EH_Series.csv) | Inverter | GoodWe | ET-EH-Series | Modbus | 79 | 79 |
| [`INV_GoodWe_ET_Series.csv`](./INV_GoodWe_ET_Series.csv) | Inverter | GoodWe | ET-Series | Modbus | 79 | 79 |
| [`INV_Goodwe_ET_Series.csv`](./INV_Goodwe_ET_Series.csv) | Inverter | GoodWe | ET-Series | Modbus | 79 | 79 |
| [`INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv`](./INV_GoodWe_ET_EH_BT_BH_EHB_AES_ABP_BTC_Series.csv) | Inverter | GoodWe | ET_EH_BT_BH_EHB_AES_ABP_BTC_Series | Modbus | 530 | 16 |
| [`COM_GoodWe_EzLogger_Pro.csv`](./COM_GoodWe_EzLogger_Pro.csv) | Logger | GoodWe | EzLogger-Pro | Modbus | 25 | 24 |
| [`COM_Goodwe_EzLogger_Pro.csv`](./COM_Goodwe_EzLogger_Pro.csv) | Logger | GoodWe | EzLogger-Pro | Modbus | 25 | 24 |
| [`COM_GoodWe_Ezlogger3000.csv`](./COM_GoodWe_Ezlogger3000.csv) | Logger | GoodWe | Ezlogger3000 | Modbus | 25 | 24 |
| [`COM_Goodwe_Ezlogger3000.csv`](./COM_Goodwe_Ezlogger3000.csv) | Logger | GoodWe | Ezlogger3000 | Modbus | 25 | 24 |
| [`INV_GoodWe_G3_Series.csv`](./INV_GoodWe_G3_Series.csv) | Inverter | GoodWe | G3_Series | Modbus | 120 | 40 |
| [`MET_GoodWe_GM1000_GM3000.csv`](./MET_GoodWe_GM1000_GM3000.csv) | Meter | GoodWe | GM1000-GM3000 | Modbus | 18 | 18 |
| [`MET_Goodwe_GM1000_GM3000.csv`](./MET_Goodwe_GM1000_GM3000.csv) | Meter | GoodWe | GM1000-GM3000 | Modbus | 18 | 18 |
| [`INV_GoodWe_GT-SERIES.csv`](./INV_GoodWe_GT-SERIES.csv) | Inverter | GoodWe | GT-SERIES | Modbus | 91 | 90 |
| [`INV_Goodwe_GT-SERIES.csv`](./INV_Goodwe_GT-SERIES.csv) | Inverter | GoodWe | GT-SERIES | Modbus | 91 | 90 |
| [`INV_GoodWe_HT_others_Series.csv`](./INV_GoodWe_HT_others_Series.csv) | Inverter | GoodWe | HT-Others-Series | Modbus | 91 | 90 |
| [`INV_Goodwe_HT_others_Series.csv`](./INV_Goodwe_HT_others_Series.csv) | Inverter | GoodWe | HT-Others-Series | Modbus | 91 | 90 |
| [`INV_GoodWe_HT_Series.csv`](./INV_GoodWe_HT_Series.csv) | Inverter | GoodWe | HT-Series | Modbus | 294 | 77 |
| [`INV_Goodwe_HT_Series.csv`](./INV_Goodwe_HT_Series.csv) | Inverter | GoodWe | HT-Series | Modbus | 294 | 77 |
| [`INV_GoodWe_HT1100.csv`](./INV_GoodWe_HT1100.csv) | Inverter | GoodWe | HT1100 | Modbus | 277 | 76 |
| [`INV_Goodwe_HT1100.csv`](./INV_Goodwe_HT1100.csv) | Inverter | GoodWe | HT1100 | Modbus | 277 | 76 |
| [`INV_GoodWe_HT1500.csv`](./INV_GoodWe_HT1500.csv) | Inverter | GoodWe | HT1500 | Modbus | 288 | 77 |
| [`INV_Goodwe_HT1500.csv`](./INV_Goodwe_HT1500.csv) | Inverter | GoodWe | HT1500 | Modbus | 288 | 77 |
| [`BAT_GoodWe_Lynx_Home_HV.csv`](./BAT_GoodWe_Lynx_Home_HV.csv) | Battery | GoodWe | Lynx-Home-HV | Modbus | 18 | 18 |
| [`BAT_Goodwe_Lynx_Home_HV.csv`](./BAT_Goodwe_Lynx_Home_HV.csv) | Battery | GoodWe | Lynx-Home-HV | Modbus | 18 | 18 |
| [`INV_GoodWe_MS_DNS_XS.csv`](./INV_GoodWe_MS_DNS_XS.csv) | Inverter | GoodWe | MS-DNS-XS | Modbus | 50 | 37 |
| [`INV_Goodwe_MS_DNS_XS.csv`](./INV_Goodwe_MS_DNS_XS.csv) | Inverter | GoodWe | MS-DNS-XS | Modbus | 50 | 37 |
| [`INV_GoodWe_MT.csv`](./INV_GoodWe_MT.csv) | Inverter | GoodWe | MT | Modbus | 95 | 46 |
| [`INV_Goodwe_MT.csv`](./INV_Goodwe_MT.csv) | Inverter | GoodWe | MT | Modbus | 95 | 46 |
| [`INV_GoodWe_MT_series.csv`](./INV_GoodWe_MT_series.csv) | Inverter | GoodWe | MT-Series | Modbus | 68 | 55 |
| [`INV_Goodwe_MT_Series.csv`](./INV_Goodwe_MT_Series.csv) | Inverter | GoodWe | MT-Series | Modbus | 68 | 55 |
| [`INV_GoodWe_MT-Series_ModbusRTU.csv`](./INV_GoodWe_MT-Series_ModbusRTU.csv) | Inverter | GoodWe | MT-Series-ModbusRTU | Modbus | 68 | 55 |
| [`INV_Goodwe_MT-Series_ModbusRTU.csv`](./INV_Goodwe_MT-Series_ModbusRTU.csv) | Inverter | GoodWe | MT-Series-ModbusRTU | Modbus | 68 | 55 |
| [`COM_GoodWe_SCB3000.csv`](./COM_GoodWe_SCB3000.csv) | Logger | GoodWe | SCB3000 | Modbus | 29 | 28 |
| [`COM_Goodwe_SCB3000.csv`](./COM_Goodwe_SCB3000.csv) | Logger | GoodWe | SCB3000 | Modbus | 29 | 28 |
| [`INV_Goodwe_SDT_G2_G3.csv`](./INV_Goodwe_SDT_G2_G3.csv) | Inverter | GoodWe | SDT-G2-G3-Series | Modbus | 60 | 47 |
| [`INV_GoodWe_SDTG2.csv`](./INV_GoodWe_SDTG2.csv) | Inverter | GoodWe | SDTG2 | Modbus | 116 | 49 |
| [`INV_Goodwe_SDTG2.csv`](./INV_Goodwe_SDTG2.csv) | Inverter | GoodWe | SDTG2 | Modbus | 116 | 49 |
| [`INV_GoodWe_SDT_G2_G3.csv`](./INV_GoodWe_SDT_G2_G3.csv) | Inverter | GoodWe | SDT_G2_G3 | Modbus | 60 | 47 |
| [`INV_GoodWe_SMT.csv`](./INV_GoodWe_SMT.csv) | Inverter | GoodWe | SMT | Modbus | 115 | 49 |
| [`INV_Goodwe_SMT.csv`](./INV_Goodwe_SMT.csv) | Inverter | GoodWe | SMT | Modbus | 115 | 49 |
| [`INV_GoodWe_SMT_Series.csv`](./INV_GoodWe_SMT_Series.csv) | Inverter | GoodWe | SMT-Series | Modbus | 68 | 55 |
| [`INV_Goodwe_SMT_Series.csv`](./INV_Goodwe_SMT_Series.csv) | Inverter | GoodWe | SMT-Series | Modbus | 68 | 55 |
| [`INV_GoodWe_SMT50K_SMT60K.csv`](./INV_GoodWe_SMT50K_SMT60K.csv) | Inverter | GoodWe | SMT50K-SMT60K | Modbus | 119 | 53 |
| [`INV_Goodwe_SMT50K_SMT60K.csv`](./INV_Goodwe_SMT50K_SMT60K.csv) | Inverter | GoodWe | SMT50K-SMT60K | Modbus | 119 | 53 |
| [`INV_GoodWe_ThreePhase_SunSpec.csv`](./INV_GoodWe_ThreePhase_SunSpec.csv) | Inverter | GoodWe | ThreePhase-SunSpec | Modbus | 29 | 25 |
| [`INV_Goodwe_ThreePhase_SunSpec.csv`](./INV_Goodwe_ThreePhase_SunSpec.csv) | Inverter | GoodWe | ThreePhase-SunSpec | Modbus | 29 | 25 |
| [`INV_GoodWe_XS-DNS-MS-SDT-G3_SERIES.csv`](./INV_GoodWe_XS-DNS-MS-SDT-G3_SERIES.csv) | Inverter | GoodWe | XS-DNS-MS-SDT-G3-Series | Modbus | 50 | 37 |
| [`INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv`](./INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv) | Inverter | GoodWe | XS-DNS-MS-SDT-G3-Series | Modbus | 50 | 37 |
| [`INV_Goodwe_XS_DNS_MS.csv`](./INV_Goodwe_XS_DNS_MS.csv) | Inverter | GoodWe | XS-DNS-MS-Series | Modbus | 50 | 37 |
| [`INV_GoodWe_XS_DNS_MS.csv`](./INV_GoodWe_XS_DNS_MS.csv) | Inverter | GoodWe | XS_DNS_MS | Modbus | 50 | 37 |
| [`INV_GoodWe_Modbus.csv`](./INV_GoodWe_Modbus.csv) | Inverter | Goodwe | Modbus | Modbus | 36 | 18 |
| [`OTH_Greenheiss_GH-IT_3-4-9M.csv`](./OTH_Greenheiss_GH-IT_3-4-9M.csv) | Other | Greenheiss | GH-IT_3-4-9M | Modbus | 385 | 10 |
| [`COM_Greenheiss_GH-IT_3-4-9M_Modbus_Communication_Protocol.csv`](./COM_Greenheiss_GH-IT_3-4-9M_Modbus_Communication_Protocol.csv) | Logger | Greenheiss | GH-IT_3-4-9M_Modbus_Communication_Protocol | Modbus | 69 | 8 |
| [`OTH_Greenheiss_GH-I_Style_GH-IT_2M.csv`](./OTH_Greenheiss_GH-I_Style_GH-IT_2M.csv) | Other | Greenheiss | GH-I_Style_GH-IT_2M | Modbus | 219 | 8 |
| [`COM_Greenheiss_GH-I_Style_GH-IT_2M_Modbus_Communication_Protocol.csv`](./COM_Greenheiss_GH-I_Style_GH-IT_2M_Modbus_Communication_Protocol.csv) | Logger | Greenheiss | GH-I_Style_GH-IT_2M_Modbus_Communication_Protocol | Modbus | 408 | 7 |
| [`INV_Growatt_125kw.csv`](./INV_Growatt_125kw.csv) | Inverter | Growatt | 125kw | Modbus | 156 | 91 |
| [`INV_Growatt_Generic.csv`](./INV_Growatt_Generic.csv) | Inverter | Growatt | Generic | Modbus | 851 | 851 |
| [`INV_Growatt_MAX_Series.csv`](./INV_Growatt_MAX_Series.csv) | Inverter | Growatt | MAX-Series | Modbus | 29 | 29 |
| [`INV_Growatt_MAX1500V_MAX-XLV_Series.csv`](./INV_Growatt_MAX1500V_MAX-XLV_Series.csv) | Inverter | Growatt | MAX1500V_MAX-XLV_Series | Modbus | 468 | 151 |
| [`INV_Growatt_MOD_TL3-XH_Series.csv`](./INV_Growatt_MOD_TL3-XH_Series.csv) | Inverter | Growatt | MOD_TL3-XH_Series | Modbus | 409 | 53 |
| [`INV_Growatt_Modbus.csv`](./INV_Growatt_Modbus.csv) | Inverter | Growatt | Modbus | Modbus | 25 | 14 |
| [`INV_Growatt_PM_5.csv`](./INV_Growatt_PM_5.csv) | Inverter | Growatt | PM_5 | Modbus | 203 | 99 |
| [`INV_Growatt_PV_Inverter_V3.14.csv`](./INV_Growatt_PV_Inverter_V3.14.csv) | Inverter | Growatt | PV_Inverter_V3.14 | Modbus | 136 | 12 |
| [`INV_Growatt_RTU_Protocol_II.csv`](./INV_Growatt_RTU_Protocol_II.csv) | Inverter | Growatt | RTU_Protocol_II | Modbus | 49 | 4 |
| [`INV_Growatt_TL-X_TL-XH_Series.csv`](./INV_Growatt_TL-X_TL-XH_Series.csv) | Inverter | Growatt | TL-X_TL-XH_Series | Modbus | 425 | 53 |
| [`INV_Growatt_TL3-X_Series.csv`](./INV_Growatt_TL3-X_Series.csv) | Inverter | Growatt | TL3-X_Series | Modbus | 371 | 95 |
| [`INV_Growatt_W_Growatt_Generic.csv`](./INV_Growatt_W_Growatt_Generic.csv) | Inverter | Growatt | W_Growatt_Generic | Modbus | 852 | 124 |
| [`INV_Growatt_W_Growatt_MAX1500V_MAX-XLV_Series.csv`](./INV_Growatt_W_Growatt_MAX1500V_MAX-XLV_Series.csv) | Inverter | Growatt | W_Growatt_MAX1500V_MAX-XLV_Series | Modbus | 466 | 151 |
| [`INV_Growatt_W_Growatt_MOD_TL3-XH_Series.csv`](./INV_Growatt_W_Growatt_MOD_TL3-XH_Series.csv) | Inverter | Growatt | W_Growatt_MOD_TL3-XH_Series | Modbus | 410 | 53 |
| [`INV_Growatt_W_Growatt_TL-X_TL-XH_Series.csv`](./INV_Growatt_W_Growatt_TL-X_TL-XH_Series.csv) | Inverter | Growatt | W_Growatt_TL-X_TL-XH_Series | Modbus | 426 | 53 |
| [`INV_Growatt_W_Growatt_TL3-X_Series.csv`](./INV_Growatt_W_Growatt_TL3-X_Series.csv) | Inverter | Growatt | W_Growatt_TL3-X_Series | Modbus | 371 | 95 |
| [`MET_Hager_Generic.csv`](./MET_Hager_Generic.csv) | Meter | HAGER | Generic | Modbus | 102 | 4 |
| [`OTH_Hager_Generic.csv`](./OTH_Hager_Generic.csv) | Other | HAGER | Generic | Modbus | 102 | 10 |
| [`COM_Hoymiles_GENERIC.csv`](./COM_Hoymiles_GENERIC.csv) | Logger | HOYMILES | GENERIC | Modbus | 4,963 | 36 |
| [`COM_Hoymiles_METER.csv`](./COM_Hoymiles_METER.csv) | Logger | HOYMILES | METER | Modbus | 291 | 18 |
| [`COM_Hoymiles_MICROINVERTER.csv`](./COM_Hoymiles_MICROINVERTER.csv) | Logger | HOYMILES | MICROINVERTER | Modbus | 4,264 | 24 |
| [`INV_Huawei_364233-A_noBF.csv`](./INV_Huawei_364233-A_noBF.csv) | Inverter | HUAWEI | 364233-A_noBF | Modbus | 192 | 34 |
| [`INV_Huawei_7550606570-M0.csv`](./INV_Huawei_7550606570-M0.csv) | Inverter | HUAWEI | 7550606570-M0 | Modbus | 73 | 37 |
| [`INV_Huawei_81012141820.csv`](./INV_Huawei_81012141820.csv) | Inverter | HUAWEI | 81012141820 | Modbus | 152 | 25 |
| [`BAT_Huawei_INV_STORAGE.csv`](./BAT_Huawei_INV_STORAGE.csv) | Battery | HUAWEI | INV_STORAGE | Modbus | 272 | 13 |
| [`BAT_Huawei_Luna_2000.csv`](./BAT_Huawei_Luna_2000.csv) | Battery | HUAWEI | Luna_2000 | Modbus | 170 | 25 |
| [`MET_Huawei_Power_Meter.csv`](./MET_Huawei_Power_Meter.csv) | Meter | HUAWEI | Power_Meter | Modbus | 44 | 4 |
| [`INV_Huawei_V4_BATTERY.csv`](./INV_Huawei_V4_BATTERY.csv) | Inverter | HUAWEI | V4_BATTERY | Modbus | 140 | 9 |
| [`INV_Huawei_V4_METER.csv`](./INV_Huawei_V4_METER.csv) | Inverter | HUAWEI | V4_METER | Modbus | 22 | 8 |
| [`INV_Hitachi_50_40.csv`](./INV_Hitachi_50_40.csv) | Inverter | Hitachi | 50_40 | Modbus | 110 | 28 |
| [`OTH_Hitachi_50_40.csv`](./OTH_Hitachi_50_40.csv) | Other | Hitachi | 50_40 | Modbus | 110 | 28 |
| [`OTH_Hopewind_Generic.csv`](./OTH_Hopewind_Generic.csv) | Other | Hopewind | Generic | Modbus | 78 | 68 |
| [`OTH_Hoymiles_GENERIC.csv`](./OTH_Hoymiles_GENERIC.csv) | Other | Hoymiles | GENERIC | Modbus | 4,963 | 36 |
| [`BAT_Hoymiles_HoyUltra2.csv`](./BAT_Hoymiles_HoyUltra2.csv) | Battery | Hoymiles | HoyUltra2 | Modbus | 127 | 6 |
| [`MET_Hoymiles_METER.csv`](./MET_Hoymiles_METER.csv) | Meter | Hoymiles | METER | Modbus | 291 | 18 |
| [`INV_Hoymiles_MICROINVERTER.csv`](./INV_Hoymiles_MICROINVERTER.csv) | Inverter | Hoymiles | MICROINVERTER | Modbus | 4,264 | 24 |
| [`OTH_Huawei_2345.csv`](./OTH_Huawei_2345.csv) | Other | Huawei | 2345 | Modbus | 47 | 22 |
| [`INV_Huawei_364233-A.csv`](./INV_Huawei_364233-A.csv) | Inverter | Huawei | 364233-A | Modbus | 192 | 35 |
| [`OTH_Huawei_364233-A.csv`](./OTH_Huawei_364233-A.csv) | Other | Huawei | 364233-A | Modbus | 192 | 35 |
| [`OTH_Huawei_364233-A_noBF.csv`](./OTH_Huawei_364233-A_noBF.csv) | Other | Huawei | 364233-A_noBF | Modbus | 192 | 34 |
| [`OTH_Huawei_60-HV.csv`](./OTH_Huawei_60-HV.csv) | Other | Huawei | 60-HV | Modbus | 343 | 29 |
| [`OTH_Huawei_7550606570-M0.csv`](./OTH_Huawei_7550606570-M0.csv) | Other | Huawei | 7550606570-M0 | Modbus | 73 | 37 |
| [`OTH_Huawei_81012141820.csv`](./OTH_Huawei_81012141820.csv) | Other | Huawei | 81012141820 | Modbus | 152 | 25 |
| [`OTH_Huawei_9095100105.csv`](./OTH_Huawei_9095100105.csv) | Other | Huawei | 9095100105 | Modbus | 73 | 37 |
| [`COM_Huawei_DEF_HUAWEI_SmartLogger_VSeries.csv`](./COM_Huawei_DEF_HUAWEI_SmartLogger_VSeries.csv) | Logger | Huawei | DEF_HUAWEI_SmartLogger_VSeries | Modbus | 176 | 20 |
| [`MET_Huawei_DTSU666-H.csv`](./MET_Huawei_DTSU666-H.csv) | Meter | Huawei | DTSU666-H | Modbus | 22 | 22 |
| [`OTH_Huawei_Etude_Huawei.csv`](./OTH_Huawei_Etude_Huawei.csv) | Other | Huawei | Etude_Huawei | Modbus | 182 | 182 |
| [`INV_Huawei_GENERIC.csv`](./INV_Huawei_GENERIC.csv) | Inverter | Huawei | GENERIC | Modbus | 256 | 80 |
| [`INV_Huawei_INV.csv`](./INV_Huawei_INV.csv) | Inverter | Huawei | INV | Modbus | 256 | 80 |
| [`MET_Huawei_INV_METER.csv`](./MET_Huawei_INV_METER.csv) | Meter | Huawei | INV_METER | Modbus | 22 | 9 |
| [`INV_Huawei_Inverter.csv`](./INV_Huawei_Inverter.csv) | Inverter | Huawei | Inverter | Modbus | 94 | 60 |
| [`BAT_Huawei_LUNA2000.csv`](./BAT_Huawei_LUNA2000.csv) | Battery | Huawei | LUNA2000 | Modbus | 170 | 25 |
| [`BAT_Huawei_LUNA2000_Storage.csv`](./BAT_Huawei_LUNA2000_Storage.csv) | Battery | Huawei | LUNA2000_Storage | Modbus | 272 | 13 |
| [`OTH_Huawei_Luna_2000.csv`](./OTH_Huawei_Luna_2000.csv) | Other | Huawei | Luna_2000 | Modbus | 170 | 25 |
| [`INV_Huawei_MODBUS_2345.csv`](./INV_Huawei_MODBUS_2345.csv) | Inverter | Huawei | MODBUS_2345 | Modbus | 47 | 22 |
| [`OTH_Huawei_MODBUS_2345.csv`](./OTH_Huawei_MODBUS_2345.csv) | Other | Huawei | MODBUS_2345 | Modbus | 47 | 22 |
| [`INV_Huawei_MODBUS_60-HV.csv`](./INV_Huawei_MODBUS_60-HV.csv) | Inverter | Huawei | MODBUS_60-HV | Modbus | 347 | 30 |
| [`OTH_Huawei_MODBUS_60-HV.csv`](./OTH_Huawei_MODBUS_60-HV.csv) | Other | Huawei | MODBUS_60-HV | Modbus | 347 | 30 |
| [`INV_Huawei_MODBUS_7550606570-M0.csv`](./INV_Huawei_MODBUS_7550606570-M0.csv) | Inverter | Huawei | MODBUS_7550606570-M0 | Modbus | 73 | 37 |
| [`OTH_Huawei_MODBUS_7550606570-M0.csv`](./OTH_Huawei_MODBUS_7550606570-M0.csv) | Other | Huawei | MODBUS_7550606570-M0 | Modbus | 73 | 37 |
| [`INV_Huawei_MODBUS_81012141820.csv`](./INV_Huawei_MODBUS_81012141820.csv) | Inverter | Huawei | MODBUS_81012141820 | Modbus | 328 | 25 |
| [`OTH_Huawei_MODBUS_81012141820.csv`](./OTH_Huawei_MODBUS_81012141820.csv) | Other | Huawei | MODBUS_81012141820 | Modbus | 328 | 25 |
| [`INV_Huawei_MODBUS_9095100105.csv`](./INV_Huawei_MODBUS_9095100105.csv) | Inverter | Huawei | MODBUS_9095100105 | Modbus | 73 | 37 |
| [`OTH_Huawei_MODBUS_9095100105.csv`](./OTH_Huawei_MODBUS_9095100105.csv) | Other | Huawei | MODBUS_9095100105 | Modbus | 73 | 37 |
| [`INV_Huawei_ModbusV3.csv`](./INV_Huawei_ModbusV3.csv) | Inverter | Huawei | ModbusV3 | Modbus | 97 | 15 |
| [`OTH_Huawei_ModbusV3.csv`](./OTH_Huawei_ModbusV3.csv) | Other | Huawei | ModbusV3 | Modbus | 97 | 15 |
| [`INV_Huawei_ModbusV3_small.csv`](./INV_Huawei_ModbusV3_small.csv) | Inverter | Huawei | ModbusV3_small | Modbus | 97 | 15 |
| [`INV_Huawei_ModbusV4.csv`](./INV_Huawei_ModbusV4.csv) | Inverter | Huawei | ModbusV4 | Modbus | 322 | 20 |
| [`OTH_Huawei_ModbusV4.csv`](./OTH_Huawei_ModbusV4.csv) | Other | Huawei | ModbusV4 | Modbus | 322 | 20 |
| [`INV_Huawei_Modbus_364233-A.csv`](./INV_Huawei_Modbus_364233-A.csv) | Inverter | Huawei | Modbus_364233-A | Modbus | 192 | 34 |
| [`OTH_Huawei_Modbus_364233-A.csv`](./OTH_Huawei_Modbus_364233-A.csv) | Other | Huawei | Modbus_364233-A | Modbus | 192 | 34 |
| [`INV_Huawei_Modbus_All_V3.csv`](./INV_Huawei_Modbus_All_V3.csv) | Inverter | Huawei | Modbus_All_V3 | Modbus | 97 | 15 |
| [`OTH_Huawei_Modbus_All_V3.csv`](./OTH_Huawei_Modbus_All_V3.csv) | Other | Huawei | Modbus_All_V3 | Modbus | 97 | 15 |
| [`MET_Huawei_PowerMeter.csv`](./MET_Huawei_PowerMeter.csv) | Meter | Huawei | PowerMeter | Modbus | 44 | 13 |
| [`COM_Huawei_SL3000.csv`](./COM_Huawei_SL3000.csv) | Logger | Huawei | SL3000 | Modbus | 82 | 7 |
| [`INV_Huawei_SUN2000-2_5KTL.csv`](./INV_Huawei_SUN2000-2_5KTL.csv) | Inverter | Huawei | SUN2000-2_5KTL | Modbus | 47 | 22 |
| [`INV_Huawei_SUN2000-30_40KTL-M3.csv`](./INV_Huawei_SUN2000-30_40KTL-M3.csv) | Inverter | Huawei | SUN2000-30_40KTL-M3 | Modbus | 64 | 64 |
| [`INV_Huawei_SUN2000-33A_36_42KTL.csv`](./INV_Huawei_SUN2000-33A_36_42KTL.csv) | Inverter | Huawei | SUN2000-33A_36_42KTL | Modbus | 192 | 35 |
| [`INV_Huawei_SUN2000-50_70KTL-M0.csv`](./INV_Huawei_SUN2000-50_70KTL-M0.csv) | Inverter | Huawei | SUN2000-50_70KTL-M0 | Modbus | 73 | 37 |
| [`INV_Huawei_SUN2000-60KTL-HV.csv`](./INV_Huawei_SUN2000-60KTL-HV.csv) | Inverter | Huawei | SUN2000-60KTL-HV | Modbus | 347 | 30 |
| [`INV_Huawei_SUN2000-8_20KTL.csv`](./INV_Huawei_SUN2000-8_20KTL.csv) | Inverter | Huawei | SUN2000-8_20KTL | Modbus | 328 | 25 |
| [`INV_Huawei_SUN2000-90_105KTL.csv`](./INV_Huawei_SUN2000-90_105KTL.csv) | Inverter | Huawei | SUN2000-90_105KTL | Modbus | 73 | 37 |
| [`INV_Huawei_SUN2000-All_V3.csv`](./INV_Huawei_SUN2000-All_V3.csv) | Inverter | Huawei | SUN2000-All_V3 | Modbus | 97 | 15 |
| [`INV_Huawei_SUN2000-M3.csv`](./INV_Huawei_SUN2000-M3.csv) | Inverter | Huawei | SUN2000-M3 | Modbus | 64 | 64 |
| [`INV_Huawei_SUN2000-Utility.csv`](./INV_Huawei_SUN2000-Utility.csv) | Inverter | Huawei | SUN2000-Utility | Modbus | 88 | 88 |
| [`INV_Huawei_SUN2000-V3.csv`](./INV_Huawei_SUN2000-V3.csv) | Inverter | Huawei | SUN2000-V3 | Modbus | 97 | 15 |
| [`INV_Huawei_SUN2000-V4_Generic.csv`](./INV_Huawei_SUN2000-V4_Generic.csv) | Inverter | Huawei | SUN2000-V4_Generic | Modbus | 322 | 20 |
| [`INV_Huawei_SUN2000C00.csv`](./INV_Huawei_SUN2000C00.csv) | Inverter | Huawei | SUN2000C00 | Modbus | 274 | 33 |
| [`INV_Huawei_SUN2000C00_MODBUS_Protocol.csv`](./INV_Huawei_SUN2000C00_MODBUS_Protocol.csv) | Inverter | Huawei | SUN2000C00_MODBUS_Protocol | Modbus | 129 | 31 |
| [`INV_Huawei_SUN2000C11.csv`](./INV_Huawei_SUN2000C11.csv) | Inverter | Huawei | SUN2000C11 | Modbus | 183 | 33 |
| [`INV_Huawei_SUN2000C11_MODBUS_Protocol.csv`](./INV_Huawei_SUN2000C11_MODBUS_Protocol.csv) | Inverter | Huawei | SUN2000C11_MODBUS_Protocol | Modbus | 129 | 31 |
| [`INV_Huawei_SUN2000MA_V100R001C00.csv`](./INV_Huawei_SUN2000MA_V100R001C00.csv) | Inverter | Huawei | SUN2000MA_V100R001C00 | Modbus | 263 | 39 |
| [`INV_Huawei_SUN2000_V4.csv`](./INV_Huawei_SUN2000_V4.csv) | Inverter | Huawei | SUN2000_V4 | Modbus | 322 | 24 |
| [`COM_Huawei_SmartLogger3000.csv`](./COM_Huawei_SmartLogger3000.csv) | Logger | Huawei | SmartLogger3000 | Modbus | 83 | 18 |
| [`INV_Huawei_SmartLogger3000.csv`](./INV_Huawei_SmartLogger3000.csv) | Inverter | Huawei | SmartLogger3000 | Modbus | 11 | 4 |
| [`COM_Huawei_SmartLogger_VSeries.csv`](./COM_Huawei_SmartLogger_VSeries.csv) | Logger | Huawei | SmartLogger_VSeries | Modbus | 176 | 20 |
| [`INV_Huawei_V2.csv`](./INV_Huawei_V2.csv) | Inverter | Huawei | V2 | Modbus | 214 | 41 |
| [`OTH_Huawei_V2.csv`](./OTH_Huawei_V2.csv) | Other | Huawei | V2 | Modbus | 214 | 46 |
| [`INV_Huawei_V3.csv`](./INV_Huawei_V3.csv) | Inverter | Huawei | V3 | Modbus | 182 | 74 |
| [`OTH_Huawei_V3.csv`](./OTH_Huawei_V3.csv) | Other | Huawei | V3 | Modbus | 97 | 15 |
| [`INV_Huawei_V4.csv`](./INV_Huawei_V4.csv) | Inverter | Huawei | V4 | Modbus | 392 | 80 |
| [`OTH_Huawei_V4.csv`](./OTH_Huawei_V4.csv) | Other | Huawei | V4 | Modbus | 322 | 24 |
| [`BAT_Huawei_V4_BATTERY.csv`](./BAT_Huawei_V4_BATTERY.csv) | Battery | Huawei | V4_BATTERY | Modbus | 140 | 9 |
| [`MET_Huawei_V4_METER.csv`](./MET_Huawei_V4_METER.csv) | Meter | Huawei | V4_METER | Modbus | 22 | 8 |
| [`OTH_Huawei_Work.csv`](./OTH_Huawei_Work.csv) | Other | Huawei | Work | Modbus | 426 | 2 |
| [`SEN_Hukseflux_Generic.csv`](./SEN_Hukseflux_Generic.csv) | Sensor | Hukseflux | Generic | Modbus | 27 | 4 |
| [`SEN_Hukseflux_SR05-D1A3.csv`](./SEN_Hukseflux_SR05-D1A3.csv) | Sensor | Hukseflux | SR05-D1A3 | Modbus | 28 | 1 |
| [`SEN_Hukseflux_SR05_D1A3.csv`](./SEN_Hukseflux_SR05_D1A3.csv) | Sensor | Hukseflux | SR05_D1A3 | Modbus | 28 | 4 |
| [`SEN_Hukseflux_SR20-D2.csv`](./SEN_Hukseflux_SR20-D2.csv) | Sensor | Hukseflux | SR20-D2 | Modbus | 12 | 4 |
| [`SEN_Hukseflux_SR20_D2.csv`](./SEN_Hukseflux_SR20_D2.csv) | Sensor | Hukseflux | SR20_D2 | Modbus | 13 | 4 |
| [`SEN_Hukseflux_SR30.csv`](./SEN_Hukseflux_SR30.csv) | Sensor | Hukseflux | SR30 | Modbus | 28 | 1 |
| [`OTH_Hypercharger_by_alpitronic_EV_Charger_ALPITRONIC_HYPERCHARGER.csv`](./OTH_Hypercharger_by_alpitronic_EV_Charger_ALPITRONIC_HYPERCHARGER.csv) | Other | Hypercharger | by_alpitronic_EV_Charger_ALPITRONIC_HYPERCHARGER | Modbus | 80 | 5 |
| [`INV_Hyxi_100-120K.csv`](./INV_Hyxi_100-120K.csv) | Inverter | Hyxi | 100-120K | Modbus | 228 | 38 |
| [`INV_Hyxi_8-50K.csv`](./INV_Hyxi_8-50K.csv) | Inverter | Hyxi | 8-50K | Modbus | 155 | 36 |
| [`INV_Hyxi_8-50K_workinprogress.csv`](./INV_Hyxi_8-50K_workinprogress.csv) | Inverter | Hyxi | 8-50K_workinprogress | Modbus | 183 | 40 |
| [`INV_Hyxi_HYXIPOWER_100-120K.csv`](./INV_Hyxi_HYXIPOWER_100-120K.csv) | Inverter | Hyxi | HYXIPOWER_100-120K | Modbus | 133 | 33 |
| [`INV_Hyxi_HYXIPOWER_8-50K.csv`](./INV_Hyxi_HYXIPOWER_8-50K.csv) | Inverter | Hyxi | HYXIPOWER_8-50K | Modbus | 155 | 36 |
| [`INV_Hyxi_HYXIPOWER_8-50K_workinprogress.csv`](./INV_Hyxi_HYXIPOWER_8-50K_workinprogress.csv) | Inverter | Hyxi | HYXIPOWER_8-50K_workinprogress | Modbus | 183 | 40 |
| [`OTH_ICPDAS_M-7000.csv`](./OTH_ICPDAS_M-7000.csv) | Other | ICPDAS | M-7000 | Modbus | 70 | 19 |
| [`OTH_ICPDAS_M7028.csv`](./OTH_ICPDAS_M7028.csv) | Other | ICPDAS | M7028 | Modbus | 70 | 18 |
| [`MET_IME_CE4DF3DTMID.csv`](./MET_IME_CE4DF3DTMID.csv) | Meter | IME | CE4DF3DTMID | Modbus | 66 | 16 |
| [`MET_IME_CE4TBDTMID.csv`](./MET_IME_CE4TBDTMID.csv) | Meter | IME | CE4TBDTMID | Modbus | 82 | 16 |
| [`MET_IME_Generic.csv`](./MET_IME_Generic.csv) | Meter | IME | GENERIC | Modbus | 85 | 16 |
| [`OTH_IME_GENERIC.csv`](./OTH_IME_GENERIC.csv) | Other | IME | GENERIC | Modbus | 82 | 16 |
| [`MET_IME_IM-MFD4ORFCDT1.csv`](./MET_IME_IM-MFD4ORFCDT1.csv) | Meter | IME | IM-MFD4ORFCDT1 | Modbus | 73 | 21 |
| [`OTH_IME_IM-MFD4ORFCDT1.csv`](./OTH_IME_IM-MFD4ORFCDT1.csv) | Other | IME | IM-MFD4ORFCDT1 | Modbus | 70 | 21 |
| [`MET_Imesys_NEMO_96HDLe.csv`](./MET_Imesys_NEMO_96HDLe.csv) | Meter | IMESYS | NEMO_96HDLe | Modbus | 797 | 30 |
| [`OTH_Imesys_NEMO_96HDLe.csv`](./OTH_Imesys_NEMO_96HDLe.csv) | Other | IMESYS | NEMO_96HDLe | Modbus | 797 | 30 |
| [`OTH_Imesys_PR139FR.csv`](./OTH_Imesys_PR139FR.csv) | Other | IMESYS | PR139FR | Modbus | 57 | 26 |
| [`INV_Ingeteam_1play_ABE1000.csv`](./INV_Ingeteam_1play_ABE1000.csv) | Inverter | INGETEAM | 1play_ABE1000 | Modbus | 45 | 12 |
| [`INV_Ingeteam_1play_ABF1000.csv`](./INV_Ingeteam_1play_ABF1000.csv) | Inverter | INGETEAM | 1play_ABF1000 | Modbus | 45 | 10 |
| [`INV_Ingeteam_3Phases.csv`](./INV_Ingeteam_3Phases.csv) | Inverter | INGETEAM | 3phases | Modbus | 111 | 12 |
| [`INV_Ingeteam_3play_ABI_1000_ABI_1004.csv`](./INV_Ingeteam_3play_ABI_1000_ABI_1004.csv) | Inverter | INGETEAM | 3play_ABI_1000_ABI_1004 | Modbus | 59 | 26 |
| [`INV_Ingeteam_3play_ABS1004_todo.csv`](./INV_Ingeteam_3play_ABS1004_todo.csv) | Inverter | INGETEAM | 3play_ABS1004_todo | Modbus | 56 | 26 |
| [`INV_Ingeteam_PowerBlock_ABK1000.csv`](./INV_Ingeteam_PowerBlock_ABK1000.csv) | Inverter | INGETEAM | PowerBlock_ABK1000 | Modbus | 76 | 19 |
| [`INV_Ingeteam_PowerMax_centraux.csv`](./INV_Ingeteam_PowerMax_centraux.csv) | Inverter | INGETEAM | PowerMax_centraux | Modbus | 31 | 3 |
| [`INV_INVT_Generic.csv`](./INV_INVT_Generic.csv) | Inverter | INVT | Generic | Modbus | 41 | 29 |
| [`INV_INVT_ModbusRTU_INVT_Generic.csv`](./INV_INVT_ModbusRTU_INVT_Generic.csv) | Inverter | INVT | ModbusRTU_INVT_Generic | Modbus | 41 | 29 |
| [`OTH_INVT_ModbusRTU_INVT_Generic.csv`](./OTH_INVT_ModbusRTU_INVT_Generic.csv) | Other | INVT | ModbusRTU_INVT_Generic | Modbus | 41 | 29 |
| [`OTH_Inaccess_MODBUS_Inaccess.csv`](./OTH_Inaccess_MODBUS_Inaccess.csv) | Other | Inaccess | MODBUS_Inaccess | Modbus | 1,098 | 2 |
| [`MET_Inepro_PRO380-MOD-CT.csv`](./MET_Inepro_PRO380-MOD-CT.csv) | Meter | Inepro | PRO380-MOD-CT | Modbus | 24 | 4 |
| [`MET_Inepro_PRO380-MOD-CT_Inepro.csv`](./MET_Inepro_PRO380-MOD-CT_Inepro.csv) | Meter | Inepro | PRO380-MOD-CT_Inepro | Modbus | 24 | 9 |
| [`OTH_Inepro_PRO380-MOD-CT_Inepro.csv`](./OTH_Inepro_PRO380-MOD-CT_Inepro.csv) | Other | Inepro | PRO380-MOD-CT_Inepro | Modbus | 24 | 9 |
| [`OTH_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv`](./OTH_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv) | Other | IngenieurBuro | MODBUS_Si-RS485-TC-2T-v | Modbus | 7 | 3 |
| [`SEN_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv`](./SEN_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv) | Sensor | IngenieurBuro | MODBUS_Si-RS485-TC-2T-v | Modbus | 7 | 3 |
| [`OTH_IngenieurBuro_MODBUS_Tx-RS485-MB.csv`](./OTH_IngenieurBuro_MODBUS_Tx-RS485-MB.csv) | Other | IngenieurBuro | MODBUS_Tx-RS485-MB | Modbus | 9 | 1 |
| [`SEN_IngenieurBuro_MODBUS_Tx-RS485-MB.csv`](./SEN_IngenieurBuro_MODBUS_Tx-RS485-MB.csv) | Sensor | IngenieurBuro | MODBUS_Tx-RS485-MB | Modbus | 9 | 1 |
| [`OTH_IngenieurBuro_Si-RS485-TC-2T-v.csv`](./OTH_IngenieurBuro_Si-RS485-TC-2T-v.csv) | Other | IngenieurBuro | Si-RS485-TC-2T-v | Modbus | 7 | 3 |
| [`SEN_IngenieurBuro_Si-RS485-TC-2T-v.csv`](./SEN_IngenieurBuro_Si-RS485-TC-2T-v.csv) | Sensor | IngenieurBuro | Si-RS485-TC-2T-v | Modbus | 7 | 3 |
| [`SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv`](./SEN_IngenieurBuro_Si-RS485TC-2T-v-MB.csv) | Sensor | IngenieurBuro | Si-RS485TC-2T-v-MB | Modbus | 7 | 0 |
| [`SEN_IngenieurBuro_Si-RS485TC-3T-MB.csv`](./SEN_IngenieurBuro_Si-RS485TC-3T-MB.csv) | Sensor | IngenieurBuro | Si-RS485TC-3T-MB | Modbus | 10 | 3 |
| [`SEN_IngenieurBuro_TA-RS485-MB.csv`](./SEN_IngenieurBuro_TA-RS485-MB.csv) | Sensor | IngenieurBuro | TA-RS485-MB | Modbus | 3 | 1 |
| [`SEN_IngenieurBuro_TM-RS485-MB.csv`](./SEN_IngenieurBuro_TM-RS485-MB.csv) | Sensor | IngenieurBuro | TM-RS485-MB | Modbus | 3 | 1 |
| [`OTH_IngenieurBuro_Tx-RS485-MB.csv`](./OTH_IngenieurBuro_Tx-RS485-MB.csv) | Other | IngenieurBuro | Tx-RS485-MB | Modbus | 6 | 1 |
| [`SEN_IngenieurBuro_TX-RS485-MB.csv`](./SEN_IngenieurBuro_TX-RS485-MB.csv) | Sensor | IngenieurBuro | Tx-RS485-MB | Modbus | 6 | 1 |
| [`OTH_Ingeteam_1play_ABE1000.csv`](./OTH_Ingeteam_1play_ABE1000.csv) | Other | Ingeteam | 1play_ABE1000 | Modbus | 45 | 12 |
| [`OTH_Ingeteam_1play_ABF1000.csv`](./OTH_Ingeteam_1play_ABF1000.csv) | Other | Ingeteam | 1play_ABF1000 | Modbus | 45 | 10 |
| [`OTH_Ingeteam_3play_ABI_1000_ABI_1004.csv`](./OTH_Ingeteam_3play_ABI_1000_ABI_1004.csv) | Other | Ingeteam | 3play_ABI_1000_ABI_1004 | Modbus | 59 | 26 |
| [`OTH_Ingeteam_3play_ABS1004_todo.csv`](./OTH_Ingeteam_3play_ABS1004_todo.csv) | Other | Ingeteam | 3play_ABS1004_todo | Modbus | 56 | 26 |
| [`INV_Ingeteam_AAA0030IMB02.csv`](./INV_Ingeteam_AAA0030IMB02.csv) | Inverter | Ingeteam | AAA0030IMB02 | Modbus | 131 | 41 |
| [`OTH_Ingeteam_MODBUS_INGETEAM_Stringbox.csv`](./OTH_Ingeteam_MODBUS_INGETEAM_Stringbox.csv) | Other | Ingeteam | MODBUS_INGETEAM_Stringbox | Modbus | 52 | 19 |
| [`INV_Ingeteam_Modbus_1play_ABE1000.csv`](./INV_Ingeteam_Modbus_1play_ABE1000.csv) | Inverter | Ingeteam | Modbus_1play_ABE1000 | Modbus | 45 | 12 |
| [`OTH_Ingeteam_Modbus_1play_ABE1000.csv`](./OTH_Ingeteam_Modbus_1play_ABE1000.csv) | Other | Ingeteam | Modbus_1play_ABE1000 | Modbus | 45 | 12 |
| [`INV_Ingeteam_Modbus_1play_ABF1000.csv`](./INV_Ingeteam_Modbus_1play_ABF1000.csv) | Inverter | Ingeteam | Modbus_1play_ABF1000 | Modbus | 42 | 10 |
| [`OTH_Ingeteam_Modbus_1play_ABF1000.csv`](./OTH_Ingeteam_Modbus_1play_ABF1000.csv) | Other | Ingeteam | Modbus_1play_ABF1000 | Modbus | 42 | 10 |
| [`INV_Ingeteam_Modbus_3play_ABI1000-1004.csv`](./INV_Ingeteam_Modbus_3play_ABI1000-1004.csv) | Inverter | Ingeteam | Modbus_3play_ABI1000-1004 | Modbus | 59 | 26 |
| [`OTH_Ingeteam_Modbus_3play_ABI1000-1004.csv`](./OTH_Ingeteam_Modbus_3play_ABI1000-1004.csv) | Other | Ingeteam | Modbus_3play_ABI1000-1004 | Modbus | 59 | 26 |
| [`OTH_Ingeteam_Modbus_Map_Rapid_Fusion_EN.csv`](./OTH_Ingeteam_Modbus_Map_Rapid_Fusion_EN.csv) | Other | Ingeteam | Modbus_Map_Rapid_Fusion_EN | Modbus | 98 | 0 |
| [`INV_Ingeteam_Modbus_PowerBlock_ABK1000.csv`](./INV_Ingeteam_Modbus_PowerBlock_ABK1000.csv) | Inverter | Ingeteam | Modbus_PowerBlock_ABK1000 | Modbus | 76 | 19 |
| [`OTH_Ingeteam_Modbus_PowerBlock_ABK1000.csv`](./OTH_Ingeteam_Modbus_PowerBlock_ABK1000.csv) | Other | Ingeteam | Modbus_PowerBlock_ABK1000 | Modbus | 76 | 19 |
| [`INV_Ingeteam_Modbus_PowerMax_centraux.csv`](./INV_Ingeteam_Modbus_PowerMax_centraux.csv) | Inverter | Ingeteam | Modbus_PowerMax_centraux | Modbus | 31 | 3 |
| [`OTH_Ingeteam_Modbus_PowerMax_centraux.csv`](./OTH_Ingeteam_Modbus_PowerMax_centraux.csv) | Other | Ingeteam | Modbus_PowerMax_centraux | Modbus | 31 | 3 |
| [`OTH_Ingeteam_PowerBlock_ABK1000.csv`](./OTH_Ingeteam_PowerBlock_ABK1000.csv) | Other | Ingeteam | PowerBlock_ABK1000 | Modbus | 76 | 19 |
| [`OTH_Ingeteam_PowerMax_centraux.csv`](./OTH_Ingeteam_PowerMax_centraux.csv) | Other | Ingeteam | PowerMax_centraux | Modbus | 31 | 3 |
| [`OTH_Intelec_annunciator_Annunciator.csv`](./OTH_Intelec_annunciator_Annunciator.csv) | Other | Intelec | annunciator_Annunciator | Modbus | 10 | 0 |
| [`INV_Jema_Generic.csv`](./INV_Jema_Generic.csv) | Inverter | JEMA | Generic | Modbus | 60 | 1 |
| [`INV_Jema_Generic_revG.csv`](./INV_Jema_Generic_revG.csv) | Inverter | JEMA | Generic_revG | Modbus | 60 | 26 |
| [`MET_Janitza_12926-UB_MODBUS_COMPTEUR_JANITZA_UMG604-PRO.csv`](./MET_Janitza_12926-UB_MODBUS_COMPTEUR_JANITZA_UMG604-PRO.csv) | Meter | Janitza | 12926-UB_MODBUS_COMPTEUR_JANITZA_UMG604-PRO | Modbus | 71 | 23 |
| [`MET_Janitza_GENERIC_TCP.csv`](./MET_Janitza_GENERIC_TCP.csv) | Meter | Janitza | GENERIC_TCP | Modbus | 61 | 14 |
| [`MET_Janitza_Generic.csv`](./MET_Janitza_Generic.csv) | Meter | Janitza | Generic | Modbus | 61 | 30 |
| [`MET_Janitza_Generic_RTU.csv`](./MET_Janitza_Generic_RTU.csv) | Meter | Janitza | Generic_RTU | Modbus | 61 | 14 |
| [`MET_Janitza_MODBUS_Janitza.csv`](./MET_Janitza_MODBUS_Janitza.csv) | Meter | Janitza | MODBUS_Janitza | Modbus | 61 | 14 |
| [`MET_Janitza_MODBUS_UMG-104_Janitza_RTU.csv`](./MET_Janitza_MODBUS_UMG-104_Janitza_RTU.csv) | Meter | Janitza | MODBUS_UMG-104_Janitza_RTU | Modbus | 61 | 11 |
| [`MET_Janitza_Modbus_Janitza_UMG103CBM.csv`](./MET_Janitza_Modbus_Janitza_UMG103CBM.csv) | Meter | Janitza | Modbus_Janitza_UMG103CBM | Modbus | 61 | 9 |
| [`MET_Janitza_UMG-604.csv`](./MET_Janitza_UMG-604.csv) | Meter | Janitza | UMG-604 | Modbus | 4 | 4 |
| [`MET_Janitza_UMG604-PRO.csv`](./MET_Janitza_UMG604-PRO.csv) | Meter | Janitza | UMG604-PRO | Modbus | 71 | 23 |
| [`MET_Janitza_UMG604_PRO.csv`](./MET_Janitza_UMG604_PRO.csv) | Meter | Janitza | UMG604-PRO | Modbus | 71 | 23 |
| [`MET_Janitza_UMG604PRO_COMPLET.csv`](./MET_Janitza_UMG604PRO_COMPLET.csv) | Meter | Janitza | UMG604PRO_COMPLET | Modbus | 5,877 | 31 |
| [`MET_Janitza_UMG604_RTU.csv`](./MET_Janitza_UMG604_RTU.csv) | Meter | Janitza | UMG604_RTU | Modbus | 61 | 10 |
| [`MET_Janitza_W_Janitza_Generic.csv`](./MET_Janitza_W_Janitza_Generic.csv) | Meter | Janitza | W_Janitza_Generic | Modbus | 61 | 30 |
| [`OTH_Jema_Generic_revG.csv`](./OTH_Jema_Generic_revG.csv) | Other | Jema | Generic_revG | Modbus | 60 | 26 |
| [`INV_Jema_IFX3_1050_RevD.csv`](./INV_Jema_IFX3_1050_RevD.csv) | Inverter | Jema | IFX3_1050_RevD | Modbus | 33 | 10 |
| [`OTH_Jema_IFX3_1050_RevD.csv`](./OTH_Jema_IFX3_1050_RevD.csv) | Other | Jema | IFX3_1050_RevD | Modbus | 33 | 12 |
| [`INV_Jema_Inverter.csv`](./INV_Jema_Inverter.csv) | Inverter | Jema | Inverter | Modbus | 57 | 27 |
| [`INV_Jema_Modbus_revG.csv`](./INV_Jema_Modbus_revG.csv) | Inverter | Jema | Modbus_revG | Modbus | 60 | 26 |
| [`OTH_Jema_Modbus_revG.csv`](./OTH_Jema_Modbus_revG.csv) | Other | Jema | Modbus_revG | Modbus | 60 | 26 |
| [`OTH_K_star.csv`](./OTH_K_star.csv) | Other | K | star | Modbus | 36 | 18 |
| [`INV_Kehua_SPI100-250K-B.csv`](./INV_Kehua_SPI100-250K-B.csv) | Inverter | KEHUA | SPI100-250K-B | Modbus | 238 | 116 |
| [`INV_Kehua_SPI_100K-250K-B.csv`](./INV_Kehua_SPI_100K-250K-B.csv) | Inverter | KEHUA | SPI_100K-250K-B | Modbus | 238 | 6 |
| [`SEN_KippZonen_RT1.csv`](./SEN_KippZonen_RT1.csv) | Sensor | KIPPZONEN | RT1 | Modbus | 15 | 0 |
| [`INV_Kostal_PIKOCI_Modbus.csv`](./INV_Kostal_PIKOCI_Modbus.csv) | Inverter | KOSTAL | PIKOCI_Modbus | Modbus | 78 | 30 |
| [`INV_Kostal_PIKO_CI.csv`](./INV_Kostal_PIKO_CI.csv) | Inverter | KOSTAL | PIKO_CI | Modbus | 78 | 8 |
| [`INV_Kostal_PLENTICORE.csv`](./INV_Kostal_PLENTICORE.csv) | Inverter | KOSTAL | PLENTICORE | Modbus | 154 | 36 |
| [`INV_Kstar_Generic.csv`](./INV_Kstar_Generic.csv) | Inverter | KSTAR | Generic | Modbus | 52 | 6 |
| [`INV_Kstar_HYBRID.csv`](./INV_Kstar_HYBRID.csv) | Inverter | KSTAR | HYBRID | Modbus | 176 | 10 |
| [`INV_Kstar_KIG1-250K.csv`](./INV_Kstar_KIG1-250K.csv) | Inverter | KSTAR | KIG1-250K | Modbus | 254 | 6 |
| [`INV_Kaco_95308-LOGNON_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.0.3.3-ec484f_modbusRTU.csv`](./INV_Kaco_95308-LOGNON_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.0.3.3-ec484f_modbusRTU.csv) | Inverter | Kaco | 95308-LOGNON_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.0.3.3-ec484f_modbusRTU | Modbus | 928 | 37 |
| [`INV_Kaco_97636-CARTOUX_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.1.10.1-1e3c4_modbusTCP.csv`](./INV_Kaco_97636-CARTOUX_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.1.10.1-1e3c4_modbusTCP.csv) | Inverter | Kaco | 97636-CARTOUX_SunSpec_inverter_KACO_new_energy_blueplanet_125_NX3_M10_IEC_5.1.10.1-1e3c4_modbusTCP | Modbus | 928 | 37 |
| [`INV_Kaco_CHAZE_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V3.52_modbusTCP.csv`](./INV_Kaco_CHAZE_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V3.52_modbusTCP.csv) | Inverter | Kaco | CHAZE_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V3.52_modbusTCP | Modbus | 101 | 34 |
| [`INV_Kaco_CHRISTIAENS_SunSpec_inverter_KACO_new_energy_Powador_60.0_TL3_V3.52_modbusTCP.csv`](./INV_Kaco_CHRISTIAENS_SunSpec_inverter_KACO_new_energy_Powador_60.0_TL3_V3.52_modbusTCP.csv) | Inverter | Kaco | CHRISTIAENS_SunSpec_inverter_KACO_new_energy_Powador_60.0_TL3_V3.52_modbusTCP | Modbus | 101 | 34 |
| [`INV_Kaco_CHRISTIANS_SunSpec_inverter_KACO_new_energy_blueplanet_15.0_TL3_V5.20_modbusTCP_1.csv`](./INV_Kaco_CHRISTIANS_SunSpec_inverter_KACO_new_energy_blueplanet_15.0_TL3_V5.20_modbusTCP_1.csv) | Inverter | Kaco | CHRISTIANS_SunSpec_inverter_KACO_new_energy_blueplanet_15.0_TL3_V5.20_modbusTCP_1 | Modbus | 638 | 36 |
| [`INV_Kaco_GOUSSARD_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V5.24_modbusTCP.csv`](./INV_Kaco_GOUSSARD_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V5.24_modbusTCP.csv) | Inverter | Kaco | GOUSSARD_SunSpec_inverter_KACO_new_energy_blueplanet_50.0_TL3_WM_V5.24_modbusTCP | Modbus | 628 | 37 |
| [`INV_Kaco_MOURLON_SunSpec_inverter_KACO_new_energy_blueplanet_20.0_TL3_V5.20_modbusTCP_1.csv`](./INV_Kaco_MOURLON_SunSpec_inverter_KACO_new_energy_blueplanet_20.0_TL3_V5.20_modbusTCP_1.csv) | Inverter | Kaco | MOURLON_SunSpec_inverter_KACO_new_energy_blueplanet_20.0_TL3_V5.20_modbusTCP_1 | Modbus | 697 | 36 |
| [`INV_Kaco_NX1.csv`](./INV_Kaco_NX1.csv) | Inverter | Kaco | NX1 | Modbus | 182 | 35 |
| [`INV_Kaco_NX1-NX3_ModbusRTU.csv`](./INV_Kaco_NX1-NX3_ModbusRTU.csv) | Inverter | Kaco | NX1-NX3_ModbusRTU | Modbus | 182 | 36 |
| [`INV_Kaco_NX1-NX3_Series.csv`](./INV_Kaco_NX1-NX3_Series.csv) | Inverter | Kaco | NX1-NX3_Series | Modbus | 182 | 9 |
| [`INV_Kaco_NX3_G2_G3_SunSpec.csv`](./INV_Kaco_NX3_G2_G3_SunSpec.csv) | Inverter | Kaco | NX3_G2_G3_SunSpec | Modbus | 458 | 30 |
| [`INV_Kaco_Powador_60.0_TL3.csv`](./INV_Kaco_Powador_60.0_TL3.csv) | Inverter | Kaco | Powador_60.0_TL3 | Modbus | 101 | 34 |
| [`OTH_Kaco_SUNSPEC_NX3_G2_G3_1xx_7xx.csv`](./OTH_Kaco_SUNSPEC_NX3_G2_G3_1xx_7xx.csv) | Other | Kaco | SUNSPEC_NX3_G2_G3_1xx_7xx | Modbus | 495 | 32 |
| [`INV_Kaco_SunSpec_inverter_KACO_new_energy_blueplanet_100_NX3_M8_IEC_5.0.3.3-ec484f_modbusRTU.csv`](./INV_Kaco_SunSpec_inverter_KACO_new_energy_blueplanet_100_NX3_M8_IEC_5.0.3.3-ec484f_modbusRTU.csv) | Inverter | Kaco | SunSpec_inverter_KACO_new_energy_blueplanet_100_NX3_M8_IEC_5.0.3.3-ec484f_modbusRTU | Modbus | 908 | 37 |
| [`INV_Kaco_Sunspec.csv`](./INV_Kaco_Sunspec.csv) | Inverter | Kaco | Sunspec | Modbus | 107 | 36 |
| [`INV_Kaco_Sunspec_11.csv`](./INV_Kaco_Sunspec_11.csv) | Inverter | Kaco | Sunspec_11 | Modbus | 107 | 7 |
| [`INV_Kaco_blueplanet_100_NX3_M8.csv`](./INV_Kaco_blueplanet_100_NX3_M8.csv) | Inverter | Kaco | blueplanet_100_NX3_M8 | Modbus | 908 | 37 |
| [`INV_Kaco_blueplanet_125_NX3_M10.csv`](./INV_Kaco_blueplanet_125_NX3_M10.csv) | Inverter | Kaco | blueplanet_125_NX3_M10 | Modbus | 928 | 37 |
| [`INV_Kaco_blueplanet_125_TL3.csv`](./INV_Kaco_blueplanet_125_TL3.csv) | Inverter | Kaco | blueplanet_125_TL3 | Modbus | 828 | 5 |
| [`INV_Kaco_blueplanet_15.0_TL3.csv`](./INV_Kaco_blueplanet_15.0_TL3.csv) | Inverter | Kaco | blueplanet_15.0_TL3 | Modbus | 638 | 36 |
| [`INV_Kaco_blueplanet_20.0_TL3.csv`](./INV_Kaco_blueplanet_20.0_TL3.csv) | Inverter | Kaco | blueplanet_20.0_TL3 | Modbus | 697 | 36 |
| [`INV_Kaco_blueplanet_50.0_NX3_M5.csv`](./INV_Kaco_blueplanet_50.0_NX3_M5.csv) | Inverter | Kaco | blueplanet_50.0_NX3_M5 | Modbus | 197 | 6 |
| [`INV_Kaco_blueplanet_50.0_TL3.csv`](./INV_Kaco_blueplanet_50.0_TL3.csv) | Inverter | Kaco | blueplanet_50.0_TL3 | Modbus | 628 | 37 |
| [`INV_Kaco_blueplanet_NX1.csv`](./INV_Kaco_blueplanet_NX1.csv) | Inverter | Kaco | blueplanet_NX1 | Modbus | 182 | 35 |
| [`INV_Kaco_blueplanet_NX3.csv`](./INV_Kaco_blueplanet_NX3.csv) | Inverter | Kaco | blueplanet_NX3 | Modbus | 182 | 36 |
| [`OTH_Kaco_sunspec_information_model_reference.csv`](./OTH_Kaco_sunspec_information_model_reference.csv) | Other | Kaco | sunspec_information_model_reference | Modbus | 226 | 7 |
| [`INV_Kaco_sunspec_information_model_reference_kaco.csv`](./INV_Kaco_sunspec_information_model_reference_kaco.csv) | Inverter | Kaco | sunspec_information_model_reference_kaco | Modbus | 28 | 0 |
| [`INV_Kehua_Modbus_SinglePhase.csv`](./INV_Kehua_Modbus_SinglePhase.csv) | Inverter | Kehua | Modbus_SinglePhase | Modbus | 159 | 29 |
| [`OTH_Kehua_Modbus_SinglePhase.csv`](./OTH_Kehua_Modbus_SinglePhase.csv) | Other | Kehua | Modbus_SinglePhase | Modbus | 159 | 29 |
| [`INV_Kehua_Modbus_ThreePhase.csv`](./INV_Kehua_Modbus_ThreePhase.csv) | Inverter | Kehua | Modbus_ThreePhase | Modbus | 33 | 22 |
| [`OTH_Kehua_Modbus_ThreePhase.csv`](./OTH_Kehua_Modbus_ThreePhase.csv) | Other | Kehua | Modbus_ThreePhase | Modbus | 33 | 22 |
| [`OTH_Kehua_SPI100-250K-B.csv`](./OTH_Kehua_SPI100-250K-B.csv) | Other | Kehua | SPI100-250K-B | Modbus | 238 | 116 |
| [`OTH_Kehua_SPI5-60B-K_Work.csv`](./OTH_Kehua_SPI5-60B-K_Work.csv) | Other | Kehua | SPI5-60B-K_Work | Modbus | 210 | 95 |
| [`INV_Kehua_SinglePhase.csv`](./INV_Kehua_SinglePhase.csv) | Inverter | Kehua | SinglePhase | Modbus | 159 | 5 |
| [`OTH_Kehua_SinglePhase_Work.csv`](./OTH_Kehua_SinglePhase_Work.csv) | Other | Kehua | SinglePhase_Work | Modbus | 149 | 26 |
| [`INV_Kehua_ThreePhase.csv`](./INV_Kehua_ThreePhase.csv) | Inverter | Kehua | ThreePhase | Modbus | 84 | 3 |
| [`OTH_Kehua_ThreePhase.csv`](./OTH_Kehua_ThreePhase.csv) | Other | Kehua | ThreePhase | Modbus | 108 | 26 |
| [`INV_Kehua_ThreePhase_nondefinitif.csv`](./INV_Kehua_ThreePhase_nondefinitif.csv) | Inverter | Kehua | ThreePhase_nondefinitif | Modbus | 84 | 18 |
| [`OTH_Kehua_ThreePhase_nondefinitif.csv`](./OTH_Kehua_ThreePhase_nondefinitif.csv) | Other | Kehua | ThreePhase_nondefinitif | Modbus | 84 | 17 |
| [`INV_Kerneli_Sistemi_ST124-22.csv`](./INV_Kerneli_Sistemi_ST124-22.csv) | Inverter | Kerneli | Sistemi_ST124-22 | Modbus | 162 | 28 |
| [`OTH_Kerneli_Sistemi_ST124-22.csv`](./OTH_Kerneli_Sistemi_ST124-22.csv) | Other | Kerneli | Sistemi_ST124-22 | Modbus | 162 | 28 |
| [`OTH_Kerneli_Sistemi_ST1_Kernel_Sistemi.csv`](./OTH_Kerneli_Sistemi_ST1_Kernel_Sistemi.csv) | Other | Kerneli | Sistemi_ST1_Kernel_Sistemi | Modbus | 117 | 2 |
| [`SEN_KippZonen_DustIQ.csv`](./SEN_KippZonen_DustIQ.csv) | Sensor | KippZonen | DustIQ | Modbus | 19 | 0 |
| [`OTH_KippZonen_IDsite_SMP_lite.csv`](./OTH_KippZonen_IDsite_SMP_lite.csv) | Other | KippZonen | IDsite_SMP_lite | Modbus | 5 | 0 |
| [`SEN_KippZonen_IDsite_SMP_lite.csv`](./SEN_KippZonen_IDsite_SMP_lite.csv) | Sensor | KippZonen | IDsite_SMP_lite | Modbus | 5 | 0 |
| [`OTH_KippZonen_K_SMP12.csv`](./OTH_KippZonen_K_SMP12.csv) | Other | KippZonen | K_SMP12 | Modbus | 53 | 4 |
| [`SEN_KippZonen_K_SMP12.csv`](./SEN_KippZonen_K_SMP12.csv) | Sensor | KippZonen | K_SMP12 | Modbus | 53 | 4 |
| [`OTH_KippZonen_SMP12.csv`](./OTH_KippZonen_SMP12.csv) | Other | KippZonen | SMP12 | Modbus | 53 | 4 |
| [`SEN_KippZonen_SMP12.csv`](./SEN_KippZonen_SMP12.csv) | Sensor | KippZonen | SMP12 | Modbus | 53 | 3 |
| [`OTH_KippZonen_SMP_full.csv`](./OTH_KippZonen_SMP_full.csv) | Other | KippZonen | SMP_full | Modbus | 35 | 1 |
| [`SEN_KippZonen_SMP_full.csv`](./SEN_KippZonen_SMP_full.csv) | Sensor | KippZonen | SMP_full | Modbus | 35 | 1 |
| [`SEN_KippZonen_Si-SMP-full.csv`](./SEN_KippZonen_Si-SMP-full.csv) | Sensor | KippZonen | Si-SMP-full | Modbus | 35 | 0 |
| [`INV_Kostal_Modbus.csv`](./INV_Kostal_Modbus.csv) | Inverter | Kostal | Modbus | Modbus | 94 | 8 |
| [`OTH_Kostal_Modbus.csv`](./OTH_Kostal_Modbus.csv) | Other | Kostal | Modbus | Modbus | 94 | 16 |
| [`OTH_Kostal_PIKOCI_Modbus.csv`](./OTH_Kostal_PIKOCI_Modbus.csv) | Other | Kostal | PIKOCI_Modbus | Modbus | 78 | 30 |
| [`OTH_Kostal_PIKO_CI_events_code-V1.1-20240516.csv`](./OTH_Kostal_PIKO_CI_events_code-V1.1-20240516.csv) | Other | Kostal | PIKO_CI_events_code-V1.1-20240516 | Modbus | 3 | 1 |
| [`OTH_Kostal_PIKO_CI_events_code-V1.2-20240516_via_modbus.csv`](./OTH_Kostal_PIKO_CI_events_code-V1.2-20240516_via_modbus.csv) | Other | Kostal | PIKO_CI_events_code-V1.2-20240516_via_modbus | Modbus | 3 | 1 |
| [`OTH_Kostal_PLENTICORE.csv`](./OTH_Kostal_PLENTICORE.csv) | Other | Kostal | PLENTICORE | Modbus | 154 | 36 |
| [`OTH_Kostal_SunSpecInformationModelReference.csv`](./OTH_Kostal_SunSpecInformationModelReference.csv) | Other | Kostal | SunSpecInformationModelReference | Modbus | 227 | 7 |
| [`INV_Kostal_SunSpec_Float.csv`](./INV_Kostal_SunSpec_Float.csv) | Inverter | Kostal | SunSpec_Float | Modbus | 204 | 35 |
| [`OTH_Kostal_SunSpec_Float.csv`](./OTH_Kostal_SunSpec_Float.csv) | Other | Kostal | SunSpec_Float | Modbus | 204 | 37 |
| [`OTH_Kstar_Generic.csv`](./OTH_Kstar_Generic.csv) | Other | Kstar | Generic | Modbus | 52 | 21 |
| [`OTH_Kstar_HYBRID.csv`](./OTH_Kstar_HYBRID.csv) | Other | Kstar | HYBRID | Modbus | 176 | 10 |
| [`INV_Kstar_KSG1-250K.csv`](./INV_Kstar_KSG1-250K.csv) | Inverter | Kstar | KSG1-250K | Modbus | 254 | 56 |
| [`INV_Kstar_Modbus.csv`](./INV_Kstar_Modbus.csv) | Inverter | Kstar | Modbus | Modbus | 36 | 18 |
| [`OTH_Kstar_Modbus.csv`](./OTH_Kstar_Modbus.csv) | Other | Kstar | Modbus | Modbus | 36 | 18 |
| [`INV_Kstar_power_reduction.csv`](./INV_Kstar_power_reduction.csv) | Inverter | Kstar | power_reduction | Modbus | 37 | 19 |
| [`OTH_Kstar_power_reduction.csv`](./OTH_Kstar_power_reduction.csv) | Other | Kstar | power_reduction | Modbus | 37 | 19 |
| [`MET_Lacecal_ITR2_0.csv`](./MET_Lacecal_ITR2_0.csv) | Meter | LACECAL | ITR2_0 | Modbus | 67 | 0 |
| [`MET_Legrand_CF_Generic.csv`](./MET_Legrand_CF_Generic.csv) | Meter | LEGRAND | CF_Generic | Modbus | 11 | 1 |
| [`INV_Legrand_Trimod_RTU.csv`](./INV_Legrand_Trimod_RTU.csv) | Inverter | LEGRAND | Trimod_RTU | Modbus | 42 | 10 |
| [`SEN_LSI_DPA870_873.csv`](./SEN_LSI_DPA870_873.csv) | Sensor | LSI | DPA870_873 | Modbus | 30 | 0 |
| [`OTH_LSI_SRL_LSILASTEM_DPA870-873.csv`](./OTH_LSI_SRL_LSILASTEM_DPA870-873.csv) | Other | LSI | SRL_LSILASTEM_DPA870-873 | Modbus | 30 | 1 |
| [`SEN_LSI_SRL_LSILASTEM_DPA870-873.csv`](./SEN_LSI_SRL_LSILASTEM_DPA870-873.csv) | Sensor | LSI | SRL_LSILASTEM_DPA870-873 | Modbus | 30 | 1 |
| [`OTH_Lacecal_ITR2_0.csv`](./OTH_Lacecal_ITR2_0.csv) | Other | Lacecal | ITR2_0 | Modbus | 67 | 0 |
| [`MET_Legrand_60177710.csv`](./MET_Legrand_60177710.csv) | Meter | Legrand | 60177710 | Modbus | 81 | 16 |
| [`MET_Legrand_CF-GENERIC.csv`](./MET_Legrand_CF-GENERIC.csv) | Meter | Legrand | CF-GENERIC | Modbus | 11 | 1 |
| [`MET_Legrand_EMDX.csv`](./MET_Legrand_EMDX.csv) | Meter | Legrand | EMDX | Modbus | 489 | 30 |
| [`MET_Legrand_EMSCX3.csv`](./MET_Legrand_EMSCX3.csv) | Meter | Legrand | EMSCX3 | Modbus | 108 | 17 |
| [`MET_Legrand_EMS_414922-24-25-27.csv`](./MET_Legrand_EMS_414922-24-25-27.csv) | Meter | Legrand | EMS_414922-24-25-27 | Modbus | 103 | 17 |
| [`MET_Legrand_IME_CE4DF3DTMID.csv`](./MET_Legrand_IME_CE4DF3DTMID.csv) | Meter | Legrand | IME_CE4DF3DTMID | Modbus | 66 | 16 |
| [`MET_Legrand_LEG412052.csv`](./MET_Legrand_LEG412052.csv) | Meter | Legrand | LEG412052 | Modbus | 489 | 4 |
| [`MET_Legrand_LG412041.csv`](./MET_Legrand_LG412041.csv) | Meter | Legrand | LG412041 | Modbus | 85 | 16 |
| [`MET_Legrand_Nemo96HDLe.csv`](./MET_Legrand_Nemo96HDLe.csv) | Meter | Legrand | Nemo96HDLe | Modbus | 151 | 28 |
| [`OTH_Legrand_Nemo96HDLe.csv`](./OTH_Legrand_Nemo96HDLe.csv) | Other | Legrand | Nemo96HDLe | Modbus | 151 | 28 |
| [`OTH_Legrand_Nemo_Work.csv`](./OTH_Legrand_Nemo_Work.csv) | Other | Legrand | Nemo_Work | Modbus | 169 | 28 |
| [`INV_Legrand_Trimod.csv`](./INV_Legrand_Trimod.csv) | Inverter | Legrand | Trimod | Modbus | 42 | 10 |
| [`OTH_Legrand_Trimod.csv`](./OTH_Legrand_Trimod.csv) | Other | Legrand | Trimod | Modbus | 42 | 10 |
| [`OTH_Legrand_Trimod_RTU.csv`](./OTH_Legrand_Trimod_RTU.csv) | Other | Legrand | Trimod_RTU | Modbus | 42 | 10 |
| [`MET_Lettel_MAP4-34RJ.csv`](./MET_Lettel_MAP4-34RJ.csv) | Meter | Lettel | MAP4-34RJ | Modbus | 405 | 4 |
| [`MET_Lettel_MAP4_34RJ.csv`](./MET_Lettel_MAP4_34RJ.csv) | Meter | Lettel | MAP4-34RJ | Modbus | 405 | 17 |
| [`MET_Lettel_MAP4_345M.csv`](./MET_Lettel_MAP4_345M.csv) | Meter | Lettel | MAP4_345M | Modbus | 300 | 30 |
| [`MET_Lettel_MCT4-45.csv`](./MET_Lettel_MCT4-45.csv) | Meter | Lettel | MCT4-45 | Modbus | 30 | 4 |
| [`MET_Lettel_MCT4-45_FLOAT.csv`](./MET_Lettel_MCT4-45_FLOAT.csv) | Meter | Lettel | MCT4-45_FLOAT | Modbus | 30 | 25 |
| [`MET_Lettel_MCX2.csv`](./MET_Lettel_MCX2.csv) | Meter | Lettel | MCX2 | Modbus | 54 | 8 |
| [`MET_Lettel_MCX4-34.csv`](./MET_Lettel_MCX4-34.csv) | Meter | Lettel | MCX4-34 | Modbus | 139 | 4 |
| [`MET_Lettel_MCX4-34V.csv`](./MET_Lettel_MCX4-34V.csv) | Meter | Lettel | MCX4-34V | Modbus | 102 | 4 |
| [`MET_Lettel_MCX4-34V_FLOAT.csv`](./MET_Lettel_MCX4-34V_FLOAT.csv) | Meter | Lettel | MCX4-34V_FLOAT | Modbus | 102 | 33 |
| [`MET_Lettel_MCX4_34V.csv`](./MET_Lettel_MCX4_34V.csv) | Meter | Lettel | MCX4_34V | Modbus | 139 | 33 |
| [`MET_Lettel_MODBUS_LETTEL.csv`](./MET_Lettel_MODBUS_LETTEL.csv) | Meter | Lettel | MODBUS_LETTEL | Modbus | 449 | 0 |
| [`MET_Lettel_W_Lettel_MCX4-34V-FLOAT.csv`](./MET_Lettel_W_Lettel_MCX4-34V-FLOAT.csv) | Meter | Lettel | W_Lettel_MCX4-34V-FLOAT | Modbus | 102 | 52 |
| [`MET_Lettel_W_Lettel_MCX4-34V-INT.csv`](./MET_Lettel_W_Lettel_MCX4-34V-INT.csv) | Meter | Lettel | W_Lettel_MCX4-34V-INT | Modbus | 139 | 52 |
| [`OTH_Lettel_map4-345m_table_modbus_fr.csv`](./OTH_Lettel_map4-345m_table_modbus_fr.csv) | Other | Lettel | map4-345m_table_modbus_fr | Modbus | 326 | 30 |
| [`MET_Lufft_MODBUS_RTU_Lufft_WS10.csv`](./MET_Lufft_MODBUS_RTU_Lufft_WS10.csv) | Meter | Lufft | MODBUS_RTU_Lufft_WS10 | Modbus | 23 | 3 |
| [`OTH_Lufft_MODBUS_RTU_Lufft_WS10.csv`](./OTH_Lufft_MODBUS_RTU_Lufft_WS10.csv) | Other | Lufft | MODBUS_RTU_Lufft_WS10 | Modbus | 23 | 4 |
| [`SEN_Lufft_WS10.csv`](./SEN_Lufft_WS10.csv) | Sensor | Lufft | WS10 | Modbus | 23 | 0 |
| [`MET_Lufft_WS600.csv`](./MET_Lufft_WS600.csv) | Meter | Lufft | WS600 | Modbus | 42 | 4 |
| [`OTH_Lufft_WS600.csv`](./OTH_Lufft_WS600.csv) | Other | Lufft | WS600 | Modbus | 42 | 4 |
| [`MET_MCI_RI-D140-C.csv`](./MET_MCI_RI-D140-C.csv) | Meter | MCI | RI-D140-C | Modbus | 87 | 11 |
| [`OTH_MCI_RI-D140-C.csv`](./OTH_MCI_RI-D140-C.csv) | Other | MCI | RI-D140-C | Modbus | 87 | 11 |
| [`OTH_MCI_contax_1m_100a_031277_mapping_Modbus_20102017.csv`](./OTH_MCI_contax_1m_100a_031277_mapping_Modbus_20102017.csv) | Other | MCI | contax_1m_100a_031277_mapping_Modbus_20102017 | Modbus | 32 | 17 |
| [`OTH_MCI_contax_d_tri_80a_mid_modbus_ref_031333_-_mapping_modbus_V3.csv`](./OTH_MCI_contax_d_tri_80a_mid_modbus_ref_031333_-_mapping_modbus_V3.csv) | Other | MCI | contax_d_tri_80a_mid_modbus_ref_031333_-_mapping_modbus_V3 | Modbus | 36 | 18 |
| [`BAT_MarsEnergy_LocalController_AllInOne.csv`](./BAT_MarsEnergy_LocalController_AllInOne.csv) | BESS | Mars Energy | MarsLocalController All-In-One V7.0 | Modbus | 16,746 | 4 |
| [`COM_MarsEnergy_EMS_Slave247.csv`](./COM_MarsEnergy_EMS_Slave247.csv) | BESS | Mars Energy | MarsLocalController EMS V7.0 | Modbus | 12 | 0 |
| [`BAT_MarsEnergy_RackCell_SlaveN.csv`](./BAT_MarsEnergy_RackCell_SlaveN.csv) | BatteryRack | Mars Energy | MarsLocalController Rack V7.0 | Modbus | 5,030 | 0 |
| [`BAT_MarsEnergy_LocalController_V7.0.csv`](./BAT_MarsEnergy_LocalController_V7.0.csv) | BESS | Mars Energy | MarsLocalController V7.0 | Modbus | 11,704 | 4 |
| [`MET_MarsEnergy_AC.csv`](./MET_MarsEnergy_AC.csv) | Meter | MarsEnergy | AC | Modbus | 1,600 | 3 |
| [`MET_MarsEnergy_AC_1Unit.csv`](./MET_MarsEnergy_AC_1Unit.csv) | Meter | MarsEnergy | AC_1Unit | Modbus | 16 | 3 |
| [`BAT_MarsEnergy_LC.csv`](./BAT_MarsEnergy_LC.csv) | Battery | MarsEnergy | LC | Modbus | 12 | 4 |
| [`MET_MarsEnergy_LC.csv`](./MET_MarsEnergy_LC.csv) | Meter | MarsEnergy | LC | Modbus | 1,800 | 2 |
| [`OTH_MarsEnergy_LC.csv`](./OTH_MarsEnergy_LC.csv) | Other | MarsEnergy | LC | Modbus | 8 | 3 |
| [`MET_MarsEnergy_LC-AC.csv`](./MET_MarsEnergy_LC-AC.csv) | Meter | MarsEnergy | LC-AC | Modbus | 34 | 4 |
| [`INV_MarsEnergy_LC_1Unit.csv`](./INV_MarsEnergy_LC_1Unit.csv) | Inverter | MarsEnergy | LC_1Unit | Modbus | 18 | 2 |
| [`INV_MarsEnergy_LC_EMS_V7.csv`](./INV_MarsEnergy_LC_EMS_V7.csv) | Inverter | MarsEnergy | LC_EMS_V7 | Modbus | 12 | 4 |
| [`INV_MarsEnergy_LC_V7.csv`](./INV_MarsEnergy_LC_V7.csv) | Inverter | MarsEnergy | LC_V7 | Modbus | 174 | 174 |
| [`INV_MarsEnergy_LiquidCooling-AC.csv`](./INV_MarsEnergy_LiquidCooling-AC.csv) | Inverter | MarsEnergy | LiquidCooling-AC | Modbus | 18 | 1 |
| [`BAT_MarsEnergy_LocalControler_Bank_Rack_Cell.csv`](./BAT_MarsEnergy_LocalControler_Bank_Rack_Cell.csv) | Battery | MarsEnergy | LocalControler_Bank_Rack_Cell | Modbus | 84 | 6 |
| [`BAT_MarsEnergy_LocalControler_PCS_Information.csv`](./BAT_MarsEnergy_LocalControler_PCS_Information.csv) | Battery | MarsEnergy | LocalControler_PCS_Information | Modbus | 111 | 31 |
| [`OTH_MarsEnergy_MARS_Commandes_Unit247_SOURCE.csv`](./OTH_MarsEnergy_MARS_Commandes_Unit247_SOURCE.csv) | Other | MarsEnergy | MARS_Commandes_Unit247_SOURCE | Modbus | 7 | 4 |
| [`COM_MarsEnergy_MARS_LocalController_V7_1_Slave.csv`](./COM_MarsEnergy_MARS_LocalController_V7_1_Slave.csv) | Logger | MarsEnergy | MARS_LocalController_V7_1_Slave | Modbus | 207 | 26 |
| [`OTH_MarsEnergy_MARS_LocalController_V7_1_Slave.csv`](./OTH_MarsEnergy_MARS_LocalController_V7_1_Slave.csv) | Other | MarsEnergy | MARS_LocalController_V7_1_Slave | Modbus | 207 | 26 |
| [`COM_MarsEnergy_MARS_LocalController_V7_Complet.csv`](./COM_MarsEnergy_MARS_LocalController_V7_Complet.csv) | Logger | MarsEnergy | MARS_LocalController_V7_Complet | Modbus | 15,125 | 26 |
| [`OTH_MarsEnergy_MARS_LocalController_V7_Complet.csv`](./OTH_MarsEnergy_MARS_LocalController_V7_Complet.csv) | Other | MarsEnergy | MARS_LocalController_V7_Complet | Modbus | 15,125 | 26 |
| [`BAT_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv`](./BAT_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv) | Battery | MarsEnergy | MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM | Modbus | 35 | 3 |
| [`COM_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv`](./COM_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv) | Logger | MarsEnergy | MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM | Modbus | 35 | 3 |
| [`OTH_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv`](./OTH_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv) | Other | MarsEnergy | MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM | Modbus | 35 | 3 |
| [`COM_MarsEnergy_MARS_LocalController_V7_Unit1_Un_seul_element_WebdynSunPM.csv`](./COM_MarsEnergy_MARS_LocalController_V7_Unit1_Un_seul_element_WebdynSunPM.csv) | Logger | MarsEnergy | MARS_LocalController_V7_Unit1_Un_seul_element_WebdynSunPM | Modbus | 146 | 40 |
| [`OTH_MarsEnergy_MARS_LocalController_V7_Unit1_Un_seul_element_WebdynSunPM.csv`](./OTH_MarsEnergy_MARS_LocalController_V7_Unit1_Un_seul_element_WebdynSunPM.csv) | Other | MarsEnergy | MARS_LocalController_V7_Unit1_Un_seul_element_WebdynSunPM | Modbus | 146 | 40 |
| [`COM_MarsEnergy_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv`](./COM_MarsEnergy_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv) | Logger | MarsEnergy | MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM | Modbus | 5 | 0 |
| [`OTH_MarsEnergy_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv`](./OTH_MarsEnergy_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv) | Other | MarsEnergy | MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM | Modbus | 5 | 0 |
| [`OTH_MarsEnergy_MarsLocalController_IO_List_ModbusTCP_V7.0.csv`](./OTH_MarsEnergy_MarsLocalController_IO_List_ModbusTCP_V7.0.csv) | Other | MarsEnergy | MarsLocalController_IO_List_ModbusTCP_V7.0 | Modbus | 251 | 39 |
| [`BAT_MarsEnergy_MarsLocalController_V7.0_EMS_Slave247.csv`](./BAT_MarsEnergy_MarsLocalController_V7.0_EMS_Slave247.csv) | Battery | MarsEnergy | MarsLocalController_V7.0_EMS_Slave247 | Modbus | 12 | 3 |
| [`OTH_MarsEnergy_MarsLocalController_V7.0_EMS_Slave247.csv`](./OTH_MarsEnergy_MarsLocalController_V7.0_EMS_Slave247.csv) | Other | MarsEnergy | MarsLocalController_V7.0_EMS_Slave247 | Modbus | 12 | 3 |
| [`BAT_MarsEnergy_MarsLocalController_V7.0_RackCell_SlaveN.csv`](./BAT_MarsEnergy_MarsLocalController_V7.0_RackCell_SlaveN.csv) | Battery | MarsEnergy | MarsLocalController_V7.0_RackCell_SlaveN | Modbus | 37 | 37 |
| [`OTH_MarsEnergy_MarsLocalController_V7.0_RackCell_SlaveN.csv`](./OTH_MarsEnergy_MarsLocalController_V7.0_RackCell_SlaveN.csv) | Other | MarsEnergy | MarsLocalController_V7.0_RackCell_SlaveN | Modbus | 37 | 3 |
| [`BAT_MarsEnergy_MarsLocalController_V7.0_webdyn.csv`](./BAT_MarsEnergy_MarsLocalController_V7.0_webdyn.csv) | Battery | MarsEnergy | MarsLocalController_V7.0_webdyn | Modbus | 192 | 37 |
| [`OTH_MarsEnergy_MarsLocalController_V7.0_webdyn.csv`](./OTH_MarsEnergy_MarsLocalController_V7.0_webdyn.csv) | Other | MarsEnergy | MarsLocalController_V7.0_webdyn | Modbus | 192 | 37 |
| [`OTH_MarsEnergy_MarsRenewable.csv`](./OTH_MarsEnergy_MarsRenewable.csv) | Other | MarsEnergy | MarsRenewable | Modbus | 174 | 41 |
| [`OTH_MarsEnergy_MarsRenewable_AC.csv`](./OTH_MarsEnergy_MarsRenewable_AC.csv) | Other | MarsEnergy | MarsRenewable_AC | Modbus | 1,600 | 3 |
| [`OTH_MarsEnergy_MarsRenewable_AC_1Unit.csv`](./OTH_MarsEnergy_MarsRenewable_AC_1Unit.csv) | Other | MarsEnergy | MarsRenewable_AC_1Unit | Modbus | 16 | 3 |
| [`OTH_MarsEnergy_MarsRenewable_LC.csv`](./OTH_MarsEnergy_MarsRenewable_LC.csv) | Other | MarsEnergy | MarsRenewable_LC | Modbus | 8 | 3 |
| [`OTH_MarsEnergy_MarsRenewable_LC-AC.csv`](./OTH_MarsEnergy_MarsRenewable_LC-AC.csv) | Other | MarsEnergy | MarsRenewable_LC-AC | Modbus | 34 | 4 |
| [`OTH_MarsEnergy_MarsRenewable_LC_1Unit.csv`](./OTH_MarsEnergy_MarsRenewable_LC_1Unit.csv) | Other | MarsEnergy | MarsRenewable_LC_1Unit | Modbus | 18 | 2 |
| [`OTH_MarsEnergy_MarsRenewable_LC_EMS_V7.csv`](./OTH_MarsEnergy_MarsRenewable_LC_EMS_V7.csv) | Other | MarsEnergy | MarsRenewable_LC_EMS_V7 | Modbus | 8 | 3 |
| [`OTH_MarsEnergy_MarsRenewable_LC_V7.csv`](./OTH_MarsEnergy_MarsRenewable_LC_V7.csv) | Other | MarsEnergy | MarsRenewable_LC_V7 | Modbus | 172 | 172 |
| [`OTH_MarsEnergy_MarsRenewable_LiquidCooling-AC.csv`](./OTH_MarsEnergy_MarsRenewable_LiquidCooling-AC.csv) | Other | MarsEnergy | MarsRenewable_LiquidCooling-AC | Modbus | 18 | 1 |
| [`BAT_MarsEnergy_MarsRenewable_LocalControler_Bank_Rack_Cell.csv`](./BAT_MarsEnergy_MarsRenewable_LocalControler_Bank_Rack_Cell.csv) | Battery | MarsEnergy | MarsRenewable_LocalControler_Bank_Rack_Cell | Modbus | 83 | 6 |
| [`OTH_MarsEnergy_MarsRenewable_LocalControler_Bank_Rack_Cell.csv`](./OTH_MarsEnergy_MarsRenewable_LocalControler_Bank_Rack_Cell.csv) | Other | MarsEnergy | MarsRenewable_LocalControler_Bank_Rack_Cell | Modbus | 84 | 6 |
| [`OTH_MarsEnergy_MarsRenewable_LocalControler_PCS_Information.csv`](./OTH_MarsEnergy_MarsRenewable_LocalControler_PCS_Information.csv) | Other | MarsEnergy | MarsRenewable_LocalControler_PCS_Information | Modbus | 111 | 31 |
| [`MET_MarsEnergy_MarsRenewable_Meter.csv`](./MET_MarsEnergy_MarsRenewable_Meter.csv) | Meter | MarsEnergy | MarsRenewable_Meter | Modbus | 24 | 2 |
| [`OTH_MarsEnergy_MarsRenewable_PCS-Bank.csv`](./OTH_MarsEnergy_MarsRenewable_PCS-Bank.csv) | Other | MarsEnergy | MarsRenewable_PCS-Bank | Modbus | 84 | 32 |
| [`OTH_MarsEnergy_MarsRenewable_PV.csv`](./OTH_MarsEnergy_MarsRenewable_PV.csv) | Other | MarsEnergy | MarsRenewable_PV | Modbus | 52 | 30 |
| [`BAT_MarsEnergy_MarsRenewable_Rack-Cell.csv`](./BAT_MarsEnergy_MarsRenewable_Rack-Cell.csv) | Battery | MarsEnergy | MarsRenewable_Rack-Cell | Modbus | 42 | 3 |
| [`BAT_MarsEnergy_MarsRenewable_Rack_V7.csv`](./BAT_MarsEnergy_MarsRenewable_Rack_V7.csv) | Battery | MarsEnergy | MarsRenewable_Rack_V7 | Modbus | 110 | 7 |
| [`OTH_MarsEnergy_MarsRenewable_Rack_V7.csv`](./OTH_MarsEnergy_MarsRenewable_Rack_V7.csv) | Other | MarsEnergy | MarsRenewable_Rack_V7 | Modbus | 110 | 7 |
| [`INV_MarsEnergy_Meter.csv`](./INV_MarsEnergy_Meter.csv) | Inverter | MarsEnergy | Meter | Modbus | 24 | 2 |
| [`BAT_MarsEnergy_PCS-Bank.csv`](./BAT_MarsEnergy_PCS-Bank.csv) | Battery | MarsEnergy | PCS-Bank | Modbus | 86 | 34 |
| [`OTH_MarsEnergy_PCS-Bank.csv`](./OTH_MarsEnergy_PCS-Bank.csv) | Other | MarsEnergy | PCS-Bank | Modbus | 84 | 32 |
| [`INV_MarsEnergy_PV.csv`](./INV_MarsEnergy_PV.csv) | Inverter | MarsEnergy | PV | Modbus | 52 | 30 |
| [`OTH_MarsEnergy_PV.csv`](./OTH_MarsEnergy_PV.csv) | Other | MarsEnergy | PV | Modbus | 52 | 30 |
| [`BAT_MarsEnergy_Rack-Cell.csv`](./BAT_MarsEnergy_Rack-Cell.csv) | Battery | MarsEnergy | Rack-Cell | Modbus | 42 | 3 |
| [`INV_MarsEnergy_Rack-Cell.csv`](./INV_MarsEnergy_Rack-Cell.csv) | Inverter | MarsEnergy | Rack-Cell | Modbus | 42 | 3 |
| [`INV_MarsEnergy_Rack_V7.csv`](./INV_MarsEnergy_Rack_V7.csv) | Inverter | MarsEnergy | Rack_V7 | Modbus | 110 | 7 |
| [`INV_MarsEnergy_AirCooling.csv`](./INV_MarsEnergy_AirCooling.csv) | Inverter | MarsRenewable | AirCooling | Modbus | 1,600 | 2 |
| [`INV_MarsEnergy_LC_AC.csv`](./INV_MarsEnergy_LC_AC.csv) | Inverter | MarsRenewable | LC-AC | Modbus | 3,400 | 3 |
| [`INV_MarsEnergy_LiquidCooling.csv`](./INV_MarsEnergy_LiquidCooling.csv) | Inverter | MarsRenewable | LiquidCooling | Modbus | 1,800 | 1 |
| [`MET_MarsEnergy_Meter.csv`](./MET_MarsEnergy_Meter.csv) | Meter | MarsRenewable | Meter | Modbus | 11 | 3 |
| [`SEN_MarsEnergy_PV.csv`](./SEN_MarsEnergy_PV.csv) | Inverter | MarsRenewable | PV | Modbus | 49 | 45 |
| [`OTH_Maximet_GMX600.csv`](./OTH_Maximet_GMX600.csv) | Other | Maximet | GMX600 | Modbus | 11 | 11 |
| [`SEN_Maximet_GMX600.csv`](./SEN_Maximet_GMX600.csv) | Sensor | Maximet | GMX600 | Modbus | 11 | 11 |
| [`SEN_MeierNT_ADL-SR.csv`](./SEN_MeierNT_ADL-SR.csv) | Sensor | MeierNT | ADL-SR | Modbus | 18 | 1 |
| [`OTH_MeierNT_MEIER_ADL-SR.csv`](./OTH_MeierNT_MEIER_ADL-SR.csv) | Other | MeierNT | MEIER_ADL-SR | Modbus | 18 | 4 |
| [`SEN_MeierNT_MEIER_ADL-SR.csv`](./SEN_MeierNT_MEIER_ADL-SR.csv) | Sensor | MeierNT | MEIER_ADL-SR | Modbus | 18 | 4 |
| [`OTH_Monichek_CombinerBox.csv`](./OTH_Monichek_CombinerBox.csv) | Other | Monichek | CombinerBox | Modbus | 524 | 56 |
| [`INV_Nidec_SOLARGATE_WebdynSunPM_complet.csv`](./INV_Nidec_SOLARGATE_WebdynSunPM_complet.csv) | Inverter | Nidec | SOLARGATE_WebdynSunPM_complet | Modbus | 73 | 12 |
| [`MET_OG_NHEOBOX_OG_NHEOBOX.csv`](./MET_OG_NHEOBOX_OG_NHEOBOX.csv) | Meter | OG | NHEOBOX_OG_NHEOBOX | Modbus | 96 | 15 |
| [`OTH_OG_NHEOBOX_OG_NHEOBOX.csv`](./OTH_OG_NHEOBOX_OG_NHEOBOX.csv) | Other | OG | NHEOBOX_OG_NHEOBOX | Modbus | 96 | 15 |
| [`MET_Polier_MTR5LMOD.csv`](./MET_Polier_MTR5LMOD.csv) | Meter | POLIER | MTR5LMOD | Modbus | 33 | 4 |
| [`MET_Polier_MTROGMOD.csv`](./MET_Polier_MTROGMOD.csv) | Meter | POLIER | MTROGMOD | Modbus | 308 | 29 |
| [`MET_Polier_PROMOD380.csv`](./MET_Polier_PROMOD380.csv) | Meter | POLIER | PROMOD380 | Modbus | 88 | 31 |
| [`INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv`](./INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv) | Inverter | POWERELECTRONICS | Freesun_MODBUS_POWERELECTRONICS_RTU | Modbus | 15 | 4 |
| [`INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv`](./INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv) | Inverter | POWERELECTRONICS | Freesun_MODBUS_POWERELECTRONICS_TCP | Modbus | 15 | 4 |
| [`MET_Panasonic_AKW8115.csv`](./MET_Panasonic_AKW8115.csv) | Meter | Panasonic | AKW8115 | Modbus | 14 | 4 |
| [`OTH_Panasonic_AKW8115.csv`](./OTH_Panasonic_AKW8115.csv) | Other | Panasonic | AKW8115 | Modbus | 14 | 4 |
| [`MET_PhoenixContact_EEM-MA600.csv`](./MET_PhoenixContact_EEM-MA600.csv) | Meter | PhoenixContact | EEM-MA600 | Modbus | 87 | 7 |
| [`OTH_PhoenixContact_EEM-MA600.csv`](./OTH_PhoenixContact_EEM-MA600.csv) | Other | PhoenixContact | EEM-MA600 | Modbus | 87 | 7 |
| [`OTH_Polier_MTR5LMOD.csv`](./OTH_Polier_MTR5LMOD.csv) | Other | Polier | MTR5LMOD | Modbus | 33 | 17 |
| [`OTH_Power_One.csv`](./OTH_Power_One.csv) | Other | Power | One | Modbus | 58 | 12 |
| [`INV_PowerElectronics_Freesun_INV_POWER_ELECTRONICS_FREESUN_HE_HEC_HES.csv`](./INV_PowerElectronics_Freesun_INV_POWER_ELECTRONICS_FREESUN_HE_HEC_HES.csv) | Inverter | PowerElectronics | Freesun_INV_POWER_ELECTRONICS_FREESUN_HE_HEC_HES | Modbus | 246 | 15 |
| [`OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv`](./OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv) | Other | PowerElectronics | Freesun_MODBUS_POWERELECTRONICS_RTU | Modbus | 15 | 4 |
| [`OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv`](./OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv) | Other | PowerElectronics | Freesun_MODBUS_POWERELECTRONICS_TCP | Modbus | 15 | 4 |
| [`INV_PowerElectronics_Freesun_POWER_ELECTRONICS_FREESUN_HE_HEC_HES.csv`](./INV_PowerElectronics_Freesun_POWER_ELECTRONICS_FREESUN_HE_HEC_HES.csv) | Inverter | PowerElectronics | Freesun_POWER_ELECTRONICS_FREESUN_HE_HEC_HES | Modbus | 246 | 15 |
| [`INV_Refusol_Generic.csv`](./INV_Refusol_Generic.csv) | Inverter | REFUSOL | GENERIC | Modbus | 536 | 36 |
| [`OTH_Refusol_GENERIC.csv`](./OTH_Refusol_GENERIC.csv) | Other | Refusol | GENERIC | Modbus | 536 | 36 |
| [`OTH_Refusol_REFU-Modbus_Protocol_English_G3_V1.17_ONLY_SOLAR.csv`](./OTH_Refusol_REFU-Modbus_Protocol_English_G3_V1.17_ONLY_SOLAR.csv) | Other | Refusol | REFU-Modbus_Protocol_English_G3_V1.17_ONLY_SOLAR | Modbus | 144 | 17 |
| [`OTH_Riello_AROS_GENERIC.csv`](./OTH_Riello_AROS_GENERIC.csv) | Other | Riello | AROS_GENERIC | Modbus | 48 | 17 |
| [`BAT_Robestec_BESS.csv`](./BAT_Robestec_BESS.csv) | Battery | Robestec | BESS | Modbus | 1,723 | 17 |
| [`COM_Robestec_SJ2025B2846SJ2025B2846yubangmodbusTcp-en_1.csv`](./COM_Robestec_SJ2025B2846SJ2025B2846yubangmodbusTcp-en_1.csv) | Logger | Robestec | SJ2025B2846SJ2025B2846yubangmodbusTcp-en_1 | Modbus | 1,457 | 17 |
| [`MET_Saci_MDL3-MID_M1DL3T-MID.csv`](./MET_Saci_MDL3-MID_M1DL3T-MID.csv) | Meter | SACI | MDL3-MID_M1DL3T-MID | Modbus | 86 | 14 |
| [`INV_SAJ_C6.csv`](./INV_SAJ_C6.csv) | Inverter | SAJ | C6 | Modbus | 138 | 48 |
| [`OTH_SAJ_C6.csv`](./OTH_SAJ_C6.csv) | Other | SAJ | C6 | Modbus | 138 | 48 |
| [`OTH_SAJ_CM2-EMS.csv`](./OTH_SAJ_CM2-EMS.csv) | Other | SAJ | CM2-EMS | Modbus | 66 | 3 |
| [`INV_SAJ_Generic.csv`](./INV_SAJ_Generic.csv) | Inverter | SAJ | Generic | Modbus | 45 | 8 |
| [`OTH_SAJ_Generic.csv`](./OTH_SAJ_Generic.csv) | Other | SAJ | Generic | Modbus | 45 | 8 |
| [`OTH_SAJ_Modbus_protocol_for_SAJ_CM2-EMS_Module-EN-2026.02.csv`](./OTH_SAJ_Modbus_protocol_for_SAJ_CM2-EMS_Module-EN-2026.02.csv) | Other | SAJ | Modbus_protocol_for_SAJ_CM2-EMS_Module-EN-2026.02 | Modbus | 66 | 3 |
| [`INV_SAJ_PLUS_R5.csv`](./INV_SAJ_PLUS_R5.csv) | Inverter | SAJ | PLUS_R5 | Modbus | 73 | 27 |
| [`OTH_SAJ_PLUS_R5.csv`](./OTH_SAJ_PLUS_R5.csv) | Other | SAJ | PLUS_R5 | Modbus | 73 | 27 |
| [`INV_SAJ_R5_R6_C6.csv`](./INV_SAJ_R5_R6_C6.csv) | Inverter | SAJ | R5_R6_C6 | Modbus | 391 | 51 |
| [`OTH_SAJ_R5_R6_C6.csv`](./OTH_SAJ_R5_R6_C6.csv) | Other | SAJ | R5_R6_C6 | Modbus | 391 | 51 |
| [`INV_SAJ_R6.csv`](./INV_SAJ_R6.csv) | Inverter | SAJ | R6 | Modbus | 313 | 8 |
| [`OTH_SAJ_R6.csv`](./OTH_SAJ_R6.csv) | Other | SAJ | R6 | Modbus | 313 | 10 |
| [`INV_SAJ_R6_v2.csv`](./INV_SAJ_R6_v2.csv) | Inverter | SAJ | R6_v2 | Modbus | 221 | 29 |
| [`OTH_SAJ_R6_v2.csv`](./OTH_SAJ_R6_v2.csv) | Other | SAJ | R6_v2 | Modbus | 221 | 29 |
| [`INV_SAJ_sunTrioTL.csv`](./INV_SAJ_sunTrioTL.csv) | Inverter | SAJ | SunTrioTL | Modbus | 45 | 8 |
| [`INV_SAJ_SunUno-Trio.csv`](./INV_SAJ_SunUno-Trio.csv) | Inverter | SAJ | SunUno-Trio | Modbus | 201 | 9 |
| [`OTH_SAJ_SunUno-Trio.csv`](./OTH_SAJ_SunUno-Trio.csv) | Other | SAJ | SunUno-Trio | Modbus | 201 | 9 |
| [`INV_SAJ_Sununo_Plus_Suntrio.csv`](./INV_SAJ_Sununo_Plus_Suntrio.csv) | Inverter | SAJ | Sununo_Plus_Suntrio | Modbus | 201 | 8 |
| [`OTH_SAJ_sunTrioTL.csv`](./OTH_SAJ_sunTrioTL.csv) | Other | SAJ | sunTrioTL | Modbus | 45 | 7 |
| [`INV_Salicru_EQUINOX2.csv`](./INV_Salicru_EQUINOX2.csv) | Inverter | SALICRU | EQUINOX2 | Modbus | 249 | 64 |
| [`INV_Satcon_Generic.csv`](./INV_Satcon_Generic.csv) | Inverter | SATCON | Generic | Modbus | 95 | 11 |
| [`INV_Satcon_PowerGate.csv`](./INV_Satcon_PowerGate.csv) | Inverter | SATCON | PowerGate | Modbus | 96 | 3 |
| [`INV_Schneider_30-33-50-CL.csv`](./INV_Schneider_30-33-50-CL.csv) | Inverter | SCHNEIDER | 30-33-50-CL | Modbus | 73 | 41 |
| [`SEN_SESAutomation_WT-485-PRO.csv`](./SEN_SESAutomation_WT-485-PRO.csv) | Sensor | SESAutomation | WT-485-PRO | Modbus | 10 | 3 |
| [`OTH_Siebert_AFFICHEUR_SIEBERT_XC4XX.csv`](./OTH_Siebert_AFFICHEUR_SIEBERT_XC4XX.csv) | Other | SIEBERT | AFFICHEUR_SIEBERT_XC4XX | Modbus | 3 | 3 |
| [`OTH_Siebert_AFFICHEUR_SIEBERT_XC4XX_TCP.csv`](./OTH_Siebert_AFFICHEUR_SIEBERT_XC4XX_TCP.csv) | Other | SIEBERT | AFFICHEUR_SIEBERT_XC4XX_TCP | Modbus | 3 | 3 |
| [`OTH_Siebert_XC4XX.csv`](./OTH_Siebert_XC4XX.csv) | Other | SIEBERT | XC4XX | Modbus | 3 | 2 |
| [`INV_Siel_10TL.csv`](./INV_Siel_10TL.csv) | Inverter | SIEL | 10TL | Modbus | 17 | 1 |
| [`INV_Siel_10_TL_SOLAR_INVERTER.csv`](./INV_Siel_10_TL_SOLAR_INVERTER.csv) | Inverter | SIEL | 10_TL_SOLAR_INVERTER | Modbus | 17 | 0 |
| [`INV_Siel_DSP-DSPX.csv`](./INV_Siel_DSP-DSPX.csv) | Inverter | SIEL | DSP-DSPX | Modbus | 21 | 1 |
| [`INV_Siel_DSPX.csv`](./INV_Siel_DSPX.csv) | Inverter | SIEL | DSPX | Modbus | 21 | 0 |
| [`INV_Siel_FUEL_CELL.csv`](./INV_Siel_FUEL_CELL.csv) | Inverter | SIEL | FUEL_CELL | Modbus | 22 | 0 |
| [`INV_Siel_Monophase.csv`](./INV_Siel_Monophase.csv) | Inverter | SIEL | Monophase | Modbus | 11 | 1 |
| [`INV_Siel_RECTIFIERS.csv`](./INV_Siel_RECTIFIERS.csv) | Inverter | SIEL | RECTIFIERS | Modbus | 19 | 0 |
| [`INV_Siel_STS.csv`](./INV_Siel_STS.csv) | Inverter | SIEL | STS | Modbus | 25 | 0 |
| [`INV_Siel_Solar_Converters.csv`](./INV_Siel_Solar_Converters.csv) | Inverter | SIEL | Solar_Converters | Modbus | 19 | 0 |
| [`INV_Siel_UPS_UNITS.csv`](./INV_Siel_UPS_UNITS.csv) | Inverter | SIEL | UPS_UNITS | Modbus | 22 | 0 |
| [`INV_Siel_WINDWAVE.csv`](./INV_Siel_WINDWAVE.csv) | Inverter | SIEL | WINDWAVE | Modbus | 21 | 0 |
| [`INV_Sigenergy_Generic.csv`](./INV_Sigenergy_Generic.csv) | Inverter | SIGENERGY | Generic | Modbus | 194 | 61 |
| [`INV_Sineng_HYBRID.csv`](./INV_Sineng_HYBRID.csv) | Inverter | SINENG | HYBRID | Modbus | 443 | 24 |
| [`BAT_SLEnergy_Battery.csv`](./BAT_SLEnergy_Battery.csv) | Battery | SLEnergy | Battery | Modbus | 188 | 12 |
| [`MET_SLEnergy_ConsoMeter.csv`](./MET_SLEnergy_ConsoMeter.csv) | Meter | SLEnergy | ConsoMeter | Modbus | 4 | 2 |
| [`INV_SLEnergy_DG.csv`](./INV_SLEnergy_DG.csv) | Inverter | SLEnergy | DG | Modbus | 615 | 42 |
| [`OTH_SLEnergy_DG.csv`](./OTH_SLEnergy_DG.csv) | Other | SLEnergy | DG | Modbus | 615 | 39 |
| [`BAT_SLEnergy_Generator.csv`](./BAT_SLEnergy_Generator.csv) | Battery | SLEnergy | Generator | Modbus | 16 | 1 |
| [`MET_SLEnergy_GridMeter.csv`](./MET_SLEnergy_GridMeter.csv) | Meter | SLEnergy | GridMeter | Modbus | 65 | 8 |
| [`INV_SLEnergy_Inverter.csv`](./INV_SLEnergy_Inverter.csv) | Inverter | SLEnergy | Inverter | Modbus | 263 | 28 |
| [`MET_SLEnergy_OffGridMeter.csv`](./MET_SLEnergy_OffGridMeter.csv) | Meter | SLEnergy | OffGridMeter | Modbus | 33 | 2 |
| [`BAT_SLEnergy_PCS.csv`](./BAT_SLEnergy_PCS.csv) | Battery | SLEnergy | PCS | Modbus | 683 | 47 |
| [`OTH_SLEnergy_PCS.csv`](./OTH_SLEnergy_PCS.csv) | Other | SLEnergy | PCS | Modbus | 683 | 45 |
| [`INV_SLEnergy_SL_Energy_Generic.csv`](./INV_SLEnergy_SL_Energy_Generic.csv) | Inverter | SLEnergy | SL_Energy_Generic | Modbus | 597 | 38 |
| [`OTH_SLEnergy_SL_Energy_Generic.csv`](./OTH_SLEnergy_SL_Energy_Generic.csv) | Other | SLEnergy | SL_Energy_Generic | Modbus | 597 | 39 |
| [`INV_SLEnergy_SL_Energy_Generic_version_1.csv`](./INV_SLEnergy_SL_Energy_Generic_version_1.csv) | Inverter | SLEnergy | SL_Energy_Generic_version_1 | Modbus | 592 | 38 |
| [`OTH_SLEnergy_SL_Energy_Generic_version_1.csv`](./OTH_SLEnergy_SL_Energy_Generic_version_1.csv) | Other | SLEnergy | SL_Energy_Generic_version_1 | Modbus | 592 | 39 |
| [`OTH_SLEnergy_Work.csv`](./OTH_SLEnergy_Work.csv) | Other | SLEnergy | Work | Modbus | 612 | 612 |
| [`BAT_SLEnergy_bms5_p1_1.csv`](./BAT_SLEnergy_bms5_p1_1.csv) | Battery | SLEnergy | bms5_p1_1 | Modbus | 213 | 9 |
| [`OTH_SLEnergy_bms5_p1_1.csv`](./OTH_SLEnergy_bms5_p1_1.csv) | Other | SLEnergy | bms5_p1_1 | Modbus | 213 | 9 |
| [`OTH_SLEnergy_bms5_p1_1_1.csv`](./OTH_SLEnergy_bms5_p1_1_1.csv) | Other | SLEnergy | bms5_p1_1_1 | Modbus | 213 | 9 |
| [`OTH_SMA_CORE2_STP_110-60_Sunspec_Modbus_Model_PICS_Overview_V12.csv`](./OTH_SMA_CORE2_STP_110-60_Sunspec_Modbus_Model_PICS_Overview_V12.csv) | Other | SMA | CORE2_STP_110-60_Sunspec_Modbus_Model_PICS_Overview_V12 | Modbus | 206 | 61 |
| [`COM_SMA_ClusterController.csv`](./COM_SMA_ClusterController.csv) | Logger | SMA | ClusterController | Modbus | 24 | 23 |
| [`COM_SMA_ComGateway.csv`](./COM_SMA_ComGateway.csv) | Logger | SMA | ComGateway | Modbus | 22 | 21 |
| [`INV_SMA_Core1.csv`](./INV_SMA_Core1.csv) | Inverter | SMA | Core1 | Modbus | 105 | 70 |
| [`INV_SMA_Core1_Core2.csv`](./INV_SMA_Core1_Core2.csv) | Inverter | SMA | Core1_Core2 | Modbus | 77 | 77 |
| [`INV_SMA_Core2.csv`](./INV_SMA_Core2.csv) | Inverter | SMA | Core2 | Modbus | 1,211 | 71 |
| [`COM_SMA_DataManager_M.csv`](./COM_SMA_DataManager_M.csv) | Logger | SMA | DataManager-M | Modbus | 28 | 27 |
| [`MET_SMA_EnergyMeter.csv`](./MET_SMA_EnergyMeter.csv) | Meter | SMA | EnergyMeter | Modbus | 26 | 25 |
| [`INV_SMA_Generic.csv`](./INV_SMA_Generic.csv) | Inverter | SMA | Generic | Modbus | 857 | 96 |
| [`MET_SMA_HomeManager.csv`](./MET_SMA_HomeManager.csv) | Meter | SMA | HomeManager | Modbus | 26 | 25 |
| [`INV_SMA_MODBUS_SMA_Sunspec.csv`](./INV_SMA_MODBUS_SMA_Sunspec.csv) | Inverter | SMA | MODBUS_SMA_Sunspec | Modbus | 45 | 20 |
| [`INV_SMA_Modbus.csv`](./INV_SMA_Modbus.csv) | Inverter | SMA | Modbus | Modbus | 29 | 29 |
| [`OTH_SMA_Revised.csv`](./OTH_SMA_Revised.csv) | Other | SMA | Revised | Modbus | 36 | 17 |
| [`INV_SMA_SC2xxx.csv`](./INV_SMA_SC2xxx.csv) | Inverter | SMA | SC2xxx | Modbus | 78 | 5 |
| [`INV_SMA_SC_Webbox.csv`](./INV_SMA_SC_Webbox.csv) | Inverter | SMA | SC_Webbox | Modbus | 350 | 52 |
| [`OTH_SMA_STP-X_STP_xx-50_Sunspec_Modbus_Model_PICS_Overview_V1.csv`](./OTH_SMA_STP-X_STP_xx-50_Sunspec_Modbus_Model_PICS_Overview_V1.csv) | Other | SMA | STP-X_STP_xx-50_Sunspec_Modbus_Model_PICS_Overview_V1 | Modbus | 820 | 24 |
| [`INV_SMA_STPCORE1_50-41.csv`](./INV_SMA_STPCORE1_50-41.csv) | Inverter | SMA | STPCORE1_50-41 | Modbus | 105 | 70 |
| [`INV_SMA_STPX.csv`](./INV_SMA_STPX.csv) | Inverter | SMA | STPX | Modbus | 509 | 46 |
| [`INV_SMA_STPX_27_30_50_60.csv`](./INV_SMA_STPX_27_30_50_60.csv) | Inverter | SMA | STPX_27_30_50_60 | Modbus | 391 | 6 |
| [`INV_SMA_STP_110-60.csv`](./INV_SMA_STP_110-60.csv) | Inverter | SMA | STP_110-60 | Modbus | 1,211 | 71 |
| [`INV_SMA_STP_125-70.csv`](./INV_SMA_STP_125-70.csv) | Inverter | SMA | STP_125-70 | Modbus | 464 | 86 |
| [`INV_SMA_STP_5-10-SE40.csv`](./INV_SMA_STP_5-10-SE40.csv) | Inverter | SMA | STP_5-10-SE40 | Modbus | 372 | 59 |
| [`INV_SMA_STP_X.csv`](./INV_SMA_STP_X.csv) | Inverter | SMA | STP_X | Modbus | 435 | 56 |
| [`INV_SMA_SUNNY_CENTRAL.csv`](./INV_SMA_SUNNY_CENTRAL.csv) | Inverter | SMA | SUNNY_CENTRAL | Modbus | 182 | 5 |
| [`INV_SMA_SunSpec_Inverter.csv`](./INV_SMA_SunSpec_Inverter.csv) | Inverter | SMA | SunSpec-Inverter | Modbus | 29 | 29 |
| [`INV_SMA_SunSpec_inverter_SMA_Solar_Inverter_2.83.03.R_modbusTCP_STP25000TL30.csv`](./INV_SMA_SunSpec_inverter_SMA_Solar_Inverter_2.83.03.R_modbusTCP_STP25000TL30.csv) | Inverter | SMA | SunSpec_inverter_SMA_Solar_Inverter_2.83.03.R_modbusTCP_STP25000TL30 | Modbus | 591 | 37 |
| [`INV_SMA_SunnyBoy.csv`](./INV_SMA_SunnyBoy.csv) | Inverter | SMA | SunnyBoy | Modbus | 36 | 35 |
| [`INV_SMA_SunnyBoyStorage.csv`](./INV_SMA_SunnyBoyStorage.csv) | Inverter | SMA | SunnyBoyStorage | Modbus | 45 | 44 |
| [`INV_SMA_SunnyCentral.csv`](./INV_SMA_SunnyCentral.csv) | Inverter | SMA | SunnyCentral | Modbus | 182 | 5 |
| [`INV_SMA_SunnyCentral_Storage.csv`](./INV_SMA_SunnyCentral_Storage.csv) | Inverter | SMA | SunnyCentral-Storage | Modbus | 193 | 16 |
| [`INV_SMA_SunnyIsland.csv`](./INV_SMA_SunnyIsland.csv) | Inverter | SMA | SunnyIsland | Modbus | 35 | 34 |
| [`INV_SMA_SunnyTripower.csv`](./INV_SMA_SunnyTripower.csv) | Inverter | SMA | SunnyTripower | Modbus | 59 | 59 |
| [`INV_SMA_SunnyTripower_SmartEnergy.csv`](./INV_SMA_SunnyTripower_SmartEnergy.csv) | Inverter | SMA | SunnyTripower-SmartEnergy | Modbus | 372 | 59 |
| [`INV_SMA_SunnyTripower_X.csv`](./INV_SMA_SunnyTripower_X.csv) | Inverter | SMA | SunnyTripower-X | Modbus | 435 | 56 |
| [`INV_SMA_Sunspec.csv`](./INV_SMA_Sunspec.csv) | Inverter | SMA | Sunspec | Modbus | 29 | 29 |
| [`INV_SNEF_Modbus.csv`](./INV_SNEF_Modbus.csv) | Inverter | SNEF | Modbus | Modbus | 16 | 13 |
| [`OTH_SNEF_Modbus.csv`](./OTH_SNEF_Modbus.csv) | Other | SNEF | Modbus | Modbus | 16 | 13 |
| [`MET_Socomec_E33_E43.csv`](./MET_Socomec_E33_E43.csv) | Meter | SOCOMEC | E33_E43 | Modbus | 73 | 4 |
| [`MET_Socomec_ElectricMotionWebdyn_modbusRTU_Meter_SOCOMEC_DIRIS_B10.csv`](./MET_Socomec_ElectricMotionWebdyn_modbusRTU_Meter_SOCOMEC_DIRIS_B10.csv) | Meter | SOCOMEC | ElectricMotionWebdyn_modbusRTU_Meter_SOCOMEC_DIRIS_B10 | Modbus | 19 | 7 |
| [`MET_Socomec_W_SOCOMEC_COUNTIS_E23.csv`](./MET_Socomec_W_SOCOMEC_COUNTIS_E23.csv) | Meter | SOCOMEC | W_SOCOMEC_COUNTIS_E23 | Modbus | 59 | 18 |
| [`MET_Socomec_W_SOCOMEC_DIRIS_B10.csv`](./MET_Socomec_W_SOCOMEC_DIRIS_B10.csv) | Meter | SOCOMEC | W_SOCOMEC_DIRIS_B10 | Modbus | 23 | 7 |
| [`MET_Socomec_W_SOCOMEC_Diris_A40.csv`](./MET_Socomec_W_SOCOMEC_Diris_A40.csv) | Meter | SOCOMEC | W_SOCOMEC_Diris_A40 | Modbus | 35 | 17 |
| [`MET_Socomec_W_SOCOMEC_E33_E43.csv`](./MET_Socomec_W_SOCOMEC_E33_E43.csv) | Meter | SOCOMEC | W_SOCOMEC_E33_E43 | Modbus | 73 | 18 |
| [`INV_SofarSolar_GENERIC.csv`](./INV_SofarSolar_GENERIC.csv) | Inverter | SOFARSOLAR | GENERIC | Modbus | 547 | 32 |
| [`INV_SofarSolar_HYD_G3.csv`](./INV_SofarSolar_HYD_G3.csv) | Inverter | SOFARSOLAR | HYD_G3 | Modbus | 538 | 32 |
| [`INV_SolarEdge_Modbus_Sunspec.csv`](./INV_SolarEdge_Modbus_Sunspec.csv) | Inverter | SOLAREDGE | Modbus_Sunspec | Modbus | 46 | 10 |
| [`MET_SolarLog_PRO380_WORK.csv`](./MET_SolarLog_PRO380_WORK.csv) | Meter | SOLARLOG | PRO380_WORK | Modbus | 148 | 30 |
| [`BAT_SolaX_EMS1000-PRO.csv`](./BAT_SolaX_EMS1000-PRO.csv) | Battery | SOLAX | EMS1000-PRO | Modbus | 347 | 37 |
| [`INV_Solplanet_Generic.csv`](./INV_Solplanet_Generic.csv) | Inverter | SOLPLANET | GENERIC | Modbus | 625 | 60 |
| [`SEN_SRegeltechnik_S_REGELTECHNIK_HYDRASGARD.csv`](./SEN_SRegeltechnik_S_REGELTECHNIK_HYDRASGARD.csv) | Sensor | SRegeltechnik | S_REGELTECHNIK_HYDRASGARD | Modbus | 8 | 0 |
| [`OTH_SRegeltechnik_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv`](./OTH_SRegeltechnik_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv) | Other | SRegeltechnik | S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD | Modbus | 8 | 1 |
| [`SEN_SRegeltechnik_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv`](./SEN_SRegeltechnik_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv) | Sensor | SRegeltechnik | S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD | Modbus | 8 | 1 |
| [`OTH_Saci_MDL3-MID_M1DL3T-MID.csv`](./OTH_Saci_MDL3-MID_M1DL3T-MID.csv) | Other | Saci | MDL3-MID_M1DL3T-MID | Modbus | 86 | 14 |
| [`OTH_Salgueirinha_Modbus_Registers.csv`](./OTH_Salgueirinha_Modbus_Registers.csv) | Other | Salgueirinha | Modbus_Registers | Modbus | 1,294 | 34 |
| [`OTH_Satcon_Generic.csv`](./OTH_Satcon_Generic.csv) | Other | Satcon | Generic | Modbus | 95 | 11 |
| [`OTH_Schneider_30-33-50-CL.csv`](./OTH_Schneider_30-33-50-CL.csv) | Other | Schneider | 30-33-50-CL | Modbus | 60 | 41 |
| [`OTH_Schneider_3255lite.csv`](./OTH_Schneider_3255lite.csv) | Other | Schneider | 3255lite | Modbus | 105 | 11 |
| [`INV_Schneider_CL30-33-50.csv`](./INV_Schneider_CL30-33-50.csv) | Inverter | Schneider | CL30-33-50 | Modbus | 58 | 41 |
| [`OTH_Schneider_CL30-33-50.csv`](./OTH_Schneider_CL30-33-50.csv) | Other | Schneider | CL30-33-50 | Modbus | 58 | 41 |
| [`INV_Schneider_CONEXT_CORE_XC.csv`](./INV_Schneider_CONEXT_CORE_XC.csv) | Inverter | Schneider | CONEXT_CORE_XC | Modbus | 210 | 42 |
| [`INV_Schneider_CONEXT_CORE_XC_RTU.csv`](./INV_Schneider_CONEXT_CORE_XC_RTU.csv) | Inverter | Schneider | CONEXT_CORE_XC_RTU | Modbus | 210 | 42 |
| [`INV_Schneider_Conext_CL.csv`](./INV_Schneider_Conext_CL.csv) | Inverter | Schneider | Conext_CL | Modbus | 61 | 44 |
| [`INV_Schneider_Conext_CL_30-33-50.csv`](./INV_Schneider_Conext_CL_30-33-50.csv) | Inverter | Schneider | Conext_CL_30-33-50 | Modbus | 61 | 5 |
| [`OTH_Schneider_DISPLAY_SCHNEIDER_DISPLAYTECHNIK.csv`](./OTH_Schneider_DISPLAY_SCHNEIDER_DISPLAYTECHNIK.csv) | Other | Schneider | DISPLAY_SCHNEIDER_DISPLAYTECHNIK | Modbus | 3 | 3 |
| [`MET_Schneider_EM1200.csv`](./MET_Schneider_EM1200.csv) | Meter | Schneider | EM1200 | Modbus | 96 | 24 |
| [`OTH_Schneider_EM1200.csv`](./OTH_Schneider_EM1200.csv) | Other | Schneider | EM1200 | Modbus | 96 | 24 |
| [`MET_Schneider_EM1200-EM1220.csv`](./MET_Schneider_EM1200-EM1220.csv) | Meter | Schneider | EM1200-EM1220 | Modbus | 96 | 4 |
| [`MET_Schneider_EM6433.csv`](./MET_Schneider_EM6433.csv) | Meter | Schneider | EM6433 | Modbus | 17 | 11 |
| [`OTH_Schneider_EM6433.csv`](./OTH_Schneider_EM6433.csv) | Other | Schneider | EM6433 | Modbus | 17 | 11 |
| [`MET_Schneider_EM6433_evo.csv`](./MET_Schneider_EM6433_evo.csv) | Meter | Schneider | EM6433_evo | Modbus | 17 | 11 |
| [`OTH_Schneider_EM6433_evo.csv`](./OTH_Schneider_EM6433_evo.csv) | Other | Schneider | EM6433_evo | Modbus | 17 | 11 |
| [`MET_Schneider_Generic.csv`](./MET_Schneider_Generic.csv) | Meter | Schneider | Generic | Modbus | 517 | 6 |
| [`OTH_Schneider_Generic.csv`](./OTH_Schneider_Generic.csv) | Other | Schneider | Generic | Modbus | 7 | 4 |
| [`MET_Schneider_IEM2050.csv`](./MET_Schneider_IEM2050.csv) | Meter | Schneider | IEM2050 | Modbus | 69 | 13 |
| [`OTH_Schneider_IEM2050.csv`](./OTH_Schneider_IEM2050.csv) | Other | Schneider | IEM2050 | Modbus | 69 | 13 |
| [`MET_Schneider_IEM_3xxx.csv`](./MET_Schneider_IEM_3xxx.csv) | Meter | Schneider | IEM_3xxx | Modbus | 55 | 4 |
| [`OTH_Schneider_ION9000_Modbus_Map_004.005.000.csv`](./OTH_Schneider_ION9000_Modbus_Map_004.005.000.csv) | Other | Schneider | ION9000_Modbus_Map_004.005.000 | Modbus | 3,945 | 18 |
| [`OTH_Schneider_MODBUS_SCHNEIDER_NSX.csv`](./OTH_Schneider_MODBUS_SCHNEIDER_NSX.csv) | Other | Schneider | MODBUS_SCHNEIDER_NSX | Modbus | 10 | 0 |
| [`INV_Schneider_Modbus_RL.csv`](./INV_Schneider_Modbus_RL.csv) | Inverter | Schneider | Modbus_RL | Modbus | 11 | 5 |
| [`OTH_Schneider_Modbus_RL.csv`](./OTH_Schneider_Modbus_RL.csv) | Other | Schneider | Modbus_RL | Modbus | 11 | 5 |
| [`MET_Schneider_PM5000_IEM3000.csv`](./MET_Schneider_PM5000_IEM3000.csv) | Meter | Schneider | PM5000-iEM3000 | Modbus | 24 | 24 |
| [`OTH_Schneider_PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0.csv`](./OTH_Schneider_PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0.csv) | Other | Schneider | PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0 | Modbus | 28 | 3 |
| [`OTH_Schneider_PM5563_PM5563RD_ModbusRegisterList_v2.3.0-ORIGINAL.csv`](./OTH_Schneider_PM5563_PM5563RD_ModbusRegisterList_v2.3.0-ORIGINAL.csv) | Other | Schneider | PM5563_PM5563RD_ModbusRegisterList_v2.3.0-ORIGINAL | Modbus | 28 | 3 |
| [`OTH_Schneider_PM556x_PublicModbusRegisterList_10th_August-23.csv`](./OTH_Schneider_PM556x_PublicModbusRegisterList_10th_August-23.csv) | Other | Schneider | PM556x_PublicModbusRegisterList_10th_August-23 | Modbus | 28 | 3 |
| [`MET_Schneider_PM5xxx.csv`](./MET_Schneider_PM5xxx.csv) | Meter | Schneider | PM5xxx | Modbus | 35 | 4 |
| [`MET_Schneider_PM5xxx_Series.csv`](./MET_Schneider_PM5xxx_Series.csv) | Meter | Schneider | PM5xxx_Series | Modbus | 875 | 32 |
| [`OTH_Schneider_PM8000_Modbus_Map.csv`](./OTH_Schneider_PM8000_Modbus_Map.csv) | Other | Schneider | PM8000_Modbus_Map | Modbus | 2,922 | 18 |
| [`MET_Schneider_PM820_SERM.csv`](./MET_Schneider_PM820_SERM.csv) | Meter | Schneider | PM820_SERM | Modbus | 11 | 4 |
| [`OTH_Schneider_PM820_SERM.csv`](./OTH_Schneider_PM820_SERM.csv) | Other | Schneider | PM820_SERM | Modbus | 11 | 4 |
| [`MET_Schneider_PMC.csv`](./MET_Schneider_PMC.csv) | Meter | Schneider | PMC | Modbus | 517 | 22 |
| [`MET_Schneider_PowerLogic_PMC.csv`](./MET_Schneider_PowerLogic_PMC.csv) | Meter | Schneider | PowerLogic_PMC | Modbus | 517 | 22 |
| [`MET_Schneider_RELAYP5.csv`](./MET_Schneider_RELAYP5.csv) | Meter | Schneider | RELAYP5 | Modbus | 57 | 18 |
| [`MET_Schneider_RELAYP5_noswap.csv`](./MET_Schneider_RELAYP5_noswap.csv) | Meter | Schneider | RELAYP5_noswap | Modbus | 57 | 18 |
| [`MET_Schneider_SCH_iem3xxx.csv`](./MET_Schneider_SCH_iem3xxx.csv) | Meter | Schneider | SCH_iem3xxx | Modbus | 58 | 17 |
| [`OTH_Schneider_SCH_iem3xxx.csv`](./OTH_Schneider_SCH_iem3xxx.csv) | Other | Schneider | SCH_iem3xxx | Modbus | 54 | 17 |
| [`MET_Schneider_SEPAM_S48.csv`](./MET_Schneider_SEPAM_S48.csv) | Meter | Schneider | SEPAM_S48 | Modbus | 31 | 6 |
| [`OTH_Schneider_SEPAM_S48.csv`](./OTH_Schneider_SEPAM_S48.csv) | Other | Schneider | SEPAM_S48 | Modbus | 31 | 6 |
| [`OTH_Schneider_SE_EVSE_EVLink-Parking-City-Wallbox.csv`](./OTH_Schneider_SE_EVSE_EVLink-Parking-City-Wallbox.csv) | Other | Schneider | SE_EVSE_EVLink-Parking-City-Wallbox | Modbus | 25 | 0 |
| [`OTH_Schneider_SE_EVSE_EVLink_ProAC.csv`](./OTH_Schneider_SE_EVSE_EVLink_ProAC.csv) | Other | Schneider | SE_EVSE_EVLink_ProAC | Modbus | 49 | 12 |
| [`INV_Schneider_SunSpec.csv`](./INV_Schneider_SunSpec.csv) | Inverter | Schneider | SunSpec | Modbus | 55 | 24 |
| [`OTH_Schneider_SunSpec.csv`](./OTH_Schneider_SunSpec.csv) | Other | Schneider | SunSpec | Modbus | 55 | 24 |
| [`OTH_Schneider_XANTREX_GENERIC.csv`](./OTH_Schneider_XANTREX_GENERIC.csv) | Other | Schneider | XANTREX_GENERIC | Modbus | 235 | 16 |
| [`MET_Schneider_iEM3000.csv`](./MET_Schneider_iEM3000.csv) | Meter | Schneider | iEM3000 | Modbus | 55 | 20 |
| [`MET_Schneider_iem3000.csv`](./MET_Schneider_iem3000.csv) | Meter | Schneider | iEM3000 | Modbus | 55 | 20 |
| [`MET_Schneider_iEM3400_3500.csv`](./MET_Schneider_iEM3400_3500.csv) | Meter | Schneider | iEM3400_3500 | Modbus | 140 | 17 |
| [`OTH_Schneider_iEM3400_3500.csv`](./OTH_Schneider_iEM3400_3500.csv) | Other | Schneider | iEM3400_3500 | Modbus | 140 | 21 |
| [`MET_Schneider_iEM3X00.csv`](./MET_Schneider_iEM3X00.csv) | Meter | Schneider | iEM3X00 | Modbus | 25 | 7 |
| [`OTH_Schneider_iEM3X00.csv`](./OTH_Schneider_iEM3X00.csv) | Other | Schneider | iEM3X00 | Modbus | 25 | 7 |
| [`OTH_Schneider_iem3000.csv`](./OTH_Schneider_iem3000.csv) | Other | Schneider | iem3000 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3100.csv`](./MET_Schneider_iem3100.csv) | Meter | Schneider | iem3100 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3100.csv`](./OTH_Schneider_iem3100.csv) | Other | Schneider | iem3100 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3110.csv`](./MET_Schneider_iem3110.csv) | Meter | Schneider | iem3110 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3110.csv`](./OTH_Schneider_iem3110.csv) | Other | Schneider | iem3110 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3115.csv`](./MET_Schneider_iem3115.csv) | Meter | Schneider | iem3115 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3115.csv`](./OTH_Schneider_iem3115.csv) | Other | Schneider | iem3115 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3135.csv`](./MET_Schneider_iem3135.csv) | Meter | Schneider | iem3135 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3135.csv`](./OTH_Schneider_iem3135.csv) | Other | Schneider | iem3135 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3150.csv`](./MET_Schneider_iem3150.csv) | Meter | Schneider | iem3150 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3150.csv`](./OTH_Schneider_iem3150.csv) | Other | Schneider | iem3150 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3155.csv`](./MET_Schneider_iem3155.csv) | Meter | Schneider | iem3155 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3155.csv`](./OTH_Schneider_iem3155.csv) | Other | Schneider | iem3155 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3250.csv`](./MET_Schneider_iem3250.csv) | Meter | Schneider | iem3250 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3250.csv`](./OTH_Schneider_iem3250.csv) | Other | Schneider | iem3250 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3255.csv`](./MET_Schneider_iem3255.csv) | Meter | Schneider | iem3255 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3255.csv`](./OTH_Schneider_iem3255.csv) | Other | Schneider | iem3255 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3350.csv`](./MET_Schneider_iem3350.csv) | Meter | Schneider | iem3350 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3350.csv`](./OTH_Schneider_iem3350.csv) | Other | Schneider | iem3350 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3355.csv`](./MET_Schneider_iem3355.csv) | Meter | Schneider | iem3355 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3355.csv`](./OTH_Schneider_iem3355.csv) | Other | Schneider | iem3355 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3455.csv`](./MET_Schneider_iem3455.csv) | Meter | Schneider | iem3455 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3455.csv`](./OTH_Schneider_iem3455.csv) | Other | Schneider | iem3455 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3555.csv`](./MET_Schneider_iem3555.csv) | Meter | Schneider | iem3555 | Modbus | 55 | 20 |
| [`OTH_Schneider_iem3555.csv`](./OTH_Schneider_iem3555.csv) | Other | Schneider | iem3555 | Modbus | 55 | 21 |
| [`MET_Schneider_iem3xxx.csv`](./MET_Schneider_iem3xxx.csv) | Meter | Schneider | iem3xxx | Modbus | 58 | 17 |
| [`OTH_Schneider_iem3xxx.csv`](./OTH_Schneider_iem3xxx.csv) | Other | Schneider | iem3xxx | Modbus | 54 | 17 |
| [`MET_Schneider_new.csv`](./MET_Schneider_new.csv) | Meter | Schneider | new | Modbus | 17 | 5 |
| [`OTH_Schneider_new.csv`](./OTH_Schneider_new.csv) | Other | Schneider | new | Modbus | 17 | 5 |
| [`OTH_Secure_Elite_440-_Modbus_Mapping.csv`](./OTH_Secure_Elite_440-_Modbus_Mapping.csv) | Other | Secure | Elite_440-_Modbus_Mapping | Modbus | 198 | 4 |
| [`SEN_Seven_3S-IS.csv`](./SEN_Seven_3S-IS.csv) | Sensor | Seven | 3S-IS | Modbus | 45 | 2 |
| [`SEN_Seven_3S-IS_8.csv`](./SEN_Seven_3S-IS_8.csv) | Sensor | Seven | 3S-IS_8 | Modbus | 41 | 8 |
| [`SEN_Seven_3S-IS_8.0.csv`](./SEN_Seven_3S-IS_8.0.csv) | Sensor | Seven | 3S-IS_8.0 | Modbus | 42 | 8 |
| [`SEN_Seven_3S-IS_SUNSPEC.csv`](./SEN_Seven_3S-IS_SUNSPEC.csv) | Sensor | Seven | 3S-IS_SUNSPEC | Modbus | 45 | 6 |
| [`SEN_Seven_SEVENSENSORSOLUTION_3SIS.csv`](./SEN_Seven_SEVENSENSORSOLUTION_3SIS.csv) | Sensor | Seven | SEVENSENSORSOLUTION_3SIS | Modbus | 8 | 6 |
| [`OTH_Siebert_XC4XX_RTU.csv`](./OTH_Siebert_XC4XX_RTU.csv) | Other | Siebert | XC4XX_RTU | Modbus | 3 | 3 |
| [`OTH_Siebert_XC4XX_TCP.csv`](./OTH_Siebert_XC4XX_TCP.csv) | Other | Siebert | XC4XX_TCP | Modbus | 3 | 3 |
| [`OTH_Siel_10TL.csv`](./OTH_Siel_10TL.csv) | Other | Siel | 10TL | Modbus | 17 | 1 |
| [`OTH_Siel_DSP-DSPX.csv`](./OTH_Siel_DSP-DSPX.csv) | Other | Siel | DSP-DSPX | Modbus | 21 | 1 |
| [`OTH_Siel_MONOPHASE.csv`](./OTH_Siel_MONOPHASE.csv) | Other | Siel | MONOPHASE | Modbus | 11 | 2 |
| [`OTH_Siel_Nouveau_Feuille_Microsoft_Office_Excel.csv`](./OTH_Siel_Nouveau_Feuille_Microsoft_Office_Excel.csv) | Other | Siel | Nouveau_Feuille_Microsoft_Office_Excel | Modbus | 16 | 0 |
| [`OTH_Siemens_DCMonitor_Monichek_CombinerBox.csv`](./OTH_Siemens_DCMonitor_Monichek_CombinerBox.csv) | Other | Siemens | DCMonitor_Monichek_CombinerBox | Modbus | 6,292 | 7 |
| [`INV_Siemens_GAMESA_Proteus.csv`](./INV_Siemens_GAMESA_Proteus.csv) | Inverter | Siemens | GAMESA_Proteus | Modbus | 743 | 28 |
| [`OTH_Siemens_GD458914_R6_PROTEUS_PV-INV-EN-ICD_MONITORING.csv`](./OTH_Siemens_GD458914_R6_PROTEUS_PV-INV-EN-ICD_MONITORING.csv) | Other | Siemens | GD458914_R6_PROTEUS_PV-INV-EN-ICD_MONITORING | Modbus | 502 | 24 |
| [`OTH_Sigenergy_Generic.csv`](./OTH_Sigenergy_Generic.csv) | Other | Sigenergy | Generic | Modbus | 194 | 61 |
| [`OTH_Sineng_central_New.csv`](./OTH_Sineng_central_New.csv) | Other | Sineng | central_New | Modbus | 43 | 10 |
| [`MET_Socomec_A10.csv`](./MET_Socomec_A10.csv) | Meter | Socomec | A10 | Modbus | 25 | 5 |
| [`MET_Socomec_B10.csv`](./MET_Socomec_B10.csv) | Meter | Socomec | B10 | Modbus | 15 | 3 |
| [`MET_Socomec_B30.csv`](./MET_Socomec_B30.csv) | Meter | Socomec | B30 | Modbus | 45 | 8 |
| [`MET_Socomec_COUNTIS_E50-E53.csv`](./MET_Socomec_COUNTIS_E50-E53.csv) | Meter | Socomec | COUNTIS_E50-E53 | Modbus | 141 | 19 |
| [`MET_Socomec_COUNTIS_E23.csv`](./MET_Socomec_COUNTIS_E23.csv) | Meter | Socomec | Countis_E23 | Modbus | 59 | 18 |
| [`MET_Socomec_Countis_E23_2448.csv`](./MET_Socomec_Countis_E23_2448.csv) | Meter | Socomec | Countis_E23_2448 | Modbus | 73 | 18 |
| [`MET_Socomec_COUNTIS_E27.csv`](./MET_Socomec_COUNTIS_E27.csv) | Meter | Socomec | Countis_E27 | Modbus | 162 | 20 |
| [`MET_Socomec_Countis_E27.csv`](./MET_Socomec_Countis_E27.csv) | Meter | Socomec | Countis_E27 | Modbus | 162 | 20 |
| [`MET_Socomec_Countis_E33_E43.csv`](./MET_Socomec_Countis_E33_E43.csv) | Meter | Socomec | Countis_E33_E43 | Modbus | 73 | 18 |
| [`MET_Socomec_Countis_E50_E53.csv`](./MET_Socomec_Countis_E50_E53.csv) | Meter | Socomec | Countis_E50_E53 | Modbus | 141 | 19 |
| [`MET_Socomec_Countis_P44.csv`](./MET_Socomec_Countis_P44.csv) | Meter | Socomec | Countis_P44 | Modbus | 37 | 19 |
| [`MET_Socomec_Diris_A_Countis_E.csv`](./MET_Socomec_Diris_A_Countis_E.csv) | Meter | Socomec | Diris-A-Countis-E | Modbus | 30 | 30 |
| [`MET_Socomec_Diris_A10.csv`](./MET_Socomec_Diris_A10.csv) | Meter | Socomec | Diris_A10 | Modbus | 75 | 18 |
| [`MET_Socomec_Diris_A40.csv`](./MET_Socomec_Diris_A40.csv) | Meter | Socomec | Diris_A40 | Modbus | 35 | 17 |
| [`MET_Socomec_Diris_Axx.csv`](./MET_Socomec_Diris_Axx.csv) | Meter | Socomec | Diris_Axx | Modbus | 25 | 5 |
| [`MET_Socomec_DIRIS_B10.csv`](./MET_Socomec_DIRIS_B10.csv) | Meter | Socomec | Diris_B10 | Modbus | 23 | 7 |
| [`MET_Socomec_Diris_B30.csv`](./MET_Socomec_Diris_B30.csv) | Meter | Socomec | Diris_B30 | Modbus | 20 | 7 |
| [`MET_Socomec_DIRIS_DIGIWARE_I35.csv`](./MET_Socomec_DIRIS_DIGIWARE_I35.csv) | Meter | Socomec | Diris_Digiware_I35 | Modbus | 26 | 9 |
| [`MET_Socomec_Diris_Digiware_I35.csv`](./MET_Socomec_Diris_Digiware_I35.csv) | Meter | Socomec | Diris_Digiware_I35 | Modbus | 26 | 9 |
| [`MET_Socomec_Diris_I30.csv`](./MET_Socomec_Diris_I30.csv) | Meter | Socomec | Diris_I30 | Modbus | 20 | 1 |
| [`MET_Socomec_Diris_I35.csv`](./MET_Socomec_Diris_I35.csv) | Meter | Socomec | Diris_I35 | Modbus | 357 | 5 |
| [`MET_Socomec_E23_2448.csv`](./MET_Socomec_E23_2448.csv) | Meter | Socomec | E23_2448 | Modbus | 73 | 18 |
| [`MET_Socomec_E33_2408.csv`](./MET_Socomec_E33_2408.csv) | Meter | Socomec | E33_2408 | Modbus | 73 | 18 |
| [`MET_Socomec_ECI2.csv`](./MET_Socomec_ECI2.csv) | Meter | Socomec | ECI2 | Modbus | 7 | 1 |
| [`MET_Socomec_I30.csv`](./MET_Socomec_I30.csv) | Meter | Socomec | I30 | Modbus | 14 | 5 |
| [`MET_Socomec_I33.csv`](./MET_Socomec_I33.csv) | Meter | Socomec | I33 | Modbus | 12 | 0 |
| [`INV_Socomec_Modbus_P100TR.csv`](./INV_Socomec_Modbus_P100TR.csv) | Inverter | Socomec | Modbus_P100TR | Modbus | 123 | 6 |
| [`OTH_Socomec_Modbus_P100TR.csv`](./OTH_Socomec_Modbus_P100TR.csv) | Other | Socomec | Modbus_P100TR | Modbus | 123 | 6 |
| [`MET_Socomec_U20.csv`](./MET_Socomec_U20.csv) | Meter | Socomec | U20 | Modbus | 7 | 0 |
| [`OTH_Socomec_W_SOCOMEC_E33_E43.csv`](./OTH_Socomec_W_SOCOMEC_E33_E43.csv) | Other | Socomec | W_SOCOMEC_E33_E43 | Modbus | 73 | 18 |
| [`MET_Socomec_W_Socomec_Diris_A10.csv`](./MET_Socomec_W_Socomec_Diris_A10.csv) | Meter | Socomec | W_Socomec_Diris_A10 | Modbus | 75 | 18 |
| [`MET_Socomec_W_Socomec_Diris_B30.csv`](./MET_Socomec_W_Socomec_Diris_B30.csv) | Meter | Socomec | W_Socomec_Diris_B30 | Modbus | 20 | 7 |
| [`MET_Socomec_W_Socomec_Diris_I30.csv`](./MET_Socomec_W_Socomec_Diris_I30.csv) | Meter | Socomec | W_Socomec_Diris_I30 | Modbus | 20 | 6 |
| [`MET_Socorel_LINKY.csv`](./MET_Socorel_LINKY.csv) | Meter | Socorel | LINKY | Modbus | 149 | 6 |
| [`OTH_Socorel_Table_modbus_TIC_V4.csv`](./OTH_Socorel_Table_modbus_TIC_V4.csv) | Other | Socorel | Table_modbus_TIC_V4 | Modbus | 1,235 | 10 |
| [`INV_SofarSolar_1-40KTL.csv`](./INV_SofarSolar_1-40KTL.csv) | Inverter | SofarSolar | 1-40KTL | Modbus | 115 | 5 |
| [`INV_SofarSolar_50-70KTL.csv`](./INV_SofarSolar_50-70KTL.csv) | Inverter | SofarSolar | 50-70KTL | Modbus | 152 | 5 |
| [`INV_SofarSolar_HYDG3.csv`](./INV_SofarSolar_HYDG3.csv) | Inverter | SofarSolar | HYDG3 | Modbus | 538 | 6 |
| [`INV_SofarSolar_MODBUS_SofarSolar_1-40KTL.csv`](./INV_SofarSolar_MODBUS_SofarSolar_1-40KTL.csv) | Inverter | SofarSolar | MODBUS_SofarSolar_1-40KTL | Modbus | 115 | 30 |
| [`INV_SofarSolar_Modbus.csv`](./INV_SofarSolar_Modbus.csv) | Inverter | SofarSolar | Modbus | Modbus | 435 | 28 |
| [`INV_SofarSolar_SOFAR-G3.External.Modbus.Protocol-Engilsh-V211103.csv`](./INV_SofarSolar_SOFAR-G3.External.Modbus.Protocol-Engilsh-V211103.csv) | Inverter | SofarSolar | SOFAR-G3.External.Modbus.Protocol-Engilsh-V211103 | Modbus | 49 | 1 |
| [`INV_SofarSolar_SOFAR_Modbus_Protocol_English_G3_V1.17.csv`](./INV_SofarSolar_SOFAR_Modbus_Protocol_English_G3_V1.17.csv) | Inverter | SofarSolar | SOFAR_Modbus_Protocol_English_G3_V1.17 | Modbus | 145 | 17 |
| [`INV_SolaX_EMS1000-PRO.csv`](./INV_SolaX_EMS1000-PRO.csv) | Inverter | SolaX | EMS1000-PRO | Modbus | 347 | 33 |
| [`INV_SolaX_EMS1000-PRO_v2.csv`](./INV_SolaX_EMS1000-PRO_v2.csv) | Inverter | SolaX | EMS1000-PRO_v2 | Modbus | 181 | 28 |
| [`INV_SolarEdge_66KW.csv`](./INV_SolarEdge_66KW.csv) | Inverter | SolarEdge | 66KW | Modbus | 38 | 10 |
| [`OTH_SolarEdge_66KW.csv`](./OTH_SolarEdge_66KW.csv) | Other | SolarEdge | 66KW | Modbus | 38 | 10 |
| [`INV_SolarEdge_ActivePowerControl.csv`](./INV_SolarEdge_ActivePowerControl.csv) | Inverter | SolarEdge | ActivePowerControl | Modbus | 16 | 6 |
| [`OTH_SolarEdge_ActivePowerControl.csv`](./OTH_SolarEdge_ActivePowerControl.csv) | Other | SolarEdge | ActivePowerControl | Modbus | 16 | 6 |
| [`INV_SolarEdge_EnhancedPowerControl.csv`](./INV_SolarEdge_EnhancedPowerControl.csv) | Inverter | SolarEdge | EnhancedPowerControl | Modbus | 18 | 9 |
| [`INV_SolarEdge_EnhancedPowerControlBlock.csv`](./INV_SolarEdge_EnhancedPowerControlBlock.csv) | Inverter | SolarEdge | EnhancedPowerControlBlock | Modbus | 18 | 9 |
| [`OTH_SolarEdge_EnhancedPowerControlBlock.csv`](./OTH_SolarEdge_EnhancedPowerControlBlock.csv) | Other | SolarEdge | EnhancedPowerControlBlock | Modbus | 18 | 9 |
| [`OTH_SolarEdge_FromIndia.csv`](./OTH_SolarEdge_FromIndia.csv) | Other | SolarEdge | FromIndia | Modbus | 38 | 9 |
| [`OTH_SolarEdge_Modbus_Sunspec.csv`](./OTH_SolarEdge_Modbus_Sunspec.csv) | Other | SolarEdge | Modbus_Sunspec | Modbus | 46 | 10 |
| [`MET_SolarEdge_SE-MTR-3Y-400V-A.csv`](./MET_SolarEdge_SE-MTR-3Y-400V-A.csv) | Meter | SolarEdge | SE-MTR-3Y-400V-A | Modbus | 246 | 7 |
| [`OTH_SolarEdge_SE-MTR-3Y-400V-A.csv`](./OTH_SolarEdge_SE-MTR-3Y-400V-A.csv) | Other | SolarEdge | SE-MTR-3Y-400V-A | Modbus | 246 | 7 |
| [`INV_SolarEdge_SE_Series.csv`](./INV_SolarEdge_SE_Series.csv) | Inverter | SolarEdge | SE-Series | Modbus | 28 | 28 |
| [`INV_SolarEdge_SE25K_SunSpec.csv`](./INV_SolarEdge_SE25K_SunSpec.csv) | Inverter | SolarEdge | SE25K_SunSpec | Modbus | 53 | 23 |
| [`INV_SolarEdge_SE33.3K_SunSpec.csv`](./INV_SolarEdge_SE33.3K_SunSpec.csv) | Inverter | SolarEdge | SE33.3K_SunSpec | Modbus | 69 | 39 |
| [`INV_SolarEdge_SE33.3K_SunSpec_regul.csv`](./INV_SolarEdge_SE33.3K_SunSpec_regul.csv) | Inverter | SolarEdge | SE33.3K_SunSpec_regul | Modbus | 69 | 39 |
| [`OTH_SolarEdge_SE33.3K_SunSpec_regul.csv`](./OTH_SolarEdge_SE33.3K_SunSpec_regul.csv) | Other | SolarEdge | SE33.3K_SunSpec_regul | Modbus | 69 | 27 |
| [`INV_SolarEdge_STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU.csv`](./INV_SolarEdge_STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU.csv) | Inverter | SolarEdge | STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU | Modbus | 54 | 24 |
| [`INV_SolarEdge_STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU_ORI.csv`](./INV_SolarEdge_STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU_ORI.csv) | Inverter | SolarEdge | STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU_ORI | Modbus | 53 | 23 |
| [`INV_SolarEdge_STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU_regul.csv`](./INV_SolarEdge_STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU_regul.csv) | Inverter | SolarEdge | STEPGUIDEL_SunSpec_inverter_SolarEdge_SE33.3K-RW00IBNJ4_0004.0016.0025_modbusRTU_regul | Modbus | 69 | 28 |
| [`INV_SolarEdge_SunSpec_Generic.csv`](./INV_SolarEdge_SunSpec_Generic.csv) | Inverter | SolarEdge | SunSpec_Generic | Modbus | 46 | 10 |
| [`INV_SolarEdge_Sunspec.csv`](./INV_SolarEdge_Sunspec.csv) | Inverter | SolarEdge | Sunspec | Modbus | 43 | 9 |
| [`OTH_SolarEdge_Sunspec.csv`](./OTH_SolarEdge_Sunspec.csv) | Other | SolarEdge | Sunspec | Modbus | 43 | 9 |
| [`MET_SolarLog_POLIER_PROMOD380.csv`](./MET_SolarLog_POLIER_PROMOD380.csv) | Meter | SolarLog | POLIER_PROMOD380 | Modbus | 88 | 31 |
| [`OTH_SolarLog_PRO380_WORK.csv`](./OTH_SolarLog_PRO380_WORK.csv) | Other | SolarLog | PRO380_WORK | Modbus | 147 | 30 |
| [`MET_SolarLog_SOLARLOG.csv`](./MET_SolarLog_SOLARLOG.csv) | Meter | SolarLog | SOLARLOG | Modbus | 17 | 5 |
| [`INV_SolarEdge_FromIndia.csv`](./INV_SolarEdge_FromIndia.csv) | Inverter | Solaredge | FromIndia | Modbus | 38 | 9 |
| [`INV_SolaX_FORTH-MEGA_GENERIC.csv`](./INV_SolaX_FORTH-MEGA_GENERIC.csv) | Inverter | Solax | FORTH-MEGA_GENERIC | Modbus | 326 | 16 |
| [`INV_SolaX_Generic.csv`](./INV_SolaX_Generic.csv) | Inverter | Solax | Generic | Modbus | 214 | 4 |
| [`INV_SolaX_MICG2-PROG2.csv`](./INV_SolaX_MICG2-PROG2.csv) | Inverter | Solax | MICG2-PROG2 | Modbus | 68 | 15 |
| [`INV_SolaX_TCP.csv`](./INV_SolaX_TCP.csv) | Inverter | Solax | TCP | Modbus | 156 | 10 |
| [`INV_SolaX_Work_MODBUS_Solax.csv`](./INV_SolaX_Work_MODBUS_Solax.csv) | Inverter | Solax | Work_MODBUS_Solax | Modbus | 139 | 3 |
| [`INV_SolaX_X1-X3-GENERIC.csv`](./INV_SolaX_X1-X3-GENERIC.csv) | Inverter | Solax | X1-X3-GENERIC | Modbus | 156 | 10 |
| [`INV_SolaX_X1-X3-GENERIC_TCP.csv`](./INV_SolaX_X1-X3-GENERIC_TCP.csv) | Inverter | Solax | X1-X3-GENERIC_TCP | Modbus | 156 | 10 |
| [`INV_SolaX_X1-X3_Generic.csv`](./INV_SolaX_X1-X3_Generic.csv) | Inverter | Solax | X1-X3_Generic | Modbus | 157 | 4 |
| [`INV_SolaX_X2.csv`](./INV_SolaX_X2.csv) | Inverter | Solax | X2 | Modbus | 57 | 6 |
| [`INV_Solis_GENERIC.csv`](./INV_Solis_GENERIC.csv) | Inverter | Solis | GENERIC | Modbus | 455 | 31 |
| [`INV_Solis_GENERIC_GridTied.csv`](./INV_Solis_GENERIC_GridTied.csv) | Inverter | Solis | GENERIC_GridTied | Modbus | 690 | 44 |
| [`INV_Solis_GENERIC_Work.csv`](./INV_Solis_GENERIC_Work.csv) | Inverter | Solis | GENERIC_Work | Modbus | 654 | 30 |
| [`INV_Solis_Inverter.csv`](./INV_Solis_Inverter.csv) | Inverter | Solis | Inverter_Standard | Modbus | 93 | 50 |
| [`INV_Solis_MODBUS_solis_50_25_sunPM.csv`](./INV_Solis_MODBUS_solis_50_25_sunPM.csv) | Inverter | Solis | MODBUS_solis_50_25_sunPM | Modbus | 47 | 14 |
| [`INV_Solis_Modbus.csv`](./INV_Solis_Modbus.csv) | Inverter | Solis | Modbus | Modbus | 57 | 14 |
| [`INV_Solis_SOLIS_GENERIC.csv`](./INV_Solis_SOLIS_GENERIC.csv) | Inverter | Solis | SOLIS_GENERIC | Modbus | 455 | 31 |
| [`INV_Solis_SOLIS_GENERIC_GridTied.csv`](./INV_Solis_SOLIS_GENERIC_GridTied.csv) | Inverter | Solis | SOLIS_GENERIC_GridTied | Modbus | 690 | 44 |
| [`INV_Solis_Solis.csv`](./INV_Solis_Solis.csv) | Inverter | Solis | Solis | Modbus | 454 | 23 |
| [`INV_Solis_50_25_SunPM.csv`](./INV_Solis_50_25_SunPM.csv) | Inverter | Solis | Solis_50_25_SunPM | Modbus | 47 | 14 |
| [`INV_Solis_SunSpec_25_66K.csv`](./INV_Solis_SunSpec_25_66K.csv) | Inverter | Solis | SunSpec_25_66K | Modbus | 308 | 34 |
| [`INV_Solis_sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15.csv`](./INV_Solis_sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15.csv) | Inverter | Solis | sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15 | Modbus | 308 | 34 |
| [`INV_Solis_table2023.csv`](./INV_Solis_table2023.csv) | Inverter | Solis | table2023 | Modbus | 454 | 22 |
| [`OTH_Solplanet_AISWEIModbusProtocol_-_V2.1.11_external_1.csv`](./OTH_Solplanet_AISWEIModbusProtocol_-_V2.1.11_external_1.csv) | Other | Solplanet | AISWEIModbusProtocol_-_V2.1.11_external_1 | Modbus | 577 | 15 |
| [`INV_Solplanet_ASW_ThreePhase.csv`](./INV_Solplanet_ASW_ThreePhase.csv) | Inverter | Solplanet | ASW_ThreePhase | Modbus | 27 | 27 |
| [`COM_Solplanet_Ai-Logger.csv`](./COM_Solplanet_Ai-Logger.csv) | Logger | Solplanet | Ai-Logger | Modbus | 161 | 161 |
| [`INV_Solplanet_Ai-Logger.csv`](./INV_Solplanet_Ai-Logger.csv) | Inverter | Solplanet | Ai-Logger | Modbus | 106 | 106 |
| [`INV_Solplanet_v0.1.csv`](./INV_Solplanet_v0.1.csv) | Inverter | Solplanet | v0.1 | Modbus | 491 | 60 |
| [`OTH_Solplanet_v0.1.csv`](./OTH_Solplanet_v0.1.csv) | Other | Solplanet | v0.1 | Modbus | 491 | 60 |
| [`INV_Solplanet_v0.2.csv`](./INV_Solplanet_v0.2.csv) | Inverter | Solplanet | v0.2 | Modbus | 560 | 58 |
| [`OTH_Solplanet_v0.2.csv`](./OTH_Solplanet_v0.2.csv) | Other | Solplanet | v0.2 | Modbus | 560 | 58 |
| [`INV_Solplanet_v0.3.csv`](./INV_Solplanet_v0.3.csv) | Inverter | Solplanet | v0.3 | Modbus | 491 | 59 |
| [`OTH_Solplanet_v0.3.csv`](./OTH_Solplanet_v0.3.csv) | Other | Solplanet | v0.3 | Modbus | 491 | 59 |
| [`OTH_Solplanet_v0.4.csv`](./OTH_Solplanet_v0.4.csv) | Other | Solplanet | v0.4 | Modbus | 298 | 63 |
| [`OTH_Solplanet_v0.5.csv`](./OTH_Solplanet_v0.5.csv) | Other | Solplanet | v0.5 | Modbus | 139 | 58 |
| [`BAT_Sungrow_Battery.csv`](./BAT_Sungrow_Battery.csv) | Battery | Sungrow | Battery | Modbus | 172 | 172 |
| [`INV_Sungrow_Battery.csv`](./INV_Sungrow_Battery.csv) | Inverter | Sungrow | Battery | Modbus | 172 | 172 |
| [`COM_Sungrow_EMS300CP.csv`](./COM_Sungrow_EMS300CP.csv) | Logger | Sungrow | EMS300CP | Modbus | 694 | 8 |
| [`BAT_Sungrow_FullMap_v2.csv`](./BAT_Sungrow_FullMap_v2.csv) | Battery | Sungrow | FullMap_v2 | Modbus | 801 | 37 |
| [`INV_Sungrow_GENERIC_V.1.1.58.csv`](./INV_Sungrow_GENERIC_V.1.1.58.csv) | Inverter | Sungrow | GENERIC_V.1.1.58 | Modbus | 200 | 95 |
| [`INV_Sungrow_Generic.csv`](./INV_Sungrow_Generic.csv) | Inverter | Sungrow | Generic | Modbus | 164 | 80 |
| [`BAT_Sungrow_LC3000.csv`](./BAT_Sungrow_LC3000.csv) | Battery | Sungrow | LC3000 | Modbus | 916 | 33 |
| [`OTH_Sungrow_LC3000.csv`](./OTH_Sungrow_LC3000.csv) | Other | Sungrow | LC3000 | Modbus | 916 | 32 |
| [`COM_Sungrow_LC300_Modbus.csv`](./COM_Sungrow_LC300_Modbus.csv) | Logger | Sungrow | LC300_Modbus | Modbus | 813 | 813 |
| [`COM_Sungrow_Logger1000_3000.csv`](./COM_Sungrow_Logger1000_3000.csv) | Logger | Sungrow | Logger1000_3000 | Modbus | 50 | 11 |
| [`INV_Sungrow_MODBUS_PM.csv`](./INV_Sungrow_MODBUS_PM.csv) | Inverter | Sungrow | MODBUS_PM | Modbus | 99 | 69 |
| [`COM_Sungrow_MODBUS_SUNGROW_Logger1000-3000.csv`](./COM_Sungrow_MODBUS_SUNGROW_Logger1000-3000.csv) | Logger | Sungrow | MODBUS_SUNGROW_Logger1000-3000 | Modbus | 50 | 11 |
| [`COM_Sungrow_MODBUS_SUNGROW_Logger1000-3000RTU.csv`](./COM_Sungrow_MODBUS_SUNGROW_Logger1000-3000RTU.csv) | Logger | Sungrow | MODBUS_SUNGROW_Logger1000-3000RTU | Modbus | 49 | 11 |
| [`INV_Sungrow_PVCombinerBox.csv`](./INV_Sungrow_PVCombinerBox.csv) | Inverter | Sungrow | PVCombinerBox | Modbus | 64 | 5 |
| [`OTH_Sungrow_PVCombinerBox.csv`](./OTH_Sungrow_PVCombinerBox.csv) | Other | Sungrow | PVCombinerBox | Modbus | 64 | 5 |
| [`BAT_Sungrow_PowerStack_ST255CS.csv`](./BAT_Sungrow_PowerStack_ST255CS.csv) | Battery | Sungrow | PowerStack_ST255CS | Modbus | 1,442 | 15 |
| [`INV_Sungrow_RTU_PM.csv`](./INV_Sungrow_RTU_PM.csv) | Inverter | Sungrow | RTU_PM | Modbus | 81 | 62 |
| [`INV_Sungrow_SG_Series.csv`](./INV_Sungrow_SG_Series.csv) | Inverter | Sungrow | SG-Series | Modbus | 29 | 29 |
| [`INV_Sungrow_SG_Series_Generic.csv`](./INV_Sungrow_SG_Series_Generic.csv) | Inverter | Sungrow | SG_Series_Generic | Modbus | 200 | 95 |
| [`INV_Sungrow_TCP_PM.csv`](./INV_Sungrow_TCP_PM.csv) | Inverter | Sungrow | TCP_PM | Modbus | 81 | 62 |
| [`COM_Sungrow_TI_20241202_EMS300CP_IEC104_Modbus_Communication_Protocol_P014_software_version_CN_EN.csv`](./COM_Sungrow_TI_20241202_EMS300CP_IEC104_Modbus_Communication_Protocol_P014_software_version_CN_EN.csv) | Logger | Sungrow | TI_20241202_EMS300CP_IEC104_Modbus_Communication_Protocol_P014_software_version_CN_EN | Modbus | 694 | 8 |
| [`COM_Sungrow_TI_20251013_LC300_IEC104_Communication_Protocol_V9_EN.csv`](./COM_Sungrow_TI_20251013_LC300_IEC104_Communication_Protocol_V9_EN.csv) | Logger | Sungrow | TI_20251013_LC300_IEC104_Communication_Protocol_V9_EN | Modbus | 1,382 | 32 |
| [`COM_Sungrow_TI_20251013_LC300_Modbus_Communication_Protocol_V8_EN.csv`](./COM_Sungrow_TI_20251013_LC300_Modbus_Communication_Protocol_V8_EN.csv) | Logger | Sungrow | TI_20251013_LC300_Modbus_Communication_Protocol_V8_EN | Modbus | 813 | 813 |
| [`INV_Sungrow_Updated.csv`](./INV_Sungrow_Updated.csv) | Inverter | Sungrow | Updated | Modbus | 76 | 65 |
| [`OTH_Sungrow_Updated.csv`](./OTH_Sungrow_Updated.csv) | Other | Sungrow | Updated | Modbus | 76 | 65 |
| [`INV_Sungrow_WORK_MODBUS_SUNGROW.csv`](./INV_Sungrow_WORK_MODBUS_SUNGROW.csv) | Inverter | Sungrow | WORK_MODBUS_SUNGROW | Modbus | 14 | 2 |
| [`INV_Sungrow_short.csv`](./INV_Sungrow_short.csv) | Inverter | Sungrow | short | Modbus | 173 | 74 |
| [`BAT_Sungrow_v2.csv`](./BAT_Sungrow_v2.csv) | Battery | Sungrow | v2 | Modbus | 801 | 37 |
| [`INV_Sungrow_v2.csv`](./INV_Sungrow_v2.csv) | Inverter | Sungrow | v2 | Modbus | 801 | 37 |
| [`BAT_Sungrow_v3.csv`](./BAT_Sungrow_v3.csv) | Battery | Sungrow | v3 | Modbus | 1,442 | 15 |
| [`INV_Tata_Generic.csv`](./INV_Tata_Generic.csv) | Inverter | TATA | Generic | Modbus | 20 | 2 |
| [`SEN_Thies_Wind_Transmitter_Thies.csv`](./SEN_Thies_Wind_Transmitter_Thies.csv) | Sensor | THIES | Wind_Transmitter_Thies | Modbus | 91 | 4 |
| [`INV_TMEIC_PVL_L1000E_Revised_41.csv`](./INV_TMEIC_PVL_L1000E_Revised_41.csv) | Inverter | TMEIC | PVL_L1000E_Revised_41 | Modbus | 60 | 32 |
| [`OTH_TMEIC_PVL_L1000E_Revised_41.csv`](./OTH_TMEIC_PVL_L1000E_Revised_41.csv) | Other | TMEIC | PVL_L1000E_Revised_41 | Modbus | 60 | 11 |
| [`SEN_Trackso_SOZ03.csv`](./SEN_Trackso_SOZ03.csv) | Sensor | TRACKSO | SOZ03 | Modbus | 2 | 0 |
| [`OTH_Tata_Generic.csv`](./OTH_Tata_Generic.csv) | Other | Tata | Generic | Modbus | 20 | 11 |
| [`INV_Tata_Modbus.csv`](./INV_Tata_Modbus.csv) | Inverter | Tata | Modbus | Modbus | 20 | 11 |
| [`OTH_Tata_Modbus.csv`](./OTH_Tata_Modbus.csv) | Other | Tata | Modbus | Modbus | 20 | 11 |
| [`SEN_Termasgard_RTM1.csv`](./SEN_Termasgard_RTM1.csv) | Sensor | Termasgard | RTM1 | Modbus | 5 | 0 |
| [`OTH_Thermokon_AGS55.csv`](./OTH_Thermokon_AGS55.csv) | Other | Thermokon | AGS55 | Modbus | 66 | 4 |
| [`SEN_Thermokon_AGS55.csv`](./SEN_Thermokon_AGS55.csv) | Sensor | Thermokon | AGS55 | Modbus | 66 | 0 |
| [`OTH_Thermokon_AGS55_tcp.csv`](./OTH_Thermokon_AGS55_tcp.csv) | Other | Thermokon | AGS55_tcp | Modbus | 66 | 4 |
| [`SEN_Thermokon_AGS55_tcp.csv`](./SEN_Thermokon_AGS55_tcp.csv) | Sensor | Thermokon | AGS55_tcp | Modbus | 66 | 4 |
| [`SEN_Thies_CLIMA.csv`](./SEN_Thies_CLIMA.csv) | Sensor | Thies | CLIMA | Modbus | 110 | 7 |
| [`OTH_Thies_Wind_Transmitter_Thies.csv`](./OTH_Thies_Wind_Transmitter_Thies.csv) | Other | Thies | Wind_Transmitter_Thies | Modbus | 91 | 4 |
| [`MET_Ticmaster_PME-PMI.csv`](./MET_Ticmaster_PME-PMI.csv) | Meter | TicMaster | PME-PMI | Modbus | 6 | 4 |
| [`MET_Ticmaster_PME-PMI_VOIE2.csv`](./MET_Ticmaster_PME-PMI_VOIE2.csv) | Meter | TicMaster | PME-PMI_VOIE2 | Modbus | 6 | 4 |
| [`MET_Ticmaster_PME-PMI_VOIE3.csv`](./MET_Ticmaster_PME-PMI_VOIE3.csv) | Meter | TicMaster | PME-PMI_VOIE3 | Modbus | 6 | 4 |
| [`MET_Ticmaster_ICE_SAPHIR.csv`](./MET_Ticmaster_ICE_SAPHIR.csv) | Meter | Ticmaster | ICE_SAPHIR | Modbus | 20 | 5 |
| [`MET_Ticmaster_ICE_SAPHIR_VOIE2.csv`](./MET_Ticmaster_ICE_SAPHIR_VOIE2.csv) | Meter | Ticmaster | ICE_SAPHIR_VOIE2 | Modbus | 7 | 4 |
| [`OTH_Ticmaster_Mapping_-_FW41.48.csv`](./OTH_Ticmaster_Mapping_-_FW41.48.csv) | Other | Ticmaster | Mapping_-_FW41.48 | Modbus | 76 | 2 |
| [`OTH_Ticmaster_Mapping_-_FW41.60.csv`](./OTH_Ticmaster_Mapping_-_FW41.60.csv) | Other | Ticmaster | Mapping_-_FW41.60 | Modbus | 75 | 2 |
| [`OTH_Universal_Algodue_UPM209.csv`](./OTH_Universal_Algodue_UPM209.csv) | Other | Universal | Algodue_UPM209 | Modbus | 103 | 28 |
| [`MET_Universal_Algodue_UPM209_309.csv`](./MET_Universal_Algodue_UPM209_309.csv) | Meter | Universal | Algodue_UPM209_309 | Modbus | 103 | 28 |
| [`OTH_Universal_Algodue_UPM209_309.csv`](./OTH_Universal_Algodue_UPM209_309.csv) | Other | Universal | Algodue_UPM209_309 | Modbus | 103 | 28 |
| [`INV_Victron_Generic.csv`](./INV_Victron_Generic.csv) | Inverter | VICTRON | Generic | Modbus | 12 | 0 |
| [`OTH_Victron_CCGX-Modbus-TCP-register-list-2.07.csv`](./OTH_Victron_CCGX-Modbus-TCP-register-list-2.07.csv) | Other | Victron | CCGX-Modbus-TCP-register-list-2.07 | Modbus | 290 | 25 |
| [`OTH_Victron_CCGX-Modbus-TCP-register-list-2.23.csv`](./OTH_Victron_CCGX-Modbus-TCP-register-list-2.23.csv) | Other | Victron | CCGX-Modbus-TCP-register-list-2.23 | Modbus | 326 | 25 |
| [`OTH_Victron_CCGX-Modbus-TCP-register-list-2.80.csv`](./OTH_Victron_CCGX-Modbus-TCP-register-list-2.80.csv) | Other | Victron | CCGX-Modbus-TCP-register-list-2.80 | Modbus | 565 | 29 |
| [`INV_Victron_Modbus_Vectron.csv`](./INV_Victron_Modbus_Vectron.csv) | Inverter | Victron | Modbus_Vectron | Modbus | 83 | 30 |
| [`OTH_Victron_Modbus_Vectron.csv`](./OTH_Victron_Modbus_Vectron.csv) | Other | Victron | Modbus_Vectron | Modbus | 83 | 30 |
| [`INV_Waaree_new.csv`](./INV_Waaree_new.csv) | Inverter | Waaree | new | Modbus | 46 | 5 |
| [`OTH_Waaree_new.csv`](./OTH_Waaree_new.csv) | Other | Waaree | new | Modbus | 46 | 5 |
| [`SEN_Webdyn_M200-03.csv`](./SEN_Webdyn_M200-03.csv) | Sensor | Webdyn | M200-03 | Modbus | 1 | 0 |
| [`SEN_Webdyn_RK100-01.csv`](./SEN_Webdyn_RK100-01.csv) | Sensor | Webdyn | RK100-01 | Modbus | 2 | 1 |
| [`OTH_Webdyn_Sens_IO-8AI_work.csv`](./OTH_Webdyn_Sens_IO-8AI_work.csv) | Other | Webdyn | Sens_IO-8AI_work | Modbus | 34 | 1 |
| [`OTH_Webdyn_Sens_IO_16DI.csv`](./OTH_Webdyn_Sens_IO_16DI.csv) | Other | Webdyn | Sens_IO_16DI | Modbus | 60 | 1 |
| [`OTH_Webdyn_Sens_IO_8AI.csv`](./OTH_Webdyn_Sens_IO_8AI.csv) | Other | Webdyn | Sens_IO_8AI | Modbus | 34 | 0 |
| [`OTH_Webdyn_Sens_IO_8DI-6DO.csv`](./OTH_Webdyn_Sens_IO_8DI-6DO.csv) | Other | Webdyn | Sens_IO_8DI-6DO | Modbus | 42 | 11 |
| [`OTH_Webdyn_Sens_IOs_Modbus_definition.csv`](./OTH_Webdyn_Sens_IOs_Modbus_definition.csv) | Other | Webdyn | Sens_IOs_Modbus_definition | Modbus | 28 | 3 |
| [`OTH_Webdyn_WebdynBridge_TableDesRegistres.csv`](./OTH_Webdyn_WebdynBridge_TableDesRegistres.csv) | Other | Webdyn | WebdynBridge_TableDesRegistres | Modbus | 24 | 1 |
| [`MET_Weidmuller_610.csv`](./MET_Weidmuller_610.csv) | Meter | Weidmuller | 610 | Modbus | 122 | 3 |
| [`OTH_Weidmuller_610.csv`](./OTH_Weidmuller_610.csv) | Other | Weidmuller | 610 | Modbus | 122 | 25 |
| [`OTH_Weidmuller_Transclinic.csv`](./OTH_Weidmuller_Transclinic.csv) | Other | Weidmuller | Transclinic | Modbus | 33 | 33 |
| [`BAT_X7_BESS-AllInOne.csv`](./BAT_X7_BESS-AllInOne.csv) | Battery | X7 | BESS-AllInOne | Modbus | 794 | 746 |
| [`BAT_X7_BMS.csv`](./BAT_X7_BMS.csv) | Battery | X7 | BMS | Modbus | 24 | 22 |
| [`BAT_X7_BMS-Cells.csv`](./BAT_X7_BMS-Cells.csv) | Battery | X7 | BMS-Cells | Modbus | 368 | 368 |
| [`OTH_X7_Cabinet.csv`](./OTH_X7_Cabinet.csv) | Other | X7 | Cabinet | Modbus | 45 | 45 |
| [`BAT_X7_EMS.csv`](./BAT_X7_EMS.csv) | Battery | X7 | EMS | Modbus | 164 | 164 |
| [`MET_X7_Meters.csv`](./MET_X7_Meters.csv) | Meter | X7 | Meters | Modbus | 66 | 66 |
| [`INV_X7_PCS.csv`](./INV_X7_PCS.csv) | Inverter | X7 | PCS | Modbus | 127 | 81 |
| [`INV_ZeverSolar_MODBUS_ZeverSolar.csv`](./INV_ZeverSolar_MODBUS_ZeverSolar.csv) | Inverter | ZEVERSOLAR | MODBUS_ZeverSolar | Modbus | 82 | 16 |
| [`INV_Zucchetti_Centro_Sistemi_S.p.a_ZUCCHETTI_V3-V4.csv`](./INV_Zucchetti_Centro_Sistemi_S.p.a_ZUCCHETTI_V3-V4.csv) | Inverter | ZUCCHETTI | Centro_Sistemi_S.p.a_ZUCCHETTI_V3-V4 | Modbus | 1,778 | 39 |
| [`MET_Zelec_EM3100PURSM.csv`](./MET_Zelec_EM3100PURSM.csv) | Meter | Zelec | EM3100PURSM | Modbus | 97 | 14 |
| [`MET_Zelec_EM35TCPURS.csv`](./MET_Zelec_EM35TCPURS.csv) | Meter | Zelec | EM35TCPURS | Modbus | 171 | 26 |
| [`OTH_ZeverSolar_MODBUS_ZeverSolar.csv`](./OTH_ZeverSolar_MODBUS_ZeverSolar.csv) | Other | ZeverSolar | MODBUS_ZeverSolar | Modbus | 82 | 15 |
| [`OTH_ZeverSolar_Modbus.csv`](./OTH_ZeverSolar_Modbus.csv) | Other | ZeverSolar | Modbus | Modbus | 75 | 15 |
| [`OTH_ZeverSolar_ZeverSolar.csv`](./OTH_ZeverSolar_ZeverSolar.csv) | Other | ZeverSolar | ZeverSolar | Modbus | 98 | 4 |
| [`INV_ZeverSolar_Modbus.csv`](./INV_ZeverSolar_Modbus.csv) | Inverter | Zeversolar | Modbus | Modbus | 75 | 15 |
| [`INV_Zucchetti_Azzuro-3PH.csv`](./INV_Zucchetti_Azzuro-3PH.csv) | Inverter | Zucchetti | Azzuro-3PH | Modbus | 503 | 2 |
| [`INV_Zucchetti_Centro_Sistemi_S.p.a_INV_ZUCCHETTI_V3-V4.csv`](./INV_Zucchetti_Centro_Sistemi_S.p.a_INV_ZUCCHETTI_V3-V4.csv) | Inverter | Zucchetti | Centro_Sistemi_S.p.a_INV_ZUCCHETTI_V3-V4 | Modbus | 1,778 | 39 |
| [`INV_Zucchetti_Centro_Sistemi_S.p.a_ZCS_Hyd_3PH.csv`](./INV_Zucchetti_Centro_Sistemi_S.p.a_ZCS_Hyd_3PH.csv) | Inverter | Zucchetti | Centro_Sistemi_S.p.a_ZCS_Hyd_3PH | Modbus | 505 | 27 |
| [`OTH_Zucchetti_Centro_Sistemi_S.p.a_ZCS_Hyd_3PH.csv`](./OTH_Zucchetti_Centro_Sistemi_S.p.a_ZCS_Hyd_3PH.csv) | Other | Zucchetti | Centro_Sistemi_S.p.a_ZCS_Hyd_3PH | Modbus | 505 | 27 |
| [`INV_Fronius_Tauro_50-3-D_1.35.4-1_Float.csv`](./INV_Fronius_Tauro_50-3-D_1.35.4-1_Float.csv) | Inverter | fronius | Tauro_50-3-D_1.35.4-1_Float | Modbus | 193 | 34 |
| [`MET_map_modbus_meter_slenergy6.csv`](./MET_map_modbus_meter_slenergy6.csv) | Meter | map | modbus_meter_slenergy6 | Modbus | 7 | 1 |


## 📚 Documents Sources & Traçabilité Constructeurs

Tous les fichiers de définition de cette bibliothèque sont rigoureusement documentés et dérivés de manuels fabricants officiels conservés localement sous [`sources/`](file:///C:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/sources/).

Pour consulter la liste complète et détaillée des sources par équipementier, gamme, modèle, paramètres de communication (baud rate, parité, codes fonctions, ordre des mots) et empreintes cryptographiques SHA-256 :
👉 **Voir le grand livre de traçabilité : [`production/SOURCES.md`](file:///C:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/SOURCES.md)** ainsi que le manifeste machine [`sources/MANIFEST.json`](file:///C:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/sources/MANIFEST.json).

### Synthèse des Équipementiers Documentés :
| Équipementier | Catégories Couvertes | Documents Disponibles | Définitions Webdyn Produites |
|---|---|:---:|:---:|
| **ABB** | EVCharger, Inverter, Logger, Meter, Sensor | `16` source(s) | `225` fichiers |
| **Acrel** | Meter | `1` source(s) | `1` fichiers |
| **CarloGavazzi** | Meter | `9` source(s) | `15` fichiers |
| **Chint** | Inverter, Meter | `5` source(s) | `22` fichiers |
| **Cyberpower** | Inverter | `1` source(s) | `2` fichiers |
| **Delta** | Inverter | `2` source(s) | `25` fichiers |
| **Dyness** | Battery | `1` source(s) | `3` fichiers |
| **FoxESS** | Inverter_Battery | `1` source(s) | `15` fichiers |
| **Fronius** | Inverter, Meter | `5` source(s) | `44` fichiers |
| **Goodwe** | Inverter | `5` source(s) | `20` fichiers |
| **Growatt** | Inverter | `1` source(s) | `5` fichiers |
| **Huawei** | Battery, Inverter, Logger, Meter | `6` source(s) | `80` fichiers |
| **Ingeteam** | EVCharger, Inverter | `2` source(s) | `7` fichiers |
| **Janitza** | Meter | `3` source(s) | `17` fichiers |
| **Kaco** | Inverter | `3` source(s) | `36` fichiers |
| **Legrand** | Meter | `2` source(s) | `7` fichiers |
| **MarsEnergy** | Controller_BMS | `1` source(s) | `20` fichiers |
| **SAJ** | Inverter | `2` source(s) | `14` fichiers |
| **SMA** | Gateway, Inverter, Meter, STP125-70 | `17` source(s) | `32` fichiers |
| **Schneider** | Inverter, Meter | `6` source(s) | `33` fichiers |
| **Socomec** | Gateway, Meter | `9` source(s) | `139` fichiers |
| **SolarEdge** | Inverter, Meter | `3` source(s) | `19` fichiers |
| **Solis** | Inverter | `3` source(s) | `17` fichiers |
| **Solplanet** | Gateway | `1` source(s) | `1` fichiers |
| **Sungrow** | Battery, Inverter, Logger | `4` source(s) | `22` fichiers |
| **X7** | Inverter | `1` source(s) | `7` fichiers |

---
