# Rapport d'Audit des Doublons — Suite de Production Webdyn

- **Répertoire analysé** : `generated_definitions`
- **Total fichiers scannés** : **2231**
- **Fichiers doublons identifiés** : **107**
- **Doublons stricts (SHA-256)** : **66**
- **Conflits de catégorie (OTH_)** : **2**
- **Alias & synonymes** : **20**
- **Familles matérielles** : **314**

---

## 1. Doublons Stricts (Empreinte SHA-256 Identique)

Ces fichiers sont rigoureusement identiques au niveau binaire (même empreinte SHA-256, même en-tête et mêmes registres).

| Fichier Canonique | Doublons Stricts | Modèle Matériel | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| `BAT_FoxESS_RSeries.csv` | `BAT_Fox_ESS_FOXESS_RSERIES_BATTERY.csv` | R SERIES Storage | 49 | Supprimer le doublon strict ou créer un alias |
| `BAT_MarsRenewable_LC.csv` | `BAT_Webdyn_MarsRenewable_LC.csv` | LC | 12 | Supprimer le doublon strict ou créer un alias |
| `BAT_MarsRenewable_PCS-Bank.csv` | `BAT_Webdyn_MarsRenewable_PCS-Bank.csv` | PCS-Bank-Rack | 86 | Supprimer le doublon strict ou créer un alias |
| `BAT_Robestec_battery_Robestec_BESS.csv` | `BAT_Robestec_battery_Robestec_BESS_-_Copie.csv` | BESS | 1723 | Supprimer le doublon strict ou créer un alias |
| `BAT_SL_Energy_SLEnergy_Generator.csv` | `OTH_SL_Energy_SLEnergy_Generator.csv` | Generator | 16 | Supprimer le doublon strict ou créer un alias |
| `BAT_Webdyn_SLEnergy_Battery.csv` | `BAT_SL_Energy_SLEnergy_Battery.csv` | Battery | 188 | Supprimer le doublon strict ou créer un alias |
| `BAT_Webdyn_slenergy_bms5_p1_1.csv` | `BAT_Webdyn_slenergy_bms5_p1_1_1.csv` | Generic12_test | 213 | Supprimer le doublon strict ou créer un alias |
| `BAT_Webdyn_SLEnergy_Generator.csv` | `OTH_Webdyn_SLEnergy_Generator.csv` | Generator | 16 | Supprimer le doublon strict ou créer un alias |
| `BAT_Webdyn_SLEnergy_PCS.csv` | `BAT_SL_Energy_SLEnergy_PCS.csv` | PCS-PCS settings-Battery Cloud Management | 683 | Supprimer le doublon strict ou créer un alias |
| `INV_Delta_WPM0118CD_SunSpec_inverter_Delta_RPI-M50A_01.45_modbusRTU.csv` | `INV_Delta_WPM0118CD_SunSpec_inverter_Delta_RPI-M50A_01.50_modbusRTU.csv`<br>`INV_Delta_WPM0118CD_SunSpec_inverter_Delta_RPI-M50A_01.52_modbusRTU.csv` | RPI-M50A | 607 | Supprimer le doublon strict ou créer un alias |
| `INV_Delta_WPM01397E_SunSpec_inverter_Delta_131200200_02.45_modbusRTU.csv` | `INV_Delta_WPM01397E_SunSpec_inverter_Delta_131200200_02.47_modbusRTU.csv` | 131200200 | 823 | Supprimer le doublon strict ou créer un alias |
| `INV_Ginlong_Solis_WPM013804_SunSpec_inverter_Ginlong_Solar_inverter_760235_modbusRTU.csv` | `INV_Ginlong_Solis_WPM013804_SunSpec_inverter_Ginlong_Solar_inverter_800236_modbusRTU.csv` | Solar_inverter | 353 | Supprimer le doublon strict ou créer un alias |
| `INV_Growatt_TL3-X_Series.csv` | `INV_Growatt_W_Growatt_TL3-X_Series.csv` | TL3-X_Series | 371 | Supprimer le doublon strict ou créer un alias |
| `INV_INVT_Generic.csv` | `INV_INVT_Solar_ModbusRTU_INVT_Generic.csv` | Generic | 41 | Supprimer le doublon strict ou créer un alias |
| `INV_JEMA_Generic_revG.csv` | `INV_Webdyn_JEMA_Modbus_revG.csv` | Generic | 60 | Supprimer le doublon strict ou créer un alias |
| `INV_MarsRenewable_Meter.csv` | `INV_Webdyn_MarsRenewable_Meter.csv` | LC_Meter | 24 | Supprimer le doublon strict ou créer un alias |
| `INV_MarsRenewable_PV.csv` | `INV_Webdyn_MarsRenewable_PV.csv` | PV | 52 | Supprimer le doublon strict ou créer un alias |
| `INV_MarsRenewable_Rack-Cell.csv` | `INV_Webdyn_MarsRenewable_Rack-Cell.csv` | LC_V7 | 42 | Supprimer le doublon strict ou créer un alias |
| `INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv` | `INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv` | FREESUN | 15 | Supprimer le doublon strict ou créer un alias |
| `INV_PowerOne_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv` | `INV_PowerOne_LM1_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv` | PVS-30-TL | 956 | Supprimer le doublon strict ou créer un alias |
| `INV_SMA_STPX.csv` | `INV_Webdyn_SMA_STPX.csv` | STP X | 509 | Supprimer le doublon strict ou créer un alias |
| `INV_SMA_WPMXXXXXX_SMA_STPCORE1_50-41_AllRegs_NewTags.csv` | `INV_SMA_WPMXXXXXX_SMA_STPCORE1_50-41_AllRegs_OldTags.csv` | STP50-41 | 568 | Supprimer le doublon strict ou créer un alias |
| `INV_SMA_WPMXXXXXX_SMA_STPX_AllRegs_NewTags.csv` | `INV_SMA_WPMXXXXXX_SMA_STPX_AllRegs_OldTags.csv` | STPX | 520 | Supprimer le doublon strict ou créer un alias |
| `INV_SolarEdge_66KW.csv` | `INV_Webdyn_INDIA_SolarEdge_66KW.csv` | 66KW | 38 | Supprimer le doublon strict ou créer un alias |
| `INV_Webdyn_Delta_Sunspec_V1.20.csv` | `INV_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv` | M88H | 640 | Supprimer le doublon strict ou créer un alias |
| `INV_Webdyn_Fronius_Sunspec_IntSF.csv` | `INV_Fronius_Modbus_SunS_Fronius_IntSF.csv` | sunspec | 55 | Supprimer le doublon strict ou créer un alias |
| `INV_Webdyn_MODBUS_SUNGROW.csv` | `INV_Webdyn_SUNGROW_MODBUS_PM.csv` | Generic | 99 | Supprimer le doublon strict ou créer un alias |
| `INV_Webdyn_Simu_RTU.csv` | `INV_Webdyn_Simu_TCP.csv` | INV_Simu | 5 | Supprimer le doublon strict ou créer un alias |
| `INV_Webdyn_SLEnergy_Inverter.csv` | `INV_SL_Energy_SLEnergy_Inverter.csv` | Inverter | 263 | Supprimer le doublon strict ou créer un alias |
| `INV_Webdyn_WPM015004-Auneuil-Injection-Reseau_modbusRTU_Inverter_Solis_Generic.csv` | `INV_Webdyn_WPM015004-Auneuil-Injection-Reseau_modbusRTU_Inverter_Solis_Generic_1.csv`<br>`INV_Webdyn_WPM015004-Auneuil-Injection-Reseau_modbusRTU_Inverter_Solis_Generic_3.csv`<br>`INV_Webdyn_WPM015004-Auneuil-Injection-Reseau_modbusRTU_Inverter_Solis_Generic_4.csv` | Generic | 96 | Supprimer le doublon strict ou créer un alias |
| `INV_Zigor_WPM_Inverter_ZIGOR_STR120-250.csv` | `INV_Webdyn_WPM_Inverter_ZIGOR_STR120-250.csv` | STR120-250 | 2608 | Supprimer le doublon strict ou créer un alias |
| `MET_CHAUVINARNOUX_ULYS-TT_M.csv` | `MET_CEFEM_DEF_GridMeterChauvinArnouxUlysTT_M.csv` | Ulys TT-M | 19 | Supprimer le doublon strict ou créer un alias |
| `MET_CHAUVINARNOUX_ULYS_FLEX.csv` | `MET_ChauvinArnoud_CHAUVINARNOUX_ULYS_FLEX.csv` | Ulys_Flex | 130 | Supprimer le doublon strict ou créer un alias |
| `MET_Enerfox_FXM-210-XXX.csv` | `MET_Enerfox_METER_ENERFOX_FXM-210-XXX.csv` | FXM-210-XXX | 64 | Supprimer le doublon strict ou créer un alias |
| `MET_IME_CE4DF3DTMID.csv` | `MET_IME_Meter_IME_CE4DF3DTMID.csv`<br>`MET_Legrand_IME_CE4DF3DTMID.csv` | CE4DF3DTMID | 66 | Supprimer le doublon strict ou créer un alias |
| `MET_IME_CE4TBDTMID.csv` | `MET_IME_Meter_IME_CE4TBDTMID.csv` | CE4TBDTMID | 82 | Supprimer le doublon strict ou créer un alias |
| `MET_IME_IM-MFD4ORFCDT1.csv` | `MET_IME_Meter_IME_IM-MFD4ORFCDT1.csv` | IM-MFD4ORFCDT1 | 73 | Supprimer le doublon strict ou créer un alias |
| `MET_Janitza_Generic.csv` | `MET_Janitza_W_Janitza_Generic.csv` | Generic | 61 | Supprimer le doublon strict ou créer un alias |
| `MET_Janitza_Generic_RTU.csv` | `MET_Janitza_Generic_TCP.csv` | Generic | 61 | Supprimer le doublon strict ou créer un alias |
| `MET_Polier_METER_POLIER_PROMOD380.csv` | `MET_SolarLog_METER_POLIER_PROMOD380.csv` | PROMOD380 | 88 | Supprimer le doublon strict ou créer un alias |
| `MET_SOCOMEC_COUNTIS_E23.csv` | `MET_Socomec_W_SOCOMEC_COUNTIS_E23.csv` | COUNTIS_E23 | 59 | Supprimer le doublon strict ou créer un alias |
| `MET_Socomec_Diris_A10.csv` | `MET_Socomec_W_Socomec_Diris_A10.csv` | Diris_A10 | 75 | Supprimer le doublon strict ou créer un alias |
| `MET_SOCOMEC_Diris_A40.csv` | `MET_Socomec_W_SOCOMEC_Diris_A40.csv` | Diris_A40 | 35 | Supprimer le doublon strict ou créer un alias |
| `MET_SOCOMEC_DIRIS_B10.csv` | `MET_Socomec_W_SOCOMEC_DIRIS_B10.csv` | DIRIS_B10 | 23 | Supprimer le doublon strict ou créer un alias |
| `MET_Socomec_Diris_B30.csv` | `MET_Socomec_W_Socomec_Diris_B30.csv` | Diris_B30 | 20 | Supprimer le doublon strict ou créer un alias |
| `MET_Socomec_Diris_I30.csv` | `MET_Socomec_W_Socomec_Diris_I30.csv` | Diris_I30 | 20 | Supprimer le doublon strict ou créer un alias |
| `MET_SOCOMEC_E33_E43.csv` | `MET_Socomec_W_SOCOMEC_E33_E43.csv` | E33_E43 | 73 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_Janitza_UMG-604.csv` | `MET_Webdyn_WPM013D96_modbusTCP_meter_Janitza_UMG-604-TCP_1.csv` | UMG-604-TCP | 4 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_METER_SCHNEIDER_PM5xxx.csv` | `MET_Schneider_METER_SCHNEIDER_PM5xxx.csv` | PM5xxx | 875 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_SCH_iem3xxx.csv` | `MET_Schneider_SCH_iem3xxx.csv`<br>`MET_Webdyn_WPM015004_SCH_iem3xxx.csv` | IEM 3255 | 58 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_SLEnergy_ConsoMeter.csv` | `MET_SL_Energy_SLEnergy_ConsoMeter.csv` | ConsoMeter | 4 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_SLEnergy_GridMeter.csv` | `MET_SL_Energy_SLEnergy_GridMeter.csv` | GridMeter | 65 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_SLEnergy_OffGridMeter.csv` | `MET_SL_Energy_SLEnergy_OffGridMeter.csv` | OffGridMeter | 33 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_test_meter.csv` | `MET_Webdyn_test_meter_1_-_Copie.csv` | SE-RGMTR-1D-240C-A | 74 | Supprimer le doublon strict ou créer un alias |
| `MET_Webdyn_WPM01A743_Janitza_Generic.csv` | `MET_Webdyn_WPM01A743_Janitza_Generic_oldTag.csv`<br>`MET_Webdyn_WPM01A779_Janitza_Generic.csv` | Generic | 61 | Supprimer le doublon strict ou créer un alias |
| `OTH_Delta_Delta-Sunspec_protocol_v1.11_20180725_part29.csv` | `OTH_Delta_Delta-Sunspec_protocol_v1.12_20190403_part29.csv` | Delta-Sunspec_protoc_p29 | 52 | Supprimer le doublon strict ou créer un alias |
| `OTH_Power_One.csv` | `OTH_Webdyn_INDIA_Power_One.csv` | One | 58 | Supprimer le doublon strict ou créer un alias |
| `OTH_Webdyn_Delta-Sunspec_protocol_v1.11_20180725_part29.csv` | `OTH_Webdyn_Delta-Sunspec_protocol_v1.12_20190403_part29.csv` | Delta-Sunspec_protoc_p29 | 52 | Supprimer le doublon strict ou créer un alias |
| `OTH_Webdyn_slenergy_bms5_p1_1.csv` | `OTH_Webdyn_slenergy_bms5_p1_1_1.csv` | Generic12_test | 213 | Supprimer le doublon strict ou créer un alias |
| `OTH_Webdyn_SLEnergy_PCS.csv` | `OTH_SL_Energy_SLEnergy_PCS.csv` | PCS-PCS settings-Battery Cloud Management | 683 | Supprimer le doublon strict ou créer un alias |
| `SEN_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv` | `SEN_IngenieurBuro_WPM_MODBUS_Si-RS485-TC-2T-v.csv` | Si-RS485TC-2T-v-MB | 7 | Supprimer le doublon strict ou créer un alias |
| `SEN_IngenieurBuro_TA-RS485-MB.csv` | `SEN_IngenieurBuro_SENSOR_INGENIEURBURO_TA-RS485-MB.csv` | TA-RS485-MB | 3 | Supprimer le doublon strict ou créer un alias |
| `SEN_IngenieurBuro_TM-RS485-MB.csv` | `SEN_IngenieurBuro_SENSOR_INGENIEURBURO_TM-RS485-MB.csv` | TM-RS485-MB | 3 | Supprimer le doublon strict ou créer un alias |
| `SEN_Webdyn_Hukseflux.csv` | `SEN_Webdyn_WPM015004_Hukseflux.csv` | SR05-D1A3 | 27 | Supprimer le doublon strict ou créer un alias |
| `SEN_Webdyn_SENSOR_SEVEN_SENSOR_3S-IS.csv` | `SEN_Seven_sensor_solutions_SENSOR_SEVEN_SENSOR_3S-IS.csv` | 3S-IS | 9 | Supprimer le doublon strict ou créer un alias |
| `SEN_Webdyn_SENSOR_SEVEN_SENSOR_3S-IS_SUNSPEC.csv` | `SEN_Seven_sensor_solutions_SENSOR_SEVEN_SENSOR_3S-IS_SUNSPEC.csv` | 3S-IS | 45 | Supprimer le doublon strict ou créer un alias |

---

## 2. Conflits de Catégorie (Classification Générique `OTH_` vs Fichier Typé)

Ces fichiers `OTH_` sont des doublons stricts d'équipements déjà qualifiés en tant qu'onduleurs (`INV_`), compteurs (`MET_`), capteurs (`SEN_`) ou batteries (`BAT_`).

| Fichier Canonique Qualifié | Doublon Redondant `OTH_` | Modèle Matériel | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| `INV_CEFEM_Table_Modbus_Trio-Sun_Trio-Top_2020_07_01.csv` | `OTH_CEFEM_Table_Modbus_Trio-Sun_Trio-Top_2020_07_01.csv` | Table_Modbus_Trio-Sun_Trio-Top | 32 | Supprimer le `OTH_` redondant |
| `INV_SAJ_SAJ_sunTrioTL.csv` | `OTH_SAJ_SAJ_sunTrioTL.csv` | SAJ_sunTrioTL | 45 | Supprimer le `OTH_` redondant |

---

## 3. Alias Matériels et Variantes Synonymes de Nommage

Fichiers représentant le même équipement ou la même gamme sous deux désignations alternatives.

| Fichier Canonique | Alias / Synonyme | Description | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| `BAT_SLEnergy_bms5_p1_1.csv` | `BAT_SLEnergy_bms5_p1_1_1.csv`<br>`OTH_SLEnergy_bms5_p1_1.csv`<br>`OTH_SLEnergy_bms5_p1_1_1.csv` | Hardware model alias / synonymous naming sharing identical register map | 213 | Enregistrer dans le registre d'alias |
| `BAT_SLEnergy_Generator.csv` | `BAT_SL_Energy_SLEnergy_Generator.csv`<br>`BAT_Webdyn_SLEnergy_Generator.csv`<br>`OTH_SLEnergy_Generator.csv` | Hardware model alias / synonymous naming sharing identical register map | 16 | Enregistrer dans le registre d'alias |
| `INV_FoxESS_R_V_H3plus_GmaxPlus.csv` | `OTH_Fox_ESS_R_V_H3plus_GmaxPlus.csv` | Hardware model alias / synonymous naming sharing identical register map | 87 | Enregistrer dans le registre d'alias |
| `INV_GOODWE_GT-SERIES.csv` | `INV_Goodwe_GT_Series.csv` | Hardware model alias / synonymous naming sharing identical register map | 305 | Enregistrer dans le registre d'alias |
| `INV_Goodwe_HT1500.csv` | `INV_Goodwe_HT_1500V.csv` | Hardware model alias / synonymous naming sharing identical register map | 288 | Enregistrer dans le registre d'alias |
| `INV_Solax_SOLAX-X1-X3-GENERIC.csv` | `INV_Solax_SOLAX-X1-X3-GENERIC_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 156 | Enregistrer dans le registre d'alias |
| `INV_SolaX_X1-X3-GENERIC.csv` | `INV_SolaX_X1-X3-GENERIC_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 156 | Enregistrer dans le registre d'alias |
| `INV_Sungrow_RTU_PM.csv` | `INV_Sungrow_TCP_PM.csv` | Hardware model alias / synonymous naming sharing identical register map | 81 | Enregistrer dans le registre d'alias |
| `INV_Webdyn_INDIA_Sungrow_RTU_PM.csv` | `INV_Webdyn_INDIA_Sungrow_TCP_PM.csv` | Hardware model alias / synonymous naming sharing identical register map | 81 | Enregistrer dans le registre d'alias |
| `INV_Webdyn_INV_Simu_RTU.csv` | `INV_Webdyn_INV_Simu_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 5 | Enregistrer dans le registre d'alias |
| `MET_ChauvinArnoux_CHAUVINARNOUX_ULYS-TT_M.csv` | `OTH_ChauvinArnoux_CHAUVINARNOUX_ULYS-TT_M.csv`<br>`OTH_ChauvinArnoux_ULYS-TT_M.csv` | Hardware model alias / synonymous naming sharing identical register map | 19 | Enregistrer dans le registre d'alias |
| `MET_Janitza_JANITZA_GENERIC.csv` | `MET_Janitza_JANITZA_GENERIC_RTU.csv`<br>`MET_Janitza_JANITZA_GENERIC_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 61 | Enregistrer dans le registre d'alias |
| `MET_Legrand_CF-GENERIC.csv` | `MET_Legrand_CF_Generic.csv`<br>`MET_Legrand_METER_LEGRAND_CF-GENERIC.csv` | Hardware model alias / synonymous naming sharing identical register map | 11 | Enregistrer dans le registre d'alias |
| `MET_Webdyn_MODBUS_Janitza.csv` | `MET_Janitza_MODBUS_UMG-104_Janitza_RTU.csv`<br>`MET_Webdyn_MODBUSRTU_Janitza.csv` | Hardware model alias / synonymous naming sharing identical register map | 61 | Enregistrer dans le registre d'alias |
| `OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv` | `OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 15 | Enregistrer dans le registre d'alias |
| `OTH_Siebert_XC4XX_RTU.csv` | `OTH_Siebert_XC4XX_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 3 | Enregistrer dans le registre d'alias |
| `OTH_SLEnergy_Work.csv` | `OTH_SL_Energy_SLEnergy_Work.csv`<br>`OTH_Webdyn_SLEnergy_Work.csv` | Hardware model alias / synonymous naming sharing identical register map | 612 | Enregistrer dans le registre d'alias |
| `OTH_Webdyn_DecoupeTrameModbusExcel.csv` | `OTH_Webdyn_DecoupeTrameModbusExcel_-_Copie.csv` | Hardware model alias / synonymous naming sharing identical register map | 4 | Enregistrer dans le registre d'alias |
| `SEN_IngenieurBuro_INGENIEURBURO_Si-RS485-TC-2T-v.csv` | `OTH_IngenieurBuro_INGENIEURBURO_Si-RS485-TC-2T-v.csv`<br>`OTH_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv`<br>`OTH_IngenieurBuro_Si-RS485-TC-2T-v.csv`<br>`OTH_IngenieurBuro_WPM_MODBUS_Si-RS485-TC-2T-v.csv` | Hardware model alias / synonymous naming sharing identical register map | 7 | Enregistrer dans le registre d'alias |
| `SEN_IngenieurBuro_INGENIEURBURO_Si-RS485TC-3T-MB.csv` | `OTH_IngenieurBuro_INGENIEURBURO_Si-RS485TC-3T-MB.csv`<br>`OTH_IngenieurBuro_Si-RS485TC-3T-MB.csv` | Hardware model alias / synonymous naming sharing identical register map | 10 | Enregistrer dans le registre d'alias |

---

## 4. Familles Matérielles Partageant la Même Carte Modbus

Équipements distincts d'un même constructeur partageant rigoureusement le même profil de registres (ex. Schneider iEM3100..iEM3555, GoodWe ET/EH). Total : **314** familles répertoriées.

| Modèle de Référence | Modèles de la Famille | Registres | Recommandation |
| :--- | :--- | :---: | :--- |
| `BAT_AlroSolar_BMS_Definition.csv` | `BAT_Alro_Solar_BMS_Definition.csv` | 114 | Conserver (modèles réels distincts) |
| `BAT_Atmoce_MC-MG.csv` | `BAT_Atmoce_BATTERY_ATMOCE_MC-MG.csv` | 47 | Conserver (modèles réels distincts) |
| `BAT_Dyness_DH100F.csv` | `BAT_Dyness_DH200F.csv` | 913 | Conserver (modèles réels distincts) |
| `BAT_FoxESS_H3_Pro.csv` | `BAT_FoxESS_P3_H3_PRO_BATTERY.csv`, `BAT_FoxESS_P3_Pro.csv` | 893 | Conserver (modèles réels distincts) |
| `BAT_FoxESS_RSeries.csv` | `BAT_FoxESS_RSERIES_BATTERY.csv` | 49 | Conserver (modèles réels distincts) |
| `BAT_Huawei_Luna2000.csv` | `BAT_Huawei_Luna_2000.csv` | 170 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_LocalControler_Bank_Rack_Cell.csv` | `BAT_MarsRenewable_LocalControler_Bank_Rack_Cell.csv` | 84 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_LocalControler_PCS_Information.csv` | `BAT_MarsRenewable_LocalControler_PCS_Information.csv` | 111 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv` | `BAT_MarsRenewable_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv` | 35 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_MarsLocalController_V7.0_EMS_Slave247.csv` | `BAT_MarsRenewable_MarsLocalController_V7.0_EMS_Slave247.csv` | 12 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_MarsLocalController_V7.0_RackCell_SlaveN.csv` | `BAT_MarsRenewable_MarsLocalController_V7.0_RackCell_SlaveN.csv` | 37 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_MarsLocalController_V7.0_webdyn.csv` | `BAT_MarsRenewable_MarsLocalController_V7.0_webdyn.csv` | 192 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_MarsRenewable_LocalControler_Bank_Rack_Cell.csv` | `BAT_MarsRenewable_MarsRenewable_LocalControler_Bank_Rack_Cell.csv` | 83 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_MarsRenewable_Rack_V7.csv` | `BAT_MarsRenewable_MarsRenewable_Rack_V7.csv` | 110 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_Rack-Cell.csv` | `BAT_MarsEnergy_MarsRenewable_Rack-Cell.csv`, `BAT_MarsRenewable_MarsRenewable_Rack-Cell.csv`, `BAT_Webdyn_MarsRenewable_Rack-Cell.csv` | 42 | Conserver (modèles réels distincts) |
| `COM_Atmoce_MC-MG.csv` | `COM_Atmoce_LOGGER_ATMOCE_MC-MG.csv` | 47 | Conserver (modèles réels distincts) |
| `COM_GoodWe_Ezlogger3000.csv` | `COM_Goodwe_EzLogger_Pro.csv` | 25 | Conserver (modèles réels distincts) |
| `COM_Huawei_SmartLogger3000.csv` | `COM_Huawei_HUAWEI_SmartLogger3000.csv` | 83 | Conserver (modèles réels distincts) |
| `COM_MarsEnergy_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv` | `COM_MarsRenewable_MARS_LocalController_V7_Rack_Un_seul_element_WebdynSunPM.csv` | 35 | Conserver (modèles réels distincts) |
| `COM_MarsEnergy_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv` | `COM_MarsRenewable_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv` | 5 | Conserver (modèles réels distincts) |
| `COM_Robestec_SJ2025B2846SJ2025B2846yubangmodbusTcp-en_1.csv` | `COM_Robestec_battery_SJ2025B2846SJ2025B2846yubangmodbusTcp-en_1.csv` | 1457 | Conserver (modèles réels distincts) |
| `INV_5_Generic.csv` | `INV_Webdyn_INDIA_Sunspec_5.csv` | 49 | Conserver (modèles réels distincts) |
| `INV_ABB_Modbus_ABB_SunSpec_Light.csv` | `INV_PowerOne_Modbus_ABB_SunSpec_Light.csv` | 89 | Conserver (modèles réels distincts) |
| `INV_ABB_MODBUS_POWERONE_Ultra_Inverter_Module.csv` | `INV_PowerOne_ABB_MODBUS_POWERONE_Ultra_Inverter_Module.csv` | 56 | Conserver (modèles réels distincts) |
| `INV_ABB_Modbus_TRIO.csv` | `INV_Webdyn_ABB_Modbus_TRIO.csv` | 17 | Conserver (modèles réels distincts) |
| `INV_ABB_Modbus_Ultra_Module.csv` | `INV_Webdyn_ABB_Modbus_Ultra_Module.csv` | 56 | Conserver (modèles réels distincts) |
| `INV_ABB_PVS-100_120_PVS1xx_TCP_PICS.csv` | `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS.csv` | 986 | Conserver (modèles réels distincts) |
| `INV_ABB_PVS50_MODBUS_light.csv` | `INV_PowerOne_ABB_PVS50_MODBUS_light.csv` | 76 | Conserver (modèles réels distincts) |
| `INV_ABB_PVS_100_120_SunspecLight.csv` | `INV_PowerOne_ABB_PVS_100_120_SunspecLight.csv` | 118 | Conserver (modèles réels distincts) |
| `INV_ABB_SunSpec_IM_1_103_160.csv` | `INV_Webdyn_ABB_SunSpec_IM_1_103_160.csv` | 127 | Conserver (modèles réels distincts) |
| `INV_ABB_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv` | `INV_ABB_LM1_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv`, `INV_PowerOne_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv` | 956 | Conserver (modèles réels distincts) |
| `INV_ABB_SUNSPEC_PVS1xx_PICS_V002_GL20181011.csv` | `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011.csv` | 1118 | Conserver (modèles réels distincts) |
| `INV_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS.csv` | 956 | Conserver (modèles réels distincts) |
| `INV_ABB_Velux_SunSpec_inverter_FIMER_-3Q15-_2201B_modbusTCP.csv` | `INV_ABB_5583-CHEVALIERP2_SunSpec_inverter_ABB_3Q4200_2117B_modbusRTU.csv`, `INV_PowerOne_5583-CHEVALIERP2_SunSpec_inverter_ABB_3Q4200_2117B_modbusRTU.csv`, `INV_PowerOne_Velux_SunSpec_inverter_FIMER_-3Q15-_2201B_modbusTCP.csv` | 1111 | Conserver (modèles réels distincts) |
| `INV_AP_system_Altenergy_Power_System_Inc._three_phase_microinverter_11-12.csv` | `INV_AP_system_Altenergy_Power_System_Inc._single_phase_MicroinverterAPpcs_11-12.csv`, `INV_APsystems_Altenergy_Power_System_Inc._single_phase_MicroinverterAPpcs_11-12.csv`, `INV_APsystems_Altenergy_Power_System_Inc._three_phase_microinverter_11-12.csv` | 113 | Conserver (modèles réels distincts) |
| `INV_CyberPower_3-Phases.csv` | `OTH_Cyberpower_Modbus_3Ph.csv` | 71 | Conserver (modèles réels distincts) |
| `INV_Cyberpower_50000ETL.csv` | `INV_Cyberpower3Ph_50000ETL.csv` | 80 | Conserver (modèles réels distincts) |
| `INV_Delta_Sunspec_V1.20.csv` | `INV_Webdyn_Delta_Sunspec_V1.20.csv` | 640 | Conserver (modèles réels distincts) |
| `INV_Deye_Generic.csv` | `INV_Deye_Inverter_INV_DEYE_GENERIC.csv` | 327 | Conserver (modèles réels distincts) |
| `INV_Deye_GENERICreg-1.csv` | `INV_Deye_Inverter_DEYE_GENERICreg-1.csv` | 86 | Conserver (modèles réels distincts) |
| `INV_Deye_GENERICreg1.csv` | `INV_Deye_Inverter_DEYE_GENERICreg1.csv` | 86 | Conserver (modèles réels distincts) |
| `INV_Eastron_SMART-X96-5E.csv` | `INV_Eastron_EASTRON_SMART-X96-5E.csv` | 125 | Conserver (modèles réels distincts) |
| `INV_Efacec_PV550_Efacec_MODBUS.csv` | `INV_Efasolar_PV550_Efacec_MODBUS.csv` | 145 | Conserver (modèles réels distincts) |
| `INV_FoxESS_H3_Pro.csv` | `INV_Fox_ESS_FOXESS_P3_H3_PRO_INVERTER.csv`, `INV_FoxESS_P3_H3_PRO_INVERTER.csv`, `INV_FoxESS_P3_Pro.csv` | 285 | Conserver (modèles réels distincts) |
| `INV_FoxESS_SunSpec_RV_H3Plus.csv` | `INV_Fox_ESS_R_V_H3plus_GmaxPlus.csv` | 87 | Conserver (modèles réels distincts) |
| `INV_Fronius_Sunspec_IntSF.csv` | `INV_Webdyn_Fronius_Sunspec_IntSF.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_Fronius_Tauro_ECO_Inverter_Register_Map_Int_SF.csv` | `INV_Webdyn_INDIA_Tauro_ECO_Inverter_Register_Map_Int_SF.csv` | 257 | Conserver (modèles réels distincts) |
| `INV_Fronius_Tauro_Inverter_Register_Map_Float.csv` | `INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float.csv` | 279 | Conserver (modèles réels distincts) |
| `INV_Fronius_Tauro_Inverter_Register_Map_Int_SF.csv` | `INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Int_SF.csv` | 287 | Conserver (modèles réels distincts) |
| `INV_Fronius_WPM018C3D_Fronius_Symo_10.0-3-M.csv` | `INV_Fronius_WPM01A481_Fronius_Symo_20.0-3-M.csv` | 177 | Conserver (modèles réels distincts) |
| `INV_Ginlong_Solis_sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15_part11.csv` | `OTH_Delta_M100A_280_COM1.15_part10.csv`, `OTH_Delta_M15A-flex_COM115_part10.csv`, `OTH_Delta_M30A-flex_COM117_part10.csv`, `OTH_Delta_M30A_COM147_part10.csv`, `OTH_Delta_M50A-flex_COM114_part10.csv` *(+4 autres)* | 59 | Conserver (modèles réels distincts) |
| `INV_Goodwe_ES_EM_SBP.csv` | `INV_GoodWe_ES-EM-SBP-BP.csv` | 49 | Conserver (modèles réels distincts) |
| `INV_Goodwe_ET_Series.csv` | `INV_GoodWe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv`, `INV_GoodWe_ET_EH_Hybrid_Series.csv`, `INV_Goodwe_ET_EH_Series.csv` | 79 | Conserver (modèles réels distincts) |
| `INV_Goodwe_MS_DNS_XS.csv` | `INV_GoodWe_XS-DNS-MS-SDT-G3_SERIES.csv`, `INV_Goodwe_XS_DNS_MS.csv` | 50 | Conserver (modèles réels distincts) |
| `INV_Goodwe_MT_series.csv` | `INV_GoodWe_MT-Series_ModbusRTU.csv`, `INV_GoodWe_SMT_Series.csv` | 68 | Conserver (modèles réels distincts) |
| `INV_Growatt_125kw.csv` | `INV_Webdyn_INDIA_growatt_125kw.csv` | 156 | Conserver (modèles réels distincts) |
| `INV_Huawei_MODBUS_9095100105.csv` | `INV_Huawei_MODBUS_7550606570-M0.csv` | 73 | Conserver (modèles réels distincts) |
| `INV_Huawei_ModbusV3_small.csv` | `INV_Webdyn_HUAWEI_ModbusV3_small.csv` | 97 | Conserver (modèles réels distincts) |
| `INV_Huawei_SmartLogger3000.csv` | `INV_Webdyn_HUAWEI_SmartLogger3000.csv` | 11 | Conserver (modèles réels distincts) |
| `INV_Huawei_SUN2000_V4.csv` | `OTH_Huawei_HUAWEI_V4.csv` | 322 | Conserver (modèles réels distincts) |
| `INV_Hyxi_8-50K_workinprogress.csv` | `INV_Hyxi_HYXIPOWER_8-50K_workinprogress.csv` | 183 | Conserver (modèles réels distincts) |
| `INV_INVT_Generic.csv` | `INV_INVT_ModbusRTU_INVT_Generic.csv` | 41 | Conserver (modèles réels distincts) |
| `INV_Jema_Inverter.csv` | `INV_Webdyn_Jema_Inverter.csv` | 57 | Conserver (modèles réels distincts) |
| `INV_Jema_Modbus_revG.csv` | `INV_JEMA_Generic_revG.csv` | 60 | Conserver (modèles réels distincts) |
| `INV_Kaco_NX1_NX3_ModbusRTU.csv` | `INV_Kaco_Kaco_NX1-NX3_ModbusRTU.csv` | 182 | Conserver (modèles réels distincts) |
| `INV_Kaco_Sunspec_11.csv` | `INV_Webdyn_INDIA_kaco_Sunspec_11.csv` | 107 | Conserver (modèles réels distincts) |
| `INV_Kehua_Modbus_ThreePhase.csv` | `INV_Webdyn_KEHUA_Modbus_ThreePhase.csv` | 33 | Conserver (modèles réels distincts) |
| `INV_Kostal_SunSpec_Float.csv` | `INV_Webdyn_KOSTAL_SunSpec_Float.csv` | 204 | Conserver (modèles réels distincts) |
| `INV_Kstar_Modbus.csv` | `OTH_K_star.csv` | 36 | Conserver (modèles réels distincts) |
| `INV_Legrand_Trimod.csv` | `INV_Webdyn_Legrand_Trimod.csv` | 42 | Conserver (modèles réels distincts) |
| `INV_MarsEnergy_LC_1Unit.csv` | `INV_MarsRenewable_LC_1Unit.csv` | 18 | Conserver (modèles réels distincts) |
| `INV_MarsEnergy_LiquidCooling-AC.csv` | `INV_MarsRenewable_LiquidCooling-AC.csv` | 18 | Conserver (modèles réels distincts) |
| `INV_MarsEnergy_Meter.csv` | `INV_MarsRenewable_Meter.csv` | 24 | Conserver (modèles réels distincts) |
| `INV_MarsEnergy_PV.csv` | `INV_MarsRenewable_PV.csv` | 52 | Conserver (modèles réels distincts) |
| `INV_MarsEnergy_Rack-Cell.csv` | `INV_MarsRenewable_Rack-Cell.csv` | 42 | Conserver (modèles réels distincts) |
| `INV_MarsEnergy_Rack_V7.csv` | `INV_MarsRenewable_Rack_V7.csv` | 110 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part10.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part11.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part11.csv` | 200 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part12.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part13.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part13.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part13.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part14.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part14.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part14.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part15.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part15.csv` | 200 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part15.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part16.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part16.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part16.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part17.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part17.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part17.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part18.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part18.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part18.csv` | `INV_PowerOne_ABB_TRIO-50.0-TM-OUTD_SunSpec_RTU_PICS_part19.csv`, `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part19.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_PowerOne_PVS-100_120_PVS1xx_TCP_PICS_part20.csv` | `INV_PowerOne_SUNSPEC_PVS1xx_PICS_V002_GL20181011_part21.csv` | 69 | Conserver (modèles réels distincts) |
| `INV_Refusol_WPM017F47_SunSpec_inverter_REFU_840R040_0004_modbusTCP.csv` | `INV_Goodwe_WPM0146BE_SunSpec_inverter_GOODWE_GW80K-MT_V1.0.1_modbusRTU.csv` | 138 | Conserver (modèles réels distincts) |
| `INV_SAJ_Generic.csv` | `INV_Webdyn_INDIA_SAJ.csv` | 45 | Conserver (modèles réels distincts) |
| `INV_Schneider_Modbus_RL.csv` | `INV_Webdyn_Schneider_Modbus_RL.csv` | 11 | Conserver (modèles réels distincts) |
| `INV_Schneider_SunSpec.csv` | `INV_Webdyn_Schneider_SunSpec.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_SIEL_10_TL_SOLAR_INVERTER.csv` | `OTH_Siel_10TL.csv` | 17 | Conserver (modèles réels distincts) |
| `INV_SMA_Core1.csv` | `INV_SMA_STPCORE1_50-41.csv` | 105 | Conserver (modèles réels distincts) |
| `INV_SMA_Core2.csv` | `INV_SMA_STP_110-60.csv` | 1211 | Conserver (modèles réels distincts) |
| `INV_SMA_INV_SMA_STPX.csv` | `INV_Webdyn_INV_SMA_STPX.csv` | 509 | Conserver (modèles réels distincts) |
| `INV_SMA_Modbus.csv` | `INV_SMA_Sunspec.csv`, `INV_SMA_SunSpec_Inverter.csv` | 29 | Conserver (modèles réels distincts) |
| `INV_SMA_STP_5-10-SE40.csv` | `INV_SMA_SunnyTripower_SmartEnergy.csv` | 372 | Conserver (modèles réels distincts) |
| `INV_SMA_STP_X.csv` | `INV_SMA_SunnyTripower_X.csv` | 435 | Conserver (modèles réels distincts) |
| `INV_SMA_SunnyCentral.csv` | `INV_SMA_SUNNY_CENTRAL.csv` | 182 | Conserver (modèles réels distincts) |
| `INV_SNEF_Modbus.csv` | `INV_Webdyn_SNEF_Modbus.csv` | 16 | Conserver (modèles réels distincts) |
| `INV_Socomec_3_SOCOMEC_SUNSYS_PRO_-_3Phase_PV_Inverter_Modbus_Definition_SZE_18_07_1..._part5.csv` | `INV_Socomec_3.SOCOMEC_SUNSYS_PRO_-_3Phase_PV_Inverter_Modbus_Definition_2011_Mar_03_v00.04_FAUX_part4.csv` | 88 | Conserver (modèles réels distincts) |
| `INV_SofarSolar_SOFAR_Modbus_Protocol_HYD_5...20K-3PH_and_HYD3-6K-EP.csv` | `INV_SofarSolar_SOFAR_Modbus_Protocol_HYD_5...20K-3PH_and_HYD3-6K-EP_part2.csv` | 764 | Conserver (modèles réels distincts) |
| `INV_SolarEdge_Sunspec.csv` | `INV_Webdyn_SolarEdge_Sunspec.csv` | 43 | Conserver (modèles réels distincts) |
| `INV_Solax_Work_MODBUS_Solax.csv` | `INV_Solax_Work_MODBUS_Solax_part2.csv` | 139 | Conserver (modèles réels distincts) |
| `INV_Solis_Modbus.csv` | `INV_Webdyn_Solis_Modbus.csv` | 57 | Conserver (modèles réels distincts) |
| `INV_Solis_MODBUS_solis_50_25_sunPM.csv` | `INV_Ginlong_Solis_MODBUS_solis_50_25_sunPM.csv` | 47 | Conserver (modèles réels distincts) |
| `INV_Solis_sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15.csv` | `INV_Ginlong_Solis_sunspec-pics-solis-25-66k-us-001-for-white-labeling_1_2_updated_7-15.csv` | 308 | Conserver (modèles réels distincts) |
| `INV_Sungrow_PVCombinerBox.csv` | `INV_Sungrow_SUNGROW_PVCombinerBox.csv` | 64 | Conserver (modèles réels distincts) |
| `INV_Sungrow_Updated.csv` | `INV_Webdyn_INDIA_MODBUS_SUNGROW_Updated.csv` | 76 | Conserver (modèles réels distincts) |
| `INV_Tata_Modbus.csv` | `INV_Webdyn_TATA_Modbus.csv` | 20 | Conserver (modèles réels distincts) |
| `INV_Webdyn_Huawei_MODBUS_9095100105.csv` | `INV_Webdyn_Huawei_MODBUS_7550606570-M0.csv` | 73 | Conserver (modèles réels distincts) |
| `INV_Webdyn_Ingeteam_Modbus_3play_ABI1000-1004.csv` | `INV_Webdyn_Ingeteam_Modbus_3play_ABS1004_-_TODO_CopieOf_ABI_1000.csv` | 59 | Conserver (modèles réels distincts) |
| `INV_Webdyn_Kstar_Modbus.csv` | `OTH_Webdyn_INDIA_Modbus_K_star.csv` | 36 | Conserver (modèles réels distincts) |
| `INV_Webdyn_WPM01A743_SMA_110-60.csv` | `INV_Webdyn_WPM01A779_SMA_110-60.csv` | 1211 | Conserver (modèles réels distincts) |
| `MET_ABB_B21_B23_B24.csv` | `MET_PowerOne_ABB_Meter_B21_B23_B24.csv` | 137 | Conserver (modèles réels distincts) |
| `MET_ABB_Modbus_mapping_Meter_ABB_B21.csv` | `MET_PowerOne_Modbus_mapping_Meter_ABB_B21.csv` | 41 | Conserver (modèles réels distincts) |
| `MET_AEberle_PQPlus_Generic.csv` | `MET_A-eberle_METER_PQPLUS_GENERIC.csv` | 60 | Conserver (modèles réels distincts) |
| `MET_CarloGavazzi_EM-ET300.csv` | `MET_Webdyn_MODBUS_CarloGavazi_EM-ET300.csv` | 70 | Conserver (modèles réels distincts) |
| `MET_ChauvinArnoux_Ulys_TT-M.csv` | `MET_CHAUVIN-ARNOUX_Ulys_TT-M.csv` | 19 | Conserver (modèles réels distincts) |
| `MET_Chint_DTSU666-H.csv` | `MET_Huawei_DTSU666-H.csv` | 22 | Conserver (modèles réels distincts) |
| `MET_EcoAdapt_MODBUS_PowerElec6.csv` | `MET_Eco_Adapt_MODBUS_PowerElec6.csv` | 221 | Conserver (modèles réels distincts) |
| `MET_EcsOptec_ECS_Generic.csv` | `MET_Ecs-Optec_METER_ECS_GENERIC_FLOAT.csv`, `MET_EcsOptec_ECS_GENERIC_FLOAT.csv` | 44 | Conserver (modèles réels distincts) |
| `MET_Emu_EMUELECTRONIC_EMUPROFESSIONAL.csv` | `MET_Emu_meter_METER_EMUELECTRONIC_EMUPROFESSIONAL.csv` | 177 | Conserver (modèles réels distincts) |
| `MET_Emu_PROFESIONNAL-II.csv` | `MET_Emu_meter_METER_EMU_PROFESIONNAL-II.csv` | 388 | Conserver (modèles réels distincts) |
| `MET_FoxESS_H3_Pro.csv` | `MET_Fox_ESS_FOXESS_P3_H3_PRO_METER.csv`, `MET_FoxESS_P3_H3_PRO_METER.csv` | 61 | Conserver (modèles réels distincts) |
| `MET_Fronius_Register_Map_Float_v1.0.csv` | `MET_Fronius_Meter_Register_Map_Float_v1.0.csv` | 145 | Conserver (modèles réels distincts) |
| `MET_Fronius_Register_Map_Int_SF_v1.0.csv` | `MET_Fronius_Meter_Register_Map_Int_SF_v1.0.csv` | 165 | Conserver (modèles réels distincts) |
| `MET_IME_Generic.csv` | `MET_IME_Meter_IME_GENERIC.csv` | 85 | Conserver (modèles réels distincts) |
| `MET_Janitza_UMG-604.csv` | `MET_Janitza_WPM_meter_Janitza_UMG-RTU.csv`, `MET_Webdyn_Janitza_UMG-604.csv` | 4 | Conserver (modèles réels distincts) |
| `MET_Janitza_UMG604-PRO.csv` | `MET_Janitza_METER_JANITZA_UMG604-PRO.csv` | 71 | Conserver (modèles réels distincts) |
| `MET_Janitza_UMG604_RTU.csv` | `MET_Webdyn_Janitza_UMG604_RTU.csv` | 61 | Conserver (modèles réels distincts) |
| `MET_Legrand_60177710.csv` | `MET_Legrand_METER_LEGRAND_60177710.csv` | 81 | Conserver (modèles réels distincts) |
| `MET_Legrand_EMS_414922-24-25-27.csv` | `MET_Legrand_METER_LEGRAND_EMS_414922-24-25-27.csv` | 103 | Conserver (modèles réels distincts) |
| `MET_Legrand_LEGRAND_Meter_EMDX.csv` | `MET_Legrand_LEGRAND_Meter_LEG412052.csv` | 488 | Conserver (modèles réels distincts) |
| `MET_Legrand_LG412041.csv` | `MET_Legrand_METER_LEGRAND_LG412041.csv` | 85 | Conserver (modèles réels distincts) |
| `MET_LEGRAND_Meter_EMDX.csv` | `MET_LEGRAND_Meter_LEG412052.csv` | 489 | Conserver (modèles réels distincts) |
| `MET_map_modbus_meter_slenergy6.csv` | `MET_Webdyn_map_modbus_meter_slenergy6.csv` | 7 | Conserver (modèles réels distincts) |
| `MET_MarsEnergy_LC.csv` | `MET_MarsRenewable_LC.csv` | 1800 | Conserver (modèles réels distincts) |
| `MET_MCI_RI-D140-C.csv` | `MET_Mci_Elec_RI-D140-C.csv` | 87 | Conserver (modèles réels distincts) |
| `MET_PhenixContact_EEM-MA600.csv` | `MET_PhoenixContact_EEM-MA600.csv` | 87 | Conserver (modèles réels distincts) |
| `MET_Polier_PROMOD380.csv` | `MET_SolarLog_POLIER_PROMOD380.csv` | 88 | Conserver (modèles réels distincts) |
| `MET_Schneider_EM6433.csv` | `MET_Schneider_EM6433_evo.csv`, `MET_Webdyn_schneider_EM6433.csv`, `MET_Webdyn_schneider_EM6433_evo.csv` | 17 | Conserver (modèles réels distincts) |
| `MET_Schneider_iem3000.csv` | `MET_Schneider_iem3100.csv`, `MET_Schneider_iem3110.csv`, `MET_Schneider_iem3115.csv`, `MET_Schneider_iem3135.csv`, `MET_Schneider_iem3150.csv` *(+20 autres)* | 55 | Conserver (modèles réels distincts) |
| `MET_Schneider_iem3xxx.csv` | `MET_Webdyn_SCH_iem3xxx.csv` | 58 | Conserver (modèles réels distincts) |
| `MET_Schneider_new.csv` | `MET_Webdyn_MODBUS_EM6400NG_new.csv` | 17 | Conserver (modèles réels distincts) |
| `MET_Schneider_RELAYP5.csv` | `MET_Schneider_METER_SCHNEIDER_RELAYP5.csv` | 57 | Conserver (modèles réels distincts) |
| `MET_Schneider_RELAYP5_noswap.csv` | `MET_Webdyn_METER_SCHNEIDER_RELAYP5_noswap.csv` | 57 | Conserver (modèles réels distincts) |
| `MET_SLEnergy_ConsoMeter.csv` | `MET_Webdyn_SLEnergy_ConsoMeter.csv` | 4 | Conserver (modèles réels distincts) |
| `MET_SLEnergy_OffGridMeter.csv` | `MET_Webdyn_SLEnergy_OffGridMeter.csv` | 33 | Conserver (modèles réels distincts) |
| `MET_SMA_EnergyMeter.csv` | `MET_SMA_HomeManager.csv` | 26 | Conserver (modèles réels distincts) |
| `MET_Socomec_A10.csv` | `MET_Webdyn_Diris_A10.csv` | 25 | Conserver (modèles réels distincts) |
| `MET_Socomec_B10.csv` | `MET_Webdyn_Diris_B10.csv` | 15 | Conserver (modèles réels distincts) |
| `MET_Socomec_B30.csv` | `MET_Webdyn_Diris_B30.csv` | 45 | Conserver (modèles réels distincts) |
| `MET_Socomec_COUNTIS_E27.csv` | `MET_Socomec_METER_SOCOMEC_COUNTIS_E27.csv` | 162 | Conserver (modèles réels distincts) |
| `MET_Socomec_COUNTIS_E50-E53.csv` | `MET_Socomec_METER_SOCOMEC_COUNTIS_E50-E53.csv` | 141 | Conserver (modèles réels distincts) |
| `MET_Socomec_E23_2448.csv` | `MET_Webdyn_Countis_E23_2448.csv` | 73 | Conserver (modèles réels distincts) |
| `MET_Socomec_E33_2408.csv` | `MET_Webdyn_Countis_E33_2408.csv` | 73 | Conserver (modèles réels distincts) |
| `MET_Socomec_ECI2.csv` | `MET_Webdyn_Countis_ECI2.csv` | 7 | Conserver (modèles réels distincts) |
| `MET_Socomec_I30.csv` | `MET_Webdyn_Diris_I30.csv` | 14 | Conserver (modèles réels distincts) |
| `MET_Socomec_I33.csv` | `MET_Webdyn_Diris_I33.csv` | 12 | Conserver (modèles réels distincts) |
| `MET_Socomec_U20.csv` | `MET_Webdyn_Diris_U20.csv` | 7 | Conserver (modèles réels distincts) |
| `MET_Socorel_LINKY.csv` | `MET_Socorel_METER_SOCOREL_LINKY.csv` | 149 | Conserver (modèles réels distincts) |
| `MET_SolarLog_SOLARLOG.csv` | `MET_SolarLog_METER_SOLARLOG.csv` | 17 | Conserver (modèles réels distincts) |
| `MET_Ticmaster_ICE_SAPHIR.csv` | `MET_Ticmaster_METER_TICMASTER_ICE_SAPHIR.csv` | 20 | Conserver (modèles réels distincts) |
| `MET_Ticmaster_ICE_SAPHIR_VOIE2.csv` | `MET_Ticmaster_METER_TICMASTER_ICE_SAPHIR_VOIE2.csv` | 7 | Conserver (modèles réels distincts) |
| `MET_Ticmaster_PME-PMI.csv` | `MET_Ticmaster_METER_TICMASTER_PME-PMI.csv` | 6 | Conserver (modèles réels distincts) |
| `MET_Ticmaster_PME-PMI_VOIE2.csv` | `MET_Ticmaster_METER_TICMASTER_PME-PMI_VOIE2.csv` | 6 | Conserver (modèles réels distincts) |
| `MET_Ticmaster_PME-PMI_VOIE3.csv` | `MET_Ticmaster_METER_TICMASTER_PME-PMI_VOIE3.csv` | 6 | Conserver (modèles réels distincts) |
| `MET_Webdyn_EM21_72RVV23XOSX.csv` | `MET_CarloGavazzi_72RVV23XOSX.csv` | 7 | Conserver (modèles réels distincts) |
| `MET_Webdyn_MarsRenewable_Meter.csv` | `MET_MarsEnergy_MarsRenewable_Meter.csv`, `MET_MarsRenewable_MarsRenewable_Meter.csv` | 24 | Conserver (modèles réels distincts) |
| `OTH_5_Generic.csv` | `OTH_Webdyn_INDIA_Sunspec_5.csv` | 49 | Conserver (modèles réels distincts) |
| `OTH_ABB_Modbus_PVI.csv` | `OTH_ABB_10.0_et_12.5_ABB_PVIMODBUS.csv` | 19 | Conserver (modèles réels distincts) |
| `OTH_ABB_SunSpec_PVS.csv` | `OTH_ABB_Modbus_ABB_SunSpec_full.csv` | 981 | Conserver (modèles réels distincts) |
| `OTH_Algodue_UPM209.csv` | `OTH_Algodue_ALGODUE_UPM209.csv` | 5 | Conserver (modèles réels distincts) |
| `OTH_AlroSolar_Alro_PCS.csv` | `OTH_Alro_Solar_Alro_PCS.csv`, `OTH_Alro_Solar_PCS_Definition_part2.csv` | 443 | Conserver (modèles réels distincts) |
| `OTH_AlroSolar_BMS_Definition.csv` | `OTH_Alro_Solar_BMS_Definition.csv` | 114 | Conserver (modèles réels distincts) |
| `OTH_AlroSolar_EMS_Definition.csv` | `OTH_Alro_Solar_EMS_Definition.csv` | 167 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference.csv` | `OTH_Kostal_KOSTAL_SunSpecInformationModelReference.csv` | 227 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part12.csv` | `OTH_Kaco_sunspec_information_model_reference_part12.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part12.csv` | 94 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part56.csv` | `OTH_Kaco_sunspec_information_model_reference_part49.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part56.csv` | 74 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part57.csv` | `OTH_Axun_SunSpecInformationModelReference_part58.csv`, `OTH_Axun_SunSpecInformationModelReference_part59.csv`, `OTH_Kaco_sunspec_information_model_reference_part50.csv`, `OTH_Kaco_sunspec_information_model_reference_part51.csv`, `OTH_Kaco_sunspec_information_model_reference_part52.csv` *(+3 autres)* | 74 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part6.csv` | `OTH_Kaco_sunspec_information_model_reference_part6.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part6.csv` | 60 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part60.csv` | `OTH_Axun_SunSpecInformationModelReference_part61.csv`, `OTH_Axun_SunSpecInformationModelReference_part62.csv`, `OTH_Axun_SunSpecInformationModelReference_part63.csv`, `OTH_Kaco_sunspec_information_model_reference_part53.csv`, `OTH_Kaco_sunspec_information_model_reference_part54.csv` *(+6 autres)* | 64 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part7.csv` | `OTH_Kaco_sunspec_information_model_reference_part7.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part7.csv` | 62 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part8.csv` | `OTH_Kaco_sunspec_information_model_reference_part8.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part8.csv` | 90 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part80.csv` | `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part80.csv` | 58 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part81.csv` | `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part81.csv` | 54 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part82.csv` | `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part82.csv` | 52 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part85.csv` | `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part85.csv` | 52 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part9.csv` | `OTH_Kaco_sunspec_information_model_reference_part9.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part9.csv` | 92 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference_part95.csv` | `OTH_Kaco_sunspec_information_model_reference_part82.csv`, `OTH_Kostal_KOSTAL_SunSpecInformationModelReference_part95.csv` | 66 | Conserver (modèles réels distincts) |
| `OTH_CarloGavazzi_EM-ET300.csv` | `OTH_Webdyn_MODBUS_CarloGavazi_EM-ET300.csv` | 69 | Conserver (modèles réels distincts) |
| `OTH_CarloGavazzi_EM24.csv` | `OTH_CarloGavazzi_CarloGavazzi_EM24.csv` | 28 | Conserver (modèles réels distincts) |
| `OTH_CarloGavazzi_EM300-ET300.csv` | `OTH_CarloGavazzi_CarloGavazzi_EM300-ET300.csv` | 146 | Conserver (modèles réels distincts) |
| `OTH_Delta_Delta-Sunspec_protocol_v1.11_20180725.csv` | `OTH_Webdyn_Delta-Sunspec_protocol_v1.11_20180725.csv` | 92 | Conserver (modèles réels distincts) |
| `OTH_Delta_Delta-Sunspec_protocol_v1.11_20180725_part29.csv` | `OTH_Webdyn_Delta-Sunspec_protocol_v1.11_20180725_part29.csv` | 52 | Conserver (modèles réels distincts) |
| `OTH_Delta_Delta-Sunspec_protocol_v1.12_20190403.csv` | `OTH_Webdyn_Delta-Sunspec_protocol_v1.12_20190403.csv` | 93 | Conserver (modèles réels distincts) |
| `OTH_Delta_Delta-Sunspec_protocol_v1_03_20170907.csv` | `OTH_Delta_SunSpec-Delta_ed00_adding_FW_versions.csv`, `OTH_Webdyn_Delta-Sunspec_protocol_v1_03_20170907.csv` | 71 | Conserver (modèles réels distincts) |
| `OTH_Delta_M15A-flex_COM115.csv` | `OTH_Delta_M30A-flex_COM117.csv` | 969 | Conserver (modèles réels distincts) |
| `OTH_Delta_M15A-flex_COM115_part20.csv` | `OTH_Delta_M30A-flex_COM117_part20.csv` | 333 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part13.csv` | `OTH_Delta_M100A_280_COM1.15_part13.csv`, `OTH_Delta_M15A-flex_COM115_part13.csv`, `OTH_Delta_M30A-flex_COM117_part13.csv`, `OTH_Delta_M50A-flex_COM114_part13.csv`, `OTH_Delta_M50A_COM139_part13.csv` *(+2 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part14.csv` | `OTH_Delta_M100A_280_COM1.15_part14.csv`, `OTH_Delta_M15A-flex_COM115_part14.csv`, `OTH_Delta_M30A-flex_COM117_part14.csv`, `OTH_Delta_M50A-flex_COM114_part14.csv`, `OTH_Delta_M50A_COM139_part14.csv` *(+2 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part15.csv` | `OTH_Delta_M100A_280_COM1.15_part15.csv`, `OTH_Delta_M15A-flex_COM115_part15.csv`, `OTH_Delta_M30A-flex_COM117_part15.csv`, `OTH_Delta_M50A-flex_COM114_part15.csv`, `OTH_Delta_M50A_COM139_part15.csv` *(+2 autres)* | 59 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part16.csv` | `OTH_Delta_M100A_280_COM1.15_part16.csv`, `OTH_Delta_M15A-flex_COM115_part16.csv`, `OTH_Delta_M30A-flex_COM117_part16.csv`, `OTH_Delta_M50A-flex_COM114_part16.csv`, `OTH_Delta_M50A_COM139_part16.csv` *(+2 autres)* | 59 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part17.csv` | `OTH_Delta_M100A_280_COM1.15_part17.csv`, `OTH_Delta_M15A-flex_COM115_part17.csv`, `OTH_Delta_M30A-flex_COM117_part17.csv`, `OTH_Delta_M50A-flex_COM114_part17.csv`, `OTH_Delta_M50A_COM139_part17.csv` *(+2 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part18.csv` | `OTH_Delta_M100A_280_COM1.15_part18.csv`, `OTH_Delta_M15A-flex_COM115_part18.csv`, `OTH_Delta_M30A-flex_COM117_part18.csv`, `OTH_Delta_M50A-flex_COM114_part18.csv`, `OTH_Delta_M50A_COM139_part18.csv` *(+2 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part19.csv` | `OTH_Delta_M50A_COM139_part19.csv` | 55 | Conserver (modèles réels distincts) |
| `OTH_Delta_M30A_COM147_part20.csv` | `OTH_Delta_M50A_COM139_part20.csv` | 55 | Conserver (modèles réels distincts) |
| `OTH_Delta_M70A_COM01.20.csv` | `OTH_Delta_M50A-flex_COM114.csv`, `OTH_Delta_M70AFLEX_COM01.28.csv` | 999 | Conserver (modèles réels distincts) |
| `OTH_Delta_M70A_COM01.20_part19.csv` | `OTH_Delta_M50A-flex_COM114_part19.csv`, `OTH_Delta_M70AFLEX_COM01.28_part19.csv` | 69 | Conserver (modèles réels distincts) |
| `OTH_Delta_M70A_COM01.20_part20.csv` | `OTH_Delta_M50A-flex_COM114_part20.csv`, `OTH_Delta_M70AFLEX_COM01.28_part20.csv` | 333 | Conserver (modèles réels distincts) |
| `OTH_Delta_protocol_v1.11_20180725.csv` | `OTH_Delta_Sunspec_protocol_v1.11_20180725.csv` | 92 | Conserver (modèles réels distincts) |
| `OTH_Delta_protocol_v1.12_20190403.csv` | `OTH_Delta_Sunspec_protocol_v1.12_20190403.csv` | 93 | Conserver (modèles réels distincts) |
| `OTH_Delta_protocol_v1_03_20170907.csv` | `OTH_Delta_Sunspec_protocol_v1_03_20170907.csv` | 71 | Conserver (modèles réels distincts) |
| `OTH_Deye_Sunspec-PICS-DEYE-SUN-0820.csv` | `OTH_Deye_Inverter_Sunspec-PICS-DEYE-SUN-0820.csv` | 367 | Conserver (modèles réels distincts) |
| `OTH_EcoAdapt_MODBUS_PowerElec6.csv` | `OTH_Eco_Adapt_MODBUS_PowerElec6.csv` | 221 | Conserver (modèles réels distincts) |
| `OTH_EDMI_Mk6.csv` | `OTH_EDMI_EDMI_Mk6.csv` | 59 | Conserver (modèles réels distincts) |
| `OTH_Efacec_Exemplo_de_Mapeamento_PV500_Reflexos.csv` | `OTH_Efasolar_Exemplo_de_Mapeamento_PV500_Reflexos.csv` | 27 | Conserver (modèles réels distincts) |
| `OTH_Efacec_ntu-pv500.mbmap-0.2.3.csv` | `OTH_Efasolar_ntu-pv500.mbmap-0.2.3.csv` | 156 | Conserver (modèles réels distincts) |
| `OTH_Efacec_PV550_Efacec_MODBUS.csv` | `OTH_Efasolar_PV550_Efacec_MODBUS.csv`, `OTH_Efasolar_PV550_Efacec_MODBUS_part1.csv` | 148 | Conserver (modèles réels distincts) |
| `OTH_EKO_MS40.csv` | `OTH_EKO_EKO_MS40.csv` | 28 | Conserver (modèles réels distincts) |
| `OTH_Fronius_WPM018C3D_Fronius_Symo_10.0-3-M.csv` | `OTH_Fronius_WPM01A481_Fronius_Symo_20.0-3-M.csv` | 177 | Conserver (modèles réels distincts) |
| `OTH_Huawei_Huawei_Work_part6.csv` | `OTH_Huawei_Huawei_Work_part7.csv` | 89 | Conserver (modèles réels distincts) |
| `OTH_Huawei_MODBUS_9095100105.csv` | `OTH_Huawei_MODBUS_7550606570-M0.csv` | 73 | Conserver (modèles réels distincts) |
| `OTH_IME_GENERIC.csv` | `OTH_IME_Meter_IME_GENERIC.csv` | 82 | Conserver (modèles réels distincts) |
| `OTH_IME_IM-MFD4ORFCDT1.csv` | `OTH_IME_Meter_IME_IM-MFD4ORFCDT1.csv` | 70 | Conserver (modèles réels distincts) |
| `OTH_Imesys_NEMO_96HDLe.csv` | `OTH_IMESYS_IMESYS_NEMO_96HDLe.csv` | 797 | Conserver (modèles réels distincts) |
| `OTH_IngenieurBuro_Tx-RS485-MB.csv` | `OTH_IngenieurBuro_INGENIEURBURO_Tx-RS485-MB.csv` | 6 | Conserver (modèles réels distincts) |
| `OTH_Ingeteam_INGETEAM_PowerMax_centraux.csv` | `OTH_Webdyn_Ingeteam_Modbus_PowerMax_centraux.csv` | 31 | Conserver (modèles réels distincts) |
| `OTH_Ingeteam_PowerMax_centraux.csv` | `OTH_Ingeteam_Modbus_PowerMax_centraux.csv` | 31 | Conserver (modèles réels distincts) |
| `OTH_INVT_ModbusRTU_INVT_Generic.csv` | `OTH_INVT_Solar_ModbusRTU_INVT_Generic.csv` | 41 | Conserver (modèles réels distincts) |
| `OTH_Jema_IFX3_1050_RevD.csv` | `OTH_Jema_JEMA_IFX3_1050_RevD.csv` | 33 | Conserver (modèles réels distincts) |
| `OTH_Jema_Modbus_revG.csv` | `OTH_Jema_Generic_revG.csv`, `OTH_Jema_JEMA_Generic_revG.csv`, `OTH_Webdyn_JEMA_Modbus_revG.csv` | 60 | Conserver (modèles réels distincts) |
| `OTH_Kehua_Modbus_ThreePhase.csv` | `OTH_Webdyn_KEHUA_Modbus_ThreePhase.csv` | 33 | Conserver (modèles réels distincts) |
| `OTH_Kehua_SinglePhase_Work.csv` | `OTH_Kehua_KEHUA_SinglePhase_Work.csv` | 149 | Conserver (modèles réels distincts) |
| `OTH_Kehua_ThreePhase.csv` | `OTH_Kehua_KEHUA_ThreePhase.csv` | 108 | Conserver (modèles réels distincts) |
| `OTH_Kostal_KOSTAL_SunSpec_Float.csv` | `OTH_Kostal_KOSTAL_SunSpec_Float_part1.csv` | 249 | Conserver (modèles réels distincts) |
| `OTH_Kostal_PIKO_CI_events_code-V1.1-20240516.csv` | `OTH_Kostal_PIKO_CI_events_code-V1.2-20240516_via_modbus.csv` | 3 | Conserver (modèles réels distincts) |
| `OTH_Lacecal_ITR2_0.csv` | `OTH_Lacecal_LACECAL_ITR2_0.csv` | 67 | Conserver (modèles réels distincts) |
| `OTH_Legrand_Nemo_Work.csv` | `OTH_Legrand_LEGRAND_Nemo_Work.csv` | 169 | Conserver (modèles réels distincts) |
| `OTH_LSI_SRL_LSILASTEM_DPA870-873.csv` | `OTH_LSI_LASTEM_SRL_LSILASTEM_DPA870-873.csv` | 30 | Conserver (modèles réels distincts) |
| `OTH_Lufft_WS600.csv` | `OTH_Lufft_LUFFT_WS600.csv` | 42 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_LC.csv` | `OTH_MarsEnergy_MarsRenewable_LC.csv`, `OTH_MarsRenewable_MarsRenewable_LC.csv`, `OTH_Webdyn_MarsRenewable_LC.csv` | 8 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv` | `OTH_MarsRenewable_MARS_LocalController_V7_Unit247_Un_seul_element_WebdynSunPM.csv` | 5 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_MarsRenewable_LC_1Unit.csv` | `OTH_MarsRenewable_MarsRenewable_LC_1Unit.csv` | 18 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_MarsRenewable_LC_EMS_V7.csv` | `OTH_MarsRenewable_MarsRenewable_LC_EMS_V7.csv` | 8 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_MarsRenewable_LiquidCooling-AC.csv` | `OTH_MarsRenewable_MarsRenewable_LiquidCooling-AC.csv` | 18 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_MarsRenewable_LocalControler_PCS_Information.csv` | `OTH_MarsRenewable_MarsRenewable_LocalControler_PCS_Information.csv` | 111 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_PCS-Bank.csv` | `OTH_MarsEnergy_MarsRenewable_PCS-Bank.csv` | 84 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_PV.csv` | `OTH_MarsEnergy_MarsRenewable_PV.csv`, `OTH_MarsRenewable_MarsRenewable_PV.csv`, `OTH_Webdyn_MarsRenewable_PV.csv` | 52 | Conserver (modèles réels distincts) |
| `OTH_MCI_contax_1m_100a_031277_mapping_Modbus_20102017.csv` | `OTH_Mci_Elec_contax_1m_100a_031277_mapping_Modbus_20102017.csv` | 32 | Conserver (modèles réels distincts) |
| `OTH_MCI_contax_d_tri_80a_mid_modbus_ref_031333_-_mapping_modbus_V3.csv` | `OTH_Mci_Elec_contax_d_tri_80a_mid_modbus_ref_031333_-_mapping_modbus_V3.csv` | 36 | Conserver (modèles réels distincts) |
| `OTH_MCI_RI-D140-C.csv` | `OTH_Mci_Elec_RI-D140-C.csv` | 87 | Conserver (modèles réels distincts) |
| `OTH_MeierNT_MEIER_ADL-SR.csv` | `OTH_Meier-nt_MEIER_ADL-SR.csv` | 18 | Conserver (modèles réels distincts) |
| `OTH_Panasonic_AKW8115.csv` | `OTH_Panasonic_PANASONIC_AKW8115.csv` | 14 | Conserver (modèles réels distincts) |
| `OTH_PhenixContact_EEM-MA600.csv` | `OTH_PhoenixContact_EEM-MA600.csv` | 87 | Conserver (modèles réels distincts) |
| `OTH_S_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv` | `OTH_SRegeltechnik_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv` | 8 | Conserver (modèles réels distincts) |
| `OTH_SAJ_C6.csv` | `OTH_SAJ_SAJ_C6.csv` | 138 | Conserver (modèles réels distincts) |
| `OTH_SAJ_Generic.csv` | `OTH_Webdyn_INDIA_SAJ.csv` | 45 | Conserver (modèles réels distincts) |
| `OTH_SAJ_PLUS_R5.csv` | `OTH_SAJ_SAJ_PLUS_R5.csv` | 73 | Conserver (modèles réels distincts) |
| `OTH_SAJ_R5_R6_C6.csv` | `OTH_SAJ_SAJ_R5_R6_C6.csv` | 391 | Conserver (modèles réels distincts) |
| `OTH_Salgueirinha_Modbus_Registers.csv` | `OTH_Salgueirinha_Salgueirinha_Modbus_Registers.csv` | 1294 | Conserver (modèles réels distincts) |
| `OTH_Satcon_Generic.csv` | `OTH_Satcon_SATCON_Generic.csv` | 95 | Conserver (modèles réels distincts) |
| `OTH_Schneider_3255lite.csv` | `OTH_Schneider_Schneider_3255lite.csv`, `OTH_Schneider_Schneider_3255lite_part1.csv` | 105 | Conserver (modèles réels distincts) |
| `OTH_Schneider_EM6433.csv` | `OTH_Schneider_EM6433_evo.csv`, `OTH_Webdyn_schneider_EM6433.csv`, `OTH_Webdyn_schneider_EM6433_evo.csv` | 17 | Conserver (modèles réels distincts) |
| `OTH_Schneider_Generic.csv` | `OTH_Webdyn_MODBUS_EM1200.csv` | 7 | Conserver (modèles réels distincts) |
| `OTH_Schneider_iem3000.csv` | `OTH_Schneider_iem3100.csv`, `OTH_Schneider_iem3110.csv`, `OTH_Schneider_iem3115.csv`, `OTH_Schneider_iem3135.csv`, `OTH_Schneider_iem3150.csv` *(+7 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Schneider_iem3xxx.csv` | `OTH_Schneider_SCH_iem3xxx.csv`, `OTH_Webdyn_SCH_iem3xxx.csv`, `OTH_Webdyn_WPM015004_SCH_iem3xxx.csv` | 54 | Conserver (modèles réels distincts) |
| `OTH_Schneider_Modbus_RL.csv` | `OTH_Webdyn_Schneider_Modbus_RL.csv` | 11 | Conserver (modèles réels distincts) |
| `OTH_Schneider_new.csv` | `OTH_Webdyn_MODBUS_EM6400NG_new.csv` | 17 | Conserver (modèles réels distincts) |
| `OTH_Schneider_PM556x_PublicModbusRegisterList_10th_August-23.csv` | `OTH_Schneider_PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0.csv`, `OTH_Schneider_PM5563_PM5563RD_ModbusRegisterList_v2.3.0-ORIGINAL.csv`, `OTH_Webdyn_PM5500_PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0.csv`, `OTH_Webdyn_PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0-ORIGINAL.csv` | 28 | Conserver (modèles réels distincts) |
| `OTH_Schneider_SEPAM_S48.csv` | `OTH_Schneider_SCHNEIDER_SEPAM_S48.csv` | 31 | Conserver (modèles réels distincts) |
| `OTH_Siel_DSP-DSPX.csv` | `OTH_Siel_SIEL_DSP-DSPX.csv` | 21 | Conserver (modèles réels distincts) |
| `OTH_Siel_MONOPHASE.csv` | `OTH_Siel_SIEL_MONOPHASE.csv` | 11 | Conserver (modèles réels distincts) |
| `OTH_Sineng_central_New.csv` | `OTH_Webdyn_INDIA_SINENG_central_New.csv` | 43 | Conserver (modèles réels distincts) |
| `OTH_SNEF_Modbus.csv` | `OTH_Webdyn_SNEF_Modbus.csv` | 16 | Conserver (modèles réels distincts) |
| `OTH_SolarEdge_66KW.csv` | `OTH_Webdyn_INDIA_SolarEdge_66KW.csv` | 38 | Conserver (modèles réels distincts) |
| `OTH_SolarEdge_EnhancedPowerControlBlock.csv` | `OTH_SolarEdge_SolarEdge_EnhancedPowerControlBlock.csv` | 18 | Conserver (modèles réels distincts) |
| `OTH_SolarEdge_FromIndia.csv` | `OTH_SolarEdge_SolarEdge_FromIndia.csv` | 38 | Conserver (modèles réels distincts) |
| `OTH_SolarLog_PRO380_WORK.csv` | `OTH_SolarLog_SOLARLOG_PRO380_WORK.csv` | 147 | Conserver (modèles réels distincts) |
| `OTH_Sungrow_Updated.csv` | `OTH_Webdyn_INDIA_MODBUS_SUNGROW_Updated.csv` | 76 | Conserver (modèles réels distincts) |
| `OTH_Tata_Modbus.csv` | `OTH_Tata_Generic.csv`, `OTH_Tata_TATA_Generic.csv`, `OTH_Webdyn_TATA_Modbus.csv` | 20 | Conserver (modèles réels distincts) |
| `OTH_Thermokon_AGS55.csv` | `OTH_Webdyn_MODBUS_THERMOKON_AGS55.csv` | 66 | Conserver (modèles réels distincts) |
| `OTH_Thermokon_AGS55_tcp.csv` | `OTH_Webdyn_MODBUS_THERMOKON_AGS55_tcp.csv` | 66 | Conserver (modèles réels distincts) |
| `OTH_Thies_Wind_Transmitter_Thies.csv` | `OTH_Thies_Clima_Wind_Transmitter_Thies.csv` | 91 | Conserver (modèles réels distincts) |
| `OTH_Ticmaster_Mapping_-_FW41.48.csv` | `OTH_Ticmaster_TicMaster-Mapping_-_FW41.48.csv` | 76 | Conserver (modèles réels distincts) |
| `OTH_Ticmaster_Mapping_-_FW41.60.csv` | `OTH_Ticmaster_TicMaster-Mapping_-_FW41.60.csv` | 75 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_ABB_Modbus_PVI.csv` | `OTH_PowerOne_10.0_et_12.5_ABB_PVIMODBUS.csv` | 19 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_ABB_SunSpec_PVS.csv` | `OTH_PowerOne_Modbus_ABB_SunSpec_full.csv` | 981 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_Delta_Sunspec_V1.20.csv` | `OTH_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv` | 640 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_EM21_72RVV23XOSX.csv` | `OTH_CarloGavazzi_72RVV23XOSX.csv` | 7 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_Fronius_Sunspec_IntSF.csv` | `OTH_Fronius_Modbus_SunS_Fronius_IntSF.csv` | 55 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_Huawei_MODBUS_9095100105.csv` | `OTH_Webdyn_Huawei_MODBUS_7550606570-M0.csv` | 73 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_Hukseflux.csv` | `OTH_Webdyn_WPM015004_Hukseflux.csv` | 27 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_Ingeteam_Modbus_3play_ABI1000-1004.csv` | `OTH_Webdyn_Ingeteam_Modbus_3play_ABS1004_-_TODO_CopieOf_ABI_1000.csv` | 59 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_MarsRenewable_PCS-Bank.csv` | `OTH_MarsRenewable_MarsRenewable_PCS-Bank.csv` | 84 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_REPORT.csv` | `OTH_Webdyn_REPPORT.csv` | 6 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_SCH_iem3000.csv` | `OTH_Webdyn_SCH_iem3100.csv`, `OTH_Webdyn_SCH_iem3110.csv`, `OTH_Webdyn_SCH_iem3115.csv`, `OTH_Webdyn_SCH_iem3135.csv`, `OTH_Webdyn_SCH_iem3150.csv` *(+7 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Webdyn_Sens_IO-8AI_work.csv` | `OTH_Webdyn_WEBDYN_Sens_IO-8AI_work.csv`, `OTH_Webdyn_WEBDYN_Sens_IO-8AI_work_part1.csv` | 34 | Conserver (modèles réels distincts) |
| `OTH_Weidmuller_610.csv` | `OTH_Weidmuller_Weidmuller_610.csv` | 122 | Conserver (modèles réels distincts) |
| `OTH_Weidmuller_Transclinic.csv` | `OTH_Weidmuller_Weidmuller_Transclinic.csv` | 33 | Conserver (modèles réels distincts) |
| `SEN_DeltaOhm_LP_PYRA.csv` | `SEN_DeltaOhm_SENSOR_LP_PYRA.csv` | 5 | Conserver (modèles réels distincts) |
| `SEN_Hukseflux_SR05_D1A3.csv` | `SEN_Hukseflux_Hukseflux_SR05-D1A3.csv` | 28 | Conserver (modèles réels distincts) |
| `SEN_Hukseflux_SR20-D2.csv` | `SEN_Hukseflux_Hukseflux_SR20-D2.csv` | 12 | Conserver (modèles réels distincts) |
| `SEN_INGENIEURBURO_Si-RS485-TC-2T-v.csv` | `SEN_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv` | 7 | Conserver (modèles réels distincts) |
| `SEN_LSI_SRL_LSILASTEM_DPA870-873.csv` | `SEN_LSI_LASTEM_SRL_LSILASTEM_DPA870-873.csv` | 30 | Conserver (modèles réels distincts) |
| `SEN_MeierNT_MEIER_ADL-SR.csv` | `SEN_Meier-nt_MEIER_ADL-SR.csv` | 18 | Conserver (modèles réels distincts) |
| `SEN_S_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv` | `SEN_SRegeltechnik_S_REGELTECHNIK_S_SREGELTECHNIK_HYDRASGARD.csv` | 8 | Conserver (modèles réels distincts) |
| `SEN_SESAutomation_WT-485-PRO.csv` | `SEN_SES_automation_SEN_SESAUTOMATION_WT-485-PRO.csv` | 10 | Conserver (modèles réels distincts) |
| `SEN_Seven_3S-IS_8.0.csv` | `SEN_Seven_sensor_solutions_Seven_Sensor_3S-IS_8.0.csv` | 42 | Conserver (modèles réels distincts) |
| `SEN_Seven_3S-IS_8.csv` | `SEN_Seven_sensor_solutions_Seven_Sensor_3S-IS_8.csv` | 41 | Conserver (modèles réels distincts) |
| `SEN_Seven_Sensor_3S-IS.csv` | `SEN_Seven_3S-IS_SUNSPEC.csv`, `SEN_Webdyn_SENSOR_SEVEN_SENSOR_3S-IS_SUNSPEC.csv` | 45 | Conserver (modèles réels distincts) |
| `SEN_Seven_SEVENSENSORSOLUTION_3SIS.csv` | `SEN_Seven_sensor_solutions_SEVENSENSORSOLUTION_3SIS.csv` | 8 | Conserver (modèles réels distincts) |
| `SEN_Thermokon_AGS55_tcp.csv` | `SEN_Webdyn_MODBUS_THERMOKON_AGS55_tcp.csv` | 66 | Conserver (modèles réels distincts) |
| `SEN_Thies_Wind_Transmitter_Thies.csv` | `SEN_Thies_Clima_Wind_Transmitter_Thies.csv` | 91 | Conserver (modèles réels distincts) |
| `SEN_Webdyn_Hukseflux.csv` | `OTH_Hukseflux_Generic.csv`, `SEN_Hukseflux_Generic.csv` | 27 | Conserver (modèles réels distincts) |