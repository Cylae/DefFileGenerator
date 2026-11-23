# Rapport d'Audit des Doublons — Suite de Production Webdyn

- **Répertoire analysé** : `production`
- **Total fichiers scannés** : **1109**
- **Fichiers doublons identifiés** : **10**
- **Doublons stricts (SHA-256)** : **0**
- **Conflits de catégorie (OTH_)** : **0**
- **Alias & synonymes** : **8**
- **Familles matérielles** : **64**

---

## 1. Doublons Stricts (Empreinte SHA-256 Identique)

Ces fichiers sont rigoureusement identiques au niveau binaire (même empreinte SHA-256, même en-tête et mêmes registres).

| Fichier Canonique | Doublons Stricts | Modèle Matériel | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| Aucun | Aucun | - | - | Conforme |

---

## 2. Conflits de Catégorie (Classification Générique `OTH_` vs Fichier Typé)

Ces fichiers `OTH_` sont des doublons stricts d'équipements déjà qualifiés en tant qu'onduleurs (`INV_`), compteurs (`MET_`), capteurs (`SEN_`) ou batteries (`BAT_`).

| Fichier Canonique Qualifié | Doublon Redondant `OTH_` | Modèle Matériel | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| Aucun | Aucun | - | - | Conforme |

---

## 3. Alias Matériels et Variantes Synonymes de Nommage

Fichiers représentant le même équipement ou la même gamme sous deux désignations alternatives.

| Fichier Canonique | Alias / Synonyme | Description | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| `BAT_SLEnergy_bms5_p1_1.csv` | `OTH_SLEnergy_bms5_p1_1.csv`<br>`OTH_SLEnergy_bms5_p1_1_1.csv` | Hardware model alias / synonymous naming sharing identical register map | 213 | Enregistrer dans le registre d'alias |
| `INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv` | `INV_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 15 | Enregistrer dans le registre d'alias |
| `INV_SolaX_X1-X3-GENERIC.csv` | `INV_SolaX_X1-X3-GENERIC_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 156 | Enregistrer dans le registre d'alias |
| `INV_Sungrow_RTU_PM.csv` | `INV_Sungrow_TCP_PM.csv` | Hardware model alias / synonymous naming sharing identical register map | 81 | Enregistrer dans le registre d'alias |
| `MET_Legrand_CF-GENERIC.csv` | `MET_Legrand_CF_Generic.csv` | Hardware model alias / synonymous naming sharing identical register map | 11 | Enregistrer dans le registre d'alias |
| `OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_RTU.csv` | `OTH_PowerElectronics_Freesun_MODBUS_POWERELECTRONICS_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 15 | Enregistrer dans le registre d'alias |
| `OTH_Siebert_XC4XX_RTU.csv` | `OTH_Siebert_XC4XX_TCP.csv` | Hardware model alias / synonymous naming sharing identical register map | 3 | Enregistrer dans le registre d'alias |
| `SEN_IngenieurBuro_Si-RS485-TC-2T-v.csv` | `OTH_IngenieurBuro_MODBUS_Si-RS485-TC-2T-v.csv`<br>`OTH_IngenieurBuro_Si-RS485-TC-2T-v.csv` | Hardware model alias / synonymous naming sharing identical register map | 7 | Enregistrer dans le registre d'alias |

---

## 4. Familles Matérielles Partageant la Même Carte Modbus

Équipements distincts d'un même constructeur partageant rigoureusement le même profil de registres (ex. Schneider iEM3100..iEM3555, GoodWe ET/EH). Total : **64** familles répertoriées.

| Modèle de Référence | Modèles de la Famille | Registres | Recommandation |
| :--- | :--- | :---: | :--- |
| `BAT_Dyness_DH100F.csv` | `BAT_Dyness_DH200F.csv` | 913 | Conserver (modèles réels distincts) |
| `BAT_FoxESS_H3_Pro.csv` | `BAT_FoxESS_P3_Pro.csv` | 893 | Conserver (modèles réels distincts) |
| `BAT_FoxESS_RSeries.csv` | `BAT_FoxESS_RSERIES_BATTERY.csv` | 49 | Conserver (modèles réels distincts) |
| `BAT_Huawei_Luna2000.csv` | `BAT_Huawei_Luna_2000.csv` | 170 | Conserver (modèles réels distincts) |
| `BAT_MarsEnergy_Rack-Cell.csv` | `BAT_MarsEnergy_MarsRenewable_Rack-Cell.csv` | 42 | Conserver (modèles réels distincts) |
| `COM_GoodWe_Ezlogger3000.csv` | `COM_Goodwe_EzLogger_Pro.csv` | 25 | Conserver (modèles réels distincts) |
| `INV_ABB_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv` | `INV_ABB_LM1_SunSpec_inverter_FIMER_PVS-30-TL_2117E_modbusTCP.csv` | 956 | Conserver (modèles réels distincts) |
| `INV_ABB_Velux_SunSpec_inverter_FIMER_-3Q15-_2201B_modbusTCP.csv` | `INV_ABB_5583-CHEVALIERP2_SunSpec_inverter_ABB_3Q4200_2117B_modbusRTU.csv` | 1111 | Conserver (modèles réels distincts) |
| `INV_APsystems_Altenergy_Power_System_Inc._three_phase_microinverter_11-12.csv` | `INV_APsystems_Altenergy_Power_System_Inc._single_phase_MicroinverterAPpcs_11-12.csv` | 113 | Conserver (modèles réels distincts) |
| `INV_Delta_Sunspec_V1.20.csv` | `INV_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv` | 640 | Conserver (modèles réels distincts) |
| `INV_FoxESS_H3_Pro.csv` | `INV_FoxESS_P3_H3_PRO_INVERTER.csv`, `INV_FoxESS_P3_Pro.csv` | 285 | Conserver (modèles réels distincts) |
| `INV_FoxESS_SunSpec_RV_H3Plus.csv` | `INV_FoxESS_R_V_H3plus_GmaxPlus.csv` | 87 | Conserver (modèles réels distincts) |
| `INV_Fronius_ECO_Inverter_Register_Map_Int_SF.csv` | `INV_Fronius_Tauro_ECO_Inverter_Register_Map_Int_SF.csv` | 257 | Conserver (modèles réels distincts) |
| `INV_Fronius_Sunspec_IntSF.csv` | `INV_Fronius_Modbus_SunS_Fronius_IntSF.csv` | 55 | Conserver (modèles réels distincts) |
| `INV_Fronius_Tauro_Inverter_Register_Map_Float.csv` | `INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Float.csv` | 279 | Conserver (modèles réels distincts) |
| `INV_Fronius_Tauro_Inverter_Register_Map_Int_SF.csv` | `INV_Fronius_Verto_Inverter_Register_Map_3MPPT_Int_SF.csv` | 287 | Conserver (modèles réels distincts) |
| `INV_Goodwe_ES_EM_SBP.csv` | `INV_Goodwe_ES-EM-SBP-BP.csv` | 49 | Conserver (modèles réels distincts) |
| `INV_Goodwe_ET_Series.csv` | `INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv`, `INV_Goodwe_ET_EH_Hybrid_Series.csv`, `INV_Goodwe_ET_EH_Series.csv` | 79 | Conserver (modèles réels distincts) |
| `INV_Goodwe_MS_DNS_XS.csv` | `INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv`, `INV_Goodwe_XS_DNS_MS.csv` | 50 | Conserver (modèles réels distincts) |
| `INV_Goodwe_MT_series.csv` | `INV_Goodwe_MT-Series_ModbusRTU.csv`, `INV_Goodwe_SMT_Series.csv` | 68 | Conserver (modèles réels distincts) |
| `INV_Huawei_MODBUS_9095100105.csv` | `INV_Huawei_MODBUS_7550606570-M0.csv` | 73 | Conserver (modèles réels distincts) |
| `INV_Huawei_SUN2000_V4.csv` | `OTH_Huawei_V4.csv` | 322 | Conserver (modèles réels distincts) |
| `INV_Hyxi_8-50K_workinprogress.csv` | `INV_Hyxi_HYXIPOWER_8-50K_workinprogress.csv` | 183 | Conserver (modèles réels distincts) |
| `INV_INVT_Generic.csv` | `INV_INVT_ModbusRTU_INVT_Generic.csv` | 41 | Conserver (modèles réels distincts) |
| `INV_Jema_Modbus_revG.csv` | `INV_Jema_Generic_revG.csv` | 60 | Conserver (modèles réels distincts) |
| `INV_Kstar_Modbus.csv` | `OTH_K_star.csv` | 36 | Conserver (modèles réels distincts) |
| `INV_SMA_Core1.csv` | `INV_SMA_STPCORE1_50-41.csv` | 105 | Conserver (modèles réels distincts) |
| `INV_SMA_Core2.csv` | `INV_SMA_STP_110-60.csv` | 1211 | Conserver (modèles réels distincts) |
| `INV_SMA_Modbus.csv` | `INV_SMA_Sunspec.csv`, `INV_SMA_SunSpec_Inverter.csv` | 29 | Conserver (modèles réels distincts) |
| `INV_SMA_STP_5-10-SE40.csv` | `INV_SMA_SunnyTripower_SmartEnergy.csv` | 372 | Conserver (modèles réels distincts) |
| `INV_SMA_STP_X.csv` | `INV_SMA_SunnyTripower_X.csv` | 435 | Conserver (modèles réels distincts) |
| `INV_SMA_SunnyCentral.csv` | `INV_SMA_SUNNY_CENTRAL.csv` | 182 | Conserver (modèles réels distincts) |
| `MET_Chint_DTSU666-H.csv` | `MET_Huawei_DTSU666-H.csv` | 22 | Conserver (modèles réels distincts) |
| `MET_EcsOptec_ECS_Generic.csv` | `MET_EcsOptec_ECS_GENERIC_FLOAT.csv` | 44 | Conserver (modèles réels distincts) |
| `MET_FoxESS_H3_Pro.csv` | `MET_FoxESS_P3_H3_PRO_METER.csv` | 61 | Conserver (modèles réels distincts) |
| `MET_IME_CE4DF3DTMID.csv` | `MET_Legrand_IME_CE4DF3DTMID.csv` | 66 | Conserver (modèles réels distincts) |
| `MET_Janitza_Generic.csv` | `MET_Janitza_W_Janitza_Generic.csv` | 61 | Conserver (modèles réels distincts) |
| `MET_Polier_PROMOD380.csv` | `MET_SolarLog_POLIER_PROMOD380.csv` | 88 | Conserver (modèles réels distincts) |
| `MET_Schneider_EM6433.csv` | `MET_Schneider_EM6433_evo.csv` | 17 | Conserver (modèles réels distincts) |
| `MET_Schneider_iem3000.csv` | `MET_Schneider_iem3100.csv`, `MET_Schneider_iem3110.csv`, `MET_Schneider_iem3115.csv`, `MET_Schneider_iem3135.csv`, `MET_Schneider_iem3150.csv` *(+7 autres)* | 55 | Conserver (modèles réels distincts) |
| `MET_Schneider_iem3xxx.csv` | `MET_Schneider_SCH_iem3xxx.csv` | 58 | Conserver (modèles réels distincts) |
| `MET_SMA_EnergyMeter.csv` | `MET_SMA_HomeManager.csv` | 26 | Conserver (modèles réels distincts) |
| `OTH_ABB_Modbus_PVI.csv` | `OTH_ABB_10.0_et_12.5_ABB_PVIMODBUS.csv` | 19 | Conserver (modèles réels distincts) |
| `OTH_ABB_SunSpec_PVS.csv` | `OTH_ABB_Modbus_ABB_SunSpec_full.csv` | 981 | Conserver (modèles réels distincts) |
| `OTH_Axun_SunSpecInformationModelReference.csv` | `OTH_Kostal_SunSpecInformationModelReference.csv` | 227 | Conserver (modèles réels distincts) |
| `OTH_Delta_M15A-flex_COM115.csv` | `OTH_Delta_M30A-flex_COM117.csv` | 969 | Conserver (modèles réels distincts) |
| `OTH_Delta_M70A_COM01.20.csv` | `OTH_Delta_M50A-flex_COM114.csv`, `OTH_Delta_M70AFLEX_COM01.28.csv` | 999 | Conserver (modèles réels distincts) |
| `OTH_Delta_protocol_v1.11_20180725.csv` | `OTH_Delta_Sunspec_protocol_v1.11_20180725.csv` | 92 | Conserver (modèles réels distincts) |
| `OTH_Delta_protocol_v1.12_20190403.csv` | `OTH_Delta_Sunspec_protocol_v1.12_20190403.csv` | 93 | Conserver (modèles réels distincts) |
| `OTH_Delta_protocol_v1_03_20170907.csv` | `OTH_Delta_SunSpec-Delta_ed00_adding_FW_versions.csv`, `OTH_Delta_Sunspec_protocol_v1_03_20170907.csv` | 71 | Conserver (modèles réels distincts) |
| `OTH_Delta_Sunspec_V1.20.csv` | `OTH_Delta_MODBUS_DELTA_M88H-COM1.20-Sunspec.csv` | 640 | Conserver (modèles réels distincts) |
| `OTH_Fronius_Sunspec_IntSF.csv` | `OTH_Fronius_Modbus_SunS_Fronius_IntSF.csv` | 55 | Conserver (modèles réels distincts) |
| `OTH_Huawei_MODBUS_9095100105.csv` | `OTH_Huawei_MODBUS_7550606570-M0.csv` | 73 | Conserver (modèles réels distincts) |
| `OTH_Ingeteam_PowerMax_centraux.csv` | `OTH_Ingeteam_Modbus_PowerMax_centraux.csv` | 31 | Conserver (modèles réels distincts) |
| `OTH_Jema_Modbus_revG.csv` | `OTH_Jema_Generic_revG.csv` | 60 | Conserver (modèles réels distincts) |
| `OTH_Kostal_PIKO_CI_events_code-V1.1-20240516.csv` | `OTH_Kostal_PIKO_CI_events_code-V1.2-20240516_via_modbus.csv` | 3 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_LC.csv` | `OTH_MarsEnergy_MarsRenewable_LC.csv` | 8 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_PCS-Bank.csv` | `OTH_MarsEnergy_MarsRenewable_PCS-Bank.csv` | 84 | Conserver (modèles réels distincts) |
| `OTH_MarsEnergy_PV.csv` | `OTH_MarsEnergy_MarsRenewable_PV.csv` | 52 | Conserver (modèles réels distincts) |
| `OTH_Schneider_EM6433.csv` | `OTH_Schneider_EM6433_evo.csv` | 17 | Conserver (modèles réels distincts) |
| `OTH_Schneider_iem3000.csv` | `OTH_Schneider_iem3100.csv`, `OTH_Schneider_iem3110.csv`, `OTH_Schneider_iem3115.csv`, `OTH_Schneider_iem3135.csv`, `OTH_Schneider_iem3150.csv` *(+7 autres)* | 55 | Conserver (modèles réels distincts) |
| `OTH_Schneider_iem3xxx.csv` | `OTH_Schneider_SCH_iem3xxx.csv` | 54 | Conserver (modèles réels distincts) |
| `OTH_Schneider_PM556x_PublicModbusRegisterList_10th_August-23.csv` | `OTH_Schneider_PM5560_PM5563_PM5563RD_ModbusRegisterList_v2.3.0.csv`, `OTH_Schneider_PM5563_PM5563RD_ModbusRegisterList_v2.3.0-ORIGINAL.csv` | 28 | Conserver (modèles réels distincts) |
| `OTH_Tata_Modbus.csv` | `OTH_Tata_Generic.csv` | 20 | Conserver (modèles réels distincts) |