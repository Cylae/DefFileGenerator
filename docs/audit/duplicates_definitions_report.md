# Rapport d'Audit des Doublons — Suite de Production Webdyn

- **Répertoire analysé** : `Definitions`
- **Total fichiers scannés** : **177**
- **Fichiers doublons identifiés** : **3**
- **Doublons stricts (SHA-256)** : **1**
- **Conflits de catégorie (OTH_)** : **0**
- **Alias & synonymes** : **2**
- **Familles matérielles** : **13**

---

## 1. Doublons Stricts (Empreinte SHA-256 Identique)

Ces fichiers sont rigoureusement identiques au niveau binaire (même empreinte SHA-256, même en-tête et mêmes registres).

| Fichier Canonique | Doublons Stricts | Modèle Matériel | Registres | Recommandation |
| :--- | :--- | :--- | :---: | :--- |
| `INV_ABB_PVS800-57B.csv` | `INV_ABB_PVS800-57B_2.csv` | PVS800-57B | 1053 | Supprimer le doublon strict ou créer un alias |

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
| `INV_Goodwe_GT-SERIES.csv` | `INV_Goodwe_GT_Series.csv` | Hardware model alias / synonymous naming sharing identical register map | 305 | Enregistrer dans le registre d'alias |
| `INV_Goodwe_HT1500.csv` | `INV_Goodwe_HT_1500V.csv` | Hardware model alias / synonymous naming sharing identical register map | 288 | Enregistrer dans le registre d'alias |

---

## 4. Familles Matérielles Partageant la Même Carte Modbus

Équipements distincts d'un même constructeur partageant rigoureusement le même profil de registres (ex. Schneider iEM3100..iEM3555, GoodWe ET/EH). Total : **13** familles répertoriées.

| Modèle de Référence | Modèles de la Famille | Registres | Recommandation |
| :--- | :--- | :---: | :--- |
| `BAT_Dyness_DH100F.csv` | `BAT_Dyness_DH200F.csv` | 913 | Conserver (modèles réels distincts) |
| `COM_GoodWe_Ezlogger3000.csv` | `COM_Goodwe_EzLogger_Pro.csv` | 25 | Conserver (modèles réels distincts) |
| `INV_Goodwe_ES_EM_SBP.csv` | `INV_Goodwe_ES-EM-SBP-BP.csv` | 49 | Conserver (modèles réels distincts) |
| `INV_Goodwe_ET_Series.csv` | `INV_Goodwe_ET-EH-BT-BH-EHB-AES-ABP-BTC_ModbusRTU.csv`, `INV_Goodwe_ET_EH_Hybrid_Series.csv`, `INV_Goodwe_ET_EH_Series.csv` | 79 | Conserver (modèles réels distincts) |
| `INV_Goodwe_MS_DNS_XS.csv` | `INV_Goodwe_XS-DNS-MS-SDT-G3_SERIES.csv`, `INV_Goodwe_XS_DNS_MS.csv` | 50 | Conserver (modèles réels distincts) |
| `INV_Goodwe_MT_series.csv` | `INV_Goodwe_MT-Series_ModbusRTU.csv`, `INV_Goodwe_SMT_Series.csv` | 68 | Conserver (modèles réels distincts) |
| `INV_SMA_Core1.csv` | `INV_SMA_STPCORE1_50-41.csv` | 105 | Conserver (modèles réels distincts) |
| `INV_SMA_Core2.csv` | `INV_SMA_STP_110-60.csv` | 1211 | Conserver (modèles réels distincts) |
| `INV_SMA_Modbus.csv` | `INV_SMA_Sunspec.csv`, `INV_SMA_SunSpec_Inverter.csv` | 29 | Conserver (modèles réels distincts) |
| `INV_SMA_STP_5-10-SE40.csv` | `INV_SMA_SunnyTripower_SmartEnergy.csv` | 372 | Conserver (modèles réels distincts) |
| `INV_SMA_STP_X.csv` | `INV_SMA_SunnyTripower_X.csv` | 435 | Conserver (modèles réels distincts) |
| `INV_SMA_SunnyCentral.csv` | `INV_SMA_SUNNY_CENTRAL.csv` | 182 | Conserver (modèles réels distincts) |
| `MET_SMA_EnergyMeter.csv` | `MET_SMA_HomeManager.csv` | 26 | Conserver (modèles réels distincts) |