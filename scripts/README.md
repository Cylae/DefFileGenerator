# 🛠️ Répertoire des Scripts & Outils CLI

Ce dossier centralise l'ensemble des scripts d'automatisation, de génération d'équipements et de validation pour le projet **WebdynSunPM DefFileGenerator**.

---

## 📋 Index des Outils Disponibles

### 1. Contrôle Qualité & Validation
| Script | Description | Commande |
| :--- | :--- | :--- |
| **`validate_production.py`** | Audite rigoureusement tous les fichiers du dossier `production/` selon les 5 Piliers Apollo. | `python scripts/validate_production.py` |
| **`build_production_suite.py`** | Reconstruit et certifie l'ensemble du corpus de production (1 118 définitions). | `python scripts/build_production_suite.py` |
| **`audit_logger.py`** | Module et CLI d'enregistrement dans le journal d'audit immuable (`EQUIPMENT_AUDIT_LOG.md` et `audit_log.jsonl`). | `python scripts/audit_logger.py` |

### 2. Générateurs Dédiés par Constructeur
| Script | Fabricant | Modèles Couverts | Sortie(s) Principale(s) |
| :--- | :--- | :--- | :--- |
| **`generate_solplanet_ailogger.py`** | Solplanet | Passerelle Ai-Logger & Inverters COM1..3 | `COM_Solplanet_Ai-Logger.csv`, `INV_Solplanet_Ai-Logger.csv` |
| **`generate_cyberpower_definitions.py`** | CyberPower | Gamme réseau CPSPV3000 à CPSPV140KETL | `INV_Cyberpower_GENERIC.csv`, `INV_Cyberpower_CPSPV120K-140K.csv` |
| **`generate_dyness_definitions.py`** | Dyness | BESS Série DH (DH100F, DH200F, DH200Y, Generic) | `BAT_Dyness_DH100F.csv`, `BAT_Dyness_DH200F.csv`, etc. |
| **`generate_foxess_definitions.py`** | FoxESS | Onduleurs et stockage P3-Pro, H3-Pro, G-MAX | `INV_FoxESS_P3_Pro.csv`, `BAT_FoxESS_GMax_Storage.csv`, etc. |
| **`generate_mars_definitions.py`** | Mars Renewable | Local Controller V7.0, PCS-Bank, Rack-Cell BMS | `OTH_MarsEnergy_LC.csv`, `BAT_MarsEnergy_Rack-Cell.csv`, etc. |

### 3. Amorçage de l'Environnement
- **`bootstrap.ps1`** : Script PowerShell d'initialisation de l'environnement virtuel sous Windows.
- **`bootstrap.sh`** : Script Bash équivalent pour environnements POSIX.
