# ⚡ Mars Renewable Energy — Bundle de Définitions Dédié

Ce dossier regroupe le bundle officiel de livraison des fichiers de définition WebdynSunPM pour la gamme de systèmes de stockage, conversion et pilotage **Mars Renewable Energy**.

---

## 📦 Contenu du Bundle

Généré à partir de la spécification constructeur officielle `MarsLocalController_IO List_ModbusTCP_V7.0.xlsx` (archivée sous [`sources/MarsEnergy/`](../sources/MarsEnergy/)) :

| Fichier | Composant | Protocole | Nombre de Registres |
|---|---|---|:---:|
| `OTH_MarsEnergy_LC.csv` | Local Controller (LC V7.0) | Modbus TCP (Port 502) | 120+ |
| `OTH_MarsEnergy_PCS-Bank.csv` | Banc d'Onduleurs PCS | Modbus TCP (Port 502) | 80+ |
| `BAT_MarsEnergy_Rack-Cell.csv` | Système BMS Rack & Cellules | Modbus TCP (Port 502) | 250+ |
| `SEN_MarsEnergy_PV.csv` | Instrumentation & Capteurs PV | Modbus TCP (Port 502) | 40+ |
| *(+13 profils étendus)* | Profils combinés et déclinaisons | Modbus TCP (Port 502) | - |

---

## 🛡️ Conformité & Synchronisation

Tous les fichiers de ce dossier sont :
- 100% conformes aux **5 Piliers Apollo**.
- Rigoureusement synchronisés avec [`production/`](../production/).
- Validés par la suite de tests automatisée [`DefFileGenerator/tests/test_mars_requested_definitions.py`](../DefFileGenerator/tests/test_mars_requested_definitions.py).
