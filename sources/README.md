# 📚 Bibliothèque des Documents Sources & Spécifications Équipements

Ce répertoire centralise les spécifications techniques officielles, protocoles Modbus et manuels fabricants servant de **vérité terrain** pour l'ingénierie et la génération des fichiers de définition Webdyn.

---

## 🗂️ Architecture des Dossiers

```
sources/
├── MANIFEST.json           # Index machine-readable avec hachage SHA-256 et métadonnées
├── README.md               # Présente documentation et guide de gouvernance
└── manuals/                # Manuels et spécifications classés par constructeur
    ├── CyberPower/         # Protocoles CyberPower Systems
    │   └── CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf
    ├── FoxESS/             # Spécifications de protocoles FoxESS
    │   └── FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf
    ├── MarsEnergy/         # Spécifications et IO List Mars Energy
    │   └── MarsLocalController_IO List_ModbusTCP_V7.0.xlsx
    └── Solplanet/          # Protocoles et manuels AISWEI / Solplanet
        └── Ailogger modbus TCP user manual V1.2 202404.pdf
```

---

## 📋 Inventaire des Documents et Définitions Associées

| ID Document | Constructeur | Titre du Document | Version | Modèles Couverts | Définitions Webdyn Produites | Empreinte SHA-256 |
|---|---|---|:---:|---|---|---|
| `SRC-CP-001` | **CyberPower** | Grid-Tied PV Inverter Modbus Protocol | V1.04 | CPSPV3000..6000ETLA (1-Ph), CPSPV8000..66000ETLA (3-Ph), CPSPV120KETL, CPSPV140KETL | `INV_Cyberpower_GENERIC.csv`<br>`INV_Cyberpower_1-Phase.csv`<br>`INV_Cyberpower_3-Phases.csv`<br>`INV_Cyberpower_CPSPV120K-140K.csv` | `085694...` |
| `SRC-SP-001` | **Solplanet** | Ai-Logger Modbus TCP User Manual | V1.2 | Ai-Logger, Ai-Logger 1000 | `COM_Solplanet_Ai-Logger.csv` | `d72093...` |
| `SRC-FX-001` | **FoxESS** | Modbus Protocol Specification | V1.05.05.00 | H3-Pro, P3-Pro, R Series, G-MAX Storage | `INV_FoxESS_H3_Pro.csv`<br>`INV_FoxESS_P3_Pro.csv`<br>`BAT_FoxESS_GMax_Storage.csv` | `5c84d7...` |
| `SRC-MR-001` | **MarsEnergy** | Local Controller IO List Modbus TCP | V7.0 | LC, PCS-Bank, Rack-Cell, PV | `OTH_MarsEnergy_LC.csv`<br>`BAT_MarsEnergy_Rack-Cell.csv` | `cfae4f...` |

---

## 🛡️ Règles de Gouvernance des Sources
1. **Intégrité Cryptographique :** Tout document entrant doit être référencé dans `MANIFEST.json` avec son empreinte SHA-256 calculée.
2. **Immutabilité :** Les fichiers sous `sources/manuals/` sont conservés dans leur état originel sans altération ni renommage destructif.
3. **Traçabilité :** Tout nouveau fichier ou mise à jour est immédiatement journalisé dans `docs/audit/EQUIPMENT_AUDIT_LOG.md` et `docs/audit/TRACEABILITY_MATRIX.md`.
