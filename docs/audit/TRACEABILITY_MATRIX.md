# 🗺️ Matrice de Traçabilité des Équipements & Définitions

La matrice de traçabilité garantit l'alignement strict entre chaque **spécification constructeur d'origine**, les **modèles d'équipements couverts**, les **fichiers de définition Webdyn générés**, et leur **suite de tests automatisée**.

---

## 🧭 Matrice Complète Spécifications ➔ Définitions ➔ Tests

| Réf. Source | Constructeur | Spécification Officielle | Modèles Physiques Couverts | Fichier Définition Webdyn | Registres | Tags SunPM | Suite de Tests & Validation | Statut |
|---|---|---|---|---|:---:|:---:|---|:---:|
| `SRC-CP-001` | **CyberPower** | `CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf` (V1.04) | CPSPV3000ETLA..6000ETLA, CPSPV8000ETLA..66000ETLA-WB, CPSPV120KETL, CPSPV140KETL | [`INV_Cyberpower_GENERIC.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/INV_Cyberpower_GENERIC.csv) | 244 | 186 | `test_cyberpower_definitions.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |
| `SRC-CP-001` | **CyberPower** | `CyberPower Grid_tied PV Inverter Modbus Protocol(V0104) 1.pdf` (V1.04) | CPSPV120KETL, CPSPV140KETL (Utility Scale 120-140 kW) | [`INV_Cyberpower_CPSPV120K-140K.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/INV_Cyberpower_CPSPV120K-140K.csv) | 244 | 186 | `test_cyberpower_definitions.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |
| `SRC-SP-001` | **Solplanet** | `Ailogger modbus TCP user manual V1.2 202404.pdf` (V1.2) | Solplanet Ai-Logger, Ai-Logger 1000 (Gateway Mode) | [`COM_Solplanet_Ai-Logger.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/COM_Solplanet_Ai-Logger.csv) | 161 | 128 | `test_solplanet_ailogger.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |
| `SRC-SP-001` | **Solplanet** | `Ailogger modbus TCP user manual V1.2 202404.pdf` (V1.2) | ASW Inverters behind Ai-Logger (Method 2 / RTU ID) | [`INV_Solplanet_Ai-Logger.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/INV_Solplanet_Ai-Logger.csv) | 106 | 85 | `test_solplanet_ailogger.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |
| `SRC-FX-001` | **FoxESS** | `FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf` | FoxESS H3-Pro, P3-Pro, All-in-One | [`INV_FoxESS_H3_Pro.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/INV_FoxESS_H3_Pro.csv)<br>[`INV_FoxESS_P3_Pro.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/INV_FoxESS_P3_Pro.csv) | ~250 ch. | ~180 ch. | `test_foxess_p3_h3_pro.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |
| `SRC-FX-001` | **FoxESS** | `FoxESS Modbus Protocol--20260701 (V1.05.05.00).pdf` | FoxESS G-MAX Storage, R-Series | [`BAT_FoxESS_GMax_Storage.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/BAT_FoxESS_GMax_Storage.csv) | ~180 | ~90 | `test_foxess_p3_h3_pro.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |
| `SRC-MR-001` | **MarsEnergy** | `MarsLocalController_IO List_ModbusTCP_V7.0.xlsx` (V7.0) | Mars Local Controller, PCS-Bank, BMS | [`OTH_MarsEnergy_LC.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/OTH_MarsEnergy_LC.csv)<br>[`BAT_MarsEnergy_Rack-Cell.csv`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/production/BAT_MarsEnergy_Rack-Cell.csv) | ~300 ch. | ~120 ch. | `test_mars_requested_definitions.py`<br>`test_production_suite.py` | ✅ CERTIFIÉ |

---

## 🔄 Politique de Synchronisation et Redondance Quadruple

Pour garantir la résilience et la disponibilité du référentiel à travers le temps, chaque fichier de définition validé fait l'objet d'une **réplication stricte et univoque** :

```
                                  [Source PDF / XLSX]
                                           │
                                           ▼
                               [scripts/generate_*.py]
                                           │
              ┌────────────────────────────┼────────────────────────────┐
              ▼                            ▼                            ▼
      [Definitions/]             [production/]           [tagged_definitions/]
    (Base Référentielle)      (Gouvernance 5 Piliers)     (Balisage Télémétrie)
              │                            │                            │
              └────────────────────────────┼────────────────────────────┘
                                           ▼
                               [generated_definitions/]
                                (Générations Dérivées)
```

1. **Zéro Divergence Tolérée :** Les fichiers portant le même nom dans `production/` et `Definitions/` doivent être strictement identiques octet par octet.
2. **Encodage Uniformisé :** Encodage obligatoire en **UTF-8 sans BOM** (`encoding="utf-8"`, rejet absolu de `\xef\xbb\xbf`).
3. **Contrôle Automatisé Continu :** La suite de tests `DefFileGenerator/tests/test_definitions_redundancy.py` bloque tout commit en cas de divergence ou de perte de redondance.
