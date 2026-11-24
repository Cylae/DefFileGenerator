# 🛠️ Boîte à Outils d'Ingénierie & Scripts d'Automatisation

Ce dossier regroupe l'ensemble des scripts opérationnels, utilitaires de validation, générateurs d'équipements et outils de gouvernance du projet **DefFileGenerator**.

---

## 🧭 Catégorisation des Outils

```
scripts/
├── 🛡️ CONTRÔLE QUALITÉ & VALIDATION
│   ├── validate_production.py          # Validation stricte des 5 Piliers Apollo sur production/
│   └── verify_web_dom.py               # Test E2E de l'interface web (DOM & API)
│
├── 🏭 RECONSTRUCTION & COMPILATION MAÎTRESSE
│   ├── build_production_suite.py       # Synchronisation et compilation intégrale du catalogue (1 194 fichiers)
│   ├── import_and_organize_sources.py  # Ingestion des manuels sources et calcul des hachages SHA-256
│   └── generate_definitions_catalog.py # Production de CATALOG.json et de l'index markdown
│
├── ⚙️ GÉNÉRATEURS PAR CONSTRUCTEUR
│   ├── generate_classic_definitions.py    # Suite multi-constructeurs (ABB, CG, SMA, Huawei, etc.)
│   ├── generate_sma_definitions.py        # Spécifique SMA (STP 125-70 Core2, Core1 STP 50-41)
│   ├── generate_goodwe_definitions.py     # Gammes GoodWe (GT, SMT, HT, Hybrides ET)
│   ├── generate_x7_definitions.py         # Solution conteneurisée X7 All-In-One BESS
│   ├── generate_foxess_definitions.py     # Onduleurs et batteries FoxESS (P3, H3 Pro, G-MAX)
│   ├── generate_dyness_definitions.py     # Systèmes de stockage Dyness (DH100F, DH200F, DH200Y)
│   ├── generate_mars_definitions.py       # Mars Renewable Local Controller V7.0
│   ├── generate_cyberpower_definitions.py # Onduleurs réseau CyberPower CPSPV
│   └── generate_solplanet_ailogger.py     # Passerelle Solplanet Ai-Logger
│
├── 📜 JOURNALISATION & TRAÇABILITÉ
│   └── audit_logger.py                 # Enregistrement dans les registres d'audit immuables
│
└── 🚀 AMORÇAGE & ENVIRONNEMENT
    ├── bootstrap.ps1                   # Script PowerShell Windows (installation .venv, dépendances, tests)
    └── bootstrap.sh                    # Script Bash POSIX (Linux, macOS)
```

---

## 📋 Référence Détaillée des Commandes

### 1. Contrôle Qualité & Certification 5 Piliers
```powershell
# Auditer l'intégralité du catalogue de production (1 194 fichiers)
python scripts/validate_production.py

# Reconstruire le catalogue et synchroniser les miroirs
python scripts/build_production_suite.py

# Vérifier le bon rendu de l'interface web (API + DOM)
python scripts/verify_web_dom.py
```

### 2. Gestion des Sources Techniques & Manifeste
```powershell
# Ingestion des documents sources depuis le partage réseau (lecture seule)
# et mise à jour du grand livre cryptographique sources/MANIFEST.json
python scripts/import_and_organize_sources.py

# Régénérer l'index machine production/CATALOG.json
python scripts/generate_definitions_catalog.py
```

### 3. Exécution d'un Générateur Spécifique
```powershell
python scripts/generate_x7_definitions.py
python scripts/generate_foxess_definitions.py
python scripts/generate_dyness_definitions.py
python scripts/generate_mars_definitions.py
python scripts/generate_cyberpower_definitions.py
python scripts/generate_solplanet_ailogger.py
```
