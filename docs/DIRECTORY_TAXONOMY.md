# 🗺️ Cartographie & Taxonomie de l'Architecture des Dossiers

Ce document décrit l'organisation structurelle, la vocation et les règles d'intégrité de chaque répertoire du projet **WebdynSunPM DefFileGenerator**.

---

## 🏗️ Vue d'Ensemble de l'Arborescence

```text
DefFileGenerator/
├── DefFileGenerator/             # Cœur applicatif Python (moteur d'extraction, génération, validation)
│   ├── tests/                   # Suite de tests unitaires, fuzzing et non-régression
│   ├── dataset_knowledge.json   # Base de connaissances sémantiques (36 270+ variables apprises)
│   ├── def_gen.py               # Générateur standard de fichiers de définition Webdyn
│   ├── extractor.py             # Extracteur multi-format (PDF, Excel, CSV, XML)
│   └── production_governor.py   # Moteur de certification des 5 Piliers Apollo
│
├── production/                  # 🏭 RÉFÉRENTIEL OFFICIEL DE PRODUCTION
│   ├── CATALOG.json             # Index complet des 1 136 définitions certifiées (191 constructeurs)
│   ├── README.md                # Documentation des 5 Piliers et statistiques
│   └── <PREFIX>_<Fabricant>_<Modèle>.csv (1 136 fichiers 100% validés)
│
├── sources/                     # 📚 SOURCES BRUTES CONSTRUCTEURS & TERRAIN
│   ├── MANIFEST.json            # Registre cryptographique (SHA-256) des documents sources
│   ├── manuals/                 # Manuels d'utilisation et spécifications officielles (PDF, XLSX)
│   └── telemetry/               # Fichiers de données et télémétries de terrain réelles (CSV)
│
├── scripts/                     # 🛠️ SCRIPTS D'AUTOMATISATION & GÉNÉRATEURS CLI
│   ├── validate_production.py   # Validateur des 5 Piliers sur le dossier production/
│   ├── build_production_suite.py # Moteur de reconstruction de la production
│   ├── audit_logger.py          # Enregistreur du journal d'audit immuable
│   ├── generate_solplanet_ailogger.py    # Générateur certifié Solplanet Ai-Logger
│   ├── generate_cyberpower_definitions.py # Générateur certifié CyberPower
│   ├── generate_dyness_definitions.py    # Générateur certifié Dyness BESS
│   ├── generate_foxess_definitions.py    # Générateur certifié FoxESS
│   └── generate_mars_definitions.py      # Générateur certifié Mars Renewable
│
├── web/                         # 🌐 STUDIO & APPLICATION WEB (FastAPI)
│   ├── app.py                   # Serveur API REST (Endpoints: convert, audit, library, pillars)
│   └── static/                  # Interface utilisateur WebdynSunPM Studio (HTML5, CSS3, JS)
│
├── docs/                        # 📖 DOCUMENTATION TECHNIQUE & AUDIT
│   ├── audit/                   # Journal d'audit chronologique & Matrice de traçabilité
│   │   ├── EQUIPMENT_AUDIT_LOG.md
│   │   └── TRACEABILITY_MATRIX.md
│   └── WEBDYN_TAGS_REFERENCE.md # Dictionnaire des tags SunPM
│
├── logs/                        # 📜 JOURNAUX SYSTÈME & AUDIT TRAILS
│   └── audit_trail/
│       └── audit_log.jsonl      # Registre JSONL append-only immuable
│
├── Definitions/                 # 🏛️ Archives des définitions historiques Webdyn
├── tagged_definitions/          # Miroir de synchronisation des définitions balisées
├── generated_definitions/       # Miroir de synchronisation des définitions générées
└── scratch/                     # 🧪 Espace de travail pour scripts d'expérimentation
```

---

## 🔒 Règles de Gouvernance des Répertoires

1. **Inviolabilité de `production/`** : Aucun fichier n'entre dans `production/` sans satisfaire les 5 Piliers Apollo vérifiés par `validate_production.py`.
2. **Traçabilité de `sources/`** : Tout document technique analysé doit être consigné dans `sources/manuals/` ou `sources/telemetry/` et indexé dans `sources/MANIFEST.json`.
3. **Pérennité de `logs/audit_trail/`** : Chaque opération de génération ou de mise à jour doit être inscrite dans `logs/audit_trail/audit_log.jsonl` et reflétée dans `docs/audit/EQUIPMENT_AUDIT_LOG.md`.
4. **Redondance Multi-Répertoires** : Les fichiers certifiés sont synchronisés de façon identique dans `production/`, `Definitions/`, `tagged_definitions/` et `generated_definitions/`.
