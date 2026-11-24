# 🗺️ Cartographie & Taxonomie de l'Architecture des Dossiers

Ce document décrit l'organisation structurelle, la vocation et les règles d'intégrité de chaque répertoire du projet **WebdynSunPM DefFileGenerator**.

Pour la documentation complète d'architecture avec diagrammes Mermaid et flux de données détaillés, consultez :
👉 **[`docs/ARCHITECTURE.md`](ARCHITECTURE.md)**

---

## 🏗️ Vue d'Ensemble des 5 Domaines Fonctionnels

```text
DefFileGenerator/
├── 🏭 1. RÉFÉRENTIEL MAÎTRE DE PRODUCTION (Single Source of Truth)
│   └── production/                      # 1 194 fichiers CSV 100% conformes aux 5 Piliers Apollo
│       ├── CATALOG.json                 # Catalogue machine complet indexant tous les équipements
│       ├── SOURCES.md                   # Matrice de traçabilité liant définitions et manuels
│       └── README.md                    # Guide de gouvernance du catalogue
│
├── 📚 2. SOURCES TECHNIQUES OFFICIELLES (Vérité Terrain)
│   └── sources/                         # 110 documents officiels (PDF, XLSX, HTML)
│       ├── MANIFEST.json                # Registre cryptographique (SHA-256) scellé
│       ├── README.md                    # Inventaire complet des documents sources
│       └── <Fabricant>/                 # CarloGavazzi, ABB, SMA, Huawei, Sungrow, GoodWe...
│           └── <Catégorie>/             # Inverter, Meter, Battery, Gateway, Logger, Sensor...
│               └── <Gamme>/             # EM530_EM540, SUN2000, STP125-70, PVS800...
│
├── ⚙️ 3. CŒUR LOGICIEL & OUTILLAGE
│   ├── DefFileGenerator/                # Package Python (extracteur, normaliseur, gouverneur)
│   │   ├── tests/                       # Suite de tests exhaustive (4 653 tests unitaires)
│   │   └── dataset_knowledge.json       # Base de connaissances sémantique SunPM
│   ├── scripts/                         # Boîte à outils CLI (validation, build, générateurs)
│   └── web/                             # Studio Web moderne & API REST FastAPI
│
├── 🔄 4. SAS DE TRANSITION & BUNDLES DÉDIÉS
│   ├── generated_definitions/           # Sas d'extraction brute des parsers (staging)
│   ├── tagged_definitions/              # Sas de qualification sémantique SunPM
│   ├── Definitions/                     # Miroir historique et rétrocompatibilité multi-sites (701 CSVs)
│   ├── definitions_history/             # Magasin CAS d'objets immuables (versioning web)
│   ├── Dyness_DH_Series_Definitions/    # Bundle dédié Dyness BESS Série DH
│   ├── FoxESS_P3_H3_Pro_Definitions/    # Bundle dédié FoxESS P3 / H3 Pro
│   ├── MarsRenewable_Definitions/       # Bundle dédié Mars Renewable Energy
│   └── X7_Definitions/                  # Bundle dédié X7 All-In-One BESS
│
└── 📖 5. CONNAISSANCES & GOUVERNANCE
    ├── docs/                            # Documentation technique et rapports d'audit
    │   ├── ARCHITECTURE.md              # Architecture globale et flux système
    │   ├── DIRECTORY_TAXONOMY.md        # Présente cartographie
    │   └── audit/                       # Journal d'audit et matrices de revue
    └── logs/                            # Registres d'audit trails append-only (audit_log.jsonl)
```

---

## 🔒 Règles de Gouvernance des Répertoires

1. **Inviolabilité de `production/`** : Seuls les fichiers validés à 100% par `validate_production.py` sont admis dans ce répertoire maître.
2. **Lecture Seule sur le Partage Réseau** : Le partage commun distant (`\\WDN-FILES.WEBDYN.COM\Support\Equipementiers`) est accédé en **lecture seule stricte**.
3. **Scellement Cryptographique dans `sources/`** : Tout document technique intégré sous `sources/` est scellé par son empreinte SHA-256 dans `sources/MANIFEST.json`.
4. **Traçabilité Immuable dans `logs/`** : Chaque opération de génération est inscrite dans `logs/audit_trail/audit_log.jsonl` et documentée dans `docs/audit/EQUIPMENT_AUDIT_LOG.md`.
