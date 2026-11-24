# 📥 Sas d'Extraction Brute (Raw Parser Staging Area)

Ce répertoire constitue le **sas d'entrée et d'extraction brute** du pipeline d'ingénierie Webdyn.

---

## 🎯 Rôle dans l'Architecture

Lorsqu'un moteur d'extraction ou un script de génération (`extractor.py`, `scripts/generate_*.py`, `DefFileGenerator/*_definitions_engine.py`) analyse une documentation technique constructeur (PDF, Excel, XML, HTML), les fichiers CSV résultants sont d'abord déposés ici.

```
Documentation Source (sources/)
             │
             ▼
   Moteurs d'Extraction
             │
             ▼
┌─────────────────────────────────┐
│ generated_definitions/ (Sas)    │ ◄── VOUS ÊTES ICI
└─────────────────────────────────┘
             │
             ▼
  Enrichissement SunPM
             │
             ▼
┌─────────────────────────────────┐
│ tagged_definitions/ (Qualif.)   │
└─────────────────────────────────┘
             │
             ▼
  Certification 5 Piliers
             │
             ▼
┌─────────────────────────────────┐
│ production/ (Golden Master)     │ ◄── CATALOGUE OFFICIEL (1 194 CSVs)
└─────────────────────────────────┘
```

---

## 🔍 Caractéristiques des Fichiers dans ce Répertoire

- **Nature :** Fichiers CSV intermédiaires d'extraction directe.
- **Rôle :** Permettre l'inspection visuelle et le débogage des tables de registres avant déduplication et enrichissement sémantique.
- **Cycle de Vie :** Les fichiers stables et validés sont promus vers [`production/`](../production/) via le script `scripts/build_production_suite.py`.
- **Référentiel Officiel :** Pour l'intégration dans des passerelles Webdyn réelles en exploitation, utilisez **exclusivement** les fichiers certifiés du dossier [`production/`](../production/).
