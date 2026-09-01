# Démarrage rapide

## Installation

```bash
python -m venv .venv
# PowerShell
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e .
```

Pour contribuer ou exécuter les contrôles qualité :

```bash
pip install -e ".[dev]"
```

## Conversion directe

Transformez un document de registres en définition WebdynSunPM :

```bash
python doc_to_webdyn.py register_map.xlsx `
  --manufacturer "Webdyn" `
  --model "DeviceModel" `
  -o definition.csv
```

## Workflow contrôlable

Utilisez la CLI principale pour séparer extraction, génération et validation :

```bash
python DefFileGenerator/main.py extract register_map.xlsx -o registers.csv
python DefFileGenerator/main.py generate registers.csv --manufacturer "Webdyn" --model "DeviceModel" -o definition.csv
python DefFileGenerator/main.py validate definition.csv
```

Exécutez `python DefFileGenerator/main.py --help` et `python DefFileGenerator/main.py <commande> --help` pour les paramètres disponibles.

## Vérifier avant livraison

```bash
python -m pytest
ruff check .
```

Les données de stress et leurs sorties ne font pas partie de ce cycle rapide : elles doivent être générées localement ou dans un job CI dédié.
