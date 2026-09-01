# Développement

## Installation

```bash
pip install -e ".[dev]"
```

## Commandes de contrôle

```bash
ruff check .
python -m pytest
python -m pytest --cov=DefFileGenerator --cov-report=term-missing
```

## Stratégie de tests

- Les tests unitaires doivent être déterministes, petits et sans dépendance réseau.
- Les tests d’intégration couvrent le chemin extraction → génération → validation.
- Chaque correction de bug ajoute une fixture minimale et un test de non-régression.
- Les tests de charge et les batteries torture ne sont pas exécutés pour chaque pull request. Ils génèrent leurs résultats hors de Git.

## Règles de contribution

Une pull request doit viser un sujet unique, inclure les tests adaptés et ne pas ajouter d’artefact généré. Préférez les itérateurs pour les grandes entrées, mais ne retournez pas un générateur dépendant d’un fichier déjà fermé. Conservez l’API CLI rétrocompatible ou documentez explicitement toute rupture.

## Performance

Mesurez avant et après une optimisation. Les métriques utiles sont le temps total, le débit de registres, la mémoire maximale et la taille de sortie. Les benchmarks doivent publier un résumé chiffré, pas des fichiers de sortie massifs.
