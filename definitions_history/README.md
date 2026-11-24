# 🏛️ Magasin d'Objets Immuables CAS (Content-Addressable Storage)

Ce répertoire constitue l'infrastructure interne de **contrôle de version et d'historique immuable** pour les fichiers de définition Webdyn.

---

## ⚙️ Fonctionnement Technique

Ce magasin fonctionne selon le modèle d'un moteur de type **Git CAS (Content-Addressable Storage)** :
- Chaque version d'un fichier de définition est stockée sous forme de **blob immuable** nommé d'après son empreinte cryptographique **SHA-256**.
- Les objets sont partitionnés dans une arborescence à deux niveaux (ex: `ab/c1d2e3...`) pour des performances d'accès optimales sur tout système de fichiers.
- Les métadonnées, messages de commits, horodatages et arbres de répertoires sont gérés programmatiquement par le module [`DefFileGenerator.version_manager`](../DefFileGenerator/version_manager.py).

---

## 🌐 Exposition API & Studio Web

L'historique complet, les diffs interactifs et les retours en arrière (rollback) sont exposés via l'application web [`web/app.py`](../web/app.py) :
- `GET /api/versions/commits` : Liste chronologique des versions publiées.
- `GET /api/versions/{commit_id}` : Inspection d'un instantané historique.
- `GET /api/versions/diff` : Comparateur visuel de registres entre deux révisions.

> [!NOTE]
> Ce répertoire est géré automatiquement par le moteur. Aucun fichier ne doit y être ajouté ou modifié manuellement.
