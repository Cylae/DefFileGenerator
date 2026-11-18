# 🏭 Référentiel Officiel de Production (`production/`)

Ce dossier constitue la **source de vérité unique et certifiée** de l'ensemble des fichiers de définition WebdynSunPM du projet.
Tous les fichiers présents dans ce dossier respectent rigoureusement les **5 Piliers d'Ingénierie Apollo**.

---

## 📊 Statistiques Globales

- **Nombre total de définitions certifiées** : **1 118 fichiers**
- **Nombre de constructeurs couverts** : **137 marques**
- **Catalogue complet indexé** : Consultez [`production/CATALOG.json`](./CATALOG.json)

### Répartition par Type d'Équipement (Préfixes)
| Préfixe | Catégorie | Nombre de Définitions | Description |
| :---: | :--- | :---: | :--- |
| **`INV`** | Onduleur solaire / central | **418** | Onduleurs photovoltaïques (string, central, hybride). |
| **`OTH`** | Autres équipements & I/O | **376** | Bornes VE, relais de protection, automates, régulateurs. |
| **`MET`** | Compteur d'énergie | **194** | Compteurs de puissance triphasés et monophasés. |
| **`SEN`** | Capteur météo | **53** | Pyranomètres, sondes de température ambiante/module, anémomètres. |
| **`BAT`** | Stockage d'énergie / BESS | **48** | Systèmes de batteries industrielles, BMS, racks cellules. |
| **`COM`** | Passerelle de communication | **29** | Data loggers, concentrateurs Modbus TCP/RTU. |

---

## 🏛️ Les 5 Piliers de Conformité Apollo

Chaque fichier de définition contenu dans ce dossier est validé par le script `scripts/validate_production.py` et doit satisfaire :
1. **Nommage Standardisé** : `<PREFIX>_<Fabricant>_<Modèle>.csv` (sans sous-tirets pendants, ni suffixes temporaires).
2. **Adressage Modbus Valide** : Adresses décimales (0..65535) et suffixes de longueur octet normalisés pour les chaînes (`STRING` -> `1002_32`).
3. **Codes Action Valides** : Valeurs normalisées Webdyn (`0, 1, 2, 4, 8, 9, 10`).
4. **Codes Fonction Modbus Valides** : `1` (Coils), `2` (Discrete Inputs), `3` (Holding Registers), `4` (Input Registers).
5. **Tags SunPM Canoniques** : Mappage sur le dictionnaire canonique SunPM sans doublons physiques non autorisés.
6. **Encodage & Schéma** : Strict UTF-8 sans BOM, séparateur point-virgule (`;`), schéma strict à 11 colonnes.

---

## 🔍 Commande de Validation

Pour vérifier l'intégrité intégrale du répertoire :
```bash
python scripts/validate_production.py
```
