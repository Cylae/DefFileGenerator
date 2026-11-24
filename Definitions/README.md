# 🏛️ Bibliothèque Historique de Référence (Legacy Reference Mirror)

Ce répertoire constitue le **miroir de compatibilité historique** des définitions d'équipements industriels au format WebdynSunPM. Comprend **291 fichiers de définition** validés couvrant plus de 60 000 registres Modbus.

---

## ⚠️ Distinction Clé : `Definitions/` vs `production/`

| Répertoire | Statut | Nombre de Modèles | Usage Recommandé |
|---|---|:---:|---|
| [`production/`](../production/) | **CANONIQUE MAÎTRE (Golden Source of Truth)** | **1 194 fichiers** | **Production active, passerelles Webdyn réelles, nouveaux déploiements.** |
| [`Definitions/`](.) | **Miroir Historique & Compatibilité** | **291 fichiers de définition** | Rétrocompatibilité multi-sites, parcs existants et suites de tests historiques. |

> [!IMPORTANT]
> Pour toute nouvelle installation, supervision ou intégration dans vos passerelles Webdyn, utilisez **toujours** les définitions validées situées dans le dossier [`production/`](../production/).
> Ce dossier `Definitions/` est maintenu pour assurer la non-régression des installations historiques déployées sur le terrain.

---

## 📊 Synthèse par Catégorie & Préfixe Normalisé

| Catégorie | Préfixe | Description | Constructeurs Notables |
|---|:---:|---|---|
| Onduleurs & Convertisseurs | `INV_` | Inverters solaires et éoliens | ABB, SMA, Huawei, Sungrow, Fronius, GoodWe, Kaco, Solis |
| Compteurs & Analyseurs Réseau | `MET_` | Energy meters & analyseurs de réseau | Carlo Gavazzi, Socomec, Schneider, Algodue, Janitza, Acrel |
| Systèmes de Stockage BESS | `BAT_` | Batteries & Battery Energy Storage Systems | Dyness, FoxESS, Mars Renewable, X7, Huawei LUNA |
| Capteurs & Instrumentation | `SEN_` | Pyranomètres, sondes de température, anémomètres | Kipp & Zonen, Delta Ohm, Hukseflux, EKO, Seven Sensor |
| Passerelles & Communication | `COM_` | Loggers, automates et passerelles télécoms | GoodWe, Huawei, SMA, Solplanet, Webdyn |
| Équipements Divers & Combinés | `OTH_` | Tableaux généraux, armoires et contacteurs | Mars Renewable, ABB, X7, Webdyn |
