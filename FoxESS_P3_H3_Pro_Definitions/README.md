# ☀️ FoxESS Gammes P3 Pro & H3 Pro — Définitions WebdynSunPM

Ce dossier contient l'ensemble des fichiers de définition normalisés au standard **WebdynSunPM** pour les onduleurs commerciaux, hybrides et systèmes de stockage **FoxESS** :
- **Fichiers Communs P3 Pro / H3 Pro** (Définition tout-en-un & définitions communes par équipement)
- **Gamme FoxESS P3 Pro** (Onduleurs solaires commerciaux 5 à 30 kW)
- **Gamme FoxESS H3 Pro** (Onduleurs hybrides triphasés haute tension 15 à 30 kW)
- **Autres Gammes FoxESS** (G-MAX Storage, R-Series Inverter & Battery, Profils SunSpec R/V/H3plus)

---

## ⚖️ 1. Arbitrage Factuel et Technique : Fichiers Communs vs Fichiers Séparés

### 1.1 Faits Matériels et Protocole (Sources Officielles)
1. **Identité Matérielle et Gamme Commerciale** :
   - D'après l'attestation officielle de conformité SunSpec de FoxESS (*FoxESS Statement Ref: FoxESS-P3-5.0-SH-P3-15.0-SH-Oswiadczenie.pdf*), les onduleurs **P3-Pro** (10 à 30 kW) et **H3-Pro** (10 à 30 kW) font partie de la même famille industrielle intitulée **`P3-Pro/H3Pro series`**.
   - **H3-Pro** est l'appellation commerciale internationale de référence, tandis que **P3-Pro** constitue la désignation OEM / régionale pour une architecture matérielle strictement équivalente.
2. **Identité Rigoureuse de la Table Modbus** :
   - Selon la spécification officielle *FoxESS Modbus Protocol (V1.05.05.00)*, la cartographie des registres holding FC03 est identique bit à bit pour les deux gammes :
     - **Onduleur** : 285 registres (39050..39426 et commandes 46000..49248).
     - **Compteur** : 61 registres (36100..36148 et 38801..38846).
     - **Batterie** : 893 registres (37000..37635 et EMS 48000..48999).
   - Les registres sont naturellement disjoints, autorisant une fusion sans le moindre conflit d'adresses.

### 1.2 Contraintes Opérationnelles du Portail WebdynSunPM
- Le champ `Model` présent dans la première ligne d'en-tête du fichier CSV (`#Format Version=1.0;DeviceType=...;Manufacturer=Fox ESS;Model=...`) est utilisé par les concentrateurs Webdyn et le portail SunPM pour classifier les équipements et afficher le nom commercial sur les dashboards.
- **Le besoin d'arbitrage** :
  - Si l'on ne fournit **que** des fichiers communs (`Model=P3 Pro / H3 Pro`), un exploitant ayant un parc identifié sous contrat comme `FoxESS P3 Pro` ou `FoxESS H3 Pro` verra une appellation hybride qui peut gêner la conformité contractuelle ou la recherche filtrée.
  - Si l'on ne fournit **que** des fichiers séparés (`Model=P3 Pro` et `Model=H3 Pro`), les intégrateurs qui gèrent des sites mixtes ou qui souhaitent un fichier universel doivent dupliquer leurs configurations.
  - De plus, sur le terrain, certains concentrateurs Webdyn interrogent l'onduleur sur **une seule adresse esclave Modbus** qui regroupe la télémétrie PV, le compteur réseau et la batterie (nécessitant le fichier Tout-en-Un), tandis que d'autres créent **3 tuiles logiques séparées** (Inverter, Meter, Battery).

### 1.3 Décision d'Arbitrage Retenue (Architecture Tri-Modale)
Pour répondre de manière exhaustive et juste à toutes les exigences opérationnelles sans aucun compromis :
1. **Dossier `Commun/`** : Fichiers universels avec en-tête `Model=P3 Pro / H3 Pro` (Onduleur, Compteur, Batterie) ainsi que la définition **Tout-en-Un** (`FOXESS_P3_H3_PRO_ALL_IN_ONE.csv`, 1 233 registres).
2. **Dossier `P3_Pro/`** : Fichiers spécifiques avec en-tête `Model=P3 Pro` pour intégration nominale stricte.
3. **Dossier `H3_Pro/`** : Fichiers spécifiques avec en-tête `Model=H3 Pro` pour intégration nominale stricte.
4. **Dossier `Autres_Gammes_FoxESS/`** : Fichiers pour les autres séries issues des sources du partage Support (G-MAX Storage, R-Series, SunSpec R/V/H3plus).
5. **Validation Automatisée par Tests** : Cet arbitrage est validé par la suite de tests `DefFileGenerator/tests/test_foxess_p3_h3_pro.py` (10 tests automatisés, vérification de parité stricte des registres, couverture All-In-One, validité 100% Webdyn).

---

## 📁 2. Structure du Répertoire Déployé

```text
FoxESS_P3_H3_Pro_Definitions/
├── Commun/                               # 📂 Fichiers communs aux deux gammes
│   ├── FOXESS_P3_H3_PRO_ALL_IN_ONE.csv   # 🌟 TOUT-EN-UN (1 233 reg., Onduleur + Compteur + Batterie)
│   ├── FOXESS_P3_H3_PRO_INVERTER.csv     # Onduleur commun (285 reg., Modèle 'P3 Pro / H3 Pro')
│   ├── FOXESS_P3_H3_PRO_METER.csv        # Compteur commun (61 reg., Modèle 'P3 Pro / H3 Pro')
│   ├── FOXESS_P3_H3_PRO_BATTERY.csv      # Batterie commune (893 reg., Modèle 'P3 Pro / H3 Pro')
│   ├── FoxESS_P3_H3_PRO.csv              # Alias historique Inverter
│   └── README.txt                        # Notice technique dédiée aux fichiers communs
│
├── P3_Pro/                                # 📂 Cas spécifique Gamme P3 Pro
│   ├── FOXESS_P3_PRO_INVERTER.csv        # Onduleur (Modèle : P3 Pro)
│   ├── FOXESS_P3_PRO_METER.csv           # Compteur de livraison (Modèle : P3 Pro)
│   ├── FOXESS_P3_PRO_BATTERY.csv         # Batterie & EMS (Modèle : P3 Pro)
│   └── README.txt                         # Notice spécifique P3 Pro
│
├── H3_Pro/                                # 📂 Cas spécifique Gamme H3 Pro
│   ├── FOXESS_H3_PRO_INVERTER.csv        # Onduleur (Modèle : H3 Pro)
│   ├── FOXESS_H3_PRO_METER.csv           # Compteur de livraison (Modèle : H3 Pro)
│   ├── FOXESS_H3_PRO_BATTERY.csv         # Batterie & EMS (Modèle : H3 Pro)
│   └── README.txt                         # Notice spécifique H3 Pro
│
├── Autres_Gammes_FoxESS/                  # 📂 Autres séries issues du partage Support
│   ├── FOXESS_GMAX_STORAGE.csv           # 🔋 Stockage G-MAX (894 reg., extrait du protocole Excel)
│   ├── FOXESS_R_SERIES_INVERTER.csv      # ☀️ Onduleur Gamme R (448 reg.)
│   ├── FOXESS_R_SERIES_BATTERY.csv       # 🔋 Batterie Gamme R (49 reg.)
│   ├── FOXESS_SUNSPEC_RV_H3PLUS.csv      # 🌐 Profil SunSpec R/V/H3plus (87 reg.)
│   └── README.txt                         # Notice descriptive autres gammes
│
├── FOXESS_P3_H3_PRO_ALL_IN_ONE.csv        # 🌟 Définition Tout-en-Un racine
├── FOXESS_P3_H3_PRO_INVERTER.csv          # 🌐 Inverter commun racine
├── FOXESS_P3_H3_PRO_METER.csv             # 🌐 Compteur commun racine
├── FOXESS_P3_H3_PRO_BATTERY.csv           # 🌐 Batterie commune racine
├── FOXESS_GMAX_STORAGE.csv                # 🔋 Stockage G-MAX racine
├── FOXESS_R_SERIES_INVERTER.csv           # ☀️ Onduleur Gamme R racine
├── FOXESS_R_SERIES_BATTERY.csv            # 🔋 Batterie Gamme R racine
├── FOXESS_SUNSPEC_RV_H3PLUS.csv           # 🌐 Profil SunSpec racine
├── FoxESS_P3_H3_PRO.csv                   # 🔄 Alias historique Inverter
├── README.txt                             # 📄 Guide global textuel
├── README.md                              # 📘 Ce document de référence
├── RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md  # 📊 Audit exhaustif Markdown
└── RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.json# 🤖 Métriques d'audit CI/CD
```

---

## ⚡ 3. Guide de Choix pour l'Intégration sur WebdynSunPM

1. **Architecture Multi-Équipements (3 Tuiles Distinctes sur le Portail)** :
   - Pour afficher des graphiques séparés pour la production solaire, les flux réseau et l'état batterie :
     - *Inverter* : `FOXESS_P3_H3_PRO_INVERTER.csv` (ou version `P3_Pro/` ou `H3_Pro/`)
     - *Meter* : `FOXESS_P3_H3_PRO_METER.csv` (ou version `P3_Pro/` ou `H3_Pro/`)
     - *Battery* : `FOXESS_P3_H3_PRO_BATTERY.csv` (ou version `P3_Pro/` ou `H3_Pro/`)

2. **Architecture Tout-en-Un (1 Seul Équipement Modbus Centralisé)** :
   - Lorsque le concentrateur Webdyn communique avec l'onduleur via une seule adresse esclave Modbus qui centralise toutes les données :
     - *Équipement unique* : `FOXESS_P3_H3_PRO_ALL_IN_ONE.csv` (1 233 registres, 0 conflit, tri strict par adresse).
