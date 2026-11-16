# 📊 RAPPORT TECHNIQUE MAJEUR — GÉNÉRATION DES FICHIERS DE DÉFINITION WEBDYNSUNPM
## Traçabilité Industrielle, Validation Modbus & Cartographie des Tags SunPM
**Projet** : `DefFileGenerator` · **Norme de référence** : `INV_HUAWEI_V4.csv` (WebdynSunPM R&D)
**Date de génération** : 2026-11-16 14:49:32 · **Répertoire audité** : `C:\Users\alexandre.calais\Documents\GitHub\DefFileGenerator\tagged_definitions`
**Statut global de conformité** : **100.0% VALIDE**

---

## 1. Synthèse Exécutive & Indicateurs Clés (KPIs)

Ce rapport technique valide la conformité totale des fichiers de définition industriels générés selon le standard **WebdynSunPM** (format CSV délimité par des points-virgules, 11 colonnes canoniques). L'ensemble des documents sources et partages constructeurs ont été accédés **strictement en lecture seule** sans altération.

```mermaid
pie title Répartition des Équipements par Catégorie
    "INV" : 584
    "OTH" : 454
    "MET" : 253
    "SEN" : 67
    "BAT" : 46
    "COM" : 36
    "BESS" : 14
    "BESS" : 13
    "BESS" : 4
    "BESS" : 1
```

| Indicateur Clé | Valeur Obtenue | Statut Webdyn |
| :--- | :--- | :--- |
| **Volume Total de Fichiers Définition** | **1 472 fichiers** | 📦 **Prêt Déploiement** |
| **Taux de Validité Schéma Webdyn** | **1472 / 1472 (100.0%)** | ✅ 100% Conforme |
| **Total Registres Modbus Validés** | **490 692 registres** | 🛡️ **Zéro Perte** |
| **Tags SunPM Canoniques Assignés** | **38 100 tags** (Densité : **7.8%**) | 🏷️ **Sémantique SunPM v1** |
| **Registres Holding / Input (FC03/FC04)** | **488 838** (99.6%) | ⚡ **Télémétrie 16/32/64 bits** |
| **Registres Coils / Discretes (FC01/FC02)** | **1 854** (0.4%) | 🔘 **États & Contacts Secs** |
| **Commandes d'Écriture (Action = 1)** | **71 359 consignes** | ✍️ **Contrôle Réseau & EMS** |
| **Alarmes & Événements (Action = 9)** | **20 298 registres** | 🚨 **Supervision Alertes** |
| **Fichiers Invalides / Rejetés** | **0** | 🛡️ **Zéro Bloquant** |

---

## 2. Analyse Détaillée par Catégorie d'Équipement

| Préfixe | Catégorie Industrielle | Fichiers | Registres Modbus | Tags SunPM | Densité | Constructeurs Représentatifs |
| :---: | :--- | :---: | :---: | :---: | :---: | :--- |
| `INV` | Onduleurs & Convertisseurs (Inverters) | **584** | 183 438 | 19 597 | 10.7% | SMA, Goodwe, Fronius, ABB |
| `OTH` | Autres Équipements (IO Modules, EV, Automates) | **454** | 147 679 | 10 133 | 6.9% | Webdyn, Delta, MarsRenewable, PowerOne |
| `MET` | Compteurs & Analyseurs Réseau (Meters) | **253** | 30 982 | 3 698 | 11.9% | Schneider, Socomec, Janitza, Lettel |
| `SEN` | Capteurs Météo & Pyranomètres (Sensors) | **67** | 1 803 | 235 | 13.0% | IngenieurBuro, Hukseflux, KippZonen, Seven Sensor |
| `BAT` | Stockage d'Énergie & Batteries (BESS / BMS) | **46** | 18 799 | 982 | 5.2% | MarsRenewable, Dyness, FOX ESS, Sungrow |
| `COM` | Passerelles & Loggers de Communication | **36** | 57 331 | 1 537 | 2.7% | HOYMILES, MarsRenewable, Sungrow, Huawei |
| `FOXESS` | FOXESS | **14** | 6 428 | 1 594 | 24.8% | FoxESS, FOX ESS, FOX_ESS |
| `MarsRenewable` | MarsRenewable | **13** | 10 455 | 119 | 1.1% | MarsRenewable |
| `MarsLocalController` | Contrôleur Central BESS (Mars Local Controller) | **4** | 33 492 | 87 | 0.3% | Mars Energy |
| `FoxESS` | FoxESS | **1** | 285 | 118 | 41.4% | FoxESS |

---

## 3. Couverture des Principaux Constructeurs Industriels

| Constructeur | Définitions | Total Registres | Tags SunPM | Part de Volume | Conformité |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Webdyn** | **123** | 13 345 | 2 483 | 2.7% | ✅ **100% Conforme** |
| **SMA** | **82** | 42 769 | 2 001 | 8.7% | ✅ **100% Conforme** |
| **MarsRenewable** | **69** | 49 369 | 1 183 | 10.1% | ✅ **100% Conforme** |
| **Schneider** | **59** | 11 874 | 1 028 | 2.4% | ✅ **100% Conforme** |
| **Huawei** | **45** | 8 656 | 1 519 | 1.8% | ✅ **100% Conforme** |
| **Fronius** | **45** | 10 967 | 1 496 | 2.2% | ✅ **100% Conforme** |
| **Delta** | **43** | 28 509 | 1 455 | 5.8% | ✅ **100% Conforme** |
| **Goodwe** | **38** | 9 328 | 1 977 | 1.9% | ✅ **100% Conforme** |
| **PowerOne** | **31** | 21 779 | 759 | 4.4% | ✅ **100% Conforme** |
| **ABB** | **30** | 13 402 | 581 | 2.7% | ✅ **100% Conforme** |
| **Sungrow** | **27** | 10 243 | 2 321 | 2.1% | ✅ **100% Conforme** |
| **SL Energy** | **26** | 9 045 | 1 247 | 1.8% | ✅ **100% Conforme** |
| **Socomec** | **26** | 1 944 | 270 | 0.4% | ✅ **100% Conforme** |
| **Janitza** | **22** | 1 201 | 319 | 0.2% | ✅ **100% Conforme** |
| **SAJ** | **20** | 3 253 | 397 | 0.7% | ✅ **100% Conforme** |
| **Lettel** | **20** | 3 502 | 511 | 0.7% | ✅ **100% Conforme** |
| **FoxESS** | **19** | 9 500 | 1 980 | 1.9% | ✅ **100% Conforme** |
| **Ingeteam** | **17** | 1 150 | 305 | 0.2% | ✅ **100% Conforme** |

---

## 4. Cartographie & Distribution des Tags SunPM Canoniques

Les tags appliqués sont issus du modèle d'apprentissage supervisé entraîné sur **3 098 règles** et respectent scrupuleusement la spécification `SunPM_DEfFile_Tags_v1.xlsx` :

| Rang | Tag Canonique SunPM | Occurrence | Description / Fonctionnalité Métier | Grandeur / Unité Type |
| :---: | :--- | :---: | :--- | :--- |
| **01** | `P_AC` | **711** | Puissance active totale injectée/consommée | W, kW |
| **02** | `SERIAL` | **662** | Numéro de série de l'équipement | Chaîne de caractères |
| **03** | `FREQ` | **556** | Fréquence du réseau électrique | Hz |
| **04** | `Q_AC` | **489** | Puissance réactive totale | Var, kVar |
| **05** | `MODEL` | **479** | Modèle ou désignation commerciale de l'appareil | Chaîne de caractères |
| **06** | `I_AC_3` | **477** | Courant efficace phase C / L3 | A |
| **07** | `I_AC_1` | **431** | Courant efficace phase A / L1 | A |
| **08** | `I_AC_2` | **422** | Courant efficace phase B / L2 | A |
| **09** | `U_AC_L1L2` | **376** | Tension composée réseau L1-L2 | V |
| **10** | `PF` | **357** | Facteur de puissance global (cos phi) | Sans unité [-1.0 à 1.0] |
| **11** | `MANUFACTURER` | **332** | Télémesure Modbus normalisée SunPM | Spécifique |
| **12** | `I_AC` | **314** | Télémesure Modbus normalisée SunPM | Spécifique |
| **13** | `E_TOT` | **309** | Énergie totale cumulée produite/consommée | kWh, MWh |
| **14** | `S_AC` | **267** | Puissance apparente totale | VA, kVA |
| **15** | `U_DC` | **261** | Tension bus continu DC | V |
| **16** | `U_AC_1` | **259** | Tension simple phase A-N / L1-N | V |
| **17** | `U_AC_L2L3` | **253** | Tension composée réseau L2-L3 | V |
| **18** | `U_AC_3` | **246** | Tension simple phase C-N / L3-N | V |
| **19** | `U_AC_2` | **232** | Tension simple phase B-N / L2-N | V |
| **20** | `U_AC_L3L1` | **215** | Tension composée réseau L3-L1 | V |
| **21** | `I_DC` | **214** | Courant continu DC total | A |
| **22** | `E_DAY` | **214** | Énergie journalière produite | kWh |
| **23** | `P_AC_1` | **206** | Télémesure Modbus normalisée SunPM | Spécifique |
| **24** | `P_AC_2` | **185** | Télémesure Modbus normalisée SunPM | Spécifique |
| **25** | `P_RATED` | **178** | Télémesure Modbus normalisée SunPM | Spécifique |

---

## 5. Focus Systèmes Complexes & Découpage Modulaire

### A. Écosystème Fox ESS (Onduleurs Hybrides, Batteries & Compteurs)
Conformément aux exigences de découpage d'équipements multi-fonctions, la gamme Fox ESS a été découpée en fichiers spécialisés indépendants :
- **Partie Onduleur** : `INV_Fox_ESS_FOXESS_P3_H3_PRO_INVERTER.csv` (Télémesures MPPT, tensions/courants AC triphasés, fréquence, état réseau).
- **Partie Batterie** : `BAT_FoxESS_P3_H3_PRO.csv` (**893 registres** - BMS1 & BMS2, tensions unitaires de cellules, températures min/max, SOC, SOH, alertes de préchauffage).
- **Partie Compteur d'Énergie** : `MET_FoxESS_P3_H3_PRO.csv` (**61 registres** - Télémétrie CT1/Meter1 et CT2/Meter2 pour tensions V, courants I, puissances P/Q/S, fréquence et facteur de puissance triphasés).
- **Gamme R-Series** : `BAT_FoxESS_RSeries.csv` (Batterie 49 registres SunSpec DER Modèle 702/703) et `INV_Fox_ESS_FOXESS_RSERIESINVERTER.csv` (Onduleur).

### B. Architecture BESS Mars Renewable Local Controller V7.0
Le contrôleur BESS multi-esclaves Mars Renewable est articulé en 4 profils dédiés :
1. **`MarsLocalController_V7.0_webdyn.csv` (Esclave 1 - Concentrateur BESS Principal)** : 11 704 registres consolidant les PCS, banques batteries, climatiseurs, compteurs et automates DIO.
2. **`MarsLocalController_V7.0_EMS_Slave247.csv` (Esclave 247 - Consignes EMS)** : 12 registres de commande de puissance avec `Action = 1` automatique.
3. **`MarsLocalController_V7.0_RackCell_SlaveN.csv` (Esclave N - Suivi Cellules)** : 5 030 registres surveillant jusqu'à 1 000 cellules individuelles.
4. **`MarsLocalController_V7.0_AllInOne.csv` (Univers Complet)** : 16 746 registres pour archivage universel.

---

## 6. Audit Qualité & Conformité Webdyn Schema

Toutes les définitions générées répondent aux exigences d'intégrité industrielle :
1. **Structure Sémantique** : Entête `modbusRTU;<Type>;<Constructeur>;<Modèle>` suivi de 11 colonnes délimitées par des points-virgules.
2. **Unicité Dynamique des Tags** : Aucune définition ne comporte de doublon de tag canonique (les tags `P_AC`, `FREQ`, `PF`, etc. sont rigoureusement uniques par fichier).
3. **Désambiguïsation Triphasée** : Les variables de phases (`sac1`, `pac1`, `phase A/B/C`, `L1/L2/L3`) sont protégées et isolées en tags unitaires (`P_AC_1`, `S_AC_1`, etc.) évitant toute écrasement par les grandeurs globales.
4. **Sécurité & Injection CSV** : Tous les champs textuels sont nettoyés contre les injections de formules (`=`, `+`, `-`, `@`) via `Generator.sanitize_csv_field`.
5. **Conformité des Actions** : Les commandes de pilotage (`CMD_*`) portent automatiquement le code `Action = 1` (écriture autorisée), tandis que les alarmes sont typées `Action = 9` et la télémétrie `Action = 4`.

---

## 7. Conclusion & État de Validation

**Validation Globale** : **1472 / 1472 définitions WebdynSunPM certifiées conformes** (0 erreur bloquante, 0 régression). La bibliothèque générée est immédiatement déployable sur les concentrateurs Webdyn et le serveur SunPM.

_Rapport généré automatiquement par le module `DefFileGenerator.report_generator`._
