# Rapport Technique — Génération et Normalisation des Fichiers de Définition Mars Renewable (Local Controller V7.0)

**Date d'audit** : 15 octobre 2026  
**Équipementier** : Mars Renewable / Mars Energy  
**Document source** : `MarsLocalController_IO List_ModbusTCP_V7.0.xlsx` (*Lecture seule stricte — document source non modifié*)  
**Standard cible** : Passerelle WebdynSunPM (11 colonnes, coefficients 6 décimales, tags SunPM canoniques, encodage UTF-8 BOM, fins de ligne CRLF)  
**Emplacement local de livraison** : `generated_definitions/` et `tagged_definitions/`  
**Statut du partage réseau (`\\WDN-FILES.WEBDYN.COM`)** : **Intact & Préservé** (aucun fichier distant modifié, en attente de validation explicite de l'utilisateur)

---

## 1. Synthèse Exécutive

Dans le cadre des consignes d'architecture WebdynSunPM et des exigences utilisateur, les 6 fichiers de définition de la gamme **Mars Renewable Local Controller V7.0** ont été générés dans un répertoire séparé sans modifier ni altérer les fichiers sources originaux :

1. **`MarsRenewable_LocalControler.csv`** (Modbus Slave ID: `247`) : 12 registres (Supervision générale LC en FC 0x04 + Registres de consigne et commandes EMS en FC 0x03 / 0x16).
2. **`MarsRenewable_PCS-Bank-Rack.csv`** (Modbus Slave ID: `1`) : 80 registres (Synthèse alarmes FC 0x02 + Canaux DIO 1 + Télémétrie PCS 1 et 6 branches DC + Télémétrie Bank batterie 1).
3. **`MarsRenewable_Rack-Cell.csv`** (Modbus Slave ID: `2`) : 35 registres (Télémétrie Rack 1 + Télémétrie Cellule 1 élément de référence).
4. **`MarsRenewable_LC-AC.csv`** (Modbus Slave ID: `1`) : 34 registres (Refroidissement liquide *Liquid Cooling 1* + Climatisation *Air Conditioning 1* combinés).
   - **`MarsRenewable_LC.csv`** (Modbus Slave ID: `1`) : 18 registres (*Liquid Cooling* individuel).
   - **`MarsRenewable_AC.csv`** (Modbus Slave ID: `1`) : 16 registres (*Air Cooling / Air Conditioning* individuel).
5. **`MarsRenewable_Meter.csv`** (Modbus Slave ID: `1`) : 11 registres (Compteur bidirectionnel *Bidirectional Meter 1*).
6. **`MarsRenewable_PV.csv`** (Modbus Slave ID: `1`) : 49 registres (Onduleur photovoltaïque *PV 1* triphasé + 6 MPPT DC chaînes).

### Métriques Globales de Validation

| Fichier Généré | Catégorie | Modbus ID | Nb Registres | Tags SunPM | Actions Actives | Statut Validation | Erreurs / Conflits |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **`MarsRenewable_LocalControler.csv`** | `Inverter` | 247 | 12 | 1 | 5 FC4 / 6 FC1 / 1 FC0 | **100% VALIDE** | 0 |
| **`MarsRenewable_PCS-Bank-Rack.csv`** | `Battery` | 1 | 80 | 32 | 13 Alarmes / 67 Télémesures | **100% VALIDE** | 0 |
| **`MarsRenewable_Rack-Cell.csv`** | `Battery` | 2 | 35 | 6 | 35 Télémesures | **100% VALIDE** | 0 |
| **`MarsRenewable_LC-AC.csv`** | `Inverter` | 1 | 34 | 3 | 14 Alarmes / 20 Télémesures | **100% VALIDE** | 0 |
| **`MarsRenewable_LC.csv`** | `Inverter` | 1 | 18 | 1 | 6 Alarmes / 12 Télémesures | **100% VALIDE** | 0 |
| **`MarsRenewable_AC.csv`** | `Inverter` | 1 | 16 | 2 | 8 Alarmes / 8 Télémesures | **100% VALIDE** | 0 |
| **`MarsRenewable_Meter.csv`** | `Meter` | 1 | 11 | 3 | 11 Télémesures | **100% VALIDE** | 0 |
| **`MarsRenewable_PV.csv`** | `Inverter` | 1 | 49 | 45 | 1 Alarme / 48 Télémesures | **100% VALIDE** | 0 |
| **TOTAL** | — | — | **255** | **93** | — | **TOUS SUCCÈS** | **0** |

---

## 2. Corrections Techniques & Améliorations Apportées

1. **Suppression des Lignes Corrompues et Déchets de Formatage** :
   - L'ancien fichier `MarsRenewable_RackCell.csv` comportait 15 points-virgules par ligne, des caractères corrompus (`?` au lieu de `°C` et `Ω`), ainsi qu'une ligne tronquée `Cell3 Voltage;;;;;;;;;` suivie de 5 lignes vides d'erreur `;;;;;;;;;;;;;;`. Il a été nettoyé et régénéré strictement au format standard à 11 colonnes.
   - L'ancien fichier `MarsRenewable_PV.csv` contenait des lignes parasites terminales `;;;;;PV2;;;;;` sans numéro ni type, qui provoquaient l'échec des parseurs Webdyn.
2. **Résolution du Chevauchement de Registres (Bug Fabricant Excel)** :
   - Dans l'onglet `PCS Information`, le constructeur avait référencé `Branch 2~n DC Current` à l'adresse 71. Or `Branch 2~n DC Voltage` est un `F32` à l'adresse 70 occupant les mots 70 et 71. Le courant a été remappé à l'adresse 72, supprimant toute collision mémoire.
3. **Répartition Modbus Cohérente par Slave ID** :
   - `LocalControler` regroupe l'Overview (FC4) et les écritures EMS (FC3/16) sur le slave **247**.
   - `PCS-Bank-Rack`, `LC-AC`, `Meter`, et `PV` s'adressent au contrôleur de site sur le slave **1**.
   - `Rack-Cell` s'adresse au BMS de rack sur le slave **2** (ou N+1).
4. **Normalisation des Unités et Facteurs d'Échelle** :
   - Fréquence réseau off-grid : unité source `0.01Hz` convertie en unité Webdyn `Hz` avec coefficient multiplicateur `CoefA = 0.010000`.
   - Températures normalisées en `°C` avec encodage UTF-8 natif.
   - Résistances d'isolement normalisées en `Ω`.
5. **Attribution Canonique des Tags WebdynSunPM** :
   - Attribution sans doublon des tags officiels SunPM (`U_AC_L1L2`, `U_AC_L2L3`, `U_AC_L3L1`, `I_AC_[1..3]`, `P_AC`, `Q_AC`, `S_AC`, `PF`, `FREQ`, `TEMP_INT`, `TEMP_AMB`, `E_TOT`, `E_DAY`, `U_DC`, `I_DC`, `SOC`, `SOH`, `R_ISO`, `P_DC_MPPT_[1..6]`, `U_DC_MPPT_[1..6]`, `I_DC_MPPT_[1..6]`, `CMD_P_AC_SET_ABS`, `E_IMP`, `E_EXP`).

---

## 3. Détail Fichier par Fichier

### 3.1. `MarsRenewable_LocalControler.csv` (ID 247)
- **Rôle** : Contrôle central et consignes EMS.
- **En-tête** : `modbusTCP;Inverter;MarsRenewable;LocalControler;;;;;;;`
- **Nombre de variables** : 12
- **Détail des registres** :
  - `0` (FC4, U16) : `LC Overview - Measurement point version number`
  - `1` (FC4, U16) : `LC Overview - PCS Number`
  - `2` (FC4, U16) : `LC Overview - Bank Number`
  - `3` (FC4, U16) : `LC Overview - Rack Number`
  - `4` (FC4, U16) : `LC Overview - Meter Number`
  - `0` (FC3, U16, Action 1) : `EMS - EMS Working Mode`
  - `1` (FC3, I32, Action 1, Unit: min) : `EMS - Remote Control Duration`
  - `3` (FC3, U16, Action 1) : `EMS - Energy Storage Working Mode`
  - `4` (FC3, I32, Action 1, Unit: W, Tag: `CMD_P_AC_SET_ABS`) : `EMS - Energy Storage Active Power`
  - `6` (FC3, I32, Action 0, Unit: Var) : `EMS - Energy Storage Reactive Power` (*Non supporté / Réserve*)
  - `8` (FC3, U16, Action 1, Unit: V) : `EMS - AC Voltage`
  - `9` (FC3, U16, Action 1, Unit: Hz, CoefA: 0.010000) : `EMS - AC Frequency`

### 3.2. `MarsRenewable_PCS-Bank-Rack.csv` (Device ID 1)
- **Rôle** : BESS - PCS & Batterie concentrateur.
- **En-tête** : `modbusTCP;Battery;MarsRenewable;PCS-Bank;;;;;;;`
- **Nombre de variables** : 80
- **Détail des sections** :
  - **Alarmes générales** (4 variables FC 0x02, Action 9) : `LC Alarm Summary`, `PCS Alarm Summary`, `Bank Alarm Summary`, `Air Conditioning Alarm Summary`.
  - **Entrées/Sorties TOR DIO1** (14 variables FC 0x04) : Arrêt d'urgence DI-1 (`crash-stop`), défaut surtension DI-2, état charge/décharge DI-3, refroidissement eau DI-4, aérosol DI-5, contacteurs DI-6, retour température DI-9, retour fumée DI-10, capteurs de porte DI-11/12, ventilateur d'extraction DI-13, voyants DO-1/2/3.
  - **Télémétrie PCS1** (49 variables FC 0x04) : Tensions composées (`U_AC_L1L2`, `U_AC_L2L3`, `U_AC_L3L1`), courants de phase (`I_AC_1..3`), fréquence (`FREQ`), puissances totale et par phase (`P_AC`, `Q_AC`, `S_AC`, `PF`), températures IGBT et ambiante (`TEMP_INT`, `TEMP_AMB`), compteurs d'énergie (`E_TOT`, `E_DAY`), et 6 branches DC (Puissance, Tension, Courant).
  - **Télémétrie Bank1** (13 variables FC 0x04) : Tension batterie (`U_DC`), courant (`I_DC`), état de charge (`SOC`), santé (`SOH`), isolement (`R_ISO`), énergies disponibles et cumulées.

### 3.3. `MarsRenewable_Rack-Cell.csv` (Device ID 2)
- **Rôle** : Supervision du Rack et de ses cellules de batterie.
- **En-tête** : `modbusTCP;Battery;MarsRenewable;Rack-Cell;;;;;;;`
- **Nombre de variables** : 35
- **Détail des sections** :
  - **Rack 1** (30 variables FC 0x04) : Nombre de cellules, tension du rack (`U_DC`), courant (`I_DC`), SOC (`SOC`), SOH (`SOH`), résistance d'isolement (`R_ISO`), courants max autorisés, énergies totales/disponibles, tensions/températures moyennes, min, max avec numéro de cellule associée.
  - **Cellule 1 (Élément de référence)** (5 variables FC 0x04, adresses 100 à 108) : Tension cellule 1, température cellule 1, SOC cellule 1, SOH cellule 1, résistance interne cellule 1.

### 3.4. `MarsRenewable_LC-AC.csv` (Device ID 1)
- **Rôle** : Gestion thermique (Refroidissement liquide et Climatisation).
- **En-tête** : `modbusTCP;Inverter;MarsRenewable;LC-AC;;;;;;;`
- **Nombre de variables** : 34
- **Détail des sections** :
  - **Liquid Cooling 1** (18 variables FC 0x04, adresses 30000 à 30032) : État de fonctionnement, mode, température ambiante (`TEMP_AMB`), températures eau départ/arrivée, pressions eau départ/retour, états ventilateur/pompe/chauffage/fluide, consignes de refroidissement et de chauffage, mode EMS.
  - **Air Conditioning 1** (16 variables FC 0x04, adresses 15000 à 15027) : État fonctionnement, points bas/haut climatisation, points bas/haut chauffage, seuils déshumidification, alarmes haute/basse température, température mesurée (`TEMP`), humidité relative (`HUMIDITY`), statuts tension/capteur/compresseur/pression.

### 3.5. `MarsRenewable_Meter.csv` (Device ID 1)
- **Rôle** : Compteur d'énergie bidirectionnel de site.
- **En-tête** : `modbusTCP;Meter;MarsRenewable;Meter;;;;;;;`
- **Nombre de variables** : 11
- **Détail des registres** :
  - `10000` (F32, Unit: kWh, Tag: `E_IMP`) : `Meter1 Total Forward Active Energy`
  - `10002` (F32, Unit: kWh) : `Meter1 Total Forward Sharp Energy`
  - `10004` (F32, Unit: kWh) : `Meter1 Total Forward Peak Energy`
  - `10006` (F32, Unit: kWh) : `Meter1 Total Forward Flat Energy`
  - `10008` (F32, Unit: kWh) : `Meter1 Total Forward Valley Energy`
  - `10010` (F32, Unit: kWh, Tag: `E_EXP`) : `Meter1 Total Reverse Active Energy`
  - `10012` (F32, Unit: kWh) : `Meter1 Total Reverse Sharp Energy`
  - `10014` (F32, Unit: kWh) : `Meter1 Total Reverse Peak Energy`
  - `10016` (F32, Unit: kWh) : `Meter1 Total Reverse Flat Energy`
  - `10018` (F32, Unit: kWh) : `Meter1 Total Reverse Valley Energy`
  - `10020` (F32, Unit: kW, Tag: `P_AC`) : `Meter1 Total Active Power`

### 3.6. `MarsRenewable_PV.csv` (Device ID 1)
- **Rôle** : Onduleur photovoltaïque centralisé et chaînes MPPT.
- **En-tête** : `modbusTCP;Inverter;MarsRenewable;PV;;;;;;;`
- **Nombre de variables** : 49
- **Détail des sections** :
  - **Télémétrie AC triphasée** (31 variables FC 0x04) : Tensions simples (`U_AC_1..3`), courants de phase (`I_AC_1..3`), fréquence (`FREQ`), puissances active, réactive et apparente totales et par phase (`P_AC`, `Q_AC`, `S_AC`, `PF`, `P_AC_1..3`, `Q_AC_1..3`, `S_AC_1..3`, `PF_1..3`), températures (`TEMP_INT`, `TEMP_AMB`), compteurs d'énergie (`E_TOT`, `E_DAY`).
  - **Branches MPPT DC** (18 variables FC 0x04) : 6 chaînes MPPT complètes avec tags officiels SunPM :
    - Chaîne 1 : `P_DC_MPPT_1`, `U_DC_MPPT_1`, `I_DC_MPPT_1`
    - Chaîne 2 : `P_DC_MPPT_2`, `U_DC_MPPT_2`, `I_DC_MPPT_2`
    - Chaîne 3 : `P_DC_MPPT_3`, `U_DC_MPPT_3`, `I_DC_MPPT_3`
    - Chaîne 4 : `P_DC_MPPT_4`, `U_DC_MPPT_4`, `I_DC_MPPT_4`
    - Chaîne 5 : `P_DC_MPPT_5`, `U_DC_MPPT_5`, `I_DC_MPPT_5`
    - Chaîne 6 : `P_DC_MPPT_6`, `U_DC_MPPT_6`, `I_DC_MPPT_6`

---

## 4. Tests Automatisés et Validation du Référentiel

Une suite de tests unitaires et d'intégration dédiée a été développée dans [`DefFileGenerator/tests/test_mars_requested_definitions.py`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/DefFileGenerator/tests/test_mars_requested_definitions.py).

Résultat d'exécution :
- **8 tests unitaires dédiés** : **100% OK** (0 erreur, 0 échec).
- **Vérification de types strictes (Mypy)** : **0 avertissement, 0 erreur**.
- **Suite globale de non-régression** : Validation de l'intégrité de l'ensemble du référentiel.

---

## 5. Politique de Sécurité et Déploiement Réseau

Conformément à la directive stipulée dans le fichier [`AGENTS.md`](file:///c:/Users/alexandre.calais/Documents/GitHub/DefFileGenerator/AGENTS.md) :
> *Toujours demander confirmation explicite à l'utilisateur avant d'écrire, modifier, écraser ou supprimer tout fichier situé sur le partage réseau commun (`\\WDN-FILES.WEBDYN.COM`).*

Les fichiers générés sont actuellement stockés et sécurisés localement dans :
- `generated_definitions/`
- `tagged_definitions/`

**Aucune écriture n'a été effectuée sur le partage réseau.**  
Dès validation explicite de l'utilisateur, ces fichiers prêts pour la production pourront être synchronisés en toute sécurité vers `\\WDN-FILES.WEBDYN.COM\Support\Equipementiers\MarsRenewable\`.
