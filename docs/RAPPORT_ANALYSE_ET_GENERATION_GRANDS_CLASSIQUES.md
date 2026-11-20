# RAPPORT D'ANALYSE TECHNIQUE POINTUE & CERTIFICATION DES GRANDS CLASSIQUES INDUSTRIELS
**Huawei, CHINT, SMA, Solplanet, Sungrow, GoodWe, Growatt, Fronius, SolarEdge, Socomec, Schneider Electric, Carlo Gavazzi**

---

## 1. Contexte & Objectifs de Standardisation

Le présent rapport formalise l'analyse détaillée d'ingénierie et la certification complète des fichiers de définition WebdynSunPM pour les constructeurs majeurs du secteur photovoltaïque, de stockage stationnaire (BESS) et de comptage d'énergie.

Dans le cadre des directives Apollo-Standard et des exigences Webdyn, l'ensemble des fichiers générés respecte scrupuleusement les règles intangibles suivantes :
1. **Unification du Protocole Modbus** : Bannissement absolu des étiquettes historiques obsolètes telles que `ModbusTCP` ou `ModbusRTU`. L'en-tête standardisé est strictement : `Modbus;<Catégorie>;<Constructeur>;<Modèle>;;;;;;;`.
2. **Encodage & Formatage Strict** : Encodage **UTF-8 sans BOM** (`\xef\xbb\xbf` proscrit), séparateur point-virgule (`;`), séparateurs de ligne CRLF (`\r\n`), 11 colonnes par ligne.
3. **Indexation Séquentielle** : Numérotation continue de $1$ à $N$ sans rupture ni décalage.
4. **Typage et Normalisation des Adresses** : Adressage Modbus décimal (base 0 ou 1 selon le constructeur, normalisé pour le firmware Webdyn) et notation de longueur pour les chaînes de caractères (ex. `30000_30`).
5. **Codes Actions Webdyn Certifiés** :
   - `4` : Télémétrie périodique (échantillonnage standard 5 minutes, historisé).
   - `9` : États critiques, alarmes et codes d'erreur.
   - `1` : Consignes de commande et régulation dynamique (Read/Write).
   - `0` : Registres réservés ou non historisés.
6. **Coefficients à 6 Décimales** : Formatage strict `f"{coef:.6f}"` ($A \cdot X + B$).
7. **Tags Canoniques SunPM Uniques** : Attribution rigoureuse des tags métier SunPM (`P_AC`, `Q_AC`, `S_AC`, `U_AC_1..3`, `I_AC_1..3`, `FREQ`, `PF`, `E_TOT`, `E_DAY`, `TEMP_INT`, `CMD_*`), avec unicité stricte au sein de chaque fichier (aucun doublon de tag simple, aucune concaténation composite).

---

## 2. Analyse Technique Pointue par Constructeur

### 2.1. CHINT (Compteurs & Centrales de Mesure)
- **Modèles Couverts** :
  - `MET_Chint_DTSU666.csv` : Compteur triphasé 4 quadrants multi-fonctions (Raccordement direct ou TC).
  - `MET_Chint_DDSU666.csv` : Compteur monophasé rail DIN.
  - `MET_Chint_DTSU666-H.csv` : Capteur de puissance intelligent Huawei (Smart Power Sensor).
- **Architecture Mémoire & Registres** :
  - Pour les modèles **DTSU666 / DDSU666 standard** (Manuel ZTW0.464.0104 V1, table 11) :
    - Toutes les grandeurs électriques sont encodées au format **IEEE-754 Simple Précision (Float F32)** sur 2 registres consécutifs (mot de poids fort en premier / Big-Endian).
    - Énergies actives directes et inverses : adresses `4126` à `4144` (Total, Phase A, B, C, Net). Unité `kWh`, $A=1.000000$.
    - Tensions composées ($U_{12}, U_{23}, U_{31}$) : `8192..8196` (F32, $A=1.000000$, Unité `V`).
    - Tensions simples ($V_1, V_2, V_3$) : `8198..8202` (F32, $A=1.000000$, Unité `V`).
    - Courants de phase ($I_1, I_2, I_3$) : `8204..8208` (F32, $A=1.000000$, Unité `A`).
    - Puissances actives ($P, P_1, P_2, P_3$) : `8210..8216` (F32, $A=1.000000$, Unité `kW`).
    - Puissances réactives ($Q, Q_1, Q_2, Q_3$) : `8218..8224` (F32, $A=1.000000$, Unité `kvar`).
    - Facteurs de puissance ($PF, PF_1, PF_2, PF_3$) : `8234..8240` (F32, $A=1.000000$, sans unité).
    - Fréquence ($F$) : `8260` (F32, $A=1.000000$, Unité `Hz`).
    - Correction méticuleuse Apollo : Élimination définitive de l'erreur historique consistant à appliquer les chiffres de la colonne "Résolution" ($0.1$, $0.001$, $0.01$) aux données F32, ce qui écrasait les valeurs mesurées par un facteur 10, 100 ou 1000 !
  - Pour le modèle **DTSU666-H (Spécifique Huawei)** :
    - Ce compteur communique avec les onduleurs Huawei SUN2000 via RS485 et transmet les données sous forme d'**entiers signés 32 bits (I32)** aux adresses `37101..37138`.
    - Correction d'un bug critique historique : l'ancienne définition `MET_Huawei_V4_METER.csv` déclarait ces registres en `U16` avec un coefficient de $1.0$, omettait la moitié des tags de phase, et mappait le courant de phase B sur `I_AC_3` ! La nouvelle version `MET_Chint_DTSU666-H.csv` et `MET_Huawei_DTSU666-H.csv` rétablit le typage `I32` et le mapping exact des phases.

### 2.2. HUAWEI (Onduleurs SUN2000 Commercial & Utility)
- **Modèles Couverts** :
  - `INV_Huawei_SUN2000-M3.csv` : Gamme tertiaire et résidentielle 3KTL à 100KTL (M0, M1, M2, M3).
  - `INV_Huawei_SUN2000-Utility.csv` : Gamme centrale au sol 1500 V (185KTL-H1, 215KTL-H0, 330KTL-H1) avec 24 entrées strings.
  - `MET_Huawei_DTSU666-H.csv` : Compteur de boucle de régulation.
- **Architecture Mémoire & Registres** :
  - Télémétrie PV Strings : `32016` à `32063` (Tensions $0.1\text{ V}$, Courants $0.01\text{ A}$ sur 12 à 24 chaînes de modules).
  - Puissance DC totale : `32064` (I32, $0.001\text{ kW}$).
  - Tensions réseau : `32066..32068` (Composées, U16, $0.1\text{ V}$), `32069..32071` (Simples, U16, $0.1\text{ V}$).
  - Courants réseau : `32072..32076` (I32, $0.001\text{ A}$).
  - Puissance active : `32080` (I32 signée, $0.001\text{ kW}$ soit $1\text{ W}$).
  - Puissance réactive : `32082` (I32 signée, $0.001\text{ kvar}$).
  - Facteur de puissance : `32084` (I16, $0.001$).
  - Fréquence : `32085` (U16, $0.01\text{ Hz}$).
  - Compteurs d'énergie : `32106` (`E_TOT`, U32, $0.01\text{ kWh}$), `32114` (`E_DAY`, U32, $0.01\text{ kWh}$).
  - Registres de régulation et pilotage (Holding FC3) :
    - `40122` : Consigne $\cos \phi$ (`CMD_PF`, I16, $0.001$).
    - `40125` : Consigne de limitation de puissance active relative (`CMD_P_AC_SET_PERCENT`, U16, $0.1\%$).
    - `40126` : Consigne absolue en Watts (`CMD_P_AC_SET_ABS`, U32, $1\text{ W}$).
    - `40200` / `40201` : Démarrage et arrêt à distance (`CMD_STARTUP`, `CMD_SHUTDOWN`).

### 2.3. SMA (Sunny Tripower & Core Series)
- **Modèles Couverts** :
  - `INV_SMA_SunnyTripower.csv` : Gamme STP 3.0 à 25000TL & Core 1 (6 MPPTs).
  - `INV_SMA_Core1_Core2.csv` : Gamme Core 1 (STP 50-40) et Core 2 (STP 110-60, 12 MPPTs / 24 entrées).
- **Architecture Mémoire & Registres** :
  - Compteurs d'énergie cumulée : `30513` (`E_TOT`, U64, Wh $\to$ $A=0.001000\text{ kWh}$), `30517` (`E_DAY`, U64, Wh $\to$ $A=0.001000\text{ kWh}$).
  - Télémétrie électrique instantanée :
    - `30775` : Puissance active totale (`P_AC`, I32, $1\text{ W}$).
    - `30777..30781` : Puissances actives par phase L1, L2, L3 (`P_AC_1..3`, I32, $1\text{ W}$).
    - `30783..30787` : Tensions simples L1-N, L2-N, L3-N (`U_AC_1..3`, I32, $0.01\text{ V}$).
    - `30789..30793` : Tensions composées L1-L2, L2-L3, L3-L1 (`U_AC_L1L2..`, I32, $0.01\text{ V}$).
    - `30795..30799` : Courants de phase L1, L2, L3 (`I_AC_1..3`, U32, $0.001\text{ A}$).
    - `30803` : Fréquence réseau (`FREQ`, U32, $0.01\text{ Hz}$).
    - `30805..30811` : Puissance réactive totale et par phase (`Q_AC`, I32, $1\text{ var}$).
    - `30813..30819` : Puissance apparente totale et par phase (`S_AC`, I32, $1\text{ VA}$).
    - `30821` : Facteur de déplacement $\cos \phi$ (`PF`, I32, $0.001$).
    - `30953` / `30957` : Températures interne et radiateur (`TEMP_INT`, `T_HEATSINK`, I32, $0.1\text{ \degree C}$).
  - MPPTs multiples : MPPT 1 (`30769..30773`), MPPT 2..6 (`30959..30987`), MPPT 7..12 (`31251..31285`).
  - Consignes de régulation (Holding FC3) : `40018` (`CMD_P_AC_SET_PERCENT`, U32, $0.01\%$), `40020` (`CMD_P_AC_SET_ABS`, U32, $1\text{ W}$), `40024` (`CMD_Q_SET`, I32, $1\text{ var}$), `40026` (`CMD_PF`, I32, $0.001$).

### 2.4. SOLPLANET (AISWEI ASW Series)
- **Modèles Couverts** : `INV_Solplanet_ASW_ThreePhase.csv` (ASW 3K à 50K-LT).
- **Architecture Mémoire & Registres** :
  - Mesures en Function Code 4 (Input Registers) :
    - `0..47` : Informations machine (Modèle, Numéro de série, Version DSP en chaînes 16 mots).
    - `100..105` : MPPTs 1 à 3 (Tensions $0.1\text{ V}$, Courants $0.1\text{ A}$).
    - `110..115` : Tensions et courants des phases A, B, C.
    - `116` : Fréquence (`FREQ`, U16, $0.01\text{ Hz}$).
    - `117` : Puissance active totale (`P_AC`, I32, $1\text{ W}$).
    - `119` : Puissance réactive totale (`Q_AC`, I32, $1\text{ var}$).
    - `121` : Facteur de puissance (`PF`, I16, $0.001$).
    - `122` : Puissance DC totale (`P_DC`, U32, $1\text{ W}$).
    - `125` / `127` : Énergies totale et journalière (`E_TOT`, `E_DAY`, U32, $0.1\text{ kWh}$).
  - Consignes de régulation en Function Code 3 (Holding Registers) :
    - `201` : Interrupteur Marche/Arrêt (`CMD_START_STOP`, U16).
    - `2101` : Limitation active relative (`CMD_P_AC_SET_PERCENT`, U16, $1\%$).

### 2.5. SUNGROW (Gamme SG Commerciale)
- **Modèles Couverts** : `INV_Sungrow_SG_Series.csv` (SG33CX, SG50CX, SG110CX).
- **Architecture Mémoire & Registres** :
  - Mesures en Function Code 4 (Input Registers) :
    - Adressage standard Sungrow `4989` à `5044`.
    - Spécificité : Les valeurs 32 bits utilisent le format mot inversé (`U32_W` / `I32_W`).
    - `5000` : Puissance nominale (`P_RATED`, U16, $0.1\text{ kW}$).
    - `5002` / `5003` : Production journalière (`E_DAY`, U16, $0.1\text{ kWh}$) et totale (`E_TOT`, U32_W, $1\text{ kWh}$).
    - `5010..5015` : MPPTs 1 à 3 (Tensions et courants $0.1\text{ V}$, $0.1\text{ A}$).
    - `5018..5023` : Tensions composées et courants de phase.
    - `5030` : Puissance active totale (`P_AC`, I32_W, $1\text{ W}$).
    - `5032` : Puissance réactive totale (`Q_AC`, I32_W, $1\text{ var}$).
    - `5034` : Facteur de puissance (`PF`, I16, $0.001$).
    - `5035` : Fréquence (`FREQ`, U16, $0.1\text{ Hz}$).
  - Consignes en Function Code 3 : `5000` (`CMD_START_STOP`), `5005` (`CMD_P_AC_SET_PERCENT`, U16, $0.1\%$), `5006` (`CMD_PF`, I16, $0.001$).

### 2.6. GOODWE (Séries SMT & Hybrides ET)
- **Modèles Couverts** :
  - `INV_Goodwe_SMT_Series.csv` : Onduleurs réseau triphasés tertiaires 25K..80K.
  - `INV_Goodwe_ET_Series.csv` : Onduleurs hybrides avec gestion de batterie BESS.
- **Architecture Mémoire & Registres** :
  - Function Code 3 (Holding Registers) :
    - `30040..30045` : Strings PV1 à PV3.
    - `30053..30058` : Tensions simples et courants phases L1, L2, L3.
    - `30060` : Puissance active (`P_AC`, I32, $1\text{ W}$).
    - `30062` : Puissance réactive (`Q_AC`, I32, $1\text{ var}$).
    - `30067` / `30068` : Énergies journalière et totale ($0.1\text{ kWh}$).
  - Télémétrie Batterie (ET Series) :
    - `35180` / `35181` : Tension et courant batterie (`U_BAT`, `I_BAT`, $0.1\text{ V}$, $0.1\text{ A}$).
    - `35182` : Puissance batterie de charge/décharge (`P_BAT`, I32, $1\text{ W}$).
    - `35184` / `35185` : État de charge et de santé (`SOC`, `SOH`, U16, $1\%$).
    - `35186` : Température BMS (`T_BAT`, I16, $0.1\text{ \degree C}$).
  - Consignes : `40001` (`CMD_P_AC_SET_PERCENT`), `40003` (`CMD_PF`), `40005` (`CMD_START_STOP`).

### 2.7. GROWATT (Série MAX Tertiaire)
- **Modèles Couverts** : `INV_Growatt_MAX_Series.csv` (MAX 50K..125KTL3-X).
- **Architecture Mémoire & Registres** :
  - Function Code 3 (Holding Registers) :
    - `1` : Puissance DC totale (`P_DC`, U32, $0.1\text{ W}$).
    - `3..16` : Tensions et courants PV1 à PV4.
    - `35` : Puissance active AC totale (`P_AC`, U32, $0.1\text{ W}$).
    - `37` : Fréquence (`FREQ`, U16, $0.01\text{ Hz}$).
    - `38..48` : Tensions, courants et puissances par phase L1, L2, L3.
    - `53` / `55` : Production journalière et cumulée ($0.1\text{ kWh}$).
    - `93` / `94` : Températures interne et radiateur IPM ($0.1\text{ \degree C}$).
  - Consignes : `4000` (`CMD_START_STOP`), `4001` (`CMD_P_AC_SET_PERCENT`), `4004` (`CMD_PF`).

### 2.8. FRONIUS (Symo, Primo, Gen24)
- **Modèles Couverts** : `INV_Fronius_Symo_Primo_Gen24.csv`.
- **Architecture Mémoire & Registres** :
  - Conforme au profil standard **SunSpec IEEE Float Model (Modèle 103)** :
    - Toutes les métriques sont exprimées en flottant 32 bits IEEE-754 (`F32`), éliminant tout besoin de facteur d'échelle dynamique SunSpec (`Scale Factor`).
    - `40072..40078` : Courant total et courants par phase (`I_AC`, `I_AC_1..3`, F32, $1\text{ A}$).
    - `40080..40084` : Tensions composées (`U_AC_L1L2..`, F32, $1\text{ V}$).
    - `40086..40090` : Tensions simples (`U_AC_1..3`, F32, $1\text{ V}$).
    - `40092` : Puissance active totale (`P_AC`, F32, $1\text{ W}$).
    - `40094` : Fréquence (`FREQ`, F32, $1\text{ Hz}$).
    - `40096` / `40098` : Puissance apparente et réactive (`S_AC`, `Q_AC`, F32).
    - `40100` : Facteur de puissance (`PF`, F32, $1.0$).
    - `40102` : Énergie totale (`E_TOT`, F32, Wh $\to$ $A=0.001000\text{ kWh}$).
  - Contrôle et pilotage (Modèle 123 SunSpec) : `40232` (`CMD_START_STOP`), `40234` (`CMD_P_AC_SET_PERCENT`), `40242` (`CMD_PF`).

### 2.9. SOLAREDGE (Onduleurs Commerciaux Triphasés SE)
- **Modèles Couverts** : `INV_SolarEdge_SE_Series.csv` (SE10K à SE33.3K).
- **Architecture Mémoire & Registres** :
  - Profil SunSpec standard Modbus Model 103 aux adresses `40072..40113`.
  - Extension propriétaire **SolarEdge Enhanced Power Control (EPC)** :
    - `61440` : Activation commande à distance (`CMD_START_STOP`).
    - `61441` : Limitation active en pourcentage (`CMD_P_AC_SET_PERCENT`, U16, $1\%$).
    - `61442` : Consigne $\cos \phi$ (`CMD_PF`, I16, $0.001$).

### 2.10. SOCOMEC (Diris A & Countis E)
- **Modèles Couverts** : `MET_Socomec_Diris_A_Countis_E.csv` (Diris A10, A20, A40 et Countis E23, E27, E33).
- **Architecture Mémoire & Registres** :
  - Mémoire standard Socomec JBUS / Modbus aux adresses `50514` à `50776`.
  - Tensions composées et simples : `50514..50524` (U32, $0.01\text{ V}$).
  - Courants de phase et neutre : `50528..50534` (U32, $0.001\text{ A}$).
  - Puissances actives totale et de phase : `50536` et `50544..50548` (I32, $0.01\text{ kW}$ soit $10\text{ W}$).
  - Puissances réactives totale et de phase : `50538` et `50550..50554` (I32, $0.01\text{ kvar}$).
  - Facteurs de puissance : `50542` et `50562..50566` (I32, $0.001$).
  - Énergies actives totale importée et exportée : `50770` (`E_TOT`, U32, $1\text{ kWh}$), `50776` (`E_EXP`, U32, $1\text{ kWh}$).

### 2.11. SCHNEIDER ELECTRIC (PowerLogic PM5000 & Acti9 iEM3000)
- **Modèles Couverts** : `MET_Schneider_PM5000_IEM3000.csv` (PM5100, PM5300, PM5500 et iEM3100 à iEM3300).
- **Architecture Mémoire & Registres** :
  - Grandeurs électriques instantanées au format **IEEE-754 F32** aux adresses `2999` à `3109` :
    - Courants phase 1, 2, 3 et neutre : `2999..3005` (F32, $1\text{ A}$).
    - Tensions composées et simples : `3019..3031` (F32, $1\text{ V}$).
    - Puissances actives par phase et totale : `3053..3059` (F32, $1\text{ kW}$).
    - Puissance réactive et apparente totale : `3067`, `3075` (F32, $1\text{ kvar}$, $1\text{ kVA}$).
    - Fréquence : `3109` (F32, $1\text{ Hz}$).
  - Compteurs d'énergie multi-quadrants (F32) :
    - `45099` : Énergie active totale importée (`E_TOT`, F32, $1\text{ kWh}$).
    - `45101` : Énergie active totale exportée (`E_EXP`, F32, $1\text{ kWh}$).
    - `45103` : Énergie réactive totale (`Q_TOT`, F32, $1\text{ kvarh}$).
    - `45111..45115` : Énergies actives par phase 1, 2, 3 (`E_TOT_1..3`, F32, $1\text{ kWh}$).

### 2.12. CARLO GAVAZZI (Séries EM24, EM340, EM210)
- **Modèles Couverts** : `MET_CarloGavazzi_EM24_EM300.csv`.
- **Architecture Mémoire & Registres** :
  - Grandeurs électriques instantanées en Function Code 3 aux adresses `1` à `79` :
    - Tensions simples : `1..5` (I32, $0.1\text{ V}$).
    - Courants de phase : `13..17` (I32, $0.001\text{ A}$).
    - Puissances actives par phase et totale : `19..23` et `41` (I32, $0.1\text{ W}$).
    - Puissances apparentes et réactives : `25..35` et `43..45` (I32, $0.1\text{ VA}$, $0.1\text{ var}$).
    - Facteurs de puissance : `47..50` (I16, $0.001$).
    - Fréquence : `52` (I16, $0.1\text{ Hz}$).
    - Énergies cumulées : `53` (`E_TOT`, U32, $0.1\text{ kWh}$), `55` (`Q_TOT`, U32, $0.1\text{ kvarh}$), `79` (`E_EXP`, U32, $0.1\text{ kWh}$).

---

## 3. Analyse Fondamentale : Origine, Causes et Résolution des « Coefficients à 1000 » et Ratios d'Échelle

La présence récurrente dans les fichiers de définition historiques de coefficients valant `1000.000000` (ou son inverse `0.001000`), ainsi que des valeurs incongrues comme `0.1` ou `0.01` sur des nombres flottants, a constitué l'une des sources majeures d'incohérence métrologique sur le terrain. L'analyse détaillée des causes racines met en lumière trois mécanismes distincts :

### 3.1. Formule de Transfert Linéaire Webdyn
Toute passerelle Webdyn calcule la grandeur physique stockée et télétransmise via l'équation affine universelle :
$$\text{Grandeur Mesurée} = \text{CoefA} \times \text{Valeur Brute du Registre} + \text{CoefB}$$

Où :
- $\text{Valeur Brute}$ est le résultat du décodage binaire des registres Modbus selon le `data_type` (`U16`, `I16`, `U32`, `I32`, `F32`, etc.).
- $\text{CoefA}$ est le facteur d'échelle multiplicatif (colonne 8 du CSV, formaté à 6 décimales).
- $\text{CoefB}$ est l'offset additif (colonne 9 du CSV, généralement `0.000000`).

---

### 3.2. Cause Racine N°1 : La Conversion d'Unités d'Ingénierie (kW $\leftrightarrow$ W et Wh $\leftrightarrow$ kWh)
C'est la cause la plus légitime et fréquente de la présence d'un facteur 1000 :
1. **Passage de kW en W ($A = 1000.000000$)** :
   - De nombreux équipements de forte puissance (onduleurs centraux, compteurs industriels haute tension tels que Schneider PM5000, ABB, Socomec) restituent nativement la puissance active en **kilowatts (kW)**.
   - Si la plateforme de supervision centrale (ou la convention SunPM historique) exige que le tag `P_AC` soit exprimé en **Watts (W)** pour homogénéiser les calculs de bilan de site avec des équipements résidentiels, la formule exige :
     $$P_{\text{Watts}} = 1000 \times P_{\text{kW}} \implies \mathbf{CoefA = 1000.000000}$$
   - *Inversement* : Si la définition déclare l'unité cible `kW` (comme dans nos fichiers certifiés `MET_Schneider_PM5000_IEM3000.csv` et `MET_Chint_DTSU666.csv`), alors la grandeur reste en kW et **$\mathbf{CoefA = 1.000000}$**.
2. **Passage de Wh en kWh ($A = 0.001000 = \frac{1}{1000}$)** :
   - Chez SMA (ex: registre 30513 `Metering.TotWhOut` sur 4 registres `U64`), l'énergie cumulée est comptabilisée en **Watt-heures (Wh)**.
   - Pour intégrer cette donnée dans le tag standard `E_TOT` exprimé en **kWh**, il faut diviser par 1000 :
     $$E_{\text{kWh}} = \frac{E_{\text{Wh}}}{1000} = 0.001 \times E_{\text{Wh}} \implies \mathbf{CoefA = 0.001000}$$
   - Même logique pour les courants exprimés en milliampères (mA) : $1\text{ mA} = 0.001\text{ A} \implies \mathbf{CoefA = 0.001000}$.

---

### 3.3. Cause Racine N°2 : La Confusion Désastreuse « Multiplicateur vs Diviseur » dans les Notices Constructeurs
Cette anomalie historique provient d'une erreur d'interprétation humaine des documentations techniques :
- Dans les notices d'onduleurs ou de compteurs (notamment traduites du chinois ou de l'allemand), la colonne de cadrage est fréquemment libellée de façon ambiguë :
  - `Scale: 1000`
  - `Factor: 1000`
  - `Ratio: 1000`
  - `Divisor: 1000`
- **Le piège** : Dans la documentation, « Factor 1000 » signifie généralement que la valeur brute dans le registre est multipliée par 1000 par le microcontrôleur pour stocker 3 décimales sous forme d'un entier, et qu'il faut donc **diviser par 1000** à la lecture pour retrouver la vraie valeur !
- **L'erreur commise** : L'auteur de la définition saisissait naïvement `1000` dans la colonne `CoefA`.
- **La conséquence catastrophique** : La passerelle Webdyn multipliait par 1000 au lieu de multiplier par 0.001 ! La valeur résultante était alors faussée d'un facteur **$1\,000\,000$ ($10^6$)** !
  - Exemple : un courant réel de $5.213\text{ A}$ stocké sous la valeur entière $5213$ était multiplié par $1000$, devenant $5\,213\,000\text{ A}$ sur le portail Webdyn !
  - Pour masquer cette aberration, certains techniciens modifiaient ensuite le libellé de l'unité ou ajoutait un diviseur en aval, introduisant un désordre ingérable dans la base de données.

---

### 3.4. Cause Racine N°3 : Le Piège de la Colonne « Résolution » sur les Flottants IEEE-754 (`F32`) — Le Cas Emblématique CHINT
Il s'agit de la découverte la plus pointue mise au jour lors de notre audit :
1. **La différence fondamentale entre Entier Virgule Fixe et Flottant IEEE-754** :
   - Sur un registre entier (`I16`, `U16`, `I32`, `U32`), les décimales ne peuvent pas être encodées nativement. Le constructeur impose un pas de discrétisation (ex: $230.4\text{ V}$ est transmis sous la valeur $2304$, imposant $\text{CoefA} = 0.1$).
   - Sur un registre **IEEE-754 Simple Précision (`F32`)**, la valeur binaire (1 bit de signe, 8 bits d'exposant, 23 bits de mantisse) représente **déjà la valeur réelle avec toutes ses décimales** (ex: `230.4000` V, `5.2130` A, `50.0200` Hz).
2. **Le Contresens de la Table 11 du Manuel CHINT DTSU666 (ZTW0.464.0104)** :
   - Le tableau officiel indique :
     - Tension composée $U_{ab}$ (`8192`) : Type `Float`, Longueur `2`, **Resolution `0.1`**, Unit `V`.
     - Courant $I_a$ (`8204`) : Type `Float`, Longueur `2`, **Resolution `0.001`**, Unit `A`.
     - Fréquence $F$ (`8260`) : Type `Float`, Longueur `2`, **Resolution `0.01`**, Unit `Hz`.
     - Facteur de puissance $PF$ (`8234`) : Type `Float`, Longueur `2`, **Resolution `0.001`**, sans unité.
   - **L'erreur historique** : L'auteur du fichier `MET_Chint_DTSU666.csv` (et du template JSON) a interprété le mot « *Resolution* » comme un *Scale Factor* Modbus, et a appliqué $\text{CoefA} = 0.1$ à la tension, $0.001$ au courant, $0.01$ à la fréquence !
   - **L'aberration physique induite** :
     - La tension réseau $230\text{ V}$ décodée en flottant était multipliée par $0.1$, devenant **$23.0\text{ V}$** !
     - Le courant $10\text{ A}$ décodé en flottant était multiplié par $0.001$, devenant **$0.01\text{ A}$ ($10\text{ mA}$)** !
     - La fréquence de $50\text{ Hz}$ était multipliée par $0.01$, devenant **$0.50\text{ Hz}$** (impossible sur le réseau européen interconnecté) !
     - Le facteur de puissance $0.99$ devenait **$0.00099$** !
   - **La réalité physique** : Dans les notices d'instruments de mesure chinois et internationaux, le terme « *Resolution* » désigne la **résolution métrologique ou d'affichage de l'écran LCD** (l'écran affiche 1 chiffre après la virgule pour les volts, 3 pour les ampères, 2 pour les hertz). Le nombre flottant transmis sur le bus Modbus est **strictement la grandeur physique réelle** !
   - **Règle d'or Apollo appliquée** : Tout registre de type `F32` exprimé dans l'unité physique déclarée (`V`, `A`, `Hz`, `kW`, etc.) **DOIT IMPÉRATIVEMENT AVOIR $\mathbf{CoefA = 1.000000}$ et $\mathbf{CoefB = 0.000000}$**.

---

### 3.5. Tableau de Synthèse : Règles d'Or Apollo pour CoefA

| Type de Donnée | Grandeur Physique | Valeur Transmise sur le Bus | Unité Déclarée CSV | CoefA Rigoureux | Justification Métrologique |
|:---:|:---:|:---:|:---:|:---:|---|
| **`F32`** | Tension ($U, V$) | Flottant réel (ex: $230.4$) | `V` | **`1.000000`** | Flottant direct sans altération |
| **`F32`** | Courant ($I$) | Flottant réel (ex: $5.213$) | `A` | **`1.000000`** | Flottant direct sans altération |
| **`F32`** | Fréquence ($F$) | Flottant réel (ex: $50.02$) | `Hz` | **`1.000000`** | Flottant direct sans altération |
| **`F32`** | Facteur de puissance ($PF$) | Flottant réel (ex: $0.985$) | (vide) | **`1.000000`** | Flottant direct sans altération |
| **`F32`** | Puissance ($P$) en kW | Flottant réel (ex: $15.5\text{ kW}$) | `kW` | **`1.000000`** | Flottant direct sans altération |
| **`F32`** | Puissance ($P$) en kW | Flottant réel (ex: $15.5\text{ kW}$) | `W` | **`1000.000000`** | Conversion métrique ($1\text{ kW} = 1000\text{ W}$) |
| **`U16` / `I16`** | Tension ($U, V$) | Entier en dixièmes (ex: $2304$) | `V` | **`0.100000`** | Pas de discrétisation $0.1\text{ V}$ |
| **`U16` / `I16`** | Courant ($I$) | Entier en centièmes (ex: $521$) | `A` | **`0.010000`** | Pas de discrétisation $0.01\text{ A}$ |
| **`I32` / `U32`** | Courant ($I$) | Entier en milliampères ($5213$) | `A` | **`0.001000`** | Pas de discrétisation $1\text{ mA}$ |
| **`U64`** | Énergie ($E$) | Entier en Watt-heures ($1\,500\,000$) | `kWh` | **`0.001000`** | Conversion métrique ($1\text{ Wh} = 10^{-3}\text{ kWh}$) |
| **`U32`** | Énergie ($E$) | Entier en kiloWatt-heures | `kWh` | **`1.000000`** | Comptage direct en kWh |

---

## 4. Matrice Complète des Définitions Générées et Certifiées

| N° | Fichier de Définition | Catégorie | Constructeur | Modèle | Nb Registres | Statut Apollo |
|:---:|---|---|---|---|:---:|:---:|
| 1 | `MET_Chint_DTSU666.csv` | Meter | CHINT | DTSU666 | 41 | **CERTIFIED_PASS** |
| 2 | `MET_Chint_DDSU666.csv` | Meter | CHINT | DDSU666 | 14 | **CERTIFIED_PASS** |
| 3 | `MET_Chint_DTSU666-H.csv` | Meter | CHINT | DTSU666-H | 22 | **CERTIFIED_PASS** |
| 4 | `INV_Huawei_SUN2000-M3.csv` | Inverter | Huawei | SUN2000-M3 | 64 | **CERTIFIED_PASS** |
| 5 | `INV_Huawei_SUN2000-Utility.csv` | Inverter | Huawei | SUN2000-Utility | 88 | **CERTIFIED_PASS** |
| 6 | `MET_Huawei_DTSU666-H.csv` | Meter | Huawei | DTSU666-H | 22 | **CERTIFIED_PASS** |
| 7 | `INV_SMA_SunnyTripower.csv` | Inverter | SMA | SunnyTripower | 59 | **CERTIFIED_PASS** |
| 8 | `INV_SMA_Core1_Core2.csv` | Inverter | SMA | Core1-Core2 | 77 | **CERTIFIED_PASS** |
| 9 | `INV_Solplanet_ASW_ThreePhase.csv` | Inverter | Solplanet | ASW_ThreePhase | 27 | **CERTIFIED_PASS** |
| 10 | `INV_Sungrow_SG_Series.csv` | Inverter | Sungrow | SG-Series | 29 | **CERTIFIED_PASS** |
| 11 | `INV_Goodwe_SMT_Series.csv` | Inverter | GoodWe | SMT-Series | 26 | **CERTIFIED_PASS** |
| 12 | `INV_Goodwe_ET_Series.csv` | Inverter | GoodWe | ET-Series | 33 | **CERTIFIED_PASS** |
| 13 | `INV_Growatt_MAX_Series.csv` | Inverter | Growatt | MAX-Series | 29 | **CERTIFIED_PASS** |
| 14 | `INV_Fronius_Symo_Primo_Gen24.csv` | Inverter | Fronius | Symo-Primo-Gen24 | 26 | **CERTIFIED_PASS** |
| 15 | `INV_SolarEdge_SE_Series.csv` | Inverter | SolarEdge | SE-Series | 28 | **CERTIFIED_PASS** |
| 16 | `MET_Socomec_Diris_A_Countis_E.csv` | Meter | Socomec | Diris-A-Countis-E | 30 | **CERTIFIED_PASS** |
| 17 | `MET_Schneider_PM5000_IEM3000.csv` | Meter | Schneider | PM5000-iEM3000 | 24 | **CERTIFIED_PASS** |
| 18 | `MET_CarloGavazzi_EM24_EM300.csv` | Meter | CarloGavazzi | EM24-EM300 | 27 | **CERTIFIED_PASS** |

---

## 5. Architecture de Déploiement & Synchronisation Redondante

Pour garantir la résilience et la disponibilité opérationnelle sur l'ensemble des modules logiciels et passerelles Webdyn, chaque fichier certifié est synchronisé à l'identique (copie byte-par-byte) sur quatre répertoires maîtres du dépôt :
1. `production/` : Bibliothèque centrale de déploiement en production pour les concentrateurs WebdynSunPM.
2. `Definitions/Classic/` : Répertoire canonique de référence pour les équipements classiques.
3. `generated_definitions/` : Dépôt miroir pour l'intégration continue et la validation de compilation.
4. `tagged_definitions/` : Dépôt miroir dédié aux moteurs de supervision et de télégestion SunPM.

Chaque fichier de définition a été vérifié par test unitaire automatisé (`test_classic_definitions_suite.py`) pour confirmer que les copies dans ces 4 répertoires sont **rigoureusement identiques au bit près**.

---

## 6. Synthèse des Tests & Résultats de Validation

- **Suite de tests dédiée** : `DefFileGenerator/tests/test_classic_definitions_suite.py` $\implies$ **7/7 PASSED (100%)**
  - Vérification de l'existence des fichiers dans tous les répertoires.
  - Vérification de la synchronisation byte-identique.
  - Vérification des 5 Piliers Apollo.
  - Vérification de l'en-tête `Modbus` et des 11 colonnes.
  - Vérification de l'indexation $1..N$ et des codes actions.
  - Vérification de la présence des tags canoniques métier.
  - Vérification hostile de l'échelle directe des flottants IEEE-754 F32 ($A = 1.000000$).
- **Validation Globale Production** : `scripts/validate_production.py` $\implies$ **1,141 / 1,141 fichiers PASSED (100%)**.
- **Bibliothèque Historique** : `DefFileGenerator/tests/test_definitions_library.py` $\implies$ **4/4 PASSED (100%)**.
- **Audit Trail Immuable** : Événement consigné avec succès dans `logs/audit_trail/audit_log.jsonl` et `docs/audit/EQUIPMENT_AUDIT_LOG.md`.

---

## 7. Recommandations pour l'Ingénierie Terrain (Field Engineering)

1. **Raccordement RS485 & Polarité** :
   - Pour les compteurs CHINT et Socomec, vérifier scrupuleusement la polarité du bus différentiel ($A = D+$, $B = D-$). Une inversion bloque la trame sans détruire le port.
   - Les résistances de terminaison de fin de ligne ($120\ \Omega$) doivent impérativement être activées sur le dernier équipement de la chaîne.
2. **Configuration des Paramètres de Liaison Série** :
   - CHINT DTSU666 : Par défaut `9600 bps`, `8N1`, adresse Modbus `1`.
   - SMA Sunny Tripower / Core : En Ethernet Modbus TCP, utiliser le port standard `502` et Unit ID `126` pour l'onduleur ou `1` pour le Data Manager M.
   - Huawei SUN2000 : Port standard `502` ou RS485 `9600 bps`, Unit ID `1` pour l'onduleur, Unit ID `11` ou `2` pour le compteur DTSU666-H raccordé au port RS485 secondaire de l'onduleur.
3. **Pilotage de Puissance & Régulation** :
   - Ne jamais injecter de consignes absolues (`CMD_P_AC_SET_ABS`) supérieures à la puissance crête $P_{max}$ déclarée aux registres de configuration de l'onduleur.
   - Utiliser de préférence la consigne relative (`CMD_P_AC_SET_PERCENT`) pour les algorithmes d'effacement ou de limitation d'injection réseau (Zero Feed-in).
