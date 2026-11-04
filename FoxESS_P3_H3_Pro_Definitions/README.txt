===============================================================================
FOXESS GAMMES P3 PRO & H3 PRO - DEFINITIONS WEBDYN SUNPM
===============================================================================
Ce dossier regroupe l'ensemble des fichiers de definition WebdynSunPM pour les onduleurs
commerciaux et hybrides FoxESS des gammes P3 Pro et H3 Pro.

1. STRUCTURE DU DOSSIER :
-------------------------------------------------------------------------------
* Sous-dossier Commun/ (Fichiers communs aux deux gammes P3 Pro & H3 Pro) :
    - FOXESS_P3_H3_PRO_ALL_IN_ONE.csv (1 233 registres : Onduleur + Compteur + Batterie tout-en-un)
    - FOXESS_P3_H3_PRO_INVERTER.csv   (285 registres : Onduleur commercial)
    - FOXESS_P3_H3_PRO_METER.csv      (61 registres : Compteur de livraison)
    - FOXESS_P3_H3_PRO_BATTERY.csv    (893 registres : Batterie haute tension & EMS)
    - FoxESS_P3_H3_PRO.csv            (Alias historique Inverter)
    - README.txt                      (Notice dediee aux definitions communes)

* Sous-dossier P3_Pro/ (Cas specifique Gamme P3 Pro) :
    - FOXESS_P3_PRO_INVERTER.csv (Modele : P3 Pro)
    - FOXESS_P3_PRO_METER.csv    (Modele : P3 Pro)
    - FOXESS_P3_PRO_BATTERY.csv  (Modele : P3 Pro)
    - README.txt                 (Documentation dediee P3 Pro)

* Sous-dossier H3_Pro/ (Cas specifique Gamme H3 Pro) :
    - FOXESS_H3_PRO_INVERTER.csv (Modele : H3 Pro)
    - FOXESS_H3_PRO_METER.csv    (Modele : H3 Pro)
    - FOXESS_H3_PRO_BATTERY.csv  (Modele : H3 Pro)
    - README.txt                 (Documentation dediee H3 Pro)

* Fichiers racine :
    - FOXESS_P3_H3_PRO_ALL_IN_ONE.csv (Tout-en-un 1 233 registres)
    - FOXESS_P3_H3_PRO_INVERTER.csv, METER.csv, BATTERY.csv
    - FOXESS_P3_PRO_*.csv, FOXESS_H3_PRO_*.csv
    - README.txt, README.md
    - RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.md
    - RAPPORT_TECHNIQUE_FOXESS_P3_H3_PRO.json

2. REPARTITION DES REGISTRES MODBUS :
-------------------------------------------------------------------------------
- Définition Tout-en-Un (ALL-IN-ONE) : 1 233 registres Modbus FC03 Holding.
    * Regroupe Onduleur + Compteur + Batterie sans aucun conflit d'adresses.
    * 127 Tags SunPM canoniques assignes.
- Onduleur (Inverter) : 285 registres Modbus FC03 Holding (111 tags, 43 consignes, 24 alarmes).
- Compteur (Meter) : 61 registres Modbus FC03 Holding (27 tags, 3 commandes, 1 alarme).
- Batterie (Battery) : 893 registres Modbus FC03 Holding (18 tags, 692 consignes, 6 alarmes).

3. CONFORMITE WEBDYN :
-------------------------------------------------------------------------------
- Format : 11 colonnes separees par des points-virgules (;).
- Colonne 5 (ScaleFactor) : strictement vide.
- Facteurs Ca / Cb : formatage a 6 decimales obligatoires (ex: 1.000000, 0.001000).
- Adresses Modbus strictement croissantes, 0 doublon de tag canonique unique.
===============================================================================
