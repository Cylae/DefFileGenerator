===============================================================================
FOXESS P3 PRO & H3 PRO - FICHIERS DE DEFINITIONS COMMUNS WEBDYN SUNPM
===============================================================================
Constructeur : FoxESS
Gammes       : P3 Pro & H3 Pro (Protocole Modbus Unifie V1.05.05.00)
Protocole    : Modbus RTU / TCP (FC03 Read Holding Registers, FC06/16 Writes)
Statut       : 100% Conforme au standard Webdyn (11 colonnes, separateur point-virgule)

Ce dossier regroupe les fichiers de definition Modbus COMMUNS et universels pour les
gammes FoxESS P3 Pro et H3 Pro.

1. FICHIERS FOURNIS DANS CE DOSSIER :
-------------------------------------------------------------------------------
1. FOXESS_P3_H3_PRO_ALL_IN_ONE.csv (1 233 registres) - DEFINITION TOUT-EN-UN :
    * Regroupe l'INTEGRALITE de la station en un seul fichier de definition :
      - Onduleur PV : 24 chaines MPPT U_DC_MPPT_1..24 & I_DC_MPPT_1..24, puissance active
        P_AC (39134), reactive Q_AC (39136), tension reseau U_AC_1..3, frequence FREQ,
        consignes de bridage et commandes CMD_ON/OFF/SET/LIM.
      - Compteur de livraison reseau (Meter 1 / CT1) : telemetrie triphasee complete U_AC,
        I_AC, P_AC, Q_AC, S_AC, PF, FREQ et alerte perte compteur.
      - Systeme de batterie haute tension BMS & EMS : SOC (37612), SOH (37624), tensions
        U_DC/I_DC, temperatures, extrêmes de cellules, 6 alarmes BMS et 96 plages horaires.
    * 0 conflit d'adresses, adresses Modbus strictement croissantes, 100% valide Webdyn.
    * Usage recommande : lorsqu'un seul equipement logique Webdyn interroge l'onduleur
      qui centralise l'onduleur, le compteur de point de livraison et la batterie.

2. FOXESS_P3_H3_PRO_INVERTER.csv (285 registres) :
    * Definition commune Onduleur (Modele : P3 Pro / H3 Pro).
    * Utilise pour declarer l'onduleur comme equipement distinct sur WebdynSunPM.

3. FOXESS_P3_H3_PRO_METER.csv (61 registres) :
    * Definition commune Compteur (Modele : P3 Pro / H3 Pro).
    * Utilise pour declarer le compteur reseau (Meter 1 / CT1) comme equipement distinct.

4. FOXESS_P3_H3_PRO_BATTERY.csv (893 registres) :
    * Definition commune Batterie (Modele : P3 Pro / H3 Pro).
    * Utilise pour declarer la batterie et le gestionnaire EMS comme equipement distinct.

5. FoxESS_P3_H3_PRO.csv (285 registres) :
    * Alias historique Inverter pour compatibilite descendante des configurations existantes.

2. CHOIX D'INTEGRATION SUR WEBDYNSUNPM :
-------------------------------------------------------------------------------
- Configuration multi-equipements (recommandee pour affichage separe sur le portail) :
    * Equipement 1 : Inverter -> FOXESS_P3_H3_PRO_INVERTER.csv
    * Equipement 2 : Meter    -> FOXESS_P3_H3_PRO_METER.csv
    * Equipement 3 : Battery  -> FOXESS_P3_H3_PRO_BATTERY.csv
- Configuration tout-en-un (recommandee pour supervision centralisee sur 1 seul ID Modbus) :
    * Equipement unique       -> FOXESS_P3_H3_PRO_ALL_IN_ONE.csv
===============================================================================
