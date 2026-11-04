===============================================================================
FOXESS GAMME P3 PRO - DEFINITIONS WEBDYN SUNPM
===============================================================================
Constructeur : FoxESS
Gamme        : P3 Pro Series (Onduleurs PV Commerciaux Triphases / Hybrides)
Protocole    : Modbus RTU / TCP (FC03 Read Holding Registers, FC06/16 Writes)
Statut       : 100% Conforme au standard Webdyn (11 colonnes, separateur point-virgule)

1. FICHIERS FOURNIS DANS CE DOSSIER :
-------------------------------------------------------------------------------
- FOXESS_P3_PRO_INVERTER.csv (285 registres) :
    * En-tete Webdyn : Modbus;Inverter;FoxESS;P3 Pro;;;;;;;
    * Identification : MODEL (30000), SERIAL (30016), MANUFACTURER (30032), P_RATED (39053, kW)
    * Production PV DC : P_DC (39118, kW), 24 chaines MPPT U_DC_MPPT_1..24 & I_DC_MPPT_1..24 (39070..39117), puissances MPPT 1..24 (39329..39421)
    * Reseau AC Triphase : U_AC_1..3 (39123..39125), I_AC_1..3 (39126..39130), P_AC (39134, kW), Q_AC (39136, kVar), S_AC (39270, VA), PF (39138), FREQ (39139, Hz)
    * Mesures par phase : P_AC_1..3, Q_AC_1..3, S_AC_1..3, PF_1..3
    * Compteurs d'energie : E_TOT (39149, kWh), E_DAY (39151, kWh), Temp. interne TEMP_INT (39141, C)
    * Pilotage et consignes (Action = 1) :
        - Allumage / Extinction : CMD_ON (49077), CMD_OFF (49078)
        - Puissance active absolue : CMD_P_AC_SET_ABS (46003, 49008)
        - Puissance active pourcentage : CMD_P_AC_SET_PERCENT (49007)
        - Limitation d'export / 0-injection : CMD_P_AC_LIM_ABS (46616)
        - Consigne reactive : CMD_Q_AC_SET_ABS (46005)
    * Alarmes et etats (Action = 9) :
        - Mots d'alarmes : ALARM_1 (39067), ALARM_2 (39068), ALARM_3 (39069)
        - Etat onduleur : STATE (39063_2_1), Defaut STATE_FAULT (39063_6_1)

- FOXESS_P3_PRO_METER.csv (61 registres) :
    * En-tete Webdyn : Modbus;Meter;FoxESS;P3 Pro;;;;;;;
    * Compteur de point de livraison reseau (Meter 1 / CT1).
    * Identification : SERIAL (36100), MANUFACTURER (36116), MODEL (36132)
    * Telemetrie triphasee : U_AC_1..3, I_AC_1..3, P_AC, P_AC_1..3, Q_AC, Q_AC_1..3, S_AC, S_AC_1..3, PF, PF_1..3, FREQ
    * Alerte perte de compteur : ALARM (39069_9_1, Action = 9)
    * Configuration et compensation : 49207, 49208, 49248 (Action = 1)

- FOXESS_P3_PRO_BATTERY.csv (893 registres) :
    * En-tete Webdyn : Modbus;Battery;FoxESS;P3 Pro;;;;;;;
    * Gestion de stockage et systeme BMS / EMS.
    * Identification : MODEL (30000), SERIAL (30016), MANUFACTURER (30032)
    * Telemetrie BMS : U_DC (37609, V), I_DC (37610, A), TEMP (37611, C), SOC (37612, %), SOH (37624, %), TEMP_MAX (37617), TEMP_MIN (37618), U_CELL_MAX (37619, mV), U_CELL_MIN (37620, mV)
    * Alarmes BMS : ALARM_1 a ALARM_6 (37626 a 37631, Action = 9)
    * 692 Consignes et plages EMS (Action = 1) :
        - 96 Creneaux de planification charge/decharge reseau (registres 48000 a 48966)
        - Prechauffage batterie BMS Warm-Up (registres 53400 a 53411)
        - Limites de courant et seuils SOC de charge/decharge (46503, 46607..46610)
===============================================================================
