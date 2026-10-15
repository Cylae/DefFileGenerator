MARS RENEWABLE - DEFINITIONS LOCALES (LOCAL CONTROLLER V7.0)
============================================================
Fichiers generes conformement aux specifications WebdynSunPM et aux instructions :

1. MarsRenewable_LocalControler.csv :
   - Modbus Slave ID : 247
   - Onglet LC Information (LC Overview FC4 + Consignes EMS FC3/16)
   - 12 registres, 0 erreur

2. MarsRenewable_PCS-Bank-Rack.csv :
   - Modbus Slave ID : 1
   - Onglets PCS Information + Bank Information (+ Alarmes FC2 & DIO 1)
   - 80 registres, 0 erreur

3. MarsRenewable_Rack-Cell.csv :
   - Modbus Slave ID : 2 (ou N+1)
   - Onglet Rack and Cell Information (Rack 1 + Cellule 1 de reference)
   - 35 registres, 0 erreur

4. MarsRenewable_LC-AC.csv :
   - Modbus Slave ID : 1
   - Onglets Liquid Cooling Information (18 reg.) + Air Conditioning Information (16 reg.)
   - 34 registres, 0 erreur

5. MarsRenewable_Meter.csv :
   - Modbus Slave ID : 1
   - Onglet Meter Information (Compteur bidirectionnel 1)
   - 11 registres, 0 erreur

6. MarsRenewable_PV.csv :
   - Modbus Slave ID : 1
   - Onglet PV Information (Onduleur PV triphase + 6 chaines MPPT DC)
   - 49 registres, 0 erreur

Rapport technique complet disponible dans : RAPPORT_TECHNIQUE_MARS_RENEWABLE.md
