MARS RENEWABLE - DEFINITIONS LOCALES (LOCAL CONTROLLER V7.0)
============================================================
Fichiers generes conformement aux specifications WebdynSunPM :

1. MarsRenewable_LC-AC.csv (3 400 registres) :
   - Regroupe l'integralite des 100 unites d'Air Conditioning (15000..19999, pas de 50, 1600 reg.)
     et des 100 unites de Liquid Cooling (30000..33999, pas de 40, 1800 reg.)
   - 0 conflit, tri strict par adresse Modbus croissante, 100% valide.
   - Version monoposte disponible : MarsRenewable_LC-AC_1Unit.csv (34 registres).

2. MarsRenewable_LC.csv / MarsRenewable_LiquidCooling.csv (1 800 registres) :
   - Les 100 unites de refroidissement liquide (LiquidCooling 1 a 100).
   - Version monoposte disponible : MarsRenewable_LC_1Unit.csv (18 registres).

3. MarsRenewable_AC.csv / MarsRenewable_AirCooling.csv (1 600 registres) :
   - Les 100 unites de climatisation / air cooling (AC 1 a 100).
   - Version monoposte disponible : MarsRenewable_AC_1Unit.csv (16 registres).

4. MarsRenewable_LocalControler.csv (12 registres) :
   - Modbus Slave ID : 247
   - Overview FC4 + Consignes EMS FC3/16.

5. MarsRenewable_PCS-Bank-Rack.csv (80 registres) :
   - Modbus Slave ID : 1
   - Alarmes FC2 + DIO 1 + PCS 1 (et 6 branches DC) + Bank 1 batterie.

6. MarsRenewable_Rack-Cell.csv (35 registres) :
   - Modbus Slave ID : 2 (ou N+1)
   - Rack 1 (30 reg.) + Cellule 1 de reference (5 reg.).

7. MarsRenewable_Meter.csv (11 registres) :
   - Modbus Slave ID : 1
   - Compteur bidirectionnel 1 (E_IMP, E_EXP, P_AC).

8. MarsRenewable_PV.csv (49 registres) :
   - Modbus Slave ID : 1
   - Onduleur PV 1 triphase + 6 chaines MPPT DC.
