Mail de Serge Communal:
Les onduleurs sont configurés généralement par défaut pour recevoir une consigne externe. Il suffit d’activer le serveur Modbus lors de la configuration.

Dans le cas du STP X, du fait de sa fonction System Manager, il est nécessaire de le configurer au moment de sa mise en service ou à posteriori, dans
les réglages Système de gestion du réseau -> Puissance active (effective sur l'interface de l'onduleur?) -> Pilotage, consigne externe Modbus.

Si vous avez plusieurs STP X sur une même installation, un seul est Maitre (System Manager) et reçoit 
la consigne et la transmet aux autres, configurés en Esclaves (Subordonnés).

Pour le pilotage des onduleurs, SMA utilise l’Unit ID 3 et le registre 40016(NON!!). (Pour le registre les tests montrent que c'est le registre 40210 qu'il faut utiliser) (Tests fait avec CEME)



Procédure Webdyn :
Pour la régulation de puissance active :
Par fichier de commande -> Ecrire la valeur 1078 sur le registre 40210.
Puis activation du script avec le tag cmdPwrPercent sur le registre 40214