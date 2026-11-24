# 🏷️ Sas de Qualification Sémantique SunPM (Tagged Definitions Staging)

Ce répertoire héberge les définitions ayant subi le traitement d'**enrichissement sémantique SunPM**.

---

## 🎯 Rôle dans l'Architecture

Les variables brutes extraites des manuels constructeurs possèdent souvent des libellés hétérogènes (`Active Power`, `W`, `PAC`, `Puissance Active`, `Wirkleistung`). 

Le module [`knowledge_miner.py`](../DefFileGenerator/knowledge_miner.py) analyse ces libellés à l'aide d'un dictionnaire sémantique entraîné sur plus de **1 280 000 registres** pour leur attribuer les **tags SunPM normalisés** de supervision Webdyn :
- `P_AC` / `P_AC_1` / `P_AC_2` / `P_AC_3` (Puissances actives)
- `Q_AC` / `S_AC` (Puissances réactives et apparentes)
- `E_TOT` / `E_DAY` / `E_YEAR` (Énergies produites)
- `U_AC_1` / `U_AC_2` / `U_AC_3` / `U_AC_L1L2` (Tensions réseau)
- `I_AC_1` / `I_AC_2` / `I_AC_3` (Courants réseau)
- `FREQ` (Fréquence réseau)
- `TEMP_INT` / `TEMP_EXT` (Températures)
- `SOC` / `SOH` / `U_DC` / `I_DC` (Grandeurs de stockage BESS)

---

## 🔄 Flux de Données

1. **Entrée :** Fichiers bruts provenant de [`generated_definitions/`](../generated_definitions/).
2. **Traitement :** Inférence sémantique, résolution des conflits de tags, calcul des facteurs d'échelle (`DIV10`, `DIV100`, etc.).
3. **Sortie & Promotion :** Intégration et validation stricte dans [`production/`](../production/) par `scripts/build_production_suite.py`.
