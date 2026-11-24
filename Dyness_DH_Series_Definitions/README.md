# 🔋 Dyness BESS — Bundle de Définitions Dédié Série DH

Ce dossier regroupe le bundle officiel de livraison des fichiers de définition WebdynSunPM pour la gamme de stockage par batterie industrielle **Dyness BESS Série DH**.

---

## 📦 Contenu du Bundle

| Fichier | Modèle Couvert | Description Technique | Document Source |
|---|---|---|---|
| `BAT_Dyness_DH100F.csv` | Dyness DH100F | Système de stockage conteneurisé C&I 100 kW / 200 kWh | Manuel V2.2 officiel |
| `BAT_Dyness_DH200F.csv` | Dyness DH200F | Système de stockage conteneurisé C&I 200 kW / 200-400 kWh | Manuel V2.2 officiel |
| `BAT_Dyness_DH200Y.csv` | Dyness DH200Y | Armoire de stockage extérieure haute puissance | Manuel V2.2 officiel |
| `BAT_Dyness_GENERIC.csv` | Gamme Générique DH | Profil maître couvrant tous les registres communs BMS & PCS | Manuel V2.2 officiel |

---

## 🛡️ Conformité & Synchronisation

Tous les fichiers de ce dossier sont :
- 100% conformes aux **5 Piliers Apollo** (11 colonnes, UTF-8 sans BOM, CRLF, index 1..N, tags SunPM uniques).
- Rigoureusement synchronisés bit-à-bit avec [`production/`](../production/).
- Documentés dans [`sources/Dyness/`](../sources/Dyness/) avec leur protocole source Excel original.
