# Format des entrées

Le générateur accepte des cartes de registres issues de PDF, XLSX, CSV ou XML. La détection des colonnes est heuristique ; pour un résultat reproductible, fournissez des en-têtes explicites ou un mapping JSON.

## Colonnes recommandées

| Champ | Obligatoire | Exemples |
|---|---:|---|
| Address | Oui | `40001`, `0x9C41`, `40001_0_8` |
| Name | Oui | `Active Power` |
| Type | Oui | `U16`, `I32`, `F32`, `STR16`, `BITS` |
| RegisterType | Recommandé | Holding Register, Input Register |
| Unit | Non | W, V, A, °C |
| Factor | Non | `0.1`, `1/10` |
| Offset | Non | `0` |
| ScaleFactor | Non | `-1`, `0`, `2` |
| Action | Non | Read, Read/Write |

## Mapping explicite

```json
{
  "Address": "Modbus_Addr",
  "Name": "Signal_Description",
  "Type": "Format_Code",
  "Unit": "Engineering_Unit"
}
```

Passez ce fichier à la commande d’extraction avec `--mapping`. Conservez un mapping par fournisseur dans le dépôt du projet intégrateur, accompagné d’une fixture de test.

## Règles de qualité

Les adresses doivent être interprétables et ne pas se chevaucher dans le même espace de registres. Les tags générés doivent rester uniques. Toute correction manuelle doit être rejouable par mapping, jamais seulement appliquée au CSV de sortie.
