# Sécurité et intégrité des exports

## Injection CSV

Une cellule CSV ouverte dans un tableur peut être interprétée comme formule lorsque son premier caractère significatif est un déclencheur. Le tableur retire les espaces de tête (espace, tabulation, CR, LF, NBSP) *avant* d'évaluer la cellule, et normalise les variantes Unicode pleine chasse : un contrôle limité au premier octet brut est donc contournable.

Jeu de déclencheurs traité par `sanitize_csv_field` :

| Catégorie | Caractères |
|--|--|
| ASCII | `=` `+` `-` `@` `\|` |
| Espaces de tête | `0x20` `0x09` (tab) `0x0D` (CR) `0x0A` (LF) `0x0B` `0x0C` `U+00A0` (NBSP) |
| Pleine chasse | `＝` `＋` `－` `＠` |

Toute valeur dont le premier caractère significatif appartient à ce jeu est préfixée d'une apostrophe (recommandation OWASP). Les nombres signés finis (`-10.5`, `+25`, `1.5e3`) conservent leur signe afin de préserver leur sens métier ; les littéraux non finis (`-inf`, `+nan`) et les formes non numériques (` -10.5`, `-1_000`) sont échappés.

Traitez toute documentation fournisseur comme donnée non fiable. Vérifiez les fichiers produits avant import dans un équipement ou une chaîne de configuration.

## XML

Les parseurs XML doivent rester protégés contre les entités externes (XXE). Utilisez les parseurs sûrs configurés par le projet et n’introduisez pas de parseur XML standard non durci. Le fichier `xxe.xml` est une fixture de régression : son traitement doit échouer sans lecture de ressource externe.

## Écriture de fichiers

Les futurs changements de génération doivent écrire vers un fichier temporaire dans le même volume puis remplacer la cible de façon atomique. Ne laissez jamais un CSV partiellement généré remplacer un fichier de production valide.

## Secrets

Aucun mot de passe, jeton, export client ou donnée sensible ne doit être ajouté aux fixtures. Utilisez des valeurs synthétiques et effectuez une analyse de secrets avant chaque publication.
