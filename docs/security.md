# Sécurité et intégrité des exports

## Injection CSV

Une cellule CSV ouverte dans un tableur peut être interprétée comme formule lorsqu’elle commence par `=`, `+`, `-` ou `@`. Les valeurs textuelles de cette forme sont échappées avant export. Les nombres signés valides doivent conserver leur signe afin de préserver leur sens métier.

Traitez toute documentation fournisseur comme donnée non fiable. Vérifiez les fichiers produits avant import dans un équipement ou une chaîne de configuration.

## XML

Les parseurs XML doivent rester protégés contre les entités externes (XXE). Utilisez les parseurs sûrs configurés par le projet et n’introduisez pas de parseur XML standard non durci. Le fichier `xxe.xml` est une fixture de régression : son traitement doit échouer sans lecture de ressource externe.

## Écriture de fichiers

Les futurs changements de génération doivent écrire vers un fichier temporaire dans le même volume puis remplacer la cible de façon atomique. Ne laissez jamais un CSV partiellement généré remplacer un fichier de production valide.

## Secrets

Aucun mot de passe, jeton, export client ou donnée sensible ne doit être ajouté aux fixtures. Utilisez des valeurs synthétiques et effectuez une analyse de secrets avant chaque publication.
