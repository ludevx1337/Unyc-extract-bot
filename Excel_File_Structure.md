# Structure du fichier Excel

Le scraper utilise un fichier Excel comme liste de travail. Chaque ligne identifie un client et un numéro mobile à contrôler dans UNYC Atlas.

## Colonnes attendues

| Colonne | Obligatoire | Rôle |
|---|---:|---|
| `Client` | oui | nom du client à rechercher dans Atlas |
| `Numéro` | oui | numéro mobile dont le RIO doit être récupéré |
| `RIO` | non | résultat ; créée automatiquement si absente |

## Exemple

| Client | Numéro | RIO |
|---|---|---|
| SOCIETE EXEMPLE | 0612345678 | |
| SOCIETE EXEMPLE | 0698765432 | |
| AUTRE SOCIETE | 0600000000 | |

Le fichier utilisé par défaut est `lmunyc.xlsx`, feuille `Feuil1`. Ces valeurs peuvent être adaptées avec les variables d'environnement décrites dans le `README.md`.

## Fonctionnement

Le script :

1. lit les couples `Client` + `Numéro` ;
2. regroupe les recherches par client ;
3. ouvre la fiche du client dans UNYC Atlas ;
4. accède à la partie téléphonie ;
5. retrouve la ligne mobile correspondant au numéro demandé ;
6. ouvre `Lignes mobiles` puis l'onglet `Contrat` ;
7. récupère le RIO ;
8. écrit le résultat dans la colonne `RIO` de la ligne correspondante.

Le scraper ne collecte pas automatiquement tous les numéros visibles dans Atlas : il traite uniquement les numéros explicitement présents dans le fichier Excel.

## Format des numéros

La comparaison est effectuée après normalisation du numéro. Les espaces et certains artefacts d'Excel sont donc supprimés avant la recherche.

Pour éviter les ambiguïtés, il est recommandé de conserver les numéros au format national français sur 10 chiffres, par exemple :

```text
0612345678
```

## Données sensibles

Les fichiers Excel de production peuvent contenir des numéros de téléphone et des RIO. Ils sont exclus du dépôt par `.gitignore` et ne doivent pas être publiés avec le code source.
