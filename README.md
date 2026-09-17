# UNYC Mobile RIO Extractor

Outil Python d'automatisation destiné à enrichir un fichier Excel de lignes mobiles avec les codes **RIO** récupérés depuis l'espace **UNYC Atlas**.

Le script lit un tableau Excel contenant des clients et des numéros de mobile, recherche chaque client dans UNYC Atlas, ouvre la partie téléphonie, retrouve la ligne mobile correspondante, lit le RIO dans l'onglet contrat puis écrit le résultat dans le fichier Excel.

> Cet outil doit être utilisé uniquement avec un compte UNYC autorisé et sur des données que vous êtes habilité à consulter.

## Principe

Flux général :

```text
Excel
  -> Client + Numéro
  -> connexion UNYC Atlas
  -> recherche du client
  -> Téléphonie
  -> Lignes mobiles
  -> ligne correspondant au numéro
  -> onglet Contrat
  -> lecture du RIO
  -> écriture du RIO dans Excel
```

L'outil ne tente pas de récupérer toutes les lignes présentes sur un compte client. Il traite uniquement les numéros fournis dans le fichier Excel, ce qui permet de conserver une correspondance explicite entre les données source et les résultats.

## Fonctionnalités

- connexion automatisée à UNYC Atlas avec Playwright ;
- prise en charge d'une authentification 2FA nécessitant une validation manuelle ;
- possibilité de réutiliser une session Playwright locale ;
- lecture d'un fichier Excel `.xlsx` ;
- regroupement et traitement des lignes par client ;
- recherche des numéros mobiles dans la fiche client ;
- navigation vers `Lignes mobiles` puis `Contrat` ;
- extraction du code RIO ;
- normalisation des numéros afin d'éviter les écarts de format Excel/HTML ;
- écriture du RIO sur la ligne Excel correspondante ;
- résumé de traitement en fin d'exécution ;
- configuration par variables d'environnement afin de ne pas stocker les identifiants dans le code.

## Prérequis

- Windows, Linux ou macOS ;
- Python **3.11 ou supérieur** ;
- un accès valide à UNYC Atlas ;
- Chromium installé via Playwright ;
- un fichier Excel au format attendu.

## Installation

Créer idéalement un environnement virtuel :

```bash
python -m venv .venv
```

Sous Windows :

```powershell
.\.venv\Scripts\Activate.ps1
```

Installer ensuite les dépendances :

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
```

Le projet contient aussi un script d'installation :

```bash
python setup.py
```

## Configuration

Copier `.env.example` vers `.env` :

```text
UNYC_USERNAME=
UNYC_PASSWORD=
UNYC_BASE_URL=https://atlas.unyc.io/
UNYC_CLIENT_URL=https://atlas.unyc.io/client
UNYC_EXCEL=lmunyc.xlsx
UNYC_HEADLESS=false
UNYC_SLOW_MO_MS=200
UNYC_STORAGE_STATE=unyc_state.json
```

### Variables disponibles

| Variable | Description | Valeur par défaut |
|---|---|---|
| `UNYC_USERNAME` | identifiant UNYC Atlas | aucune |
| `UNYC_PASSWORD` | mot de passe UNYC Atlas | aucune |
| `UNYC_BASE_URL` | URL principale Atlas | `https://atlas.unyc.io/` |
| `UNYC_CLIENT_URL` | URL de recherche clients | `https://atlas.unyc.io/client` |
| `UNYC_EXCEL` | fichier Excel à traiter | `lmunyc.xlsx` |
| `UNYC_HEADLESS` | navigateur sans interface graphique | `false` |
| `UNYC_SLOW_MO_MS` | délai Playwright entre actions | `200` |
| `UNYC_STORAGE_STATE` | fichier local de session | `unyc_state.json` |

Si les identifiants ne sont pas définis dans `.env`, le script les demande au démarrage. Le mot de passe est saisi de façon masquée.

## Format Excel

Par défaut, le script lit :

- fichier : `lmunyc.xlsx` ;
- feuille : `Feuil1`.

Les colonnes utilisées sont :

| Colonne | Obligatoire | Description |
|---|---:|---|
| `Client` | oui | nom du client tel qu'il peut être recherché dans Atlas |
| `Numéro` | oui pour extraire un RIO | numéro mobile à retrouver |
| `RIO` | non | créée automatiquement si elle n'existe pas |

Exemple :

| Client | Numéro | RIO |
|---|---|---|
| SOCIETE EXEMPLE | 0612345678 | |
| SOCIETE EXEMPLE | 0698765432 | |
| AUTRE SOCIETE | 0600000000 | |

Après traitement :

| Client | Numéro | RIO |
|---|---|---|
| SOCIETE EXEMPLE | 0612345678 | 12CARACTERES |
| SOCIETE EXEMPLE | 0698765432 | 12CARACTERES |

Les numéros sont normalisés avant comparaison afin de limiter les problèmes liés aux espaces, aux valeurs Excel interprétées comme nombres ou aux formats d'affichage différents.

## Utilisation

```bash
python unyc_automation.py
```

Déroulement :

1. chargement des variables d'environnement ;
2. ouverture de Chromium ;
3. connexion à UNYC Atlas ;
4. validation manuelle du 2FA si demandé ;
5. chargement du fichier Excel ;
6. recherche de chaque client ;
7. ouverture de la fiche client ;
8. lecture des lignes de téléphonie ;
9. recherche du numéro demandé ;
10. ouverture de `Lignes mobiles` puis `Contrat` ;
11. récupération du RIO ;
12. écriture du résultat dans Excel ;
13. affichage d'un récapitulatif.

## Gestion du 2FA

Lorsque UNYC demande une authentification à deux facteurs, le navigateur reste visible et le script attend que l'utilisateur termine l'authentification. Une fois la page authentifiée détectée, le traitement reprend automatiquement.

## Session Playwright

Le fichier `unyc_state.json` peut contenir des cookies et informations de session permettant de réutiliser une authentification existante.

**Ne partagez jamais ce fichier et ne le versionnez pas.**

Il est exclu par `.gitignore`.

## Sécurité

Les éléments suivants ne doivent pas être publiés :

- `.env` ;
- `unyc_state.json` ;
- fichiers Excel de production ;
- exports contenant des numéros ou RIO ;
- identifiants, mots de passe ou cookies de session.

Le dépôt fournit `.env.example` uniquement comme modèle de configuration.

## Structure du projet

```text
Unyc extract bot/
├─ unyc_automation.py       # automatisation principale
├─ requirements.txt         # dépendances Python
├─ setup.py                 # installation simplifiée
├─ .env.example             # modèle de configuration
├─ .gitignore               # secrets, sessions et données locales
├─ Excel_File_Structure.md  # rappel du format Excel
├─ LICENSE                  # licence MIT
└─ README.md
```

Les fichiers `.env`, `unyc_state.json` et les feuilles Excel de production sont volontairement absents de cette structure publique.

## Dépannage

### Chromium n'est pas installé

```bash
python -m playwright install chromium
```

### Le client n'est pas trouvé

Vérifier que la valeur `Client` correspond suffisamment au libellé visible dans UNYC Atlas.

### Le numéro n'est pas trouvé

Vérifier :

- qu'il appartient bien au client ;
- qu'il est présent dans la colonne `Numéro` ;
- qu'il s'agit d'une ligne mobile visible dans Atlas.

### Le RIO reste vide

Le portail UNYC peut évoluer. Les sélecteurs HTML utilisés par Playwright peuvent alors nécessiter une adaptation dans `unyc_automation.py`.

## Dépendances principales

- Playwright : navigation et automatisation du portail ;
- pandas : lecture et manipulation du tableau ;
- openpyxl : moteur Excel `.xlsx` ;
- python-dotenv : chargement de la configuration `.env`.

## Licence

Ce projet est distribué sous licence **MIT**. Voir [`LICENSE`](LICENSE).

La licence couvre le code de ce dépôt. Elle ne donne aucun droit particulier sur UNYC Atlas, les comptes utilisateurs ou les données consultées via le portail.
