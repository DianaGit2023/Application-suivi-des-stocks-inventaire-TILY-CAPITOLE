# Inventaire du matériel scout

Application Excel/VBA de gestion du matériel d'un groupe scout : inventaire historisé, import de listes, bordereaux de sortie en PDF et suivi des retours.

## Fonctionnalités

- **Inventaire (CRUD)** : créer, consulter, modifier et supprimer un objet, avec recherche.
- **Historisation** : aucune donnée n'est effacée ; chaque modification conserve l'ancienne version.
- **Import** : notes de téléphone collées en texte libre ou fichier CSV/TXT, avec contrôle avant intégration.
- **Bordereau de matériel** : préparation depuis l'inventaire, numérotation automatique (`BOD001`, `BOD002`…), export PDF avec zones de signature.
- **Suivi des sorties et retours** : journal des prêts, quantité disponible calculée, indicateurs sur l'accueil.
- **Protection** : feuille de données masquée, accessible par mot de passe.

## Prérequis

- Excel pour Windows (version bureau), macros activées.
- Les macros ne fonctionnent pas sur Excel mobile ni sur Excel en ligne. Excel pour Mac n'est pas pris en charge (import de fichier et analyse des notes indisponibles).

## Contenu du projet

| Fichier | Rôle |
| --- | --- |
| `Inventaire_Scout.xlsm` | Le classeur (livré en `.xlsx`, à enregistrer en `.xlsm` après installation) |
| `InventaireScout.bas` | Module principal : inventaire, historisation, import, pré-import |
| `BordereauScout.bas` | Bordereaux, journal des sorties, retours, indicateurs |
| `ProtectionData.bas` | Masquage de la feuille Data par mot de passe |
| `import_liste_scout.csv` | Exemple de fichier d'import |

## Installation

1. Ouvrir le classeur, puis l'éditeur VBA (Alt + F11, ou onglet Développeur > Visual Basic).
2. Fichier > Importer un fichier : importer les trois modules `.bas`.
3. Exécuter les macros d'installation, dans cet ordre :

| Ordre | Macro | Effet |
| --- | --- | --- |
| 1 | `InstallerApplication` | Pose les boutons de l'accueil, de l'inventaire et du pré-import |
| 2 | `InstallerBordereau` | Crée les feuilles Bordereau et Prets, leurs boutons et les indicateurs |
| 3 | `InstallerProtectionData` | Masque la feuille Data et ajoute le bouton d'accès |

4. Enregistrer au format **Classeur Excel prenant en charge les macros (\*.xlsm)**.

Attention : relancer `InstallerBordereau` reconstruit la feuille Bordereau et efface le bordereau en cours de préparation (le journal Prets est conservé). Pour réinstaller uniquement les indicateurs, utiliser `LancerInstallationIndicateurs`.

### Mettre à jour un module

1. Dans l'éditeur, clic droit sur le module > Supprimer (répondre Non à l'export).
2. Importer la nouvelle version du fichier `.bas`.
3. Ne pas relancer l'installation, sauf indication contraire : les boutons existants restent liés aux macros.

## Structure du classeur

| Feuille | Rôle |
| --- | --- |
| Accueil | Formulaire de saisie, zone de notes, import, indicateurs |
| Inventaire | Vue des lignes actives, reconstruite par macro |
| PreImport | Sas de contrôle avant intégration |
| Bordereau | Bordereau en cours (colonnes A à E imprimées) |
| Prets | Journal des sorties et des retours |
| Referentiels | Listes déroulantes : types, états, emplacements |
| Data | Table historisée, masquée. Ne pas la modifier à la main |

Les noms des feuilles sont utilisés par le code : ne pas les renommer.

## Modèle de données

### Feuille Data

| Colonne | Contenu |
| --- | --- |
| `ID_Ligne` | Identifiant unique de la version (auto-incrémenté) |
| `ID_Objet` | Identifiant stable de l'objet, commun à toutes ses versions |
| `Type`, `Nom`, `Quantite`, `Etat`, `Emplacement`, `Commentaire` | Attributs de l'objet |
| `Date_Debut` | Date de création de la version |
| `Date_Fin` | `31/12/9999` tant que la version est en cours |
| `Actif` | `1` = version en cours, `0` = ancienne version ou objet supprimé |
| `Modifie_Par` | Nom d'utilisateur Excel |

### Règles d'historisation

- **Création** : nouvelle ligne, `Date_Debut` = aujourd'hui, `Date_Fin` = 31/12/9999, `Actif` = 1.
- **Modification** : la ligne en cours est fermée (`Date_Fin` = aujourd'hui, `Actif` = 0) et une nouvelle ligne est créée avec le même `ID_Objet`. Sans changement réel, rien n'est écrit.
- **Suppression** : la ligne en cours est fermée, sans nouvelle ligne.

### Feuille Prets

Une ligne par objet sorti : n° de bordereau, date de sortie, responsable, groupe local, destination, ID objet, type, désignation, quantité, état au départ, observations, date de retour, état au retour. Une ligne sans date de retour est une sortie en cours.

## Règles de gestion

- **Doublons** : deux noms sont considérés identiques sans tenir compte des majuscules ni du « s » final (`tente` = `Tentes`), à type égal.
- **Import d'un objet existant** : la quantité importée remplace la quantité en stock (logique de comptage), elle ne s'y ajoute pas.
- **Sortie de matériel** : le stock n'est pas diminué. Quantité disponible = quantité en stock − quantités sorties non rendues. On ne peut pas sortir plus que le disponible.
- **Bordereau validé** : il n'est plus modifiable ; 30 lignes maximum.
- **Retour** : l'état au retour est noté dans Prets ; il n'est pas reporté automatiquement dans l'inventaire.
- **Sauvegarde** : une copie `Sauvegarde_AAAAMMJJ_HHMMSS_…` du classeur est créée dans son dossier avant chaque intégration d'import.

### Formats d'import

- Fichier CSV/TXT, séparateur `;`, `,` ou tabulation, encodage UTF-8 ou Windows-1252.
- Ordre par défaut : `Type;Nom;Quantité;État;Emplacement;Commentaire`. Si la première ligne contient des titres, les colonnes sont repérées par leur nom.
- Notes libres reconnues : `3 tentes`, `tente x3`, `Tentes : 3`, `tentes 3`. Une ligne `Cuisine :` définit le type des lignes suivantes.

## Paramétrage

| Paramètre | Où le modifier |
| --- | --- |
| Types, états, emplacements | Feuille Referentiels (40 valeurs maximum par liste) |
| Préfixe des bordereaux | Constante `PREFIXE_NUM` en haut de `BordereauScout` |
| Mot de passe de la feuille Data | Constante `MOT_DE_PASSE` en haut de `ProtectionData` |


## Documentation utilisateur

Le mode opératoire pas à pas destiné aux utilisateurs est disponible ici : [Mode opératoire — Inventaire du matériel scout](https://claude.ai/artifact/KshiGevqwE7MGVvBZQbEAV).
