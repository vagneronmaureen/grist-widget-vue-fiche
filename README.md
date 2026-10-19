# Fiche Produit — Widget Grist

Widget Grist personnalisable permettant d'afficher les enregistrements sous forme de fiche structurée.

## Fonctionnalités

- Affichage en fiche de tous les types de champs Grist (texte, nombre, booléen, choix, référence…)
- Mise en page configurable par glisser-déposer
- Ajout / suppression / renommage de champs à la volée
- Sauvegarde du layout dans les options du widget
- Enregistrement des modifications directement dans Grist
- **Mode de modification au choix** (Configurer → ⚙ Paramètres) : via le bouton « Éditer », ou directement dans les champs
- **Champs en lecture seule** (Configurer → survol d'un champ → 🔒) : jamais modifiables, dans les deux modes
- **Recherche produit multi-champs** (vue non liée, Configurer → ⚙ Paramètres) : libellé affiché, champ secondaire grisé entre parenthèses (ex. ID), et choix des colonnes sur lesquelles porte la recherche

## Installation

1. Dans Grist, ajouter un **widget personnalisé**
2. Coller l'URL du widget :
```
   https://vagneronmaureen.github.io/sdpc-widget/
```
3. Accorder l'accès **complet** au widget (requis pour lire et écrire les données)

## Utilisation

| Action | Comment |
|---|---|
| Changer de fiche | Sélectionnez un enregistrement dans le sélecteur en haut |
| Modifier un champ | Cliquez directement dans la valeur |
| Sauvegarder | Bouton **Enregistrer** (apparaît dès qu'une modification est détectée) |
| Configurer la mise en page | Bouton **Configurer** (mode glisser-déposer) |
| Ajouter un élément | En mode Configurer → barre d'ajout en bas |

## Structure du projet
```
sdpc-widget/
├── index.html       # Structure HTML + styles CSS
├── main.js          # Logique JavaScript + intégration Grist API
├── manifest.json    # Métadonnées du widget
└── README.md        # Cette documentation
```

## Compatibilité

Nécessite un accès **full** à l'API Grist (`grist.ready({ requiredAccess: 'full' })`).  

## Versions

| Version | Description |
|---|---|
| v1.0 | Version initiale — layout configurable, tous types de champs |
| v2.0 | Version avec un style Notion plus prononcé |
| v3.0 | Mode d'édition configurable (bouton / direct), champs en lecture seule, recherche produit multi-champs |
