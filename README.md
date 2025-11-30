# Scanner d'arborescence Google-Drive

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## Description
Ce script génère une arborescence structurée du contenu de votre Google Drive directement dans un Google Sheet. Cette version "Avancée" intègre un moteur de filtrage natif permettant de ne sélectionner que certains types de fichiers (PDF, Sheets, Docs, Images, etc.) tout en conservant la hiérarchie des dossiers.

## Fonctionnalités clés
* **Filtrage par Type MIME** : Configuration simple via la constante `TYPES_CIBLES` pour inclure uniquement les fichiers pertinents.
* **Scan Récursif Intelligent** : Traverse tous les sous-dossiers même si le dossier parent ne contient pas de fichiers cibles immédiats.
* **Batch Processing** : Optimisé pour traiter des milliers de fichiers sans timeout immédiat.
* **Données Enrichies** : Ajout de la colonne "Type MIME" technique pour faciliter les tris ultérieurs dans le tableur.

## Configuration du filtre
Dans le fichier `Code.gs`, modifiez la constante au début du script :

```javascript
// Pour ne lister que les Google Sheets et les PDF :
const TYPES_CIBLES = [MimeType.GOOGLE_SHEETS, MimeType.PDF];

// Pour tout lister (comportement par défaut) :
const TYPES_CIBLES = [];
