# Scanner d'arborescence Google Drive

![License MIT](https://img.shields.io/badge/License-MIT-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Google%20Apps%20Script-green)
![Runtime](https://img.shields.io/badge/Google%20Apps%20Script-V8-green)
![Author](https://img.shields.io/badge/Auteur-Fabrice%20Faucheux-orange)

## Description
Ce projet transforme votre Google Sheet en un outil d'audit de fichiers. Il combine un scanneur récursif de Google Drive avec un module de visualisation de données (Dashboard). Le script génère automatiquement des statistiques sur la répartition des types de fichiers (PDF, Google Docs, Images, etc.) sous forme de tableau et de graphique.

## Fonctionnalités
### 1. Scan & Data (`Données_Drive`)
* Exploration récursive complète.
* Renommage automatique de la feuille de destination.
* Capture des métadonnées (ID, URL, Date).

### 2. Dashboard (`Tableau_de_Bord`)
* **Agrégation automatique** : Calcul des occurrences par type MIME.
* **Nettoyage des étiquettes** : Transformation des types techniques (`application/vnd.google-apps.spreadsheet`) en noms lisibles (`G-SUITE (SPREADSHEET)`).
* **Visualisation** : Génération d'un **Pie Chart 3D** dynamique intégré à la feuille.

## Installation
1.  Copiez le code `Code.js` dans l'éditeur de script.
2.  Sauvegardez et rafraîchissez le fichier Sheet.
3.  Utilisez le menu **Gestion Drive & Analytics** :
    * Étape 1 : Cliquez sur **Lister les Fichiers (Scan)**.
    * Étape 2 : Une fois le scan fini, cliquez sur **Générer le Dashboard**.

