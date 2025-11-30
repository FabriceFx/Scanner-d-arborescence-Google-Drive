/**
 * @OnlyCurrentDoc
 * Ce script gère le scan de fichiers Drive et la génération de rapports graphiques.
 */

// --- CONFIGURATION ---
const CONFIG = {
  TYPES_CIBLES: [], // Laisser vide [] pour tout lister, ou ex: [MimeType.PDF]
  NOM_FEUILLE_DATA: "Données_Drive",
  NOM_FEUILLE_DASHBOARD: "Tableau_de_Bord"
};

/**
 * Crée le menu personnalisé au chargement.
 * Auteur : Fabrice Faucheux
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Gestion Drive & Analytics')
      .addItem('1. Lister les Fichiers (Scan)', 'listerFichiersEtDossiers')
      .addSeparator()
      .addItem('2. Générer le Dashboard', 'genererDashboard')
      .addToUi();
}

/**
 * PHASE 1 : Scan et Listage
 * Scanne le dossier parent et remplit la feuille de données.
 */
function listerFichiersEtDossiers() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  let feuille = classeur.getSheetByName(CONFIG.NOM_FEUILLE_DATA);

  // Création ou récupération de la feuille de données
  if (!feuille) {
    feuille = classeur.insertSheet(CONFIG.NOM_FEUILLE_DATA);
  } else {
    feuille.clear();
  }
  
  feuille.activate(); // Focus sur la feuille
  const ui = SpreadsheetApp.getUi();

  try {
    const fichierParent = DriveApp.getFileById(classeur.getId());
    const dossierParent = fichierParent.getParents().next();

    ui.alert(`Début du scan dans "${dossierParent.getName()}"...`);

    const enTetes = ['Nom', 'Catégorie', 'Type MIME', 'Lien', 'ID', 'Date de création'];
    const donneesCompilees = [];

    // Lancement de la récursion
    scannerDossier(dossierParent, donneesCompilees, "", CONFIG.TYPES_CIBLES);

    // Écriture Batch
    if (donneesCompilees.length > 0) {
      feuille.getRange(1, 1, 1, enTetes.length)
        .setValues([enTetes])
        .setFontWeight("bold")
        .setBackground("#e0f7fa");

      feuille.getRange(2, 1, donneesCompilees.length, donneesCompilees[0].length)
        .setValues(donneesCompilees);
        
      feuille.setFrozenRows(1);
      feuille.autoResizeColumns(1, enTetes.length);
      
      ui.alert(`Scan terminé : ${donneesCompilees.length} éléments. Vous pouvez maintenant générer le Dashboard.`);
    } else {
      ui.alert("Aucune donnée trouvée.");
    }

  } catch (erreur) {
    console.error(erreur);
    ui.alert(`Erreur : ${erreur.message}`);
  }
}

/**
 * PHASE 2 : Analyse et Visualisation
 * Génère une feuille de statistiques avec un graphique camembert.
 */
function genererDashboard() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const feuilleData = classeur.getSheetByName(CONFIG.NOM_FEUILLE_DATA);
  const ui = SpreadsheetApp.getUi();

  if (!feuilleData || feuilleData.getLastRow() < 2) {
    ui.alert("Erreur : Veuillez d'abord lancer le scan (Option 1) pour générer des données.");
    return;
  }

  try {
    // 1. Récupération des données brutes (uniquement la colonne Type MIME - index 2)
    // On ignore l'en-tête (slice(1))
    const valeurs = feuilleData.getRange("C2:C" + feuilleData.getLastRow()).getValues().flat();
    
    // 2. Calcul des statistiques (Aggrégation via reduce)
    const stats = valeurs.reduce((acc, typeMime) => {
      // Simplification du nom pour le graphique (ex: 'application/pdf' -> 'pdf')
      let label = typeMime.split('/').pop().toUpperCase();
      if (label.includes("GOOGLE-APPS")) label = "G-SUITE (" + label.split('.').pop() + ")";
      
      acc[label] = (acc[label] || 0) + 1;
      return acc;
    }, {});

    // Transformation en tableau pour écriture [[Type, Nombre]]
    const tableauStats = Object.entries(stats).map(([type, count]) => [type, count]);
    // Tri décroissant
    tableauStats.sort((a, b) => b[1] - a[1]);
    
    // Ajout en-tête
    tableauStats.unshift(["Type de Fichier", "Quantité"]);

    // 3. Préparation de la feuille Dashboard
    let feuilleDash = classeur.getSheetByName(CONFIG.NOM_FEUILLE_DASHBOARD);
    if (feuilleDash) {
      feuilleDash.clear();
    } else {
      feuilleDash = classeur.insertSheet(CONFIG.NOM_FEUILLE_DASHBOARD);
    }
    feuilleDash.activate();

    // 4. Écriture des statistiques
    feuilleDash.getRange(1, 1, tableauStats.length, 2).setValues(tableauStats);
    feuilleDash.getRange("A1:B1").setFontWeight("bold").setBackground("#fff9c4"); // Jaune pâle
    
    // 5. Création du Graphique (Chart)
    const chart = feuilleDash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(feuilleDash.getRange(1, 1, tableauStats.length, 2))
      .setPosition(2, 4, 0, 0) // Placement à droite du tableau (Ligne 2, Col D)
      .setOption('title', 'Répartition des Types de Fichiers')
      .setOption('pieSliceText', 'percentage') // Affiche % sur le camembert
      .setOption('is3D', true)
      .setOption('height', 400)
      .setOption('width', 600)
      .build();

    feuilleDash.insertChart(chart);
    
    ui.alert("Tableau de bord généré avec succès !");

  } catch (e) {
    console.error(e);
    ui.alert(`Erreur Dashboard : ${e.message}`);
  }
}

/**
 * Helper : Scan récursif (identique version précédente)
 */
function scannerDossier(dossier, tableauDonnees, prefixe, filtres) {
  try {
    // Dossiers (Pour la structure)
    // On met un type MIME spécifique pour les dossiers
    tableauDonnees.push([`${prefixe}📁 ${dossier.getName()}`, 'Dossier', 'application/vnd.google-apps.folder', dossier.getUrl(), dossier.getId(), dossier.getDateCreated()]);

    const sousDossiers = dossier.getFolders();
    while (sousDossiers.hasNext()) scannerDossier(sousDossiers.next(), tableauDonnees, `${prefixe}  `, filtres);

    const fichiers = dossier.getFiles();
    while (fichiers.hasNext()) {
      const f = fichiers.next();
      const mime = f.getMimeType();
      if (filtres.length === 0 || filtres.includes(mime)) {
        tableauDonnees.push([`${prefixe}  📄 ${f.getName()}`, 'Fichier', mime, f.getUrl(), f.getId(), f.getDateCreated()]);
      }
    }
  } catch (e) { console.warn(e.message); }
}
