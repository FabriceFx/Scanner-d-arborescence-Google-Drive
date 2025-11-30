/**
 * @OnlyCurrentDoc
 * Ce script est conçu pour être utilisé uniquement avec le document Google Sheet actuel.
 */

// --- CONFIGURATION ---
/**
 * Liste des types MIME à inclure dans le rapport.
 * Laissez le tableau vide [] pour tout lister.
 * Exemples : MimeType.GOOGLE_SHEETS, MimeType.PDF, "image/jpeg"
 */
const TYPES_CIBLES = [
  // MimeType.GOOGLE_SHEETS, 
  // MimeType.PDF 
]; 

/**
 * Crée un menu personnalisé dans l'interface utilisateur de Google Sheets.
 * Auteur : Fabrice Faucheux
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Gestion des fichiers')
      .addItem('Lister les Fichiers (Filtré)', 'listerFichiersEtDossiers')
      .addToUi();
}

/**
 * Fonction principale : Orchestre le scan et l'écriture en batch.
 * @returns {void}
 */
function listerFichiersEtDossiers() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = classeur.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. Initialisation
    const fichierParent = DriveApp.getFileById(classeur.getId());
    const dossierParent = fichierParent.getParents().next();

    ui.alert(`Scan du dossier "${dossierParent.getName()}" en cours avec filtrage...`);

    // 2. Préparation des données (Mémoire)
    const enTetes = ['Nom', 'Type (Visuel)', 'Type MIME (Tech)', 'Lien', 'ID', 'Date de création'];
    const donneesCompilees = [];

    // 3. Exécution du scan
    scannerDossier(dossierParent, donneesCompilees, "", TYPES_CIBLES);

    // 4. Écriture Batch (Performance)
    if (donneesCompilees.length > 0) {
      feuille.clear();
      
      // En-têtes
      feuille.getRange(1, 1, 1, enTetes.length)
        .setValues([enTetes])
        .setFontWeight("bold")
        .setBackground("#e0f7fa") // Bleu clair pour distinguer
        .setBorder(true, true, true, true, true, true);

      // Données
      feuille.getRange(2, 1, donneesCompilees.length, donneesCompilees[0].length)
        .setValues(donneesCompilees);
        
      // Mise en page
      feuille.setFrozenRows(1);
      feuille.autoResizeColumns(1, enTetes.length);
      
      ui.alert(`Terminé : ${donneesCompilees.length} éléments trouvés correspondant aux critères.`);
    } else {
      ui.alert("Aucun fichier correspondant aux filtres n'a été trouvé.");
    }

  } catch (erreur) {
    console.error(`Erreur critique : ${erreur.message}`);
    ui.alert(`Erreur : ${erreur.message}`);
  }
}

/**
 * Scanne récursivement les dossiers et filtre les fichiers.
 * @param {GoogleAppsScript.Drive.Folder} dossier - Dossier actuel.
 * @param {Array} tableauDonnees - Accumulateur de résultats.
 * @param {string} prefixe - Indentation visuelle.
 * @param {Array<string>} filtres - Liste des MimeTypes autorisés.
 */
function scannerDossier(dossier, tableauDonnees, prefixe, filtres) {
  try {
    // Toujours ajouter le dossier pour garder la structure visuelle
    tableauDonnees.push([
      `${prefixe}📁 ${dossier.getName()}`,
      'Dossier',
      'application/vnd.google-apps.folder',
      dossier.getUrl(),
      dossier.getId(),
      dossier.getDateCreated()
    ]);

    // Scan des sous-dossiers (Récursion)
    const sousDossiers = dossier.getFolders();
    while (sousDossiers.hasNext()) {
      scannerDossier(sousDossiers.next(), tableauDonnees, `${prefixe}  `, filtres);
    }

    // Scan et Filtrage des fichiers
    const fichiers = dossier.getFiles();
    while (fichiers.hasNext()) {
      const fichier = fichiers.next();
      const mimeType = fichier.getMimeType();

      // LOGIQUE DE FILTRAGE
      // Si filtres n'est pas vide ET que le mimeType n'est pas dedans, on passe au suivant.
      if (filtres.length > 0 && !filtres.includes(mimeType)) {
        continue; 
      }

      tableauDonnees.push([
        `${prefixe}  📄 ${fichier.getName()}`,
        'Fichier',
        mimeType,
        fichier.getUrl(),
        fichier.getId(),
        fichier.getDateCreated()
      ]);
    }
  } catch (e) {
    console.warn(`Erreur d'accès (fichier/dossier ignoré) : ${e.message}`);
  }
}
