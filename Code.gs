/**
 * @fileoverview Script Audit Shared Drives
 * * @author Fabrice Faucheux
 */

// --- CONFIGURATION ---
const CONFIG = {
  NOM_ONGLET_PRINCIPAL: "Liste_Drives_Partages",
  TEMPS_MAX_EXECUTION: 1000 * 60 * 4.5, // 4.5 minutes (Marge de sécurité)
  BATCH_SIZE: 10, // Sauvegarde intermédiaire
  COLS: { // Mapping des colonnes (Index Base 0)
    ID: 0,
    LIEN_ONGLET: 1,
    DATE_EXPORT: 2,
    NOM: 3,
    USER: 4,
    DATE_MODIF: 5,
    LIEN_FICHIER: 6
  }
};

/**
 * Crée le menu personnalisé à l'ouverture.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Scanner les Drives partagés')
    .addItem('1. Importer mes Drives partagés', 'importerMesSharedDrives')
    .addSeparator()
    .addItem('2. Lancer l\'audit', 'listerDossiersSharedDrives')
    .addToUi();
}

/**
 * Importe la liste des Shared Drives.
 * ACTION : Efface tout le contenu (sauf entêtes) avant l'import.
 */
function importerMesSharedDrives() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = obtenirOuCreerOngletPrincipal(ss);

  // 1. Confirmation de sécurité
  const reponse = ui.alert(
    'Réinitialisation complète', 
    '⚠️ Attention : Cette action va effacer TOUTES les données actuelles du tableau (Lignes 2+) pour réimporter une liste propre.\n\nVoulez-vous continuer ?', 
    ui.ButtonSet.YES_NO
  );

  if (reponse !== ui.Button.YES) return;

  // 2. Nettoyage de la zone de travail
  const maxRows = feuille.getMaxRows();
  if (maxRows > 1) {
    // On efface le contenu de A2 jusqu'à la dernière ligne/colonne possible
    feuille.getRange(2, 1, maxRows - 1, 7).clearContent(); 
    SpreadsheetApp.flush(); // Mise à jour visuelle immédiate
  }

  ss.toast("Interrogation de l'API Drive...", "En cours");
  
  const liste = [];
  let pageToken = null;

  try {
    do {
      const res = Drive.Drives.list({
        maxResults: 50,
        pageToken: pageToken,
        useDomainAdminAccess: false // Mettre true si vous êtes Super Admin et voulez TOUT voir
      });
      if (res.items) {
        res.items.forEach(d => liste.push([d.id, d.name]));
      }
      pageToken = res.nextPageToken;
    } while (pageToken);

  } catch (e) {
    console.error("Erreur Import", e);
    ui.alert(`Erreur API Drive : ${e.message}\nVerifiez que le service 'Drive' est activé.`);
    return;
  }

  // 3. Écriture des résultats
  if (liste.length > 0) {
    // Tri alphabétique respectant les accents français
    liste.sort((a, b) => a[1].localeCompare(b[1], 'fr'));

    const ids = liste.map(x => [x[0]]);
    const noms = liste.map(x => [x[1]]);

    feuille.getRange(2, 1, ids.length, 1).setValues(ids);
    feuille.getRange(2, 4, noms.length, 1).setValues(noms);
    
    ui.alert(`${liste.length} Drives importés. Tableau réinitialisé.`);
  } else {
    ui.alert("Aucun Drive partagé trouvé.");
  }
}

/**
 * Fonction principale d'audit et de maillage.
 * Reprend automatiquement là où elle s'est arrêtée.
 */
function listerDossiersSharedDrives() {
  const tempsDebut = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const feuille = obtenirOuCreerOngletPrincipal(ss);

  const lastRow = feuille.getLastRow();
  if (lastRow < 2) {
    ss.toast("Liste vide. Veuillez lancer l'étape 1.", "Stop");
    return;
  }

  // Lecture en masse
  const rangeDonnees = feuille.getRange(2, 1, lastRow - 1, 7);
  const values = rangeDonnees.getValues();
  const formulas = rangeDonnees.getFormulas(); // Nécessaire pour vérifier les liens existants
  
  let modifications = false;

  for (let i = 0; i < values.length; i++) {
    
    // Check Timeout
    if (new Date().getTime() - tempsDebut > CONFIG.TEMPS_MAX_EXECUTION) {
      sauvegarderEtat(rangeDonnees, values);
      ss.toast(`Pause Timeout (${i} traités). Reprise dans 1 min...`, "Script", 20);
      declencherRepriseAuto();
      return; 
    }

    const driveId = String(values[i][CONFIG.COLS.ID]).trim();
    if (!driveId) continue;

    // A. Gestion de l'Onglet et du Lien
    // On vérifie si un lien valide existe déjà pour ne pas le casser
    const formuleActuelle = formulas[i][CONFIG.COLS.LIEN_ONGLET];
    const aBesoinDeReparation = !formuleActuelle || !formuleActuelle.includes("HYPERLINK");

    if (aBesoinDeReparation) {
      try {
        const nomDrive = recupererNomDriveRapide(driveId);
        values[i][CONFIG.COLS.NOM] = nomDrive;
        
        // Génération de l'onglet et récupération de la formule (avec point-virgule)
        const formule = genererOngletArborescence(ss, driveId, nomDrive);
        values[i][CONFIG.COLS.LIEN_ONGLET] = formule;
        values[i][CONFIG.COLS.DATE_EXPORT] = new Date();
        
        modifications = true;
      } catch (e) {
        values[i][CONFIG.COLS.LIEN_ONGLET] = `Erreur: ${e.message}`;
      }
    }

    // B. Mise à jour Activité (Si vide ou erreur)
    const userActuel = values[i][CONFIG.COLS.USER];
    if (!userActuel || userActuel === "Erreur API" || userActuel === "-") {
      const infos = recupererDerniereActivite(driveId);
      values[i][CONFIG.COLS.USER] = infos.user;
      values[i][CONFIG.COLS.DATE_MODIF] = infos.date;
      values[i][CONFIG.COLS.LIEN_FICHIER] = infos.link;
      modifications = true;
    }

    // Sauvegarde Batch
    if (modifications && i % CONFIG.BATCH_SIZE === 0) {
      sauvegarderEtat(rangeDonnees, values);
      modifications = false;
    }
  }

  // Sauvegarde Finale
  sauvegarderEtat(rangeDonnees, values);
  nettoyerTriggers();
  ss.toast("Audit complet terminé.", "Succès", 5);
}

/**
 * Initialise l'onglet principal avec formatage FR strict.
 */
function obtenirOuCreerOngletPrincipal(ss) {
  let f = ss.getSheetByName(CONFIG.NOM_ONGLET_PRINCIPAL);
  if (!f) {
    f = ss.insertSheet(CONFIG.NOM_ONGLET_PRINCIPAL);
    f.getRange("A1:G1")
      .setValues([["ID Drive", "Lien Onglet", "Date Export", "Nom Drive", "Modifié Par", "Date Modif", "Lien Fichier"]])
      .setFontWeight("bold")
      .setBackground("#e0e0e0")
      .setBorder(true, true, true, true, null, null);
    f.setFrozenRows(1);
    
    // Largeurs
    f.setColumnWidth(1, 140); // ID
    f.setColumnWidth(2, 130); // Lien
    f.setColumnWidth(3, 130); // Date Export
    f.setColumnWidth(4, 200); // Nom
    f.setColumnWidth(6, 130); // Date Modif
    f.setColumnWidth(7, 250); // URL
    
    // --- FORMATAGE LOCALE FR ---
    // Force le format jour/mois/année heure:minute pour les colonnes dates (C et F)
    f.getRange("C:C").setNumberFormat("dd/MM/yyyy HH:mm");
    f.getRange("F:F").setNumberFormat("dd/MM/yyyy HH:mm");
  }
  return f;
}

/**
 * Récupère le nom via API (Rapide).
 */
function recupererNomDriveRapide(id) {
  try { return Drive.Drives.get(id).name; } 
  catch (e) { return "Inconnu"; }
}

/**
 * Crée l'onglet enfant et retourne la formule de lien compatible FR (;).
 */
function genererOngletArborescence(ss, driveId, nomDrive) {
  const suffixe = driveId.substring(driveId.length - 4);
  const cleanName = nomDrive.replace(/[:\/\\?*\[\]]/g, "_").substring(0, 30);
  const sheetName = `📦 ${cleanName} [${suffixe}]`;

  const oldSheet = ss.getSheetByName(sheetName);
  if (oldSheet) ss.deleteSheet(oldSheet);

  const ns = ss.insertSheet(sheetName, ss.getNumSheets());
  const mainSheetId = ss.getSheetByName(CONFIG.NOM_ONGLET_PRINCIPAL).getSheetId();
  
  // CORRECTION LOCALE : Point-virgule
  ns.getRange("A1").setFormula(`=HYPERLINK("#gid=${mainSheetId}"; "⬅️ Retour Liste principale")`);
  
  const data = [["Nom du Dossier", "ID Dossier", "Lien"]];
  
  try {
    let pageToken = null;
    do {
      const res = Drive.Files.list({
        q: `'${driveId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
        maxResults: 1000,
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        fields: "nextPageToken, items(id, title, alternateLink)"
      });
      if (res.items) {
        res.items.forEach(item => data.push([item.title, item.id, item.alternateLink]));
      }
      pageToken = res.nextPageToken;
    } while (pageToken);
  } catch (e) {
    data.push(["Erreur scan", e.message, ""]);
  }

  if (data.length > 0) {
    ns.getRange(3, 1, data.length, 3).setValues(data);
    ns.getRange(3, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
    ns.setColumnWidth(1, 300);
  } else {
    ns.getRange(3, 1).setValue("Aucun sous-dossier racine.");
  }

  // CORRECTION LOCALE : Point-virgule
  return `=HYPERLINK("#gid=${ns.getSheetId()}"; "📂 Voir onglet")`;
}

/**
 * Récupère user + date via API.
 */
function recupererDerniereActivite(driveId) {
  const result = { user: "Aucun", date: "-", link: "-" };
  const queryParams = {
    corpora: 'drive', driveId: driveId, includeItemsFromAllDrives: true,
    supportsAllDrives: true, orderBy: 'modifiedDate desc', maxResults: 1, q: "trashed = false"
  };

  try {
    let res = Drive.Files.list(queryParams);
    // Fallback Legacy TeamDrive
    if (!res.items || res.items.length === 0) {
       res = Drive.Files.list({ ...queryParams, corpora: 'teamDrive', teamDriveId: driveId, includeTeamDriveItems: true, supportsTeamDrives: true });
    }

    if (res.items && res.items.length > 0) {
      const f = res.items[0];
      let userStr = "Inconnu";
      if (f.lastModifyingUser) userStr = f.lastModifyingUser.emailAddress || f.lastModifyingUser.displayName || "Nom masqué";
      else if (f.lastModifyingUserName) userStr = f.lastModifyingUserName;

      result.user = userStr;
      result.date = new Date(f.modifiedDate);
      result.link = f.alternateLink || "";
    }
  } catch (e) {
    result.user = "Erreur API";
    result.link = e.message;
  }
  return result;
}

function sauvegarderEtat(range, values) {
  range.setValues(values);
  SpreadsheetApp.flush();
}

function declencherRepriseAuto() {
  nettoyerTriggers();
  ScriptApp.newTrigger('listerDossiersSharedDrives').timeBased().after(1000 * 60).create();
}

function nettoyerTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'listerDossiersSharedDrives') ScriptApp.deleteTrigger(t);
  });
}
