/**
 * RSVP → Google Sheets (version Mairie / Houppa)
 * ================================================
 * IMPORTANT : ouvre ce script DEPUIS la Google Sheet
 * (Extensions → Apps Script), pas un projet Apps Script isolé.
 *
 * 1. Dans la Sheet, crée un onglet nommé exactement : RSVP
 * 2. Ligne 1 (en-têtes) :
 *    Timestamp | Nom | Mairie | Nb Mairie | Houppa | Nb Houppa | Détail
 * 3. Colle TOUT ce fichier (remplace l’ancien code)
 * 4. Enregistre (Ctrl/Cmd+S)
 * 5. Déployer → Gérer les déploiements → Modifier (crayon)
 *    OU Nouveau déploiement → Application Web
 *      - Exécuter en tant que : Moi
 *      - Qui a accès : Tout le monde
 * 6. Déploie, copie l’URL …/exec
 * 7. Colle l’URL dans index.html (RSVP_SCRIPT_URL)
 *
 * Après chaque modification du script : redéploie une NOUVELLE version.
 */

var SHEET_NAME = 'RSVP';

/**
 * Si getActiveSpreadsheet() échoue (script non lié à la Sheet),
 * colle ici l’ID du tableur (dans l’URL : /d/CET_ID/edit).
 */
var SPREADSHEET_ID = '';

function parseBody_(e) {
  var data = {};
  if (e && e.postData && e.postData.contents) {
    var raw = String(e.postData.contents).trim();
    try {
      data = JSON.parse(raw);
    } catch (err) {
      // fallback form-urlencoded éventuel
      if (e.parameter) data = e.parameter;
    }
  } else if (e && e.parameter) {
    data = e.parameter;
  }
  return data || {};
}

function getSheet_() {
  var ss = null;
  if (SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error(
      'Aucune feuille liée. Ouvre Apps Script depuis la Sheet, ' +
      'ou renseigne SPREADSHEET_ID en haut du fichier.'
    );
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // En-têtes si onglet vide
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Timestamp',
      'Nom',
      'Mairie',
      'Nb Mairie',
      'Houppa',
      'Nb Houppa',
      'Détail'
    ]);
  }

  return sheet;
}

function doPost(e) {
  try {
    var data = parseBody_(e);
    var sheet = getSheet_();

    var name = data.name || '';
    var mairie = data.mairie || '';
    var houppa = data.houppa || '';
    var mairieGuests = data.mairieGuests != null ? data.mairieGuests : '';
    var houppaGuests = data.houppaGuests != null ? data.houppaGuests : '';
    var detail = data.message || data.presence || '';

    // Compat ancien formulaire (une seule présence)
    if (!mairie && !houppa && data.presence) {
      detail = data.presence;
      mairie = data.presence;
      mairieGuests = data.guests || '';
    }

    sheet.appendRow([
      new Date(),
      name,
      mairie,
      mairieGuests,
      houppa,
      houppaGuests,
      detail
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Test : Exécuter doGet → doit afficher {"ok":true,"service":"rsvp"} */
function doGet() {
  try {
    getSheet_(); // vérifie l’accès Sheet
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, service: 'rsvp', sheet: SHEET_NAME }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
