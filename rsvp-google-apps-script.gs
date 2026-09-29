/**
 * RSVP → Google Sheets
 * --------------------
 * 1. Crée une Google Sheet (ex. "RSVP Ilana Jonathan")
 * 2. Ligne 1 (en-têtes) : Timestamp | Nom | Presence | Invites | Message
 * 3. Extensions → Apps Script → colle TOUT ce fichier
 * 4. Déployer → Nouveau déploiement → Type : Application Web
 *    - Exécuter en tant que : Moi
 *    - Qui a accès : Tout le monde
 * 5. Copie l'URL du déploiement
 * 6. Colle-la dans index.html : RSVP_SCRIPT_URL = 'https://script.google.com/...'
 */

var SHEET_NAME = 'RSVP'; // nom de l'onglet (crée-le ou adapte)

function doPost(e) {
  try {
    var data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      data = e.parameter;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Timestamp', 'Nom', 'Presence', 'Invites', 'Message']);
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.presence || '',
      data.guests || '',
      data.message || ''
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

/** Test rapide : Exécuter doGet dans l'éditeur pour vérifier que le script charge */
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'rsvp' }))
    .setMimeType(ContentService.MimeType.JSON);
}
