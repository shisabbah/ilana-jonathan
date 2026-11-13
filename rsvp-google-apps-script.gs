/**
 * RSVP → Google Sheets (Mairie / Houppa)
 * =======================================
 * 1. Colle TOUT ce fichier dans Apps Script
 * 2. Renseigne SPREADSHEET_ID (URL Sheet : /d/ID/edit)
 * 3. Onglet nommé RSVP
 * 4. Enregistre → Déployer → Nouvelle version (même URL /exec)
 *
 * Le site envoie du form-urlencoded → champs dans e.parameter
 */

var SHEET_NAME = 'RSVP';

/** ID du tableur (obligatoire si le script n’est pas lié à la Sheet) */
var SPREADSHEET_ID = '';

function parseBody_(e) {
  // 1) Champs formulaire (URLSearchParams / form) — priorité
  if (e && e.parameter && (e.parameter.name || e.parameter.mairie || e.parameter.presence)) {
    return e.parameter;
  }

  // 2) JSON brut (ancien mode)
  if (e && e.postData && e.postData.contents) {
    var raw = String(e.postData.contents).trim();
    try {
      return JSON.parse(raw);
    } catch (err) {
      // 3) form-urlencoded dans le body
      var out = {};
      raw.split('&').forEach(function (pair) {
        var parts = pair.split('=');
        if (parts.length >= 2) {
          out[decodeURIComponent(parts[0].replace(/\+/g, ' '))] =
            decodeURIComponent(parts.slice(1).join('=').replace(/\+/g, ' '));
        }
      });
      return out;
    }
  }

  return (e && e.parameter) || {};
}

function getSheet_() {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Renseigne SPREADSHEET_ID en haut du script.');
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

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

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.mairie || '',
      data.mairieGuests != null ? data.mairieGuests : '',
      data.houppa || '',
      data.houppaGuests != null ? data.houppaGuests : '',
      data.message || data.presence || ''
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

function doGet(e) {
  // Test navigateur : ouvrir l’URL /exec → doit afficher ok
  // Si ?name=test est passé, écrit une ligne de test
  try {
    if (e && e.parameter && e.parameter.name) {
      var sheet = getSheet_();
      sheet.appendRow([
        new Date(),
        e.parameter.name,
        e.parameter.mairie || 'test',
        e.parameter.mairieGuests || '1',
        e.parameter.houppa || 'test',
        e.parameter.houppaGuests || '1',
        'test-get'
      ]);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, wrote: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    getSheet_();
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, service: 'rsvp', sheet: SHEET_NAME }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
