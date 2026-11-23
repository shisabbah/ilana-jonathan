/**
 * RSVP → Google Sheets (Mairie / Henné / Houppa)
 * =======================================
 * 1. Colle TOUT ce fichier dans Apps Script
 * 2. Renseigne SPREADSHEET_ID (URL Sheet : /d/ID/edit)
 * 3. Onglet nommé RSVP
 * 4. Enregistre → Déployer → Nouvelle version (même URL /exec)
 *
 * Le site envoie du form-urlencoded → champs dans e.parameter
 */

var SHEET_NAME = 'RSVP';

var HEADERS = [
  'Timestamp',
  'Nom',
  'Mairie',
  'Nb Mairie',
  'Henné',
  'Nb Henné',
  'Houppa',
  'Nb Houppa',
  'Détail'
];

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

  ensureHeaders_(sheet);

  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }

  var width = Math.max(sheet.getLastColumn(), HEADERS.length);
  var current = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (cell) {
    return String(cell || '').trim();
  });

  if (HEADERS.every(function (header, index) { return current[index] === header; })) {
    return;
  }

  // Feuille déjà remplie : Mairie puis Houppa, sans colonne Henné.
  var legacyWithoutHenne =
    current[0] === 'Timestamp' &&
    current[2] === 'Mairie' &&
    current[4] === 'Houppa' &&
    current.indexOf('Henné') === -1;

  if (legacyWithoutHenne) {
    sheet.insertColumnsAfter(4, 2);
    sheet.getRange(1, 5, 1, 2).setValues([['Henné', 'Nb Henné']]);
    return;
  }

  if (sheet.getLastRow() === 1) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function invalidReason_(data) {
  var name = String(data.name || '').replace(/\s+/g, ' ').trim();
  if (!name) return 'Nom manquant.';

  function guestsOk(answer, count) {
    var present = String(answer || '').toLowerCase() === 'oui';
    var absent = String(answer || '').toLowerCase() === 'non';
    if (!present && !absent) return false;
    var n = Number(count);
    if (present) return isFinite(n) && n >= 1 && n <= 20;
    return count === '' || count == null || (isFinite(n) && n >= 0 && n <= 20);
  }

  if (!guestsOk(data.mairie, data.mairieGuests)) return 'Réponse Mairie invalide.';
  if (!guestsOk(data.henne, data.henneGuests)) return 'Réponse Henné invalide.';
  if (!guestsOk(data.houppa, data.houppaGuests)) return 'Réponse Houppa invalide.';
  return '';
}

function doPost(e) {
  try {
    var data = parseBody_(e);
    var reason = invalidReason_(data);
    if (reason) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: reason }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    data.name = String(data.name).replace(/\s+/g, ' ').trim();
    var sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.mairie || '',
      data.mairieGuests != null ? data.mairieGuests : '',
      data.henne || '',
      data.henneGuests != null ? data.henneGuests : '',
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
        e.parameter.henne || 'test',
        e.parameter.henneGuests || '1',
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
