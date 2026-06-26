// ─────────────────────────────────────────────────────────────
//  Script Google Apps Script — Parrainage BDE
//  À coller dans Extensions > Apps Script de ton Google Sheet
// ─────────────────────────────────────────────────────────────

// Remplace par l'ID de ton Google Sheet (visible dans l'URL)
// Exemple : https://docs.google.com/spreadsheets/d/SHEET_ID/edit
const SHEET_ID = 'REMPLACE_PAR_TON_SHEET_ID';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = getOrCreateSheet(ss, data.nomFamille);

    // En-têtes si la feuille est vide
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Nom de la famille', 'Membres 2A', 'Hâte de rencontrer les 1A', 'Date de soumission']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    }

    // Ligne de données
    const membres = Array.isArray(data.membres2A) ? data.membres2A.join(', ') : data.membres2A;
    sheet.appendRow([
      data.nomFamille,
      membres,
      data.hateRencontrer,
      data.timestamp || new Date().toLocaleString('fr-FR')
    ]);

    // Auto-resize colonnes
    sheet.autoResizeColumns(1, 4);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Crée un onglet "Familles" s'il n'existe pas encore
function getOrCreateSheet(ss, nomFamille) {
  let sheet = ss.getSheetByName('Familles');
  if (!sheet) {
    sheet = ss.insertSheet('Familles');
  }
  return sheet;
}

// Fonction de test à lancer manuellement depuis l'éditeur Apps Script
function testDoPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        nomFamille: 'Les Aigles',
        membres2A: ['Alice', 'Bob', 'Clara'],
        hateRencontrer: 'OUIIIIII !!',
        timestamp: new Date().toLocaleString('fr-FR')
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
