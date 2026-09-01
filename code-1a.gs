// ─────────────────────────────────────────────────────────────
//  Script Google Apps Script — Inscriptions 1A Parrainage BDE
//  À coller dans Extensions > Apps Script de ton Google Sheet 1A
// ─────────────────────────────────────────────────────────────

const SHEET_ID_1A = 'REMPLACE_PAR_TON_SHEET_ID_1A';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss    = SpreadsheetApp.openById(SHEET_ID_1A);
    const sheet = getOrCreateSheet1A(ss);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Prénom et Nom',
        'Filière',
        'Associations souhaitées',
        'Sports souhaités',
        'Note soirées (1-10)',
        'À découvrir / apprendre',
        'Hâte de rencontrer la famille',
        'Date de soumission'
      ]);
      sheet.getRange(1, 1, 1, 8)
        .setFontWeight('bold')
        .setBackground('#16a34a')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    const associations = Array.isArray(data.associations) ? data.associations.join(', ') : (data.associations || '');
    const sports       = Array.isArray(data.sports)       ? data.sports.join(', ')       : (data.sports || '');

    sheet.appendRow([
      data.prenomNom         || '',
      data.filiere           || '',
      associations,
      sports,
      data.soirees           || '',
      data.decouvrir         || '',
      data.hateRencontrer    || '',
      data.timestamp         || new Date().toLocaleString('fr-FR')
    ]);

    sheet.autoResizeColumns(1, 8);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet1A(ss) {
  let sheet = ss.getSheetByName('Inscriptions 1A');
  if (!sheet) sheet = ss.insertSheet('Inscriptions 1A');
  return sheet;
}

function testDoPost1A() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        prenomNom:      'Alice Martin',
        filiere:        'Math',
        associations:   ['BDE', 'Forum', 'Spectra'],
        sports:         ['Badminton', 'Natation'],
        soirees:        '7',
        decouvrir:      'Découvrir la vie associative et le réseau ENSAE.',
        hateRencontrer: 'OUIIIIII !!',
        timestamp:      new Date().toLocaleString('fr-FR')
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
