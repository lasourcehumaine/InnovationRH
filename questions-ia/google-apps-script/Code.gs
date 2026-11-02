/**
 * Innovation RH & Cie : réception anonyme des questions.
 * Coller dans Extensions > Apps Script du Google Sheets associé,
 * puis déployer comme application web : exécuter en tant que "Moi",
 * accès "Tout le monde". Ne pas partager le tableur avec le public.
 */
const SHEET_ID = '1NowoEJb8qT4Mt6QFsgc1qbupIQ6YH_kSAb3JN2jYmmk';
const TAB = 'Questions reçues';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const params = e && e.parameter ? e.parameter : {};
    console.log('doPost reçu ; paramètre question présent : ' + Object.prototype.hasOwnProperty.call(params, 'question'));
    const q = String(params.question || '').trim();
    if (q.length < 5 || q.length > 500) throw new Error('Question invalide');
    // Empêcher l'injection de formules à partir du texte des participants.
    const safe = /^[=+@\-]/.test(q) ? "'" + q : q;
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
    if (!sheet) throw new Error('Onglet introuvable : ' + TAB);
    sheet.appendRow([new Date(), safe, 'À traiter', '']);
    console.log('Enregistrement effectué dans ' + SHEET_ID);
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('ECHEC doPost : ' + String(err && err.stack || err));
    throw err;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/** Diagnostic sans appel de l'application : choisir testWrite dans Apps Script puis Exécuter.
 * Insère une ligne de test dans le tableur cible, ou affiche l'erreur exacte dans Exécutions.
 */
function testWrite() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tab = ss.getSheetByName(TAB);
  if (!tab) throw new Error('Onglet introuvable : ' + TAB);
  tab.appendRow([new Date(), 'TEST DIRECT APPS SCRIPT', 'À traiter', '']);
  Logger.log('Écriture confirmée dans : ' + ss.getUrl());
}

/** Diagnostic en ouvrant l’URL /exec dans le navigateur, sans écrire de question. */
function doGet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (!ss.getSheetByName(TAB)) throw new Error('Onglet introuvable : ' + TAB);
    return ContentService.createTextOutput('OK - tableur accessible : ' + ss.getName());
  } catch (e) {
    return ContentService.createTextOutput('ERREUR : ' + e.message);
  }
}
