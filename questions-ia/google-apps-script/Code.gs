/**
 * Innovation RH & Cie : réception anonyme des questions.
 * Coller dans Extensions > Apps Script du Google Sheets associé,
 * puis déployer comme application web : exécuter en tant que "Moi",
 * accès "Tout le monde". Ne pas partager le tableur avec le public.
 */
const SHEET_ID = '1i6BD9DRaJoGXFKHRAlBb6pPcwx2e1ZZ1y5mXg-d0eWQ';
const TAB = 'Questions reçues';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const params = e && e.parameter ? e.parameter : {};
    const q = String(params.question || '').trim();
    if (q.length < 5 || q.length > 500) throw new Error('Question invalide');
    // Empêcher l'injection de formules à partir du texte des participants.
    const safe = /^[=+@\-]/.test(q) ? "'" + q : q;
    lock.waitLock(10000);
    SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB)
      .appendRow([new Date(), safe, 'À traiter', '']);
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
