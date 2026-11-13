// E4 – SPŠ a VOŠ Kladno: backend pro stránku se seznamem spolužáků
const SHEET_ID = '19Kfx9TbLBxSGICdhXdzWk9NVWzdclwTLNTxjKurvVw4';

function sheet_() { return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; }

function rows_() {
  const s = sheet_();
  const last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 4).getDisplayValues().map(r => ({
    prijmeni: r[0], jmeno: r[1], ve: r[2], pozn: r[3]
  }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() { return json_({ ok: true, rows: rows_() }); }

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const s = sheet_();
      const names = s.getRange(2, 1, s.getLastRow() - 1, 2).getDisplayValues();
      const i = names.findIndex(r => r[0] === b.prijmeni && r[1] === b.jmeno);
      if (i < 0) return json_({ ok: false, error: 'Jméno nenalezeno' });
      const row = i + 2;
      if (s.getRange(row, 3).getDisplayValue() === '–') return json_({ ok: false, error: 'Nelze měnit' });
      if ('ve' in b) s.getRange(row, 3).setValue(b.ve ? 'ANO' : '');
      if ('pozn' in b) {
        let t = String(b.pozn || '').slice(0, 300);
        if (/^[=+\-@]/.test(t)) t = "'" + t;   // ochrana proti vzorcům
        s.getRange(row, 4).setValue(t);
      }
    } finally { lock.releaseLock(); }
    return json_({ ok: true, rows: rows_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
