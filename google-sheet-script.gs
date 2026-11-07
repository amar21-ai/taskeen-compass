/**
 * بوصلة التسكين – استقبال ردود المتطوعين في Google Sheet
 * الصق الكود ده في: Extensions > Apps Script جوه الشيت، وبعدين Deploy > New deployment > Web app
 * (Execute as: Me — Who has access: Anyone).
 *
 * الصفحة بتبعت قايمة [اسم العمود، القيمة] بالترتيب.
 * السكريبت بيعمل العمود لو مش موجود، فلو سؤال اتغير أو اتضاف، الشيت بيتظبط لوحده.
 */
const SHEET_NAME = "الردود";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const pairs = Array.isArray(d.row) ? d.row : [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.setRightToLeft(true); }

    // headers: first column is always the date
    let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    if (!headers.length) headers = ["التاريخ"];
    pairs.forEach(p => { const h = String(p[0]).slice(0, 120); if (headers.indexOf(h) === -1) headers.push(h); });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#EAF1F7");
    sh.setFrozenRows(1); sh.setFrozenColumns(2);

    const row = new Array(headers.length).fill("");
    row[0] = new Date();
    pairs.forEach(p => {
      const i = headers.indexOf(String(p[0]).slice(0, 120));
      let v = p[1] == null ? "" : p[1];
      if (typeof v === "string") { v = v.slice(0, 1000); if (/^[=+\-@]/.test(v) || /^0\d+$/.test(v)) v = "'" + v; }
      row[i] = v;
    });
    sh.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
