/**
 * Migration Progress Explorer — Apps Script Web App
 * ===================================================
 * Backend penyimpanan buat mode "Seleksi Kolom" pas app di-deploy ke
 * hosting gratis (storage container-nya sementara). Script ini dibuat
 * BOUND ke satu Google Sheet (Extensions > Apps Script dari sheet itu),
 * jadi tidak perlu ID sheet terpisah — SpreadsheetApp.getActiveSpreadsheet()
 * otomatis merujuk ke sheet yang lagi dibuka pas nge-buat script ini.
 *
 * Cara setup & deploy: baca README.md bagian 7.
 */

// Ganti dengan token rahasia bebas (string apa aja, semakin acak semakin
// aman). Ini dipakai buat verifikasi request, karena Web App yang di-deploy
// dengan akses "Anyone" bisa dipanggil siapa saja yang tahu URL-nya.
const SECRET_TOKEN = "GANTI_DENGAN_TOKEN_RAHASIA_BEBAS_KAMU_SENDIRI";

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== SECRET_TOKEN) {
      return jsonResponse({ ok: false, error: "unauthorized: token salah" });
    }

    if (body.action === "read") {
      return jsonResponse(handleRead(body.sheet));
    }
    if (body.action === "write") {
      return jsonResponse(handleWrite(body.sheet, body.headers, body.rows));
    }
    return jsonResponse({ ok: false, error: "action tidak dikenal: " + body.action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function handleRead(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: true, rows: [] };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, rows: [] };

  const headers = data[0];
  const rows = data.slice(1)
    // buang baris yang seluruh kolomnya kosong
    .filter((row) => row.some((cell) => cell !== "" && cell !== null))
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      return obj;
    });

  return { ok: true, rows: rows };
}

function handleWrite(sheetName, headers, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  if (!headers || headers.length === 0) return { ok: true };

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return { ok: true };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
