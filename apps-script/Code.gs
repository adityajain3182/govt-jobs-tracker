// Web-app endpoint that lets applications.html write to the "Applied" tab.
// One-time setup (from the Google Sheet):
//   1. Extensions -> Apps Script, replace the default code with this file.
//   2. Project Settings -> Script Properties -> add property TOKEN = <any long random string>.
//   3. Deploy -> New deployment -> type "Web app":
//        Execute as: Me    |    Who has access: Anyone
//   4. Copy the web app URL, then on the applications page click "Connect sheet"
//      and paste the URL + the same token.
// Redeploy (Manage deployments -> edit -> new version) after any code change.

const SHEET_NAME = "Applied";
// Headers the web page reads; created automatically in row 1 if missing.
const MANAGED_HEADERS = ["Status", "Exam Date", "Exam Centre", "Admit Card", "Result", "Notes"];
const FIELD_TO_HEADER = {
  dept: "Dept",
  post: "Post",
  deadline: "Application Deadline",
  reg: "Registration No.",
  link: "Link",
  others: "Others",
  status: "Status",
  examDate: "Exam Date",
  examCentre: "Exam Centre",
  admitCard: "Admit Card",
  result: "Result",
  notes: "Notes"
};

function doGet() {
  return json({ ok: true, service: "applications-sheet-bridge" });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty("TOKEN");
    if (!token) return json({ ok: false, error: "TOKEN script property not set" });
    if (body.token !== token) return json({ ok: false, error: "Invalid token" });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return json({ ok: false, error: "Sheet '" + SHEET_NAME + "' not found" });

    const headers = ensureHeaders(sheet);
    if (body.action === "add") return json(addRow(sheet, headers, body.data || {}));
    if (body.action === "update") return json(updateRow(sheet, headers, body.rowNumber, body.data || {}));
    return json({ ok: false, error: "Unknown action: " + body.action });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function ensureHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const have = {};
  row1.forEach(function (h, i) { have[h.trim().toLowerCase()] = i + 1; });

  MANAGED_HEADERS.forEach(function (h) {
    if (have[h.toLowerCase()]) return;
    // Reuse a junk "Column N" header slot if present, else append a new column.
    let col = 0;
    for (let i = 0; i < row1.length; i++) {
      if (/^column \d+$/i.test(row1[i].trim()) && !row1[i]._used) { col = i + 1; row1[i] = h; break; }
    }
    if (!col) { col = sheet.getLastColumn() + 1; row1.push(h); }
    sheet.getRange(1, col).setValue(h);
    have[h.toLowerCase()] = col;
  });

  const finalRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const map = {};
  finalRow.forEach(function (h, i) { map[h.trim().toLowerCase()] = i + 1; });
  return map;
}

function colFor(headers, field) {
  const h = FIELD_TO_HEADER[field];
  return h ? headers[h.toLowerCase()] || 0 : 0;
}

function addRow(sheet, headers, data) {
  const rowNumber = sheet.getLastRow() + 1;
  // Next serial number.
  const snoCol = headers["sno."] || headers["sno"] || 0;
  if (snoCol) {
    let maxSno = 0;
    sheet.getRange(2, snoCol, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().forEach(function (r) {
      const n = parseInt(r[0], 10);
      if (!isNaN(n) && n > maxSno) maxSno = n;
    });
    sheet.getRange(rowNumber, snoCol).setValue(maxSno + 1);
  }
  writeFields(sheet, headers, rowNumber, data);
  return { ok: true, rowNumber: rowNumber };
}

function updateRow(sheet, headers, rowNumber, data) {
  rowNumber = parseInt(rowNumber, 10);
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return { ok: false, error: "Invalid rowNumber: " + rowNumber };
  }
  writeFields(sheet, headers, rowNumber, data);
  return { ok: true, rowNumber: rowNumber };
}

function writeFields(sheet, headers, rowNumber, data) {
  Object.keys(data).forEach(function (field) {
    const col = colFor(headers, field);
    if (col && data[field] !== undefined && data[field] !== null) {
      sheet.getRange(rowNumber, col).setValue(String(data[field]));
    }
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
