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

// ---- One-time cleanup / migration -------------------------------------------
// Run this ONCE from the Apps Script editor: pick "cleanupSheet" in the function
// dropdown and click Run. It reads the messy "Applied" tab and writes a tidy
// "Applied_clean" tab. Nothing is deleted — review the new tab, then copy it over
// "Applied" when you're happy. No redeploy needed (this doesn't affect the web app).
const CLEAN_SHEET_NAME = "Applied_clean";
// The columns, in a sensible order, with the three timeline dates grouped up front so
// you can see at a glance what's approaching. No "Result" column — Status covers the
// outcome. "Others" holds the latest notice/update from the exam conductor.
// Column indexes (1-based) referenced later: Deadline=4, Exam Date=5, Result Date=6, Status=7.
const CLEAN_COLUMNS = [
  "Sno.", "Dept", "Post", "Application Deadline", "Exam Date", "Result Date",
  "Status", "Registration No.", "Password", "Link", "Exam Centre",
  "Admit Card", "Others", "Notes"
];
const DATE_COLS = [4, 5, 6];   // Application Deadline, Exam Date, Result Date
const STATUS_COL = 7;
const CLEAN_STATUSES = [
  "applied", "admit card out", "exam scheduled", "exam done",
  "interview", "qualified", "selected", "not qualified", "withdrawn"
];

function cleanupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SHEET_NAME);
  if (!src) throw new Error("Source sheet '" + SHEET_NAME + "' not found");

  const values = src.getDataRange().getValues();
  if (values.length < 2) throw new Error("Nothing to clean in '" + SHEET_NAME + "'");

  // Map normalized header -> source column index, skipping blank/junk "Column N".
  const idx = {};
  values[0].forEach(function (h, i) {
    const name = String(h).trim();
    if (!name || /^column \d+$/i.test(name)) return;
    idx[normKey(name)] = i;
  });

  function pick(row, aliases) {
    for (let i = 0; i < aliases.length; i++) {
      const c = idx[aliases[i]];
      if (c === undefined) continue;
      const v = String(row[c] == null ? "" : row[c]).trim();
      if (v && !/^(n\/?a|-{1,2}|na)$/i.test(v)) return v;
    }
    return "";
  }

  const out = [CLEAN_COLUMNS.slice()];
  let sno = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dept = pick(row, ["dept", "department", "org", "organisation"]);
    const post = pick(row, ["post", "position", "role"]);
    if (!dept && !post) continue; // drop blank rows

    // Normalize Status; if blank, best-guess from Others/old Result (review these!).
    const oldResult = pick(row, ["result", "outcome"]);
    let status = normStatus(pick(row, ["status", "stage"]));
    if (!status) {
      const cands = [pick(row, ["others"]), oldResult];
      for (let k = 0; k < cands.length; k++) {
        const g = normStatus(cands[k]);
        if (CLEAN_STATUSES.indexOf(g) !== -1) { status = g; break; }
      }
    }

    // The Result column is gone; keep any old result text that ISN'T just a status
    // word we already captured, so nothing is lost — fold it into Notes.
    let notes = pick(row, ["notes", "remark", "remarks"]);
    if (oldResult && CLEAN_STATUSES.indexOf(normStatus(oldResult)) === -1) {
      notes = notes ? notes + " | Result: " + oldResult : "Result: " + oldResult;
    }

    sno++;
    out.push([
      sno,
      dept,
      post,
      pick(row, ["applicationdeadline", "deadline"]),
      pick(row, ["examdate"]),
      pick(row, ["resultdate"]),                       // new, usually blank at first
      status,
      pick(row, ["registrationno", "regno", "registration"]),
      pick(row, ["password", "pwd"]),
      pick(row, ["link", "url", "portal"]),
      pick(row, ["examcentre", "examcenter", "centre", "center"]),
      pick(row, ["admitcard", "admitcardlink"]),
      pick(row, ["others"]),
      notes
    ]);
  }

  let dst = ss.getSheetByName(CLEAN_SHEET_NAME);
  if (dst) dst.clear(); else dst = ss.insertSheet(CLEAN_SHEET_NAME);
  dst.getRange(1, 1, out.length, CLEAN_COLUMNS.length).setValues(out);
  dst.getRange(1, 1, 1, CLEAN_COLUMNS.length).setFontWeight("bold");
  dst.setFrozenRows(1);

  const rows = out.length - 1;
  if (rows > 0) {
    // Date-only display for the three timeline columns (text like "Aug 2026 (tent.)"
    // is left untouched; only real dates get formatted).
    DATE_COLS.forEach(function (c) { dst.getRange(2, c, rows, 1).setNumberFormat("dd MMM yyyy"); });
    applyStatusDropdown(dst, rows);
  }
  dst.autoResizeColumns(1, CLEAN_COLUMNS.length);

  SpreadsheetApp.getActive().toast(
    rows + " rows written to '" + CLEAN_SHEET_NAME + "'. Review, then copy over '" + SHEET_NAME + "'.",
    "Cleanup done", 8
  );
}

// Put a dropdown (data validation) on the Status column with the app's exact options.
function applyStatusDropdown(sheet, rows) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CLEAN_STATUSES, true)
    .setAllowInvalid(false)
    .setHelpText("Pick one of: " + CLEAN_STATUSES.join(", "))
    .build();
  sheet.getRange(2, STATUS_COL, rows, 1).setDataValidation(rule);
}

function normKey(h) { return String(h || "").toLowerCase().replace(/[^a-z]/g, ""); }

function normStatus(raw) {
  const s = String(raw || "").toLowerCase().trim();
  if (!s) return "";
  if (/select/.test(s)) return "selected";
  if (/withdraw/.test(s)) return "withdrawn";
  if (/not\s*qualif|reject|fail|negative|unsuccess/.test(s)) return "not qualified";
  if (/qualif|clear|pass|shortlist/.test(s)) return "qualified";
  if (/interview/.test(s)) return "interview";
  if (/exam.*done|appeared|attempted|written/.test(s)) return "exam done";
  if (/exam|scheduled|phase\s*2|mains|prelims/.test(s)) return "exam scheduled";
  if (/admit|hall\s*ticket/.test(s)) return "admit card out";
  if (/appl|submit|registered/.test(s)) return "applied";
  return ""; // unrecognized -> leave blank for manual review
}

// ---- Auto-update timelines from the internet ---------------------------------
// A plain sheet can't browse the web, so this asks the Gemini API (with Google
// Search grounding) for the latest tentative schedule of each exam and fills the
// three date columns + "Others" (latest notice). Best-effort — always eyeball dates.
//
// Setup (one time):
//   1. Get a free key at https://aistudio.google.com/apikey
//   2. Apps Script -> Project Settings -> Script Properties ->
//        add GEMINI_API_KEY = <your key>
//   3. Run "refreshTimelines" once (approve the URL-fetch permission).
//   4. Optional: run "installTimelineTrigger" once to auto-refresh every morning.
//
// It only fills cells that are blank or marked tentative ("(tent.)"/"approx"/"?"),
// so any firm date you type by hand is never overwritten. Rows with a concluded
// Status (selected / not qualified / withdrawn / qualified) are skipped.
const GEMINI_MODEL = "gemini-2.5-flash";

function refreshTimelines() {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key) throw new Error("Set a GEMINI_API_KEY script property first (free key: aistudio.google.com/apikey).");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CLEAN_SHEET_NAME) || ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("No sheet to update.");

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const c = {};
  values[0].forEach(function (h, i) { c[normKey(h)] = i; });
  if (c["dept"] == null) throw new Error("No 'Dept' column found — run cleanupSheet first.");

  let checked = 0, changed = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const dept = String(row[c["dept"]] || "").trim();
    const post = String(row[c["post"]] || "").trim();
    if (!dept && !post) continue;
    if (/selected|not qualified|withdrawn|qualified/.test(String(row[c["status"]] || "").toLowerCase())) continue;

    const info = fetchExamTimeline(key, dept, post);
    checked++;
    if (!info) continue;

    let touched = false;
    touched = fillIfOpen(sheet, r, c["applicationdeadline"], row, info.applicationDeadline) || touched;
    touched = fillIfOpen(sheet, r, c["examdate"], row, info.examDate) || touched;
    touched = fillIfOpen(sheet, r, c["resultdate"], row, info.resultDate) || touched;
    if (info.latest && c["others"] != null) { sheet.getRange(r + 1, c["others"] + 1).setValue(info.latest); touched = true; }
    if (touched) changed++;
    Utilities.sleep(400); // be gentle with the API
  }
  SpreadsheetApp.getActive().toast("Checked " + checked + " apps, updated " + changed + ".", "Timelines refreshed", 6);
}

// Write value only into a blank or tentative cell, so firm hand-typed dates survive.
function fillIfOpen(sheet, r, colIdx, row, value) {
  if (colIdx == null || !value) return false;
  const cur = String(row[colIdx] || "").trim();
  if (cur && !/tent|approx|\?/i.test(cur)) return false;
  if (cur === String(value).trim()) return false;
  sheet.getRange(r + 1, colIdx + 1).setValue(value);
  return true;
}

function fetchExamTimeline(key, dept, post) {
  const exam = (dept + " " + post).trim();
  const prompt =
    'For the Indian government recruitment exam "' + exam + '", find the most recent official or ' +
    'credible tentative schedule as of today. Reply with ONLY a JSON object (no markdown fences), keys: ' +
    'applicationDeadline, examDate, resultDate, latest. ' +
    'Each date is date-only: "DD Mon YYYY" (e.g. "25 Jul 2026") if a specific date is known, or ' +
    '"Mon YYYY (tent.)" (e.g. "Aug 2026 (tent.)") if only a tentative month is known or it is inferred ' +
    'from last year\'s cycle, otherwise "". No time of day. ' +
    '"latest" = one short sentence (max 20 words) with the newest official notice/update, else "". ' +
    'Use empty strings for anything genuinely unknown; do not guess wildly.';
  const payload = { contents: [{ parts: [{ text: prompt }] }], tools: [{ google_search: {} }] };
  try {
    const res = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/" + GEMINI_MODEL +
        ":generateContent?key=" + encodeURIComponent(key),
      { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true }
    );
    const data = JSON.parse(res.getContentText());
    if (!data.candidates || !data.candidates[0] || !data.candidates[0].content) return null;
    const text = data.candidates[0].content.parts.map(function (p) { return p.text || ""; }).join("");
    const m = text.match(/\{[\s\S]*\}/);
    return m ? JSON.parse(m[0]) : null;
  } catch (err) {
    return null; // network/parse hiccup on one row shouldn't stop the rest
  }
}

// Run once to auto-refresh every morning (~6am). Re-running replaces the old trigger.
function installTimelineTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "refreshTimelines") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("refreshTimelines").timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getActive().toast("Daily timeline auto-refresh installed (~6am).", "Scheduled", 5);
}
