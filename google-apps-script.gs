/**
 * Arkan Group — Tour Expenses Sheet
 * Google Sheets Auto-Save Backend (Google Apps Script)
 *
 * SETUP (1 Minute):
 * 1. Open https://sheets.new to create a new Google Sheet (e.g. "Arkan Tour Expenses DB").
 * 2. Top menu -> Extensions -> Apps Script.
 * 3. Delete any existing code, paste ALL of this code, and press Ctrl+S to save.
 * 4. Click "Deploy" (top right) -> "New deployment".
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click "Deploy", authorize permissions, and copy the Web App URL (ends with /exec).
 * 6. Paste that URL into the Tour Expense Web App -> "Google Sheet" button -> "Test & Connect".
 */

// Optional security PIN (leave empty "" for easy setup, or set e.g. "1234" and enter the same PIN in the app)
const SYNC_KEY = "";
const DB_SHEET_NAME = "_APP_DATABASE";
const CHUNK_SIZE = 40000;

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function isAuthorized(key) {
  if (!SYNC_KEY) return true;
  return String(key || "") === String(SYNC_KEY);
}

function getOrCreateDbSheet(ss) {
  let sheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DB_SHEET_NAME);
    sheet.getRange("A1:C1").setValues([["JSON_CHUNK (Do Not Edit)", "UPDATED_AT", "CHUNK_COUNT"]]);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f1f5f9");
  }
  return sheet;
}

function readStateFromSheet(ss) {
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!sheet) return { state: null, updatedAt: 0 };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { state: null, updatedAt: 0 };

  const updatedAt = Number(sheet.getRange("B2").getValue()) || 0;
  const chunkCount = Number(sheet.getRange("C2").getValue()) || (lastRow - 1);
  if (chunkCount <= 0) return { state: null, updatedAt: 0 };

  const values = sheet.getRange(2, 1, chunkCount, 1).getValues();
  const raw = values.map(function (r) { return String(r[0] || ""); }).join("");
  if (!raw) return { state: null, updatedAt: 0 };

  return { state: JSON.parse(raw), updatedAt: updatedAt };
}

function writeStateToSheet(ss, stateObj, updatedAt) {
  const sheet = getOrCreateDbSheet(ss);
  const jsonStr = JSON.stringify(stateObj);
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {
    chunks.push([jsonStr.slice(i, i + CHUNK_SIZE)]);
  }
  if (!chunks.length) chunks.push(["{}"]);

  const prevLast = sheet.getLastRow();
  if (prevLast >= 2) {
    sheet.getRange(2, 1, prevLast - 1, 3).clearContent();
  }
  sheet.getRange(2, 1, chunks.length, 1).setValues(chunks);
  sheet.getRange("B2").setValue(updatedAt);
  sheet.getRange("C2").setValue(chunks.length);
}

function sanitizeTabName(name) {
  const cleaned = String(name || "Tour Expenses")
    .replace(/[:\\/?*\[\]]/g, " ")
    .trim()
    .slice(0, 90);
  if (!cleaned || cleaned === DB_SHEET_NAME) return "Tour Expenses";
  return cleaned;
}

function parseIsoToDate(iso) {
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso || "";
  const parts = iso.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function numOrBlank(v) {
  if (v === "" || v === null || v === undefined) return "";
  const n = Number(v);
  return (isFinite(n) && n !== 0) ? n : "";
}

function syncReadableMonthTab(ss, stateObj, monthId) {
  if (!stateObj || !Array.isArray(stateObj.months)) return;
  const month = stateObj.months.find(function (m) { return m.id === monthId; }) || stateObj.months[0];
  if (!month) return;

  const profile = stateObj.profile || {};
  const tabName = sanitizeTabName(month.label || month.id);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName, 0);
  } else {
    sheet.clear();
  }

  // Build rows array cleanly in memory for fast single-batch write
  const days = (month.days && month.days.length) ? month.days : [];
  const rows = [];

  rows.push([profile.companyName || "Arkan Group", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  rows.push([profile.reportTitle || "Tour Expenses (Out of Khulna)", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  rows.push(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]);
  rows.push([
    "Destination: " + (month.destination || profile.destination || ""), "", "", "",
    "Duration:", parseIsoToDate(month.startDate), "to", parseIsoToDate(month.endDate),
    "", "", "", "", "", "", "", ""
  ]);
  rows.push([
    "Purpose: " + (month.purpose || profile.purpose || ""), "", "", "",
    "Total Day:", "=H4-F4+1", "", "",
    "Travellers Designation: " + (month.designation || profile.designation || ""),
    "", "", "", "", "", "", ""
  ]);
  rows.push([
    "Date", "Destination", "", "Transportation Expenses", "", "", "", "", "",
    "Meal Expenses", "", "", "", "", "Living Cost", "Grand Total"
  ]);
  rows.push([
    "", "From", "To", "Bus", "Metro/ Rail", "Auto", "Bike", "Rickshaw", "Transport Total",
    "Breakfast", "Lunch", "Dinner", "Tea Break", "Meal Total", "Hotel", "Total Day Expense"
  ]);

  const dayRanges = [];
  let curRow = 8;

  for (let d = 0; d < days.length; d++) {
    const day = days[d];
    const routes = (day.routes && day.routes.length) ? day.routes : [{}];
    const startR = curRow;
    const endR = curRow + routes.length - 1;
    dayRanges.push({ startR: startR, endR: endR });

    for (let r = 0; r < routes.length; r++) {
      const rt = routes[r] || {};
      const rowArr = [
        r === 0 ? parseIsoToDate(day.date) : "",
        rt.from || "",
        rt.to || "",
        numOrBlank(rt.bus),
        numOrBlank(rt.metro),
        numOrBlank(rt.auto),
        numOrBlank(rt.bike),
        numOrBlank(rt.rickshaw),
        r === 0 ? ("=SUM(D" + startR + ":H" + endR + ")") : "",
        r === 0 ? numOrBlank(day.breakfast) : "",
        r === 0 ? numOrBlank(day.lunch) : "",
        r === 0 ? numOrBlank(day.dinner) : "",
        r === 0 ? numOrBlank(day.teaBreak) : "",
        r === 0 ? ("=SUM(J" + startR + ":M" + startR + ")") : "",
        r === 0 ? numOrBlank(day.hotel) : "",
        r === 0 ? ("=SUM(I" + startR + ",N" + startR + ",O" + startR + ")") : ""
      ];
      rows.push(rowArr);
    }
    curRow = endR + 1;
  }

  const lastDataRow = Math.max(8, curRow - 1);
  const totalRow = curRow;
  const cols = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"];
  const totalArr = ["Total", "", ""];
  for (let c = 0; c < cols.length; c++) {
    totalArr.push(days.length ? ("=SUM(" + cols[c] + "8:" + cols[c] + lastDataRow + ")") : 0);
  }
  rows.push(totalArr);

  // 3 blank rows + signature row
  rows.push(new Array(16).fill(""));
  rows.push(new Array(16).fill(""));
  rows.push(new Array(16).fill(""));
  const sigLabels = profile.signatureLabels || ["Prepared by", "Dept. Head", "Audit Dept.", "Accounts", "GM (Operation)", "Director"];
  const sigRowArr = new Array(16).fill("");
  sigRowArr[1] = sigLabels[0] || "Prepared by";
  sigRowArr[3] = sigLabels[1] || "Dept. Head";
  sigRowArr[5] = sigLabels[2] || "Audit Dept.";
  sigRowArr[7] = sigLabels[3] || "Accounts";
  sigRowArr[9] = sigLabels[4] || "GM (Operation)";
  sigRowArr[13] = sigLabels[5] || "Director";
  rows.push(sigRowArr);

  // Write values in one batch
  sheet.getRange(1, 1, rows.length, 16).setValues(rows);

  // Styling
  sheet.getRange(1, 1, rows.length, 16).setFontFamily("Calibri").setFontSize(11);
  sheet.getRange("A1:P1").merge().setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A2:P2").merge().setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("A4:P5").setFontSize(12);
  sheet.getRange("F4").setNumberFormat("d-m-yyyy").setHorizontalAlignment("right");
  sheet.getRange("G4").setHorizontalAlignment("center");
  sheet.getRange("H4").setNumberFormat("d-m-yyyy");
  sheet.getRange("F5").setHorizontalAlignment("left");

  // Header merges & style
  sheet.getRange("A6:A7").merge();
  sheet.getRange("B6:C6").merge();
  sheet.getRange("D6:I6").merge();
  sheet.getRange("J6:N6").merge();
  sheet.getRange("A6:P7")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  // Table borders & alignments
  sheet.getRange(6, 1, totalRow - 5, 16)
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

  if (days.length) {
    sheet.getRange(8, 1, lastDataRow - 7, 1).setNumberFormat("d/m/yyyy").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(8, 2, lastDataRow - 7, 2).setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheet.getRange(8, 4, lastDataRow - 7, 13).setHorizontalAlignment("center").setVerticalAlignment("middle");

    for (let i = 0; i < dayRanges.length; i++) {
      const rg = dayRanges[i];
      if (rg.endR > rg.startR) {
        for (let colIdx = 9; colIdx <= 16; colIdx++) {
          sheet.getRange(rg.startR, colIdx, rg.endR - rg.startR + 1, 1).merge();
        }
      }
    }
  }

  sheet.getRange(totalRow, 1, 1, 16).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange(totalRow + 4, 1, 1, 16).setHorizontalAlignment("center");
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    if (!isAuthorized(params.key)) {
      return jsonResponse({ ok: false, error: "Invalid Secret PIN / Key" });
    }
    const action = params.action || "ping";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "ping") {
      const existing = readStateFromSheet(ss);
      return jsonResponse({
        ok: true,
        sheetTitle: ss.getName(),
        hasData: Boolean(existing.state),
        updatedAt: existing.updatedAt || 0
      });
    }

    if (action === "load") {
      const existing = readStateFromSheet(ss);
      return jsonResponse({
        ok: true,
        sheetTitle: ss.getName(),
        state: existing.state,
        updatedAt: existing.updatedAt || 0
      });
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const bodyText = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const payload = JSON.parse(bodyText);

    if (!isAuthorized(payload.key)) {
      return jsonResponse({ ok: false, error: "Invalid Secret PIN / Key" });
    }

    const action = payload.action || "save";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "save") {
      if (!payload.state || !Array.isArray(payload.state.months)) {
        return jsonResponse({ ok: false, error: "Missing valid state payload" });
      }
      const updatedAt = Number(payload.updatedAt) || Date.now();
      writeStateToSheet(ss, payload.state, updatedAt);

      // Also update the active month's formatted Google Sheet tab!
      if (payload.syncReadable !== false) {
        try {
          if (payload.syncAllMonths) {
            for (let i = 0; i < payload.state.months.length; i++) {
              syncReadableMonthTab(ss, payload.state, payload.state.months[i].id);
            }
          } else {
            syncReadableMonthTab(ss, payload.state, payload.state.activeMonthId);
          }
        } catch (tabErr) {
          // Core JSON DB already saved safely; ignore non-fatal tab format error
        }
      }

      return jsonResponse({
        ok: true,
        sheetTitle: ss.getName(),
        updatedAt: updatedAt
      });
    }

    return jsonResponse({ ok: false, error: "Unknown POST action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
