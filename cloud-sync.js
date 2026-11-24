/**
 * Google Sheets Cloud Auto-Save Module
 * Uses CORS-safelisted text/plain POST + GET to sync state & formatted monthly tabs
 * directly with the user's own free Google Sheet via Google Apps Script.
 */
(function () {
  'use strict';

  const CLOUD_CFG_KEY = 'arkan_tour_cloud_cfg_v1';
  const APPS_SCRIPT_CODE = "/**\n * Arkan Group — Tour Expenses Sheet\n * Google Sheets Auto-Save Backend (Google Apps Script)\n *\n * SETUP (1 Minute):\n * 1. Open https://sheets.new to create a new Google Sheet (e.g. \"Arkan Tour Expenses DB\").\n * 2. Top menu -> Extensions -> Apps Script.\n * 3. Delete any existing code, paste ALL of this code, and press Ctrl+S to save.\n * 4. Click \"Deploy\" (top right) -> \"New deployment\".\n *    - Type: Web app\n *    - Execute as: Me\n *    - Who has access: Anyone\n * 5. Click \"Deploy\", authorize permissions, and copy the Web App URL (ends with /exec).\n * 6. Paste that URL into the Tour Expense Web App -> \"Google Sheet\" button -> \"Test & Connect\".\n */\n\n// Optional security PIN (leave empty \"\" for easy setup, or set e.g. \"1234\" and enter the same PIN in the app)\nconst SYNC_KEY = \"\";\nconst DB_SHEET_NAME = \"_APP_DATABASE\";\nconst CHUNK_SIZE = 40000;\n\nfunction jsonResponse(obj) {\n  return ContentService\n    .createTextOutput(JSON.stringify(obj))\n    .setMimeType(ContentService.MimeType.JSON);\n}\n\nfunction isAuthorized(key) {\n  if (!SYNC_KEY) return true;\n  return String(key || \"\") === String(SYNC_KEY);\n}\n\nfunction getOrCreateDbSheet(ss) {\n  let sheet = ss.getSheetByName(DB_SHEET_NAME);\n  if (!sheet) {\n    sheet = ss.insertSheet(DB_SHEET_NAME);\n    sheet.getRange(\"A1:C1\").setValues([[\"JSON_CHUNK (Do Not Edit)\", \"UPDATED_AT\", \"CHUNK_COUNT\"]]);\n    sheet.getRange(\"A1:C1\").setFontWeight(\"bold\").setBackground(\"#f1f5f9\");\n  }\n  return sheet;\n}\n\nfunction readStateFromSheet(ss) {\n  const sheet = ss.getSheetByName(DB_SHEET_NAME);\n  if (!sheet) return { state: null, updatedAt: 0 };\n\n  const lastRow = sheet.getLastRow();\n  if (lastRow < 2) return { state: null, updatedAt: 0 };\n\n  const updatedAt = Number(sheet.getRange(\"B2\").getValue()) || 0;\n  const chunkCount = Number(sheet.getRange(\"C2\").getValue()) || (lastRow - 1);\n  if (chunkCount <= 0) return { state: null, updatedAt: 0 };\n\n  const values = sheet.getRange(2, 1, chunkCount, 1).getValues();\n  const raw = values.map(function (r) { return String(r[0] || \"\"); }).join(\"\");\n  if (!raw) return { state: null, updatedAt: 0 };\n\n  return { state: JSON.parse(raw), updatedAt: updatedAt };\n}\n\nfunction writeStateToSheet(ss, stateObj, updatedAt) {\n  const sheet = getOrCreateDbSheet(ss);\n  const jsonStr = JSON.stringify(stateObj);\n  const chunks = [];\n  for (let i = 0; i < jsonStr.length; i += CHUNK_SIZE) {\n    chunks.push([jsonStr.slice(i, i + CHUNK_SIZE)]);\n  }\n  if (!chunks.length) chunks.push([\"{}\"]);\n\n  const prevLast = sheet.getLastRow();\n  if (prevLast >= 2) {\n    sheet.getRange(2, 1, prevLast - 1, 3).clearContent();\n  }\n  sheet.getRange(2, 1, chunks.length, 1).setValues(chunks);\n  sheet.getRange(\"B2\").setValue(updatedAt);\n  sheet.getRange(\"C2\").setValue(chunks.length);\n}\n\nfunction sanitizeTabName(name) {\n  const cleaned = String(name || \"Tour Expenses\")\n    .replace(/[:\\\\/?*\\[\\]]/g, \" \")\n    .trim()\n    .slice(0, 90);\n  if (!cleaned || cleaned === DB_SHEET_NAME) return \"Tour Expenses\";\n  return cleaned;\n}\n\nfunction parseIsoToDate(iso) {\n  if (!iso || !/^\\d{4}-\\d{2}-\\d{2}$/.test(iso)) return iso || \"\";\n  const parts = iso.split(\"-\").map(Number);\n  return new Date(parts[0], parts[1] - 1, parts[2]);\n}\n\nfunction numOrBlank(v) {\n  if (v === \"\" || v === null || v === undefined) return \"\";\n  const n = Number(v);\n  return (isFinite(n) && n !== 0) ? n : \"\";\n}\n\nfunction syncReadableMonthTab(ss, stateObj, monthId) {\n  if (!stateObj || !Array.isArray(stateObj.months)) return;\n  const month = stateObj.months.find(function (m) { return m.id === monthId; }) || stateObj.months[0];\n  if (!month) return;\n\n  const profile = stateObj.profile || {};\n  const tabName = sanitizeTabName(month.label || month.id);\n  let sheet = ss.getSheetByName(tabName);\n  if (!sheet) {\n    sheet = ss.insertSheet(tabName, 0);\n  } else {\n    sheet.clear();\n  }\n\n  // Build rows array cleanly in memory for fast single-batch write\n  const days = (month.days && month.days.length) ? month.days : [];\n  const rows = [];\n\n  rows.push([profile.companyName || \"Arkan Group\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\"]);\n  rows.push([profile.reportTitle || \"Tour Expenses (Out of Khulna)\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\"]);\n  rows.push([\"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\"]);\n  rows.push([\n    \"Destination: \" + (month.destination || profile.destination || \"\"), \"\", \"\", \"\",\n    \"Duration:\", parseIsoToDate(month.startDate), \"to\", parseIsoToDate(month.endDate),\n    \"\", \"\", \"\", \"\", \"\", \"\", \"\", \"\"\n  ]);\n  rows.push([\n    \"Purpose: \" + (month.purpose || profile.purpose || \"\"), \"\", \"\", \"\",\n    \"Total Day:\", \"=H4-F4+1\", \"\", \"\",\n    \"Travellers Designation: \" + (month.designation || profile.designation || \"\"),\n    \"\", \"\", \"\", \"\", \"\", \"\", \"\"\n  ]);\n  rows.push([\n    \"Date\", \"Destination\", \"\", \"Transportation Expenses\", \"\", \"\", \"\", \"\", \"\",\n    \"Meal Expenses\", \"\", \"\", \"\", \"\", \"Living Cost\", \"Grand Total\"\n  ]);\n  rows.push([\n    \"\", \"From\", \"To\", \"Bus\", \"Metro/ Rail\", \"Auto\", \"Bike\", \"Rickshaw\", \"Transport Total\",\n    \"Breakfast\", \"Lunch\", \"Dinner\", \"Tea Break\", \"Meal Total\", \"Hotel\", \"Total Day Expense\"\n  ]);\n\n  const dayRanges = [];\n  let curRow = 8;\n\n  for (let d = 0; d < days.length; d++) {\n    const day = days[d];\n    const routes = (day.routes && day.routes.length) ? day.routes : [{}];\n    const startR = curRow;\n    const endR = curRow + routes.length - 1;\n    dayRanges.push({ startR: startR, endR: endR });\n\n    for (let r = 0; r < routes.length; r++) {\n      const rt = routes[r] || {};\n      const rowArr = [\n        r === 0 ? parseIsoToDate(day.date) : \"\",\n        rt.from || \"\",\n        rt.to || \"\",\n        numOrBlank(rt.bus),\n        numOrBlank(rt.metro),\n        numOrBlank(rt.auto),\n        numOrBlank(rt.bike),\n        numOrBlank(rt.rickshaw),\n        r === 0 ? (\"=SUM(D\" + startR + \":H\" + endR + \")\") : \"\",\n        r === 0 ? numOrBlank(day.breakfast) : \"\",\n        r === 0 ? numOrBlank(day.lunch) : \"\",\n        r === 0 ? numOrBlank(day.dinner) : \"\",\n        r === 0 ? numOrBlank(day.teaBreak) : \"\",\n        r === 0 ? (\"=SUM(J\" + startR + \":M\" + startR + \")\") : \"\",\n        r === 0 ? numOrBlank(day.hotel) : \"\",\n        r === 0 ? (\"=SUM(I\" + startR + \",N\" + startR + \",O\" + startR + \")\") : \"\"\n      ];\n      rows.push(rowArr);\n    }\n    curRow = endR + 1;\n  }\n\n  const lastDataRow = Math.max(8, curRow - 1);\n  const totalRow = curRow;\n  const cols = [\"D\", \"E\", \"F\", \"G\", \"H\", \"I\", \"J\", \"K\", \"L\", \"M\", \"N\", \"O\", \"P\"];\n  const totalArr = [\"Total\", \"\", \"\"];\n  for (let c = 0; c < cols.length; c++) {\n    totalArr.push(days.length ? (\"=SUM(\" + cols[c] + \"8:\" + cols[c] + lastDataRow + \")\") : 0);\n  }\n  rows.push(totalArr);\n\n  // 3 blank rows + signature row\n  rows.push(new Array(16).fill(\"\"));\n  rows.push(new Array(16).fill(\"\"));\n  rows.push(new Array(16).fill(\"\"));\n  const sigLabels = profile.signatureLabels || [\"Prepared by\", \"Dept. Head\", \"Audit Dept.\", \"Accounts\", \"GM (Operation)\", \"Director\"];\n  const sigRowArr = new Array(16).fill(\"\");\n  sigRowArr[1] = sigLabels[0] || \"Prepared by\";\n  sigRowArr[3] = sigLabels[1] || \"Dept. Head\";\n  sigRowArr[5] = sigLabels[2] || \"Audit Dept.\";\n  sigRowArr[7] = sigLabels[3] || \"Accounts\";\n  sigRowArr[9] = sigLabels[4] || \"GM (Operation)\";\n  sigRowArr[13] = sigLabels[5] || \"Director\";\n  rows.push(sigRowArr);\n\n  // Write values in one batch\n  sheet.getRange(1, 1, rows.length, 16).setValues(rows);\n\n  // Styling\n  sheet.getRange(1, 1, rows.length, 16).setFontFamily(\"Calibri\").setFontSize(11);\n  sheet.getRange(\"A1:P1\").merge().setFontSize(14).setFontWeight(\"bold\").setHorizontalAlignment(\"center\");\n  sheet.getRange(\"A2:P2\").merge().setFontSize(12).setFontWeight(\"bold\").setHorizontalAlignment(\"center\");\n  sheet.getRange(\"A4:P5\").setFontSize(12);\n  sheet.getRange(\"F4\").setNumberFormat(\"d-m-yyyy\").setHorizontalAlignment(\"right\");\n  sheet.getRange(\"G4\").setHorizontalAlignment(\"center\");\n  sheet.getRange(\"H4\").setNumberFormat(\"d-m-yyyy\");\n  sheet.getRange(\"F5\").setHorizontalAlignment(\"left\");\n\n  // Header merges & style\n  sheet.getRange(\"A6:A7\").merge();\n  sheet.getRange(\"B6:C6\").merge();\n  sheet.getRange(\"D6:I6\").merge();\n  sheet.getRange(\"J6:N6\").merge();\n  sheet.getRange(\"A6:P7\")\n    .setFontWeight(\"bold\")\n    .setHorizontalAlignment(\"center\")\n    .setVerticalAlignment(\"middle\")\n    .setWrap(true);\n\n  // Table borders & alignments\n  sheet.getRange(6, 1, totalRow - 5, 16)\n    .setBorder(true, true, true, true, true, true, \"#000000\", SpreadsheetApp.BorderStyle.SOLID);\n\n  if (days.length) {\n    sheet.getRange(8, 1, lastDataRow - 7, 1).setNumberFormat(\"d/m/yyyy\").setHorizontalAlignment(\"center\").setVerticalAlignment(\"middle\");\n    sheet.getRange(8, 2, lastDataRow - 7, 2).setHorizontalAlignment(\"left\").setVerticalAlignment(\"middle\");\n    sheet.getRange(8, 4, lastDataRow - 7, 13).setHorizontalAlignment(\"center\").setVerticalAlignment(\"middle\");\n\n    for (let i = 0; i < dayRanges.length; i++) {\n      const rg = dayRanges[i];\n      if (rg.endR > rg.startR) {\n        for (let colIdx = 9; colIdx <= 16; colIdx++) {\n          sheet.getRange(rg.startR, colIdx, rg.endR - rg.startR + 1, 1).merge();\n        }\n      }\n    }\n  }\n\n  sheet.getRange(totalRow, 1, 1, 16).setFontWeight(\"bold\").setHorizontalAlignment(\"center\").setVerticalAlignment(\"middle\");\n  sheet.getRange(totalRow + 4, 1, 1, 16).setHorizontalAlignment(\"center\");\n}\n\nfunction doGet(e) {\n  try {\n    const params = (e && e.parameter) ? e.parameter : {};\n    if (!isAuthorized(params.key)) {\n      return jsonResponse({ ok: false, error: \"Invalid Secret PIN / Key\" });\n    }\n    const action = params.action || \"ping\";\n    const ss = SpreadsheetApp.getActiveSpreadsheet();\n\n    if (action === \"ping\") {\n      const existing = readStateFromSheet(ss);\n      return jsonResponse({\n        ok: true,\n        sheetTitle: ss.getName(),\n        hasData: Boolean(existing.state),\n        updatedAt: existing.updatedAt || 0\n      });\n    }\n\n    if (action === \"load\") {\n      const existing = readStateFromSheet(ss);\n      return jsonResponse({\n        ok: true,\n        sheetTitle: ss.getName(),\n        state: existing.state,\n        updatedAt: existing.updatedAt || 0\n      });\n    }\n\n    return jsonResponse({ ok: false, error: \"Unknown action: \" + action });\n  } catch (err) {\n    return jsonResponse({ ok: false, error: String(err) });\n  }\n}\n\nfunction doPost(e) {\n  const lock = LockService.getScriptLock();\n  try {\n    lock.waitLock(15000);\n    const bodyText = (e && e.postData && e.postData.contents) ? e.postData.contents : \"{}\";\n    const payload = JSON.parse(bodyText);\n\n    if (!isAuthorized(payload.key)) {\n      return jsonResponse({ ok: false, error: \"Invalid Secret PIN / Key\" });\n    }\n\n    const action = payload.action || \"save\";\n    const ss = SpreadsheetApp.getActiveSpreadsheet();\n\n    if (action === \"save\") {\n      if (!payload.state || !Array.isArray(payload.state.months)) {\n        return jsonResponse({ ok: false, error: \"Missing valid state payload\" });\n      }\n      const updatedAt = Number(payload.updatedAt) || Date.now();\n      writeStateToSheet(ss, payload.state, updatedAt);\n\n      // Also update the active month's formatted Google Sheet tab!\n      if (payload.syncReadable !== false) {\n        try {\n          if (payload.syncAllMonths) {\n            for (let i = 0; i < payload.state.months.length; i++) {\n              syncReadableMonthTab(ss, payload.state, payload.state.months[i].id);\n            }\n          } else {\n            syncReadableMonthTab(ss, payload.state, payload.state.activeMonthId);\n          }\n        } catch (tabErr) {\n          // Core JSON DB already saved safely; ignore non-fatal tab format error\n        }\n      }\n\n      return jsonResponse({\n        ok: true,\n        sheetTitle: ss.getName(),\n        updatedAt: updatedAt\n      });\n    }\n\n    return jsonResponse({ ok: false, error: \"Unknown POST action: \" + action });\n  } catch (err) {\n    return jsonResponse({ ok: false, error: String(err) });\n  } finally {\n    try { lock.releaseLock(); } catch (e2) {}\n  }\n}\n";

  const $ = (id) => document.getElementById(id);

  // Keep empty in public repositories so your private Google Sheet endpoint is never exposed.
  // Enter your Web App URL + Secret PIN inside the app's "Google Sheet" modal (stored only in your browser).
  const DEFAULT_WEB_APP_URL = '';

  let cfg = {
    webAppUrl: DEFAULT_WEB_APP_URL,
    secretKey: '',
    autoSave: true,
    autoLoad: true,
    syncReadable: true,
    sheetTitle: '',
    lastSyncedAt: 0
  };

  let callbacks = {
    getState: null,
    getLocalUpdatedAt: null,
    applyCloudState: null,
    showToast: null
  };

  let debounceTimer = null;
  let inFlight = false;
  let pendingSave = null; // { state, updatedAt, syncAllMonths }

  function loadConfig() {
    try {
      const raw = localStorage.getItem(CLOUD_CFG_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        cfg = Object.assign({}, cfg, parsed);
        if (!cfg.webAppUrl && !parsed.disconnectedManually) {
          cfg.webAppUrl = DEFAULT_WEB_APP_URL;
        }
      }
    } catch (e) {}
  }

  function saveConfig() {
    try {
      localStorage.setItem(CLOUD_CFG_KEY, JSON.stringify(cfg));
    } catch (e) {}
  }

  function isConnected() {
    return Boolean(cfg.webAppUrl && cfg.webAppUrl.trim());
  }

  function fmtTime(ts) {
    if (!ts) return '';
    try {
      return new Date(ts).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    } catch (e) {
      return '';
    }
  }

  function setStatus(status, detailText) {
    const dot = $('cloudStatusDot');
    const btnText = $('cloudBtnText');
    const modalStatus = $('cloudModalStatus');
    const saveStatus = $('saveStatus');

    if (dot) {
      dot.className = 'cloud-dot cloud-dot-' + status;
    }

    if (btnText) {
      if (status === 'disconnected') btnText.textContent = 'Google Sheet';
      else if (status === 'syncing') btnText.textContent = 'Syncing...';
      else if (status === 'synced') btnText.textContent = 'Sheet Synced';
      else if (status === 'error') btnText.textContent = 'Sync Error';
    }

    if (modalStatus) {
      modalStatus.className = 'cloud-status-banner cloud-banner-' + status;
      if (detailText) {
        modalStatus.textContent = detailText;
      } else if (status === 'disconnected') {
        modalStatus.textContent = 'Not connected — নিচের ৪টি ধাপ অনুসরণ করে আপনার Google Sheet কানেক্ট করুন।';
      } else if (status === 'synced') {
        const t = fmtTime(cfg.lastSyncedAt);
        modalStatus.textContent = `Connected to "${cfg.sheetTitle || 'Google Sheet'}"${t ? ' • Last synced ' + t : ''}`;
      } else if (status === 'syncing') {
        modalStatus.textContent = 'Google Sheet-এ ডেটা সিঙ্ক হচ্ছে...';
      } else if (status === 'error') {
        modalStatus.textContent = 'Sync failed — ইন্টারনেট বা Web App URL চেক করুন।';
      }
    }

    if (saveStatus && isConnected()) {
      if (status === 'syncing') {
        saveStatus.textContent = 'Saving to Google Sheet...';
        saveStatus.style.color = '#1d4ed8';
      } else if (status === 'synced') {
        const t = fmtTime(cfg.lastSyncedAt);
        saveStatus.textContent = `Synced to Google Sheet${t ? ' (' + t + ')' : ''}`;
        saveStatus.style.color = '#059669';
      } else if (status === 'error') {
        saveStatus.textContent = 'Saved locally (Cloud sync error)';
        saveStatus.style.color = '#d97706';
      }
    }
  }

  function validateUrl(url) {
    const u = (url || '').trim();
    if (!u) return { ok: false, error: 'অনুগ্রহ করে Apps Script Web App URL পেস্ট করুন।' };
    if (u.includes('docs.google.com/spreadsheets')) {
      return {
        ok: false,
        error: 'এটি সরাসরি Google Sheet-এর লিংক! অনুগ্রহ করে Extensions → Apps Script → Deploy → Web App থেকে পাওয়া "https://script.google.com/macros/s/.../exec" লিংকটি দিন।'
      };
    }
    if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec/i.test(u)) {
      return {
        ok: false,
        error: 'URL টি অবশ্যই "https://script.google.com/macros/s/.../exec" দিয়ে শেষ হতে হবে।'
      };
    }
    return { ok: true, url: u };
  }

  async function apiGet(url, key, action) {
    const sep = url.includes('?') ? '&' : '?';
    const fullUrl = `${url}${sep}action=${encodeURIComponent(action)}&key=${encodeURIComponent(key || '')}&_t=${Date.now()}`;
    const res = await fetch(fullUrl, { method: 'GET', redirect: 'follow' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data || !data.ok) throw new Error((data && data.error) || 'Unknown Google Apps Script error');
    return data;
  }

  async function apiPostSave(url, key, stateObj, updatedAt, syncAllMonths) {
    const res = await fetch(url, {
      method: 'POST',
      redirect: 'follow',
      // text/plain avoids CORS preflight OPTIONS which Apps Script doesn't support
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'save',
        key: key || '',
        updatedAt: updatedAt || Date.now(),
        syncReadable: cfg.syncReadable !== false,
        syncAllMonths: Boolean(syncAllMonths),
        state: stateObj
      })
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data || !data.ok) throw new Error((data && data.error) || 'Failed to save to Google Sheet');
    return data;
  }

  async function flushSave(stateObj, updatedAt, syncAllMonths) {
    if (!isConnected()) return;
    if (inFlight) {
      pendingSave = { state: stateObj, updatedAt, syncAllMonths: Boolean(syncAllMonths) };
      return;
    }

    inFlight = true;
    setStatus('syncing');

    try {
      const data = await apiPostSave(cfg.webAppUrl, cfg.secretKey, stateObj, updatedAt, syncAllMonths);
      cfg.sheetTitle = data.sheetTitle || cfg.sheetTitle;
      cfg.lastSyncedAt = data.updatedAt || Date.now();
      saveConfig();
      setStatus('synced');
      return data;
    } catch (err) {
      console.error('Google Sheet save error:', err);
      setStatus('error', 'Sync error: ' + err.message);
      throw err;
    } finally {
      inFlight = false;
      if (pendingSave) {
        const next = pendingSave;
        pendingSave = null;
        flushSave(next.state, next.updatedAt, next.syncAllMonths).catch(() => {});
      }
    }
  }

  function scheduleAutoSave(stateObj, updatedAt) {
    if (!isConnected() || !cfg.autoSave) return;
    if (debounceTimer) clearTimeout(debounceTimer);
    const saveStatus = $('saveStatus');
    if (saveStatus) {
      saveStatus.textContent = 'Queued for Google Sheet...';
      saveStatus.style.color = '#1d4ed8';
    }
    debounceTimer = setTimeout(() => {
      flushSave(stateObj, updatedAt, false).catch(() => {});
    }, 1800);
  }

  async function loadFromCloud(forceApply = false) {
    if (!isConnected()) return null;
    setStatus('syncing', 'Google Sheet থেকে ডেটা লোড হচ্ছে...');
    try {
      const data = await apiGet(cfg.webAppUrl, cfg.secretKey, 'load');
      cfg.sheetTitle = data.sheetTitle || cfg.sheetTitle;
      cfg.lastSyncedAt = data.updatedAt || cfg.lastSyncedAt;
      saveConfig();

      if (data.state && Array.isArray(data.state.months) && data.state.months.length) {
        const localUpdated = callbacks.getLocalUpdatedAt ? callbacks.getLocalUpdatedAt() : 0;
        if (forceApply || (data.updatedAt && data.updatedAt > localUpdated)) {
          if (callbacks.applyCloudState) {
            callbacks.applyCloudState(data.state, data.updatedAt || Date.now());
          }
          setStatus('synced');
          if (callbacks.showToast) {
            callbacks.showToast(`Loaded from Google Sheet ("${cfg.sheetTitle || 'DB'}")`);
          }
          return { loaded: true, data };
        }
      } else if (!forceApply && callbacks.getState) {
        // Cloud sheet is blank on first startup; push local state up automatically
        const curState = callbacks.getState();
        const ts = Date.now();
        await flushSave(curState, ts, true);
        return { loaded: false, pushedInitial: true };
      }

      setStatus('synced');
      return { loaded: false, data };
    } catch (err) {
      console.error('Google Sheet load error:', err);
      setStatus('error', 'Load error: ' + err.message);
      throw err;
    }
  }

  function syncModalFields() {
    $('cloudUrlInput').value = cfg.webAppUrl || '';
    $('cloudPinInput').value = cfg.secretKey || '';
    $('cloudAutoSaveCheck').checked = cfg.autoSave !== false;
    $('cloudAutoLoadCheck').checked = cfg.autoLoad !== false;
    $('cloudReadableCheck').checked = cfg.syncReadable !== false;
    $('cloudScriptCodeBox').textContent = APPS_SCRIPT_CODE;

    const connected = isConnected();
    $('btnCloudSaveNow').disabled = !connected;
    $('btnCloudLoadNow').disabled = !connected;
    $('btnCloudDisconnect').hidden = !connected;
    setStatus(connected ? 'synced' : 'disconnected');
  }

  async function copyTextSafe(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      try {
        await navigator.clipboard.writeText(text);
        return true;
      } catch (e) {}
    }
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) {}
    document.body.removeChild(ta);
    return ok;
  }

  function bindModalEvents() {
    $('btnCloudSync').addEventListener('click', () => {
      syncModalFields();
      $('cloudModal').hidden = false;
    });

    $('btnCopyAppsScript').addEventListener('click', async () => {
      const ok = await copyTextSafe(APPS_SCRIPT_CODE);
      if (callbacks.showToast) {
        callbacks.showToast(ok ? 'Apps Script কোড কপি হয়েছে! এখন Apps Script-এ পেস্ট করুন।' : 'Copy failed — নিচের বক্স থেকে ম্যানুয়ালি কপি করুন।');
      }
    });

    $('cloudAutoSaveCheck').addEventListener('change', (e) => {
      cfg.autoSave = e.target.checked;
      saveConfig();
    });
    $('cloudAutoLoadCheck').addEventListener('change', (e) => {
      cfg.autoLoad = e.target.checked;
      saveConfig();
    });
    $('cloudReadableCheck').addEventListener('change', (e) => {
      cfg.syncReadable = e.target.checked;
      saveConfig();
    });

    $('btnCloudTestConnect').addEventListener('click', async () => {
      const urlVal = $('cloudUrlInput').value.trim();
      const pinVal = $('cloudPinInput').value.trim();
      const check = validateUrl(urlVal);
      if (!check.ok) {
        setStatus('error', check.error);
        return;
      }

      setStatus('syncing', 'কানেকশন পরীক্ষা করা হচ্ছে...');
      try {
        const pingData = await apiGet(check.url, pinVal, 'ping');
        cfg.webAppUrl = check.url;
        cfg.secretKey = pinVal;
        cfg.sheetTitle = pingData.sheetTitle || 'Google Sheet';
        cfg.autoSave = $('cloudAutoSaveCheck').checked;
        cfg.autoLoad = $('cloudAutoLoadCheck').checked;
        cfg.syncReadable = $('cloudReadableCheck').checked;
        cfg.disconnectedManually = false;
        saveConfig();
        syncModalFields();

        if (pingData.hasData) {
          const wantLoad = confirm(
            `Google Sheet "${cfg.sheetTitle}" সফলভাবে কানেক্ট হয়েছে এবং সেখানে আগে থেকে সেভ করা ডেটা আছে।\n\nOK চাপলে শিট থেকে ডেটা লোড হবে।\nCancel চাপলে বর্তমান স্ক্রিনের ডেটা শিটে সেভ হয়ে যাবে।`
          );
          if (wantLoad) {
            await loadFromCloud(true);
          } else if (callbacks.getState) {
            await flushSave(callbacks.getState(), Date.now(), true);
            if (callbacks.showToast) callbacks.showToast(`Connected & saved to "${cfg.sheetTitle}"`);
          }
        } else if (callbacks.getState) {
          await flushSave(callbacks.getState(), Date.now(), true);
          if (callbacks.showToast) callbacks.showToast(`Connected & synced to "${cfg.sheetTitle}"`);
        }
        syncModalFields();
      } catch (err) {
        setStatus('error', 'কানেক্ট হয়নি: ' + err.message + ' — Deploy করার সময় "Who has access: Anyone" দিয়েছেন কি না চেক করুন।');
      }
    });

    $('btnCloudSaveNow').addEventListener('click', async () => {
      if (!isConnected() || !callbacks.getState) return;
      try {
        await flushSave(callbacks.getState(), Date.now(), true);
        if (callbacks.showToast) callbacks.showToast(`Saved all months to "${cfg.sheetTitle || 'Google Sheet'}"`);
      } catch (err) {
        alert('Google Sheet-এ সেভ ব্যর্থ হয়েছে: ' + err.message);
      }
    });

    $('btnCloudLoadNow').addEventListener('click', async () => {
      if (!isConnected()) return;
      if (!confirm('Google Sheet থেকে সর্বশেষ সেভ করা ডেটা লোড করবেন? এতে লোকাল আনসিঙ্কড পরিবর্তন থাকলে তা রিপ্লেস হবে।')) return;
      try {
        const res = await loadFromCloud(true);
        if (res && !res.loaded) {
          alert('Google Sheet-এ এখনো কোনো সেভ করা ডেটা পাওয়া যায়নি।');
        }
      } catch (err) {
        alert('Google Sheet থেকে লোড ব্যর্থ হয়েছে: ' + err.message);
      }
    });

    $('btnCloudDisconnect').addEventListener('click', () => {
      if (!confirm('Google Sheet কানেকশন বিচ্ছিন্ন করবেন? আপনার লোকাল ডেটা ও Google Sheet-এর ডেটা দুটোই নিরাপদ থাকবে।')) return;
      cfg.webAppUrl = '';
      cfg.secretKey = '';
      cfg.sheetTitle = '';
      cfg.lastSyncedAt = 0;
      cfg.disconnectedManually = true;
      saveConfig();
      syncModalFields();
      const saveStatus = $('saveStatus');
      if (saveStatus) {
        saveStatus.textContent = 'Saved locally';
        saveStatus.style.color = '#059669';
      }
      if (callbacks.showToast) callbacks.showToast('Disconnected from Google Sheet');
    });

    const pinToggleBtn = $('btnTogglePinVis');
    if (pinToggleBtn) {
      pinToggleBtn.addEventListener('click', () => {
        const inp = $('cloudPinInput');
        if (!inp) return;
        const showing = inp.type === 'text';
        inp.type = showing ? 'password' : 'text';
        pinToggleBtn.textContent = showing ? 'Show' : 'Hide';
      });
    }

    const clearDeviceBtn = $('btnCloudClearDevice');
    if (clearDeviceBtn) {
      clearDeviceBtn.addEventListener('click', () => {
        if (!confirm('এই ব্রাউজার থেকে সব লোকাল ডেটা, সিগনেচার, Google Sheet লিংক ও PIN মুছে ফেলবেন?\n\n(আপনার Google Sheet-এ সেভ থাকা ডেটা অক্ষত থাকবে।)')) return;
        try {
          localStorage.removeItem('arkan_tour_expenses_v1');
          localStorage.removeItem(CLOUD_CFG_KEY);
        } catch (e) {}
        location.reload();
      });
    }
  }

  window.TourCloudSync = {
    init(cbs) {
      callbacks = Object.assign(callbacks, cbs || {});
      loadConfig();
      bindModalEvents();
      if (isConnected()) {
        setStatus('synced');
        if (cfg.autoLoad) {
          loadFromCloud(false).catch(() => {});
        }
      } else {
        setStatus('disconnected');
      }
    },
    notifyLocalChange(stateObj, updatedAt) {
      scheduleAutoSave(stateObj, updatedAt);
    },
    isConnected,
    getAppsScriptCode: () => APPS_SCRIPT_CODE
  };
})();
