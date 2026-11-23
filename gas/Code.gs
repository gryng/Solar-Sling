/**
 * SOLAR SLING — Zero-Cost Online Ranking Backend (Google Apps Script + Google Sheets)
 *
 * Setup (3 minutes):
 * 1. Create a new Google Spreadsheet (https://sheets.new).
 * 2. Open Extensions > Apps Script.
 * 3. Paste this entire file into Code.gs and save.
 * 4. Click Deploy > New deployment > Type: Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL (https://script.google.com/macros/s/.../exec)
 *    and paste it into the game's "GAS設定" or CONFIG.GAS_ENDPOINT in index.html.
 */

const SHEET_NAME = 'Scores';
const MAX_STORED_ROWS = 1000;
const DEFAULT_LIMIT = 50;
const MAX_SCORE = 5000000;
const SALT = 'SOLAR_SLING_V1';

/**
 * GET /exec?action=top&period=all|weekly&limit=50
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const period = params.period === 'weekly' ? 'weekly' : 'all';
    const limit = Math.min(Math.max(parseInt(params.limit, 10) || DEFAULT_LIMIT, 1), 100);

    const sheet = getOrCreateSheet_();
    const rows = readScores_(sheet);
    const filtered = filterByPeriod_(rows, period);

    filtered.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.ts - b.ts;
    });

    const top = filtered.slice(0, limit).map(function (r, idx) {
      return {
        rank: idx + 1,
        id: r.id,
        name: r.name,
        score: r.score,
        wave: r.wave,
        maxCombo: r.maxCombo,
        ts: r.ts
      };
    });

    return jsonResponse_({
      ok: true,
      period: period,
      total: filtered.length,
      entries: top,
      serverTime: Date.now()
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

/**
 * POST /exec  (Send with Content-Type: text/plain;charset=utf-8 to avoid CORS preflight)
 * Body JSON: { name, score, wave, maxCombo, bankShots, device, sig }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: 'Empty payload' });
    }

    const payload = JSON.parse(e.postData.contents);
    const name = sanitizeName_(payload.name);
    const score = Math.floor(Number(payload.score) || 0);
    const wave = Math.max(1, Math.min(999, Math.floor(Number(payload.wave) || 1)));
    const maxCombo = Math.max(0, Math.min(9999, Math.floor(Number(payload.maxCombo) || 0)));
    const bankShots = Math.max(0, Math.min(9999, Math.floor(Number(payload.bankShots) || 0)));
    const device = String(payload.device || 'web').slice(0, 16);
    const sig = String(payload.sig || '');

    if (!name) {
      return jsonResponse_({ ok: false, error: 'Invalid player name' });
    }
    if (score <= 0 || score > MAX_SCORE) {
      return jsonResponse_({ ok: false, error: 'Score out of bounds' });
    }

    // Lightweight integrity check matching client hash
    const expectedSig = computeSig_(name, score, wave, maxCombo);
    if (sig !== expectedSig) {
      return jsonResponse_({ ok: false, error: 'Signature mismatch' });
    }

    const sheet = getOrCreateSheet_();
    const now = Date.now();
    const id = Utilities.getUuid().slice(0, 12);

    sheet.appendRow([id, now, name, score, wave, maxCombo, bankShots, device]);

    // Trim oldest low-ranking rows if sheet grows too large
    const lastRow = sheet.getLastRow();
    if (lastRow > MAX_STORED_ROWS + 1) {
      pruneSheet_(sheet);
    }

    // Compute rank
    const allRows = readScores_(sheet);
    allRows.sort(function (a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.ts - b.ts;
    });

    const rank = allRows.findIndex(function (r) { return r.id === id; }) + 1;
    const top = allRows.slice(0, DEFAULT_LIMIT).map(function (r, idx) {
      return {
        rank: idx + 1,
        id: r.id,
        name: r.name,
        score: r.score,
        wave: r.wave,
        maxCombo: r.maxCombo,
        ts: r.ts
      };
    });

    return jsonResponse_({
      ok: true,
      id: id,
      rank: rank,
      total: allRows.length,
      entries: top
    });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'timestamp', 'name', 'score', 'wave', 'maxCombo', 'bankShots', 'device']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readScores_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const list = [];
  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const score = Number(r[3]);
    if (!score || score <= 0) continue;
    list.push({
      id: String(r[0] || ''),
      ts: Number(r[1]) || 0,
      name: String(r[2] || 'PLAYER'),
      score: score,
      wave: Number(r[4]) || 1,
      maxCombo: Number(r[5]) || 0,
      bankShots: Number(r[6]) || 0,
      device: String(r[7] || 'web')
    });
  }
  return list;
}

function filterByPeriod_(rows, period) {
  if (period !== 'weekly') return rows.slice();
  const cutoff = Date.now() - 7 * 24 * 60 * 60 * 1000;
  return rows.filter(function (r) { return r.ts >= cutoff; });
}

function pruneSheet_(sheet) {
  const rows = readScores_(sheet);
  rows.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return b.ts - a.ts;
  });
  const keep = rows.slice(0, MAX_STORED_ROWS);
  sheet.clearContents();
  sheet.appendRow(['id', 'timestamp', 'name', 'score', 'wave', 'maxCombo', 'bankShots', 'device']);
  if (keep.length > 0) {
    const data = keep.map(function (r) {
      return [r.id, r.ts, r.name, r.score, r.wave, r.maxCombo, r.bankShots, r.device];
    });
    sheet.getRange(2, 1, data.length, 8).setValues(data);
  }
}

function sanitizeName_(raw) {
  const cleaned = String(raw || '')
    .replace(/[\r\n\t<>]/g, '')
    .trim()
    .slice(0, 16);
  return cleaned || 'GUEST';
}

function computeSig_(name, score, wave, maxCombo) {
  const input = [SALT, name, score, wave, maxCombo].join('|');
  let h = 2166136261;
  for (let i = 0; i < input.length; i++) {
    h ^= input.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return (h >>> 0).toString(16);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
