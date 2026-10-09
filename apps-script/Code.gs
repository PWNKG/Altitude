/**
 * Altitude — leaderboard backend (Google Apps Script bound to a Google Sheet).
 * Deploy: Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 * Each player's "id" is a secret device token; it never leaves this sheet.
 */
const SHEET = 'players';
const REQ = [50, 80, 105, 130, 150, 170, 190, 210, 230, 255, 280, 310, 340]; // points to summit each peak
const MAX_SCORE = 500;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['id', 'username', 'total', 'stars', 'best', 'created', 'updated']);
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function rows_() {
  return sheet_().getDataRange().getValues().slice(1).filter(r => r[0]);
}
function parse_(v) { try { return JSON.parse(v || '{}'); } catch (e) { return {}; } }
function summits_(stars) { let n = 0; while (n < REQ.length && (stars[n] || 0) > 0) n++; return n; }
function board_(rows) {
  return rows.map(r => {
    const st = parse_(r[3]);
    return { u: r[1], t: Number(r[2]) || 0, s: Object.values(st).reduce((a, b) => a + Number(b), 0), p: summits_(st) };
  }).sort((a, b) => b.t - a.t);
}
function me_(r) { return { u: r[1], total: Number(r[2]) || 0, stars: parse_(r[3]), best: parse_(r[4]) }; }

function doGet(e) {
  return out_({ ok: true, board: board_(rows_()) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const b = JSON.parse(e.postData.contents);
    const id = String(b.id || '');
    if (id.length < 16 || id.length > 64) throw new Error('Bad device id');
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    let idx = -1;
    for (let i = 1; i < rows.length; i++) if (rows[i][0] === id) { idx = i; break; }

    if (b.action === 'register') {
      if (idx > 0) return out_({ ok: true, me: me_(rows[idx]), board: board_(rows.slice(1)) });
      const u = String(b.username || '').trim();
      if (!/^[A-Za-z0-9_]{3,16}$/.test(u)) throw new Error('Use 3–16 letters, numbers or _');
      if (rows.slice(1).some(r => String(r[1]).toLowerCase() === u.toLowerCase())) throw new Error('That name is taken');
      const now = new Date();
      sh.appendRow([id, u, 0, '{}', '{}', now, now]);
      return out_({ ok: true, me: { u, total: 0, stars: {}, best: {} }, board: board_(rows_()) });
    }

    if (idx < 0) throw new Error('Unknown player');

    if (b.action === 'me') return out_({ ok: true, me: me_(rows[idx]), board: board_(rows.slice(1)) });

    if (b.action === 'climb') {
      const level = Number(b.level), score = Number(b.score);
      if (!Number.isInteger(level) || level < 0 || level >= REQ.length) throw new Error('Bad level');
      if (!Number.isInteger(score) || score < 0 || score > MAX_SCORE) throw new Error('Bad score');
      const r = rows[idx], stars = parse_(r[3]), best = parse_(r[4]);
      if (level > summits_(stars)) throw new Error('Peak is locked');
      const req = REQ[level];
      const s = score >= Math.min(Math.round(req * 1.6), 450) ? 3 : score >= Math.round(req * 1.3) ? 2 : score >= req ? 1 : 0;
      if (score > (best[level] ?? -1)) best[level] = score;
      if (s > (stars[level] || 0)) stars[level] = s;
      const total = Object.values(best).reduce((a, v) => a + Number(v), 0);
      sh.getRange(idx + 1, 3, 1, 3).setValues([[total, JSON.stringify(stars), JSON.stringify(best)]]);
      sh.getRange(idx + 1, 7).setValue(new Date());
      const fresh = sh.getDataRange().getValues();
      return out_({ ok: true, me: me_(fresh[idx]), board: board_(fresh.slice(1)) });
    }
    throw new Error('Unknown action');
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
