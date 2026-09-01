/**
 * FPS RICE REGISTER — STEP 2 BACKEND
 * CMAY + NFSA | September 2026 → March 2027
 *
 * Source-of-truth sheets only:
 *   CMAY AAY / CMAY PHH / CMAY SFSS
 *   NFSA AAY / NFSA PHH / NFSA SFSS
 *
 * Output sheets are never written by this API.
 * They remain formula/report driven from the six source sheets.
 *
 * Deploy as a Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 */

const CONFIG = Object.freeze({
  version: '2.0.0',
  periodStart: '2026-09',
  periodEnd: '2027-03',
  registers: ['CMAY', 'NFSA'],
  schemes: ['AAY', 'PHH', 'SFSS'],
  sourceSheets: {
    CMAY: { AAY: 'CMAY AAY', PHH: 'CMAY PHH', SFSS: 'CMAY SFSS' },
    NFSA: { AAY: 'NFSA AAY', PHH: 'NFSA PHH', SFSS: 'NFSA SFSS' }
  },
  // Each monthly block is 9 columns:
  // DAY | Date | Manual Opening | Opening | Received | Total | Selling | Closing | Remarks
  blocks: {
    '2026-09': { startCol: 1,  dayCol: 1, dateCol: 2, manualOpenCol: 3, openCol: 4, receiveCol: 5, totalCol: 6, soldCol: 7, closeCol: 8, remarksCol: 9, days: 30 },
    '2026-10': { startCol: 10, dayCol: 10, dateCol: 11, manualOpenCol: 12, openCol: 13, receiveCol: 14, totalCol: 15, soldCol: 16, closeCol: 17, remarksCol: 18, days: 31 },
    '2026-11': { startCol: 19, dayCol: 19, dateCol: 20, manualOpenCol: 21, openCol: 22, receiveCol: 23, totalCol: 24, soldCol: 25, closeCol: 26, remarksCol: 27, days: 30 },
    '2026-12': { startCol: 28, dayCol: 28, dateCol: 29, manualOpenCol: 30, openCol: 31, receiveCol: 32, totalCol: 33, soldCol: 34, closeCol: 35, remarksCol: 36, days: 31 },
    '2027-01': { startCol: 37, dayCol: 37, dateCol: 38, manualOpenCol: 39, openCol: 40, receiveCol: 41, totalCol: 42, soldCol: 43, closeCol: 44, remarksCol: 45, days: 31 },
    '2027-02': { startCol: 46, dayCol: 46, dateCol: 47, manualOpenCol: 48, openCol: 49, receiveCol: 50, totalCol: 51, soldCol: 52, closeCol: 53, remarksCol: 54, days: 28 },
    '2027-03': { startCol: 55, dayCol: 55, dateCol: 56, manualOpenCol: 57, openCol: 58, receiveCol: 59, totalCol: 60, soldCol: 61, closeCol: 62, remarksCol: 63, days: 31 }
  },
  firstDataRow: 4
});

function doGet(e) {
  const action = String(e && e.parameter && e.parameter.action || 'ping').toLowerCase();
  try {
    switch (action) {
      case 'ping':
        return json_({ ok: true, version: CONFIG.version, period: [CONFIG.periodStart, CONFIG.periodEnd], registers: CONFIG.registers, schemes: CONFIG.schemes });
      case 'profile':
        return json_({ ok: true, profile: readProfile_() });
      case 'readentries':
      case 'entries':
        return json_({ ok: true, version: CONFIG.version, entries: readEntries_(e.parameter || {}) });
      default:
        return json_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err), stack: String(err && err.stack || '') });
  }
}

function doPost(e) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const body = parseBody_(e);
    const action = String(body.action || '').toLowerCase();
    if (action === 'profile') return json_({ ok: true, profile: saveProfile_(body.profile || {}) });
    if (action === 'upsert') return json_(upsertBatch_([body.row || body.entry || body]));
    if (action === 'upsertbatch' || action === 'batch') return json_(upsertBatch_(Array.isArray(body.rows) ? body.rows : Array.isArray(body.entries) ? body.entries : []));
    if (action === 'delete') return json_(deleteEntry_(body));
    return json_({ ok: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err), stack: String(err && err.stack || '') });
  } finally {
    lock.releaseLock();
  }
}

function parseBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  try { return JSON.parse(raw); } catch (_) {
    // Also support application/x-www-form-urlencoded fallback.
    const p = e && e.parameter ? e.parameter : {};
    if (p.payload) return JSON.parse(p.payload);
    return p;
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const missing = [];
  CONFIG.registers.forEach(r => CONFIG.schemes.forEach(s => {
    const name = CONFIG.sourceSheets[r][s];
    if (!ss.getSheetByName(name)) missing.push(name);
  }));
  if (missing.length) throw new Error('Missing source sheets: ' + missing.join(', '));

  const props = PropertiesService.getDocumentProperties();
  if (!props.getProperty('FPS_API_VERSION')) props.setProperty('FPS_API_VERSION', CONFIG.version);
  if (!props.getProperty('FPS_PROFILE')) props.setProperty('FPS_PROFILE', JSON.stringify({ fpsName:'', fpsCode:'', place:'', dealerName:'' }));

  return 'OK — six source sheets verified. Backend will write ONLY to source sheets.';
}

function saveProfile_(p) {
  const profile = {
    fpsName: cleanText_(p.fpsName),
    fpsCode: cleanText_(p.fpsCode),
    place: cleanText_(p.place),
    dealerName: cleanText_(p.dealerName),
    updatedAt: new Date().toISOString()
  };
  PropertiesService.getDocumentProperties().setProperty('FPS_PROFILE', JSON.stringify(profile));
  return profile;
}

function readProfile_() {
  const raw = PropertiesService.getDocumentProperties().getProperty('FPS_PROFILE');
  if (!raw) return { fpsName:'', fpsCode:'', place:'', dealerName:'' };
  try { return JSON.parse(raw); } catch (_) { return { fpsName:'', fpsCode:'', place:'', dealerName:'' }; }
}

function upsertBatch_(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return { ok: false, error: 'No rows supplied.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  const metadata = PropertiesService.getDocumentProperties();

  rows.forEach(input => {
    try {
      const row = normalizeEntry_(input);
      validateEntry_(row);

      const key = entryKey_(row);
      const versionKey = 'VER_' + safePropertyKey_(key);
      const previous = metadata.getProperty(versionKey);
      const oldTs = previous ? Number(previous) : 0;
      const incomingTs = Date.parse(row.updatedAt) || Date.now();

      // Prevent an older delayed offline retry from overwriting newer data.
      if (oldTs && incomingTs < oldTs) {
        results.push({ ok: false, key, entryId: row.entryId, conflict: true, error: 'Older update rejected; newer version already saved.', serverUpdatedAt: new Date(oldTs).toISOString() });
        return;
      }

      const block = CONFIG.blocks[row.riceMonth];
      const sheetName = CONFIG.sourceSheets[row.register][row.scheme];
      const sh = ss.getSheetByName(sheetName);
      if (!sh) throw new Error('Source sheet not found: ' + sheetName);

      const targetRow = CONFIG.firstDataRow + row.day - 1;
      const dateCell = sh.getRange(targetRow, block.dateCol).getValue();
      if (!sameDayOfMonth_(dateCell, row.day)) {
        throw new Error('Template day mismatch at ' + sheetName + ' row ' + targetRow + '. Expected day ' + row.day + '.');
      }

      // IMPORTANT: only write source/input columns. Formula columns remain untouched.
      // C = Manual Opening, E = Received, G = Selling, I = Remarks.
      sh.getRange(targetRow, block.manualOpenCol).setValue(row.openingMode === 'manual' ? row.opening : '');
      sh.getRange(targetRow, block.receiveCol).setValue(row.received === null ? '' : row.received);
      sh.getRange(targetRow, block.soldCol).setValue(row.sold === null ? '' : row.sold);
      sh.getRange(targetRow, block.remarksCol).setValue(row.remarks || '');

      SpreadsheetApp.flush();
      metadata.setProperty(versionKey, String(incomingTs));

      const computed = readSourceRow_(sh, block, targetRow, row);
      results.push({ ok: true, key, entryId: row.entryId, register: row.register, scheme: row.scheme, riceMonth: row.riceMonth, physicalDate: row.physicalDate, data: computed, serverUpdatedAt: new Date(incomingTs).toISOString() });
    } catch (err) {
      results.push({ ok: false, entryId: input && (input.entryId || input.id || ''), error: String(err && err.message || err) });
    }
  });

  return { ok: results.every(r => r.ok || r.conflict), results };
}

function deleteEntry_(input) {
  const row = normalizeEntry_(input);
  validateEntry_(row);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const block = CONFIG.blocks[row.riceMonth];
  const sheetName = CONFIG.sourceSheets[row.register][row.scheme];
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Source sheet not found: ' + sheetName);

  const targetRow = CONFIG.firstDataRow + row.day - 1;
  const dateCell = sh.getRange(targetRow, block.dateCol).getValue();
  if (!sameDayOfMonth_(dateCell, row.day)) throw new Error('Template day mismatch.');

  sh.getRange(targetRow, block.manualOpenCol).clearContent();
  sh.getRange(targetRow, block.receiveCol).clearContent();
  sh.getRange(targetRow, block.soldCol).clearContent();
  sh.getRange(targetRow, block.remarksCol).clearContent();
  SpreadsheetApp.flush();

  PropertiesService.getDocumentProperties().deleteProperty('VER_' + safePropertyKey_(entryKey_(row)));
  return { ok: true, key: entryKey_(row), deleted: true };
}

function readEntries_(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wantRegister = params.register ? String(params.register).toUpperCase() : null;
  const wantScheme = params.scheme ? String(params.scheme).toUpperCase() : null;
  const out = [];

  CONFIG.registers.filter(r => !wantRegister || r === wantRegister).forEach(register => {
    CONFIG.schemes.filter(s => !wantScheme || s === wantScheme).forEach(scheme => {
      const sheetName = CONFIG.sourceSheets[register][scheme];
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return;

      Object.keys(CONFIG.blocks).forEach(month => {
        const b = CONFIG.blocks[month];
        const values = sh.getRange(CONFIG.firstDataRow, b.dateCol, b.days, 8).getValues();
        // Returned columns start at Date; we need Date, Manual Opening, Opening, Received, Total, Sold, Closing, Remarks.
        values.forEach((r, i) => {
          const d = r[0];
          if (!(d instanceof Date) && d === '') return;
          const physicalDate = formatDate_(d);
          const manual = numOrNull_(r[1]);
          const opening = numOrZero_(r[2]);
          const received = numOrNull_(r[3]);
          const total = numOrZero_(r[4]);
          const sold = numOrNull_(r[5]);
          const closing = numOrZero_(r[6]);
          const remarks = cleanText_(r[7]);

          // Only return meaningful records; empty formula rows are omitted.
          if (manual === null && received === null && sold === null && !remarks && Number(opening) === 0 && Number(total) === 0 && Number(closing) === 0) return;
          out.push({
            entryId: buildEntryId_(register, scheme, month, physicalDate),
            key: register + '|' + scheme + '|' + month + '|' + physicalDate,
            register, scheme, riceMonth: month, physicalDate,
            day: i + 1,
            openingMode: manual === null ? 'auto' : 'manual',
            opening: manual,
            received,
            sold,
            remarks,
            computedOpening: opening,
            total,
            closing
          });
        });
      });
    });
  });

  return out;
}

function readSourceRow_(sh, b, rowNumber, row) {
  const vals = sh.getRange(rowNumber, b.manualOpenCol, 1, 7).getValues()[0];
  // C..I = Manual Opening, Opening, Received, Total, Selling, Closing, Remarks
  return {
    manualOpening: numOrNull_(vals[0]),
    opening: numOrZero_(vals[1]),
    received: numOrNull_(vals[2]),
    total: numOrZero_(vals[3]),
    sold: numOrNull_(vals[4]),
    closing: numOrZero_(vals[5]),
    remarks: cleanText_(vals[6])
  };
}

function normalizeEntry_(x) {
  const physicalDate = normalizeDate_(x.physicalDate || x.date);
  const riceMonth = normalizeMonth_(x.riceMonth || x.month);
  const register = String(x.register || '').toUpperCase();
  const scheme = String(x.scheme || x.commodity || '').toUpperCase();
  const openingMode = String(x.openingMode || (x.manualOpening !== undefined && x.manualOpening !== null && x.manualOpening !== '' ? 'manual' : 'auto')).toLowerCase();
  const openingRaw = x.opening !== undefined ? x.opening : x.manualOpening;

  return {
    entryId: cleanText_(x.entryId || x.id || buildEntryId_(register, scheme, riceMonth, physicalDate)),
    register,
    scheme,
    riceMonth,
    physicalDate,
    day: getDayFromIso_(physicalDate),
    openingMode: openingMode === 'manual' ? 'manual' : 'auto',
    opening: openingMode === 'manual' ? numOrZero_(openingRaw) : null,
    received: blankNum_(x.received),
    sold: blankNum_(x.sold !== undefined ? x.sold : x.selling),
    remarks: cleanText_(x.remarks),
    updatedAt: normalizeTimestamp_(x.updatedAt || x.clientUpdatedAt || x.timestamp)
  };
}

function validateEntry_(r) {
  if (!CONFIG.registers.includes(r.register)) throw new Error('Invalid register: ' + r.register);
  if (!CONFIG.schemes.includes(r.scheme)) throw new Error('Invalid scheme: ' + r.scheme);
  if (!CONFIG.blocks[r.riceMonth]) throw new Error('Rice month must be September 2026 to March 2027.');
  if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(r.physicalDate)) throw new Error('Invalid physical date.');

  const expectedMonth = r.riceMonth;
  const actualPhysicalMonth = r.physicalDate.slice(0, 7);
  // Physical date may be from a different calendar month than the rice month.
  // But it must still fall inside the supported period.
  if (!CONFIG.blocks[actualPhysicalMonth]) throw new Error('Physical date is outside September 2026 to March 2027.');
  if (r.day < 1 || r.day > CONFIG.blocks[expectedMonth].days) {
    throw new Error('Date day is not valid for rice month ' + expectedMonth + '.');
  }
  if (r.received !== null && r.received < 0) throw new Error('Received cannot be negative.');
  if (r.sold !== null && r.sold < 0) throw new Error('Selling cannot be negative.');
}

function entryKey_(r) {
  return [r.register, r.scheme, r.riceMonth, r.physicalDate].join('|');
}

function buildEntryId_(register, scheme, riceMonth, physicalDate) {
  return [register, scheme, riceMonth, physicalDate].join('_');
}

function normalizeMonth_(value) {
  const s = String(value || '').trim();
  if (/^\\d{4}-\\d{2}$/.test(s)) return s;
  const map = {
    'September 2026':'2026-09','October 2026':'2026-10','November 2026':'2026-11',
    'December 2026':'2026-12','January 2027':'2027-01','February 2027':'2027-02','March 2027':'2027-03'
  };
  return map[s] || s;
}

function normalizeDate_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\\d{4}-\\d{2}-\\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (isNaN(d)) throw new Error('Invalid date: ' + s);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normalizeTimestamp_(v) {
  if (v === undefined || v === null || v === '') return new Date().toISOString();
  const n = Date.parse(String(v));
  return isNaN(n) ? new Date().toISOString() : new Date(n).toISOString();
}

function formatDate_(d) {
  if (d instanceof Date && !isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(d || '');
}

function sameDate_(cellDate, isoDate) { return formatDate_(cellDate) === isoDate; }

function sameDayOfMonth_(cellDate, expectedDay) {
  if (cellDate instanceof Date && !isNaN(cellDate)) return cellDate.getDate() === Number(expectedDay);
  const s = String(cellDate || '').trim();
  const m = s.match(/(?:^|[-\/\s])([0-9]{1,2})(?:$|[-\/\s])/);
  return !!m && Number(m[1]) === Number(expectedDay);
}

function getDayFromIso_(iso) { return Number(iso.slice(8, 10)); }

function blankNum_(v) {
  if (v === undefined || v === null || v === '') return null;
  const n = Number(v);
  return isFinite(n) ? n : null;
}

function numOrNull_(v) {
  if (v === undefined || v === null || v === '') return null;
  const n = Number(v);
  return isFinite(n) ? n : null;
}

function numOrZero_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function cleanText_(v) { return String(v === undefined || v === null ? '' : v).trim(); }

function safePropertyKey_(s) { return Utilities.base64EncodeWebSafe(String(s)).replace(/=+$/,''); }
