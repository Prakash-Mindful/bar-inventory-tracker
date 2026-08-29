/**
 * Bar count -> Google Sheet
 *
 * Writes each counted item into the Counts tab, works out the change against that
 * item's previous count, stores it, and hands the change back to the phone.
 *
 * Setup, once:
 *  1. Open your tracker in Google Sheets (File > Save as Google Sheets if it is still .xlsx).
 *  2. Extensions > Apps Script. Delete whatever is there, paste this file in, Save.
 *  3. Deploy > New deployment > Web app.
 *       Execute as:      Me
 *       Who has access:  Anyone
 *  4. Authorise when Google asks. Copy the URL ending in /exec.
 *  5. In the count page, tap the sync line at the top, paste the URL, Save and test.
 *
 * After any edit here: Deploy > Manage deployments > pencil icon > Version: New version.
 * Without that step the old code keeps running.
 */

var SHEET_NAME = 'Counts';
var HEAD = ['Date', 'Shift', 'No', 'Item', 'Category', 'Qty', 'Unit', 'Counted by',
            'Key', 'Prev qty', 'Change'];
var COLS = HEAD.length;

/* ---------------- write ---------------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.ping) return json({ ok: true, ping: true });

    var rows = body.rows || [];
    if (!rows.length) return json({ ok: true, saved: 0, deltas: [] });

    var sheet = getSheet();
    var index = keyIndex(sheet);
    var append = [];
    var keys = [];

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var key = String(row[8] || (row[0] + '|' + row[1] + '|' + row[2]));
      keys.push(key);
      var out = [String(row[0]), String(row[1]), Number(row[2]), String(row[3]),
                 String(row[4]), Number(row[5]), String(row[6]), String(row[7] || ''), key];
      if (index[key]) sheet.getRange(index[key], 1, 1, 9).setValues([out]);
      else append.push(out);
    }
    if (append.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, append.length, 9).setValues(append);
    }

    // Prev qty and Change are recomputed from the whole sheet, in row order, so a
    // correction to an older count fixes its own delta and the one after it.
    var map = recomputeDeltas(sheet);

    var deltas = [];
    for (var k = 0; k < keys.length; k++) {
      var m = map[keys[k]] || {};
      deltas.push({
        no: Number(rows[k][2]), item: String(rows[k][3]), unit: String(rows[k][6]),
        qty: Number(rows[k][5]),
        prev: m.prev === '' || m.prev === undefined ? null : m.prev,
        prevDate: m.prevDate || null, prevShift: m.prevShift || null,
        delta: m.change === '' || m.change === undefined ? null : m.change
      });
    }
    return json({ ok: true, saved: rows.length, added: append.length, deltas: deltas });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Walks Counts top to bottom. For every row, the previous count is the closest
 * earlier row for the same item number. Writes columns J and K in one pass and
 * returns key -> {prev, change, prevDate, prevShift}.
 */
function recomputeDeltas(sheet) {
  var last = sheet.getLastRow();
  var map = {};
  if (last < 2) return map;

  var data = sheet.getRange(2, 1, last - 1, 9).getValues();
  var block = [];
  var seen = {};

  for (var i = 0; i < data.length; i++) {
    var no = Number(data[i][2]);
    var qty = Number(data[i][5]);
    var key = String(data[i][8] || '');
    if (!no) { block.push(['', '']); continue; }

    var prev = seen[no];
    var pQty = prev ? prev.qty : '';
    var chg = prev ? round2(qty - prev.qty) : '';
    block.push([pQty, chg]);

    if (key) map[key] = { prev: pQty, change: chg,
                          prevDate: prev ? prev.date : null,
                          prevShift: prev ? prev.shift : null };

    seen[no] = { qty: qty, date: dateStr(data[i][0]), shift: String(data[i][1]) };
  }
  sheet.getRange(2, 10, block.length, 2).setValues(block);
  return map;
}

function keyIndex(sheet) {
  var index = {};
  var last = sheet.getLastRow();
  if (last < 2) return index;
  var keys = sheet.getRange(2, 9, last - 1, 1).getValues();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i][0]) index[String(keys[i][0])] = i + 2;
  }
  return index;
}

/* ---------------- read ---------------- */

function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : '';
    if (action !== 'latest') return json({ ok: true, ping: true });

    var sheet = getSheet();
    var last = sheet.getLastRow();
    var latest = {};
    if (last > 1) {
      var data = sheet.getRange(2, 1, last - 1, COLS).getValues();
      for (var i = 0; i < data.length; i++) {
        var no = Number(data[i][2]);
        if (!no) continue;
        var here = { qty: Number(data[i][5]), date: dateStr(data[i][0]), shift: String(data[i][1]) };
        var before = latest[no];
        latest[no] = {
          qty: here.qty, date: here.date, shift: here.shift,
          prev: before ? before.qty : null,
          prevDate: before ? before.date : null,
          prevShift: before ? before.shift : null
        };
      }
    }
    return json({ ok: true, latest: latest });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ---------------- helpers ---------------- */

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLS).setValues([HEAD]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function dateStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function round2(n) { return Math.round(n * 100) / 100; }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
