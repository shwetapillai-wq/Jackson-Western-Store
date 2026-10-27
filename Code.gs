/**
 * Weekly client email for JW Store.
 * Reads the "Daily Data" tab and the "Client Config" tab of this Google Sheet.
 * Setup steps are in README.md.
 */

// ---------- Settings ----------
var TEST_MODE = true;               // true = emails go only to you. Set to false to go live.
var SENDER_NAME = 'Vserve Weekly Update';
var CLIENT_LABEL = 'JW Store';
var DATA_SHEET = 'Daily Data';
var CONFIG_SHEET = 'Client Config';
var LOG_SHEET = 'Email Log';
var UPDATES_SHEET = 'Team Updates';   // Dev and SEO rows (optional)
var KEYWORDS_SHEET = 'SEO Keywords';  // keyword rankings (optional)

// ---------- One-time setup ----------
function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  var cfg = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  if (cfg.getLastRow() === 0) {
    cfg.appendRow(['Client Name', 'To Emails (comma separated)', 'CC Emails', 'Active (Yes/No)', 'Show Daily Notes (Yes/No)']);
    cfg.appendRow(['Client name here', 'client@example.com', '', 'No', 'No']);
    cfg.getRange(1, 1, 1, 5).setFontWeight('bold');
    cfg.setFrozenRows(1);
  }
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  if (log.getLastRow() === 0) {
    log.appendRow(['Sent at', 'Sent to', 'Status', 'SKUs completed in report']);
    log.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
}

// Sends every Monday around 9 AM. Run once.
function createWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importNewFiles') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importNewFiles').timeBased().everyDays(1).atHour(19).create();   // picks up new Dev/DM files daily
}

// ---------- Entry points ----------
function sendWeeklyReport() { run_(TEST_MODE); }
function sendTestToMe() { run_(true); }   // always sends to you only

function run_(testOnly) {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
  var cfg = ss.getSheetByName(CONFIG_SHEET).getDataRange().getValues();
  if (cfg.length < 2) { log.appendRow([new Date(), '', 'Error: Client Config is empty', 0]); return; }
  var c = cfg[1];
  if (String(c[3]).trim().toLowerCase() !== 'yes') { log.appendRow([new Date(), '', 'Skipped: Active is not Yes', 0]); return; }
  try {
    try { importNewFiles(); } catch (ie) { log.appendRow([new Date(), '', 'Import error: ' + ie.message, 0]); }
    var att = attachmentsForWeek_(iso_(weekRange_(new Date()).start));
    var rep = buildReport_(loadRows_(), weekRange_(new Date()), String(c[4]).trim().toLowerCase() === 'yes', String(c[0]).trim(), loadUpdates_(), loadKeywords_());
    if (!rep) { log.appendRow([new Date(), '', 'Skipped: no activity last week', 0]); return; }
    var me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    var to = testOnly ? me : String(c[1]).trim();
    var cc = testOnly ? '' : String(c[2]).trim();
    var html = att.names.length ? rep.html.replace('Please reply to this email', 'Attached: ' + esc_(att.names.join(', ')) + '.<br>Please reply to this email') : rep.html;
    MailApp.sendEmail({ to: to, cc: cc, subject: (testOnly ? '[TEST] ' : '') + rep.subject, htmlBody: html, name: SENDER_NAME, attachments: att.blobs });
    if (!testOnly) markEmailed_(att.ids);
    log.appendRow([new Date(), to, testOnly ? 'Sent (test)' : 'Sent', rep.completed]);
  } catch (e) {
    log.appendRow([new Date(), '', 'Error: ' + e.message, 0]);
  }
}

// ---------- Data ----------
// Finds the header row by name, so blank rows or moved columns do not matter.
function loadRows_() {
  var vals = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET).getDataRange().getValues();
  var h = -1;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].map(function (x) { return String(x).trim(); }).indexOf('SKUs Worked') >= 0) { h = i; break; }
  }
  if (h < 0) throw new Error("Header 'SKUs Worked' not found in " + DATA_SHEET);
  var head = vals[h].map(function (x) { return String(x).trim(); });
  var ix = function (n) { return head.indexOf(n); };
  var iD = ix('Date'), iW = ix('SKUs Worked'), iC = ix('SKUs Completed'), iH = ix('Hours'), iN = ix('Comments');
  var seen = {}, rows = [];
  for (var r = h + 1; r < vals.length; r++) {
    var v = vals[r];
    var d = v[iD] instanceof Date ? v[iD] : new Date(v[iD]);
    if (isNaN(d.getTime())) continue;
    var key = [iso_(d), v[iW], v[iC], v[iH], iN >= 0 ? v[iN] : ''].join('|');
    if (seen[key]) continue;       // skip exact duplicate rows
    seen[key] = 1;
    rows.push({ date: d, worked: Number(v[iW]) || 0, completed: Number(v[iC]) || 0, hours: Number(v[iH]) || 0,
                note: iN >= 0 ? String(v[iN]).trim() : '' });
  }
  return rows;
}

// Reads an optional tab by header name. Returns [] if the tab does not exist.
function readTable_(sheetName, mustHave) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues(), h = -1;
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].map(function (x) { return String(x).trim(); }).indexOf(mustHave) >= 0) { h = i; break; }
  }
  if (h < 0) return [];
  var head = vals[h].map(function (x) { return String(x).trim(); });
  return vals.slice(h + 1).map(function (v) {
    var o = {}; head.forEach(function (n, j) { o[n] = v[j]; }); return o;
  });
}
function mondayKey_(v) {
  var d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return null;
  return iso_(addDays_(d, -((d.getDay() + 6) % 7)));
}
function loadUpdates_() {
  return readTable_(UPDATES_SHEET, 'Activity').map(function (o) {
    return { week: mondayKey_(o['Week Start']), team: String(o['Team'] || '').trim(), activity: String(o['Activity'] || '').trim(),
             count: o['Count'] === '' || o['Count'] == null ? '' : Number(o['Count']), hours: o['Hours'] === '' || o['Hours'] == null ? '' : Number(o['Hours']),
             details: String(o['Details'] || '').trim() };
  }).filter(function (u) { return u.week && u.activity; });
}
function loadKeywords_() {
  return readTable_(KEYWORDS_SHEET, 'Keyword').map(function (o) {
    return { week: mondayKey_(o['Week Start']), keyword: String(o['Keyword'] || '').trim(), rank: Number(o['Rank']) };
  }).filter(function (k) { return k.week && k.keyword; });
}

// Last full Monday to Sunday week before "today"
function weekRange_(today) {
  var t = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var start = addDays_(t, -(((t.getDay() + 6) % 7) + 7));
  return { start: start, end: addDays_(start, 6) };
}

// ---------- Report ----------
function buildReport_(rows, range, showNotes, clientName, updates, keywords) {
  function inWeek(start) {
    var s = iso_(start), e = iso_(addDays_(start, 6));
    return rows.filter(function (r) { var k = iso_(r.date); return k >= s && k <= e; });
  }
  var cur = inWeek(range.start);
  var worked = sum_(cur, 'worked'), completed = sum_(cur, 'completed'), hours = sum_(cur, 'hours');
  var days = cur.filter(function (r) { return r.worked > 0 || r.hours > 0; }).length;
  var weekKey = iso_(range.start);
  var upd = (updates || []).filter(function (u) { return u.week === weekKey; });
  if (!days && !upd.length) return null;
  var prev = inWeek(addDays_(range.start, -7));
  var pc = sum_(prev, 'completed');
  var change = pc > 0 ? Math.round((completed - pc) / pc * 100) + '%' : 'n/a';
  if (pc > 0 && completed >= pc) change = '+' + change;
  var rate = worked > 0 ? Math.round(completed / worked * 100) + '%' : 'n/a';

  var dayRows = cur.sort(function (a, b) { return a.date - b.date; })
    .filter(function (r) { return r.worked > 0 || r.hours > 0; })
    .map(function (r) { return [fmt_(r.date, 'EEE d MMM'), r.worked, r.completed, r1_(r.hours)]; });
  var trend = [];
  for (var w = 3; w >= 0; w--) {
    var ws = addDays_(range.start, -7 * w), wk = inWeek(ws), wW = sum_(wk, 'worked'), wC = sum_(wk, 'completed');
    trend.push(['Week of ' + fmt_(ws, 'd MMM'), wW, wC, wW > 0 ? Math.round(wC / wW * 100) + '%' : '-']);
  }
  var notes = showNotes ? cur.filter(function (r) { return r.note; }).map(function (r) { return [fmt_(r.date, 'EEE d MMM'), r.note]; }) : [];

  var period = fmt_(range.start, 'd MMM yyyy') + ' to ' + fmt_(range.end, 'd MMM yyyy');
  var html = '<div style="font-family:Arial,sans-serif;color:#1c2b2a;max-width:640px">' +
    '<h2 style="margin:0 0 4px">Weekly update: ' + esc_(CLIENT_LABEL) + '</h2>' +
    '<p style="margin:0 0 16px;color:#667573">' + period + '</p>' +
    (clientName ? '<p>Hi ' + esc_(clientName) + ',</p>' : '') +
    '<p>Here is a summary of the work completed on your account last week.</p>' +
    (days ? '<h3 style="margin:18px 0 6px;font-size:15px">Catalogue work</h3>' +
      kpis_([['SKUs worked', worked], ['SKUs completed', completed], ['Completion rate', rate], ['Completed vs last week', change]]) +
      kpis_([['Hours', r1_(hours)], ['Days worked', days]]) +
      table_('Day by day', ['Day', 'Worked', 'Completed', 'Hours'], dayRows) +
      table_('Four week trend', ['Week', 'Worked', 'Completed', 'Rate'], trend) +
      table_('Notes', ['Day', 'Note'], notes) : '') +
    teamsHtml_(upd) + keywordsHtml_(keywords, weekKey, iso_(addDays_(range.start, -7))) +
    '<p style="color:#667573;font-size:12px;margin-top:20px">Please reply to this email with any questions.</p></div>';
  return { subject: 'Weekly update: ' + CLIENT_LABEL + ' (' + period + ')', html: html, completed: completed };
}

function teamsHtml_(upd) {
  var order = [], by = {};
  upd.forEach(function (u) { if (!by[u.team]) { by[u.team] = []; order.push(u.team); } by[u.team].push(u); });
  return order.map(function (t) {
    var L = by[t], hasC = L.some(function (u) { return u.count !== ''; }), hasH = L.some(function (u) { return u.hours !== ''; }), hasD = L.some(function (u) { return u.details; });
    var head = ['Activity'].concat(hasC ? ['Count'] : [], hasH ? ['Hours'] : [], hasD ? ['Details'] : []);
    var rows = L.map(function (u) { return [u.activity].concat(hasC ? [u.count] : [], hasH ? [u.hours] : [], hasD ? [u.details] : []); });
    return table_((t || 'Team') + ' team', head, rows);
  }).join('');
}

function keywordsHtml_(kw, wk, prevWk) {
  var cur = (kw || []).filter(function (k) { return k.week === wk; });
  if (!cur.length) return '';
  var prev = {}; (kw || []).forEach(function (k) { if (k.week === prevWk) prev[k.keyword] = k.rank; });
  var rows = cur.map(function (k) {
    var p = prev[k.keyword], ch = p == null ? 'new' : (k.rank < p ? 'up ' + (p - k.rank) : (k.rank > p ? 'down ' + (k.rank - p) : 'no change'));
    return [k.keyword, k.rank, ch];
  });
  return table_('SEO keyword rankings', ['Keyword', 'Rank', 'vs last week'], rows);

}

// ---------- Helpers ----------
function sum_(list, f) { return list.reduce(function (s, r) { return s + r[f]; }, 0); }
function kpis_(items) {
  return '<table style="width:100%;border-collapse:separate;border-spacing:6px;margin-bottom:4px"><tr>' + items.map(function (k) {
    return '<td style="background:#f2f6f5;border-radius:6px;padding:10px;text-align:center"><div style="font-size:22px;font-weight:bold">' +
      esc_(k[1]) + '</div><div style="font-size:12px;color:#667573">' + esc_(k[0]) + '</div></td>';
  }).join('') + '</tr></table>';
}
function table_(title, head, rows) {
  if (!rows.length) return '';
  var right = head.map(function (h, i) {
    return i > 0 && rows.every(function (r) { return r[i] === '' || r[i] == null || /^[-+]?[\d.,]+%?$/.test(String(r[i])) || String(r[i]) === 'n/a'; });
  });
  var th = head.map(function (h, i) { return '<th style="text-align:' + (right[i] ? 'right' : 'left') + ';padding:6px 8px;border-bottom:2px solid #0f6b63">' + esc_(h) + '</th>'; }).join('');
  var tr = rows.map(function (r) {
    return '<tr>' + r.map(function (c, i) { return '<td style="text-align:' + (right[i] ? 'right' : 'left') + ';padding:6px 8px;border-bottom:1px solid #dde3e1;vertical-align:top">' + cell_(c) + '</td>'; }).join('') + '</tr>';
  }).join('');
  return '<h3 style="margin:18px 0 6px;font-size:15px">' + esc_(title) + '</h3><table style="width:100%;border-collapse:collapse;font-size:14px"><tr>' + th + '</tr>' + tr + '</table>';
}
function cell_(c) {
  return esc_(c == null ? '' : c).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1">$1</a>').replace(/\n/g, '<br>');
}
function iso_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function fmt_(d, f) { return Utilities.formatDate(d, Session.getScriptTimeZone(), f); }
function addDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function r1_(n) { return Math.round(n * 10) / 10; }
