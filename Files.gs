/**
 * Reads the weekly Dev (Excel) and DM/SEO (PDF) files that the teams drop into one Google Drive folder,
 * adds their numbers to the "Team Updates" and "SEO Keywords" tabs, and keeps the original files
 * so they can be attached to the weekly client email.
 *
 * Needs: Services (+) > Drive API > version v3 > Add.  See README.md.
 */

var FOLDER_ID = 'PASTE-DRIVE-FOLDER-ID-HERE';  // the ID in the folder's web address after /folders/
var ATTACH_FILES = true;                       // attach the original Dev and DM files to the client email

// ---------- Entry points ----------
// Runs daily (see createWeeklyTrigger). Safe to run by hand any time.
function importNewFiles() {
  if (/^PASTE/.test(FOLDER_ID)) return;   // folder not set up yet
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var state = getState_(), it = folder.getFiles(), log = logSheet_();
  while (it.hasNext()) {
    var f = it.next(), id = f.getId();
    if (state[id]) continue;
    var kind = kindOf_(f.getName());
    if (!kind) continue;
    var week = weekOfUpload_(f.getDateCreated());
    var rec = { week: iso_(week), name: f.getName(), kind: kind, emailed: false, note: '' };
    try {
      var parsed = kind === 'dev' ? { rows: parseDev_(xlsxValues_(f)), keywords: [] } : parseDm_(pdfText_(f));
      var n = writeUpdates_(week, parsed.rows, parsed.keywords);
      rec.note = 'read ' + parsed.rows.length + ' rows, ' + parsed.keywords.length + ' keywords (' + n + ' new)';
    } catch (e) {
      rec.note = 'could not read numbers: ' + e.message;   // the file can still be attached
    }
    state[id] = rec;
    log.appendRow([new Date(), '', 'Import: ' + f.getName() + ' for week of ' + rec.week, rec.note]);
  }
  saveState_(state);
}

// Shows what the script would read, without writing anything. Check the Execution log.
function previewFolderFiles() {
  var it = DriveApp.getFolderById(FOLDER_ID).getFiles();
  while (it.hasNext()) {
    var f = it.next(), kind = kindOf_(f.getName());
    if (!kind) { Logger.log('Ignored: ' + f.getName()); continue; }
    try {
      var p = kind === 'dev' ? { rows: parseDev_(xlsxValues_(f)), keywords: [] } : parseDm_(pdfText_(f));
      Logger.log(f.getName() + ' (' + kind + ', week of ' + iso_(weekOfUpload_(f.getDateCreated())) + ')\n' + JSON.stringify(p, null, 1));
    } catch (e) { Logger.log(f.getName() + ': ERROR ' + e.message); }
  }
}

// ---------- Used by the weekly email ----------
function attachmentsForWeek_(weekKey) {
  var out = { blobs: [], ids: [], names: [] };
  if (!ATTACH_FILES) return out;
  var state = getState_();
  Object.keys(state).forEach(function (id) {
    var r = state[id];
    if (r.week === weekKey && !r.emailed) {
      try { out.blobs.push(DriveApp.getFileById(id).getBlob()); out.ids.push(id); out.names.push(r.name); } catch (e) {}
    }
  });
  return out;
}
function markEmailed_(ids) {
  var state = getState_();
  ids.forEach(function (id) { if (state[id]) state[id].emailed = true; });
  saveState_(state);
}

// ---------- File type and week ----------
function kindOf_(name) {
  var n = String(name).toLowerCase();
  if (/\.xlsx?$/.test(n) && /dev/.test(n)) return 'dev';   // keep "Dev" in the Dev file name
  if (/\.pdf$/.test(n)) return 'dm';
  return null;
}
// Files uploaded Thursday to Sunday belong to that week; Monday to Wednesday uploads belong to the week before.
function weekOfUpload_(created) {
  var d = addDays_(new Date(created.getFullYear(), created.getMonth(), created.getDate()), -3);
  return addDays_(d, -((d.getDay() + 6) % 7));
}

// ---------- Reading Google-converted files ----------
function xlsxValues_(f) {
  var tmp = Drive.Files.create({ name: 'tmp-import', mimeType: MimeType.GOOGLE_SHEETS }, f.getBlob());
  try { return SpreadsheetApp.openById(tmp.id).getSheets()[0].getDataRange().getValues(); }
  finally { Drive.Files.remove(tmp.id); }
}
function pdfText_(f) {
  var tmp = Drive.Files.create({ name: 'tmp-import', mimeType: MimeType.GOOGLE_DOCS }, f.getBlob(), { ocrLanguage: 'en' });
  try { return DocumentApp.openById(tmp.id).getBody().getText(); }
  finally { Drive.Files.remove(tmp.id); }
}

// ---------- Parsers (plain functions, easy to check) ----------
// Dev Excel: a list of work descriptions, then a line like "Total: 2 hours".
// Returns rows as [team, activity, count, hours, details].
function parseDev_(values) {
  var rows = [], total = '';
  values.forEach(function (r) {
    var cells = r.map(function (c) { return String(c).trim(); });
    var first = cells.filter(function (c) { return c; })[0];
    if (!first) return;
    var m = first.match(/^total\s*:?\s*([\d.]+)\s*h/i);
    if (m) { total = Number(m[1]); return; }
    if (/^work description$/i.test(first)) return;
    var next = cells[cells.indexOf(first) + 1];
    var hrs = next !== undefined && next !== '' && !isNaN(Number(next)) ? Number(next) : '';
    rows.push(['Dev', first, '', hrs, '']);
  });
  if (total !== '') rows.push(['Dev', 'Total hours', '', total, '']);
  return rows;
}

// DM/SEO PDF text. Any line like "90 Backlinks Created" becomes a counted activity, so new activity names
// are picked up too. URLs and short titles under it become its details.
function parseDm_(text) {
  var t = String(text).replace(/\r/g, '');
  var kwStart = t.search(/Top Ranking Keywords/i);
  var after = kwStart >= 0 ? t.slice(kwStart) : '';
  var kwEndRel = after.search(/Off-?\s?Page|Link Building|Page Indexing/i);
  var kwText = kwStart >= 0 ? (kwEndRel > 0 ? after.slice(0, kwEndRel) : after) : '';
  var rest = kwStart >= 0 && kwEndRel > 0 ? after.slice(kwEndRel) : t;
  return { rows: dmCounts_(rest), keywords: dmKeywords_(kwText) };
}
function dmKeywords_(kwText) {
  var lines = kwText.split('\n').map(function (l) { return l.trim(); }).filter(Boolean), out = [], m;
  lines.forEach(function (l) { if ((m = l.match(/^(\d+)\s+(.+?)\s+(\d+)$/))) out.push([m[2], Number(m[3])]); });
  if (!out.length) {   // table cells on separate lines
    for (var i = 0; i + 2 < lines.length; i++) {
      if (/^\d+$/.test(lines[i]) && !/^\d+$/.test(lines[i + 1]) && /^\d+$/.test(lines[i + 2])) { out.push([lines[i + 1], Number(lines[i + 2])]); i += 2; }
    }
  }
  return out;
}
function dmCounts_(text) {
  var raw = text.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
  var lines = raw.map(function (l) { return l.replace(/^[\s●•○◦▪*\-]+/, '').trim(); });
  var rows = [], cur = null, m;
  lines.forEach(function (l, i) {
    if ((m = l.match(/^(\d+)\s+([A-Za-z].{2,70})$/)) && !/^https?:/i.test(l)) {
      cur = ['SEO', m[2].trim(), Number(m[1]), '', []]; rows.push(cur); return;
    }
    if (/:\s*$/.test(l)) { cur = null; return; }          // a section heading ends the current activity
    if (!cur) return;
    var url = l.match(/^https?:\/\/\S+/);
    if (url) {
      var u = url[0], nxt = lines[i + 1];
      if (nxt && /^[\w\-./%?=&]+$/.test(nxt) && nxt.length < 40 && !/^\d+\s/.test(nxt)) u += nxt;  // URL wrapped onto the next line
      cur[4].push(u); return;
    }
    if (/^[\w\-./%?=&]+$/.test(l) && cur[4].length && /^https?:/.test(cur[4][cur[4].length - 1])) return;   // wrapped part already joined
    if (/^[○◦]/.test(raw[i]) || titleLike_(l)) cur[4].push(l);   // sub-bullets and title-like lines, e.g. a blog title
  });
  rows.forEach(function (r) { r[4] = r[4].join('\n'); });
  if (/Backlink Opportunities Identified/i.test(text)) rows.push(['SEO', 'Competitor Backlink Analysis', '', '', 'Backlink opportunities identified']);
  return rows;
}
// A short line with no commas or full stop where most words start with a capital letter.
function titleLike_(l) {
  if (l.length > 90 || /[.,;:]/.test(l)) return false;
  var w = l.split(/\s+/);
  if (w.length < 2) return false;
  return w.filter(function (x) { return /^["'A-Z0-9]/.test(x); }).length / w.length >= 0.6;
}

// ---------- Writing into the tracker tabs ----------
function ensureTab_(name, head) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function writeUpdates_(week, rows, keywords) {
  var added = 0, u = ensureTab_(UPDATES_SHEET, ['Week Start', 'Team', 'Activity', 'Count', 'Hours', 'Details']);
  var have = {}; readTable_(UPDATES_SHEET, 'Activity').forEach(function (o) { have[mondayKey_(o['Week Start']) + '|' + o['Team'] + '|' + o['Activity']] = 1; });
  rows.forEach(function (r) {
    if (have[iso_(week) + '|' + r[0] + '|' + r[1]]) return;
    u.appendRow([week, r[0], r[1], r[2], r[3], r[4]]); u.getRange(u.getLastRow(), 1).setNumberFormat('yyyy-mm-dd'); added++;
  });
  if (keywords.length) {
    var k = ensureTab_(KEYWORDS_SHEET, ['Week Start', 'Keyword', 'Rank']), kh = {};
    readTable_(KEYWORDS_SHEET, 'Keyword').forEach(function (o) { kh[mondayKey_(o['Week Start']) + '|' + o['Keyword']] = 1; });
    keywords.forEach(function (x) { if (!kh[iso_(week) + '|' + x[0]]) { k.appendRow([week, x[0], x[1]]); k.getRange(k.getLastRow(), 1).setNumberFormat('yyyy-mm-dd'); added++; } });
  }
  return added;
}

// ---------- Remembering which files were handled ----------
function getState_() { return JSON.parse(PropertiesService.getScriptProperties().getProperty('files') || '{}'); }
function saveState_(s) { PropertiesService.getScriptProperties().setProperty('files', JSON.stringify(s)); }
function logSheet_() { return SpreadsheetApp.getActive().getSheetByName(LOG_SHEET) || SpreadsheetApp.getActive().insertSheet(LOG_SHEET); }
