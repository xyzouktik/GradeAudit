/**
 * GradeAudit.gs
 *
 * Reads a Google Doc, finds every grading entry (person + rating 1-10 + who
 * graded them), keeps only the entries that YOU graded, and emails you the list.
 *
 * The people's names are never hardcoded: they are scanned from the document.
 *
 * SETUP
 *  1. Go to https://script.google.com and create a New project.
 *  2. Paste this whole file over the default Code.gs.
 *  3. Set DOC_URL_OR_ID below (and EMAIL if you want it somewhere else).
 *  4. Run the function `emailGradedList` once and accept the permission prompt.
 *
 * HOW "GRADED BY" IS READ
 *  - Table with a header row, e.g. columns: Person | Rating | Graded by
 *    (the script auto-detects the name / rating / graded-by columns).
 *  - Inline lines, e.g.:  Alice - 8 - graded by john@example.com
 *                         Bob: 7/10 (by John Doe)
 *  - An entry is only reported if its "graded by" value matches you.
 */

var DOC_URL_OR_ID = '';
var EMAIL = '';
var MIN_RATING = 1;
var MAX_RATING = 10;
var MY_ALIASES = [];

function onOpen() {
  try {
    DocumentApp.getUi()
      .createMenu('Grade Checker')
      .addItem('Email list graded by me', 'emailGradedList')
      .addToUi();
  } catch (e) {
    // Not bound to a document; ignore.
  }
}

function emailGradedList() {
  var doc = openTargetDoc_();
  var me = buildMe_();
  var graded = scanDoc_(doc).filter(function (e) { return isMe_(e.grader, me); });
  if (!graded.length) {
    throw new Error('Found no entries in "' + doc.getName() + '" whose "graded by" matches you (' + me.display + ').');
  }
  var to = EMAIL || firstEmail_(me) || Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    subject: 'People graded by you in "' + doc.getName() + '"',
    htmlBody: buildHtml_(doc, graded, me)
  });
  Logger.log('Emailed ' + graded.length + ' entries graded by ' + me.display + ' to ' + to);
}

function buildHtml_(doc, graded, me) {
  var rows = graded.map(function (e) {
    return '<tr><td style="padding:4px 12px;border:1px solid #ccc">' + escapeHtml_(e.name) +
      '</td><td style="padding:4px 12px;border:1px solid #ccc;text-align:center">' + e.rating + '/' + MAX_RATING +
      '</td><td style="padding:4px 12px;border:1px solid #ccc">' + escapeHtml_(e.grader) + '</td></tr>';
  }).join('');
  return '<h2>People graded by you in "' + escapeHtml_(doc.getName()) + '"</h2>' +
    '<p>' + graded.length + ' entr' + (graded.length === 1 ? 'y' : 'ies') +
    ' matched you (' + escapeHtml_(me.display) + ').</p>' +
    '<table style="border-collapse:collapse;font-family:Arial,sans-serif">' +
    '<tr><th style="padding:4px 12px;border:1px solid #ccc;background:#f0f0f0">Person</th>' +
    '<th style="padding:4px 12px;border:1px solid #ccc;background:#f0f0f0">Rating</th>' +
    '<th style="padding:4px 12px;border:1px solid #ccc;background:#f0f0f0">Graded by</th></tr>' +
    rows + '</table>';
}

function openTargetDoc_() {
  if (DOC_URL_OR_ID) return DocumentApp.openById(extractId_(DOC_URL_OR_ID));
  var active = DocumentApp.getActiveDocument();
  if (active) return active;
  throw new Error('Set DOC_URL_OR_ID at the top of the script (or bind the script to the document).');
}

function extractId_(urlOrId) {
  var m = String(urlOrId).match(/\/d\/([-\w]+)/);
  return m ? m[1] : String(urlOrId).trim();
}

function scanDoc_(doc) {
  var entries = [];
  var body = doc.getBody();
  body.getParagraphs().forEach(function (p) {
    var e = parseInline_(p.getText());
    if (e) entries.push(e);
  });
  body.getTables().forEach(function (t) { scanTable_(t, entries); });
  return entries;
}

function scanTable_(table, entries) {
  var nRows = table.getNumRows();
  if (!nRows) return;
  var cols = detectColumns_(readRow_(table.getRow(0)));
  var start = 0;
  if (cols.name !== -1 && cols.rating !== -1 && cols.grader !== -1) {
    start = 1;
  } else {
    cols = guessColumns_(table);
    start = 0;
  }
  for (var r = start; r < nRows; r++) {
    var cells = readRow_(table.getRow(r));
    var name = cols.name !== -1 ? cells[cols.name] : null;
    var rating = cols.rating !== -1 ? parseRatingOnly_(cells[cols.rating]) : null;
    var grader = cols.grader !== -1 ? cells[cols.grader] : null;
    if (name && rating !== null && grader) entries.push({ name: name, rating: rating, grader: grader });
  }
}

function readRow_(row) {
  var cells = [];
  for (var c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText().trim());
  return cells;
}

function detectColumns_(header) {
  var cols = { name: -1, rating: -1, grader: -1 };
  header.forEach(function (h, i) {
    var t = norm_(h);
    if (cols.grader === -1 && /graded\s*by|grader|reviewer|evaluat|assessed\s*by|marked\s*by|^by$|who/.test(t)) cols.grader = i;
    else if (cols.rating === -1 && /rating|score|\bgrade\b|\bmark\b|points|result/.test(t)) cols.rating = i;
    else if (cols.name === -1 && /name|person|student|participant|candidate|member/.test(t)) cols.name = i;
  });
  return cols;
}

function guessColumns_(table) {
  var cols = { name: -1, rating: -1, grader: -1 };
  if (!table.getNumRows()) return cols;
  var cells = readRow_(table.getRow(0));
  var others = [];
  cells.forEach(function (cell, i) {
    if (cols.rating === -1 && parseRatingOnly_(cell) !== null) cols.rating = i;
    else others.push(i);
  });
  if (others.length) cols.name = others[0];
  if (others.length > 1) cols.grader = others[others.length - 1];
  return cols;
}

function parseInline_(text) {
  if (!text || !text.trim()) return null;
  var m = text.match(/^\s*(.+?)\s*[-:–—\t]\s*(\d{1,2})(?![\d\/])(?:\s*\/\s*10|\s*out\s*of\s*10)?\s*(.*)$/i);
  if (!m) return null;
  var rating = Number(m[2]);
  if (rating < MIN_RATING || rating > MAX_RATING) return null;
  var name = cleanName_(m[1]);
  var grader = parseGrader_(m[3]);
  if (!name || !grader || parseRatingOnly_(name) !== null) return null;
  return { name: name, rating: rating, grader: grader };
}

function parseGrader_(tail) {
  var s = String(tail || '').trim();
  if (!s) return null;
  s = s.replace(/^[-–—:,\s]+/, '');
  s = s.replace(/^\(?\s*(?:graded\s*by|grade\s*by|by|grader|marked\s*by|reviewed\s*by|assessed\s*by)?\s*[:：]?\s*/i, '');
  s = s.replace(/[)\s]+$/, '').trim();
  return s || null;
}

function parseRatingOnly_(text) {
  var m = String(text || '').match(/^\s*(\d{1,2})(?![\d\/])\s*(?:\/\s*10|out\s*of\s*10)?\s*$/i);
  if (!m) return null;
  var n = Number(m[1]);
  return (n >= MIN_RATING && n <= MAX_RATING) ? n : null;
}

function cleanName_(s) {
  return String(s || '').replace(/[\s\-:–—]+$/, '').replace(/^[\s\-:–—]+/, '').trim();
}

function buildMe_() {
  var emails = {}, names = {};
  function addEmail(v) { v = String(v || '').trim().toLowerCase(); if (v) emails[v] = true; }
  function addName(v) { v = norm_(v); if (v) names[v] = true; }

  var active = '';
  try { active = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  addEmail(active);
  try { addEmail(Session.getEffectiveUser().getEmail()); } catch (e) {}

  try {
    var c = active ? ContactsApp.getContact(active) : null;
    if (c) {
      addName(c.getFullName());
      addName(c.getGivenName() + ' ' + c.getFamilyName());
      addName(c.getGivenName());
      addName(c.getFamilyName());
      addName(c.getNickname());
    }
  } catch (e) {}

  (MY_ALIASES || []).forEach(function (a) {
    if (String(a).indexOf('@') !== -1) addEmail(a); else addName(a);
  });

  var list = Object.keys(emails).concat(Object.keys(names));
  return { emails: Object.keys(emails), names: Object.keys(names), display: list.join(', ') || '(unknown user)' };
}

function isMe_(grader, me) {
  var g = norm_(grader);
  if (!g) return false;
  for (var i = 0; i < me.emails.length; i++) {
    if (g.indexOf(me.emails[i]) !== -1) return true;
  }
  if (g.indexOf('@') !== -1) return false;
  for (var j = 0; j < me.names.length; j++) {
    var n = me.names[j];
    if (n.length < 3) continue;
    var re = new RegExp('(^|[^a-z])' + escapeRegExp_(n) + '([^a-z]|$)');
    if (re.test(g)) return true;
  }
  return false;
}

function firstEmail_(me) {
  return me.emails.length ? me.emails[0] : '';
}

function norm_(s) {
  return String(s || '').toLowerCase().replace(/[\u00a0\s]+/g, ' ').trim();
}

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
