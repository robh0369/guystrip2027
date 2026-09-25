// Trip vote backend for index.html. Paste into Extensions > Apps Script on a Google Sheet,
// then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).

var VOTERS = ['Matt', 'Jerry', 'Jeff', 'Mike', 'Brian', 'Rob'];
var TRIPS = ['boston', 'sanjuan', 'montreal'];
var POINTS = [3, 2, 1];

function doGet() {
  return json_(state_());
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad request.' });
  }
  var name = String(body.name || '');
  var ranking = body.ranking;
  if (VOTERS.indexOf(name) < 0) return json_({ ok: false, error: 'Unknown voter.' });
  if (!isValidRanking_(ranking)) return json_({ ok: false, error: 'Rank all three trips, each once.' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var now = new Date();
    var votes = sheet_('Votes', ['Name', '1st', '2nd', '3rd', 'Updated']);
    var rows = votes.getDataRange().getValues();
    var row = [name, ranking[0], ranking[1], ranking[2], now];
    var found = false;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === name) {
        votes.getRange(i + 1, 1, 1, row.length).setValues([row]);
        found = true;
        break;
      }
    }
    if (!found) votes.appendRow(row);
    sheet_('History', ['Timestamp', 'Name', '1st', '2nd', '3rd']).appendRow([now, name, ranking[0], ranking[1], ranking[2]]);
  } finally {
    lock.releaseLock();
  }
  return json_(state_());
}

function isValidRanking_(r) {
  if (!Array.isArray(r) || r.length !== TRIPS.length) return false;
  var seen = {};
  for (var i = 0; i < r.length; i++) {
    if (TRIPS.indexOf(r[i]) < 0 || seen[r[i]]) return false;
    seen[r[i]] = true;
  }
  return true;
}

function state_() {
  var rows = sheet_('Votes', ['Name', '1st', '2nd', '3rd', 'Updated']).getDataRange().getValues().slice(1);
  var tally = {};
  TRIPS.forEach(function (t) { tally[t] = { points: 0, first: 0 }; });
  var voters = [];
  rows.forEach(function (r) {
    if (VOTERS.indexOf(r[0]) < 0) return;
    var ranking = [r[1], r[2], r[3]];
    if (!isValidRanking_(ranking)) return;
    ranking.forEach(function (t, i) { tally[t].points += POINTS[i]; });
    tally[ranking[0]].first += 1;
    voters.push({ name: r[0], ranking: ranking, updated: r[4] instanceof Date ? r[4].toISOString() : String(r[4]) });
  });
  return { ok: true, voters: voters, tally: tally, roster: VOTERS };
}

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
