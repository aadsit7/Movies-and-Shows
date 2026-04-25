/* ═══════════════════════════════════════════════════════════
   Family Media Tracker — Apps Script backend
   Sheet layout:
     Content_Master  — movies (content_type="Movie") and
                        shows  (content_type="TV Show")
     Live_TV_Channels — live TV entries
════════════════════════════════════════════════════════════ */

var CONTENT_MASTER  = 'Content_Master';
var LIVE_TV_SHEET   = 'Live_TV_Channels';

/* Fields projected for each content type */
var CONTENT_FIELDS = [
  'title', 'content_type', 'genre_primary', 'age_rating',
  'description', 'year_started', 'seasons_count', 'tone', 'family_safe'
];
var LIVE_TV_FIELDS = [
  'favorite_team_or_channel', 'live_tv_type', 'league',
  'default_channel_or_provider', 'profile_name'
];

/* ── Routing ─────────────────────────────────────────────── */
function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : 'getAllMedia';
    if (action === 'getAllMedia') return respondJson(readAllMedia());
    return respondJson({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respondJson({ error: err.message });
  }
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'addRow')      return respondJson(handleAddRow(body.sheetName, body.rowData));
    if (action === 'updateRow')   return respondJson(handleUpdateRow(body.sheetName, body.rowIndex, body.rowData));
    if (action === 'claudeSearch') return respondJson(handleClaudeSearch(body.query, body.sheetName));

    return respondJson({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respondJson({ error: err.message });
  }
}

/* ── Read all media ──────────────────────────────────────── */
function readAllMedia() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var movies  = [];
  var shows   = [];
  var liveTV  = [];

  /* Content_Master → split by content_type */
  var contentSheet = ss.getSheetByName(CONTENT_MASTER);
  if (contentSheet) {
    var rawData = contentSheet.getDataRange().getValues();
    if (rawData.length > 1) {
      var headers = normalizeHeaders(rawData[0]);
      for (var i = 1; i < rawData.length; i++) {
        var row = buildObj(headers, rawData[i]);
        var item = projectFields(row, CONTENT_FIELDS);
        item.rowIndex = i + 1;
        var ct = String(row['content_type'] || '').trim();
        if (ct === 'Movie')   movies.push(item);
        else if (ct === 'TV Show') shows.push(item);
      }
    }
  }

  /* Live_TV_Channels */
  var liveSheet = ss.getSheetByName(LIVE_TV_SHEET);
  if (liveSheet) {
    var liveRaw = liveSheet.getDataRange().getValues();
    if (liveRaw.length > 1) {
      var liveHeaders = normalizeHeaders(liveRaw[0]);
      for (var j = 1; j < liveRaw.length; j++) {
        var liveRow = buildObj(liveHeaders, liveRaw[j]);
        var liveItem = projectFields(liveRow, LIVE_TV_FIELDS);
        liveItem.rowIndex = j + 1;
        liveTV.push(liveItem);
      }
    }
  }

  return { movies: movies, shows: shows, liveTV: liveTV };
}

/* ── Add row ─────────────────────────────────────────────── */
function handleAddRow(sheetName, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (isLiveTVSheet(sheetName)) {
    var sheet = ss.getSheetByName(LIVE_TV_SHEET);
    if (!sheet) return { error: LIVE_TV_SHEET + ' sheet not found' };
    appendByHeaders(sheet, rowData);

  } else {
    /* Movies and Shows both go into Content_Master */
    var sheet = ss.getSheetByName(CONTENT_MASTER);
    if (!sheet) return { error: CONTENT_MASTER + ' sheet not found' };

    /* Ensure content_type is set from the caller's sheetName if absent */
    if (!rowData['content_type']) {
      rowData = shallowCopy(rowData);
      rowData['content_type'] = isShowsSheet(sheetName) ? 'TV Show' : 'Movie';
    }
    appendByHeaders(sheet, rowData);
  }

  return { success: true };
}

/* ── Update row ──────────────────────────────────────────── */
function handleUpdateRow(sheetName, rowIndex, rowData) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = isLiveTVSheet(sheetName)
    ? ss.getSheetByName(LIVE_TV_SHEET)
    : ss.getSheetByName(CONTENT_MASTER);

  if (!sheet) return { error: 'Sheet not found for: ' + sheetName };

  var headers = normalizeHeaders(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  var row     = headers.map(function(h) { return rowData[h] !== undefined ? rowData[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  return { success: true };
}

/* ── Claude search (placeholder — wire up your API key here) */
function handleClaudeSearch(query, sheetName) {
  /* TODO: call Claude API and return a JSON object with media metadata */
  return { error: 'Claude search not yet configured on the server' };
}

/* ── Helpers ─────────────────────────────────────────────── */

function respondJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeHeaders(headerRow) {
  return headerRow.map(function(h) { return String(h).trim(); });
}

function buildObj(headers, rowValues) {
  var obj = {};
  headers.forEach(function(h, i) { obj[h] = rowValues[i]; });
  return obj;
}

function projectFields(obj, fields) {
  var out = {};
  fields.forEach(function(f) { out[f] = obj[f] !== undefined ? obj[f] : ''; });
  return out;
}

function appendByHeaders(sheet, rowData) {
  var headers = normalizeHeaders(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  var row = headers.map(function(h) { return rowData[h] !== undefined ? rowData[h] : ''; });
  sheet.appendRow(row);
}

function isLiveTVSheet(name) {
  if (!name) return false;
  var n = name.toLowerCase().replace(/[\s_]/g, '');
  return n === 'livetv' || n === 'livetchannels' || n === 'live_tv_channels';
}

function isShowsSheet(name) {
  if (!name) return false;
  var n = name.toLowerCase().trim();
  return n === 'shows';
}

function shallowCopy(obj) {
  var copy = {};
  for (var k in obj) { if (obj.hasOwnProperty(k)) copy[k] = obj[k]; }
  return copy;
}
