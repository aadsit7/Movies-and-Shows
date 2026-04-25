/* ═══════════════════════════════════════════════════════════
   Family Media Tracker — Apps Script backend
   Sheet layout:
     Content_Master  — movies (content_type="Movie") and
                        shows  (content_type="TV Show")
     Live_TV_Channels — live TV entries
════════════════════════════════════════════════════════════ */

var CONTENT_MASTER  = 'Content_Master';
var LIVE_TV_SHEET   = 'Live_TV_Channels';

/* Anthropic API key — set this before deploying */
var ANTHROPIC_API_KEY = '';
var ANTHROPIC_MODEL   = 'claude-sonnet-4-6';

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

  return { success: true, movies: movies, shows: shows, liveTV: liveTV };
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

/* ── Claude search ───────────────────────────────────────── */
function handleClaudeSearch(query, sheetName) {
  var prompt = 'You are a media database assistant. The user searched for: "' + query + '"\n\n' +
    'Return ONLY valid JSON, no markdown, no explanation.\n\n' +
    'For a Movie use:\n' +
    '{"type":"Movie","title":"","year":"","genre":"","rating":"","description":"","director":"","cast":"","streamingOn":"","imdbScore":""}\n\n' +
    'For a TV Show use:\n' +
    '{"type":"Show","title":"","year":"","genre":"","rating":"","description":"","network":"","seasons":"","latestEpisode":"","status":"","cast":"","streamingOn":"","imdbScore":""}\n\n' +
    'For Live TV or Sports use:\n' +
    '{"type":"LiveTV","channel":"","network":"","league":"","genre":"","description":"","streamingOn":"","nextGame":"","tvChannel":""}\n\n' +
    'Be accurate. Real data only. JSON only.';

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model: 'claude-opus-4-5',
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());
  var text = result.content && result.content[0] ? result.content[0].text : '{}';

  try {
    var mediaData = JSON.parse(text);
    if (sheetName) handleAddRow(sheetName, mediaData);
    return { success: true, data: mediaData };
  } catch(e) {
    return { error: 'Could not parse response', raw: text };
  }
}

function claudeEnrichSearch(query, sheetName) {
  var contentKind  = inferContentKind(sheetName);
  var systemPrompt = buildClaudeSystemPrompt(contentKind);

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 1024,
    system:     systemPrompt,
    messages: [
      { role: 'user', content: 'Look up: ' + query }
    ]
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'x-api-key':         ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    return { error: 'Anthropic API error ' + code + ': ' + body };
  }

  var data = JSON.parse(body);
  var text = (data.content && data.content[0] && data.content[0].text) || '';
  return { result: text };
}

function buildClaudeSystemPrompt(contentKind) {
  if (contentKind === 'liveTV') {
    return 'You look up live TV channel, team, or league metadata. ' +
      'Reply with ONLY a single JSON object using these keys: ' +
      LIVE_TV_FIELDS.join(', ') + '. ' +
      'Use empty strings for unknown fields. ' +
      'No prose, no commentary, no markdown code fences.';
  }
  var typeLabel  = contentKind === 'TV Show' ? 'TV show' : 'movie';
  var typeValue  = contentKind === 'TV Show' ? 'TV Show' : 'Movie';
  return 'You look up ' + typeLabel + ' metadata. ' +
    'Reply with ONLY a single JSON object using these keys: ' +
    CONTENT_FIELDS.join(', ') + '. ' +
    'Set content_type to "' + typeValue + '". ' +
    'family_safe should be "Yes" or "No". ' +
    'Use empty strings for unknown fields. ' +
    'No prose, no commentary, no markdown code fences.';
}

function inferContentKind(sheetName) {
  if (isLiveTVSheet(sheetName)) return 'liveTV';
  if (isShowsSheet(sheetName))  return 'TV Show';
  return 'Movie';
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
