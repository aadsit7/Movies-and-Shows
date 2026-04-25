/* ═══════════════════════════════════════════════════════════
   Family Media Tracker — Apps Script backend
   Sheet layout:
     Content_Master  — movies (content_type="Movie") and
                        shows  (content_type="TV Show")
     Live_TV_Channels — live TV entries
════════════════════════════════════════════════════════════ */

var CONTENT_MASTER  = 'Content_Master';
var LIVE_TV_SHEET   = 'Live_TV_Channels';

/* Anthropic API key — set in Apps Script Project Settings → Script Properties.
   Property name: ANTHROPIC_API_KEY
   The literal below is a fallback for local testing only; leave it empty in
   the deployed copy. */
var ANTHROPIC_API_KEY = '';
var ANTHROPIC_MODEL   = 'claude-opus-4-7';

function getAnthropicKey() {
  var fromProps = '';
  try {
    fromProps = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || '';
  } catch (_) {}
  return fromProps || ANTHROPIC_API_KEY || '';
}

/* Fields projected for each content type. Optional columns
   (streaming_on, imdb_score, etc.) are surfaced if the sheet has them; if
   not, projectFields just emits an empty string for that key. */
var CONTENT_FIELDS = [
  'title', 'content_type', 'genre_primary', 'age_rating',
  'description', 'year_started', 'seasons_count', 'tone', 'family_safe',
  'streaming_on', 'imdb_score', 'cast', 'director',
  'network', 'status', 'latest_episode', 'next_airs'
];
var LIVE_TV_FIELDS = [
  'favorite_team_or_channel', 'live_tv_type', 'league',
  'default_channel_or_provider', 'profile_name',
  'network', 'genre', 'description', 'streaming_on',
  'next_game', 'tv_channel'
];

/* ── Cache ───────────────────────────────────────────────── */
var CACHE_KEY = 'allMedia_v1';
var CACHE_TTL = 300; // seconds (5 minutes)

function invalidateCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (_) {}
}

/* ── Routing ─────────────────────────────────────────────── */
function doGet(e) {
  try {
    var action = e && e.parameter ? e.parameter.action : 'getAllMedia';
    var force  = e && e.parameter && e.parameter.forceRefresh === 'true';
    if (action === 'getAllMedia') {
      if (force) invalidateCache();
      return respondJson(readAllMedia());
    }
    return respondJson({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respondJson({ error: err.message });
  }
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'addRow')            return respondJson(handleAddRow(body.sheetName, body.rowData));
    if (action === 'updateRow')         return respondJson(handleUpdateRow(body.sheetName, body.rowIndex, body.rowData));
    if (action === 'deleteRow')         return respondJson(handleDeleteRow(body.sheetName, body.rowIndex));
    if (action === 'claudeSearch')      return respondJson(handleClaudeSearch(body.query, body.sheetName));
    if (action === 'removeDuplicates')  return respondJson(removeDuplicatesFromSheet(body.sheetName));

    return respondJson({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respondJson({ error: err.message });
  }
}

/* ── Read all media ──────────────────────────────────────── */

/* Returns cached JSON when available so repeated fetches are near-instant.
   The cache is invalidated whenever rows are written or duplicates removed. */
function readAllMedia() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  var result = fetchAllMediaFromSheet();

  try { cache.put(CACHE_KEY, JSON.stringify(result), CACHE_TTL); } catch (_) {}

  return result;
}

/* Reads both sheets and returns deduplicated arrays.
   Duplicates are detected case-insensitively on title (Content_Master)
   and favorite_team_or_channel (Live_TV_Channels). Only the first
   occurrence of each key is kept so rowIndex remains valid for updates. */
function fetchAllMediaFromSheet() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var movies  = [];
  var shows   = [];
  var liveTV  = [];

  /* Content_Master → split by content_type, dedup by title */
  var contentSheet = ss.getSheetByName(CONTENT_MASTER);
  if (contentSheet) {
    var rawData = contentSheet.getDataRange().getValues();
    if (rawData.length > 1) {
      var headers     = normalizeHeaders(rawData[0]);
      var seenContent = {};
      for (var i = 1; i < rawData.length; i++) {
        var row = buildObj(headers, rawData[i]);
        var key = String(row['title'] || '').toLowerCase().trim();
        if (!key) continue;            // skip blank rows
        if (seenContent[key]) continue; // skip duplicates
        seenContent[key] = true;
        var item = projectFields(row, CONTENT_FIELDS);
        item.rowIndex = i + 1;
        var ct = String(row['content_type'] || '').trim();
        if (ct === 'Movie')        movies.push(item);
        else if (ct === 'TV Show') shows.push(item);
      }
    }
  }

  /* Live_TV_Channels — dedup by favorite_team_or_channel */
  var liveSheet = ss.getSheetByName(LIVE_TV_SHEET);
  if (liveSheet) {
    var liveRaw = liveSheet.getDataRange().getValues();
    if (liveRaw.length > 1) {
      var liveHeaders = normalizeHeaders(liveRaw[0]);
      var seenLive    = {};
      for (var j = 1; j < liveRaw.length; j++) {
        var liveRow = buildObj(liveHeaders, liveRaw[j]);
        var liveKey = String(liveRow['favorite_team_or_channel'] || '').toLowerCase().trim();
        if (!liveKey) continue;        // skip blank rows
        if (seenLive[liveKey]) continue; // skip duplicates
        seenLive[liveKey] = true;
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
    var liveRow  = mapToSheetRow(rowData, 'liveTV');
    var liveTitle = (liveRow.favorite_team_or_channel || '').toLowerCase().trim();
    if (liveTitle && hasDuplicate(sheet, 'favorite_team_or_channel', liveTitle)) {
      return { success: true, duplicate: true };
    }
    appendByHeaders(sheet, liveRow);

  } else {
    /* Movies and Shows both go into Content_Master */
    var sheet = ss.getSheetByName(CONTENT_MASTER);
    if (!sheet) return { error: CONTENT_MASTER + ' sheet not found' };

    var kind       = isShowsSheet(sheetName) ? 'TV Show' : 'Movie';
    var contentRow = mapToSheetRow(rowData, kind);
    var titleVal   = (contentRow.title || '').toLowerCase().trim();
    if (titleVal && hasDuplicate(sheet, 'title', titleVal)) {
      return { success: true, duplicate: true };
    }
    appendByHeaders(sheet, contentRow);
  }

  invalidateCache();
  return { success: true };
}

/* Returns true if the sheet already has a row whose titleHeader column
   matches newTitle (case-insensitive). */
function hasDuplicate(sheet, titleHeader, newTitle) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  var headers  = normalizeHeaders(data[0]);
  var titleIdx = headers.indexOf(titleHeader);
  if (titleIdx === -1) return false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][titleIdx]).toLowerCase().trim() === newTitle) return true;
  }
  return false;
}

/* ── Remove duplicates from a sheet ─────────────────────── */
/* Scans the sheet for rows whose key column (title for Content_Master,
   favorite_team_or_channel for Live_TV_Channels) appears more than once
   (case-insensitive). All occurrences after the first are deleted.
   Rows are deleted from the bottom up so indices don't shift mid-loop.
   Returns { success, removed } where removed is the count of deleted rows. */
function removeDuplicatesFromSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet, keyCol;

  if (isLiveTVSheet(sheetName)) {
    sheet  = ss.getSheetByName(LIVE_TV_SHEET);
    keyCol = 'favorite_team_or_channel';
  } else {
    sheet  = ss.getSheetByName(CONTENT_MASTER);
    keyCol = 'title';
  }
  if (!sheet) return { error: 'Sheet not found: ' + sheetName };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, removed: 0 };

  var headers = normalizeHeaders(data[0]);
  var keyIdx  = headers.indexOf(keyCol);
  if (keyIdx === -1) return { error: 'Key column not found: ' + keyCol };

  var seen         = {};
  var rowsToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][keyIdx]).toLowerCase().trim();
    if (!key) continue;
    if (seen[key]) {
      rowsToDelete.push(i + 1); // 1-based sheet row number
    } else {
      seen[key] = true;
    }
  }

  /* Delete from bottom to top so earlier row indices stay valid */
  for (var d = rowsToDelete.length - 1; d >= 0; d--) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  if (rowsToDelete.length > 0) invalidateCache();

  return { success: true, removed: rowsToDelete.length };
}

/* ── Map Claude's output keys → sheet header keys ────────── */
/* appendByHeaders only writes columns whose headers exist on the sheet, so
   any keys returned here that the sheet doesn't have are silently dropped.
   That means you can add columns like streaming_on, imdb_score, cast,
   director, network, status, latest_episode, next_airs to Content_Master
   and they will start populating without any code change. */
function mapToSheetRow(data, kind) {
  data = data || {};
  if (kind === 'liveTV') {
    return {
      favorite_team_or_channel:   firstOf(data, ['favorite_team_or_channel', 'channel', 'title']),
      live_tv_type:               firstOf(data, ['live_tv_type']),
      league:                     firstOf(data, ['league']),
      default_channel_or_provider: firstOf(data, ['default_channel_or_provider', 'streamingOn', 'tvChannel', 'network']),
      profile_name:               firstOf(data, ['profile_name']),
      /* optional richer columns */
      network:       firstOf(data, ['network']),
      genre:         firstOf(data, ['genre']),
      description:   firstOf(data, ['description']),
      streaming_on:  firstOf(data, ['streamingOn', 'streaming_on']),
      next_game:     firstOf(data, ['nextGame', 'next_game']),
      tv_channel:    firstOf(data, ['tvChannel', 'tv_channel'])
    };
  }

  return {
    title:          firstOf(data, ['title']),
    content_type:   kind === 'TV Show' ? 'TV Show' : 'Movie',
    genre_primary:  firstOf(data, ['genre_primary', 'genre']),
    age_rating:     firstOf(data, ['age_rating', 'rating']),
    description:    firstOf(data, ['description']),
    year_started:   firstOf(data, ['year_started', 'year']),
    seasons_count:  firstOf(data, ['seasons_count', 'seasons']),
    tone:           firstOf(data, ['tone']),
    family_safe:    firstOf(data, ['family_safe']),
    /* optional richer columns */
    streaming_on:   firstOf(data, ['streamingOn', 'streaming_on']),
    imdb_score:     firstOf(data, ['imdbScore', 'imdb_score', 'imdb']),
    cast:           firstOf(data, ['cast']),
    director:       firstOf(data, ['director']),
    network:        firstOf(data, ['network']),
    status:         firstOf(data, ['status']),
    latest_episode: firstOf(data, ['latestEpisode', 'latest_episode']),
    next_airs:      firstOf(data, ['nextAirs', 'next_airs', 'nextAiring'])
  };
}

function firstOf(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    var v = obj[keys[i]];
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return '';
}

/* ── Update row ──────────────────────────────────────────── */
function handleUpdateRow(sheetName, rowIndex, rowData) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = isLiveTVSheet(sheetName)
    ? ss.getSheetByName(LIVE_TV_SHEET)
    : ss.getSheetByName(CONTENT_MASTER);

  if (!sheet) return { error: 'Sheet not found for: ' + sheetName };

  var lastCol  = sheet.getLastColumn();
  var headers  = normalizeHeaders(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  var existing = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  /* Normalize Claude-style keys (streamingOn → streaming_on, etc.)
     so a refresh always persists the freshest values from Claude. */
  var kind       = inferContentKind(sheetName);
  var mappedKind = kind === 'liveTV' ? 'liveTV' : (kind === 'TV Show' ? 'TV Show' : 'Movie');
  var normalized = mapToSheetRow(rowData, mappedKind);

  var row = headers.map(function(h, i) {
    /* Prefer normalized value, then direct key match, then keep existing cell. */
    var v = normalized[h];
    if (v !== undefined && v !== '') return v;
    v = rowData[h];
    if (v !== undefined && v !== '') return v;
    return existing[i] !== undefined ? existing[i] : '';
  });

  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  invalidateCache();
  return { success: true };
}

/* ── Delete row ──────────────────────────────────────────── */
function handleDeleteRow(sheetName, rowIndex) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = isLiveTVSheet(sheetName)
    ? ss.getSheetByName(LIVE_TV_SHEET)
    : ss.getSheetByName(CONTENT_MASTER);

  if (!sheet) return { error: 'Sheet not found for: ' + sheetName };

  var rowNum = parseInt(rowIndex, 10);
  if (isNaN(rowNum) || rowNum < 2) return { error: 'Invalid rowIndex: ' + rowIndex };
  if (rowNum > sheet.getLastRow()) return { error: 'Row out of range: ' + rowIndex };

  sheet.deleteRow(rowNum);
  invalidateCache();
  return { success: true };
}

/* ── Claude search (with web_search tool) ────────────────── */
function handleClaudeSearch(query, sheetName) {
  var apiKey = getAnthropicKey();
  if (!apiKey) {
    return { error: 'Missing ANTHROPIC_API_KEY — set it in Apps Script → Project Settings → Script Properties' };
  }

  var prompt =
    'You are a media database assistant. The user searched for: "' + query + '"\n\n' +
    'Use the web_search tool to look up current, accurate information from credible sources ' +
    '(IMDb, Rotten Tomatoes, Wikipedia, official network and streaming-service pages). ' +
    'Then return ONLY a single JSON object, no markdown, no explanation, no prose around it.\n\n' +
    'For a Movie use:\n' +
    '{"type":"Movie","title":"","year":"","genre":"","rating":"","description":"","director":"","cast":"","streamingOn":"","imdbScore":""}\n\n' +
    'For a TV Show use:\n' +
    '{"type":"Show","title":"","year":"","genre":"","rating":"","description":"","network":"","seasons":"","latestEpisode":"","status":"","nextAirs":"","cast":"","streamingOn":"","imdbScore":""}\n\n' +
    'For Live TV or Sports use:\n' +
    '{"type":"LiveTV","channel":"","network":"","league":"","genre":"","description":"","streamingOn":"","nextGame":"","tvChannel":""}\n\n' +
    'Be accurate. Real data only. JSON only.';

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 2048,
    tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 5 }],
    messages: [{ role: 'user', content: prompt }]
  };

  var response;
  try {
    response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:      'post',
      contentType: 'application/json',
      headers: {
        'x-api-key':         apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (netErr) {
    return { error: 'Network error contacting Anthropic: ' + netErr.message };
  }

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code < 200 || code >= 300) {
    var apiErr;
    try { apiErr = JSON.parse(body); } catch (_) {}
    var msg = (apiErr && apiErr.error && apiErr.error.message) || body;
    return { error: 'Anthropic API ' + code + ': ' + msg };
  }

  var result;
  try { result = JSON.parse(body); }
  catch (e) { return { error: 'Bad API response: ' + body.substring(0, 200) }; }

  var text = extractTextFromContent(result.content);
  if (!text) return { error: 'Empty response from Claude' };

  var mediaData = parseJsonFromText(text);
  if (!mediaData) {
    return { error: 'Could not parse JSON from response', raw: text.substring(0, 400) };
  }

  if (sheetName) {
    try { handleAddRow(sheetName, mediaData); } catch (writeErr) { /* don't fail search on write */ }
  }

  return { success: true, data: mediaData };
}

/* Pull the concatenated text out of an Anthropic content array. With the
   web_search tool, content can include server_tool_use and
   web_search_tool_result blocks; we only want the model's final text. */
function extractTextFromContent(content) {
  if (!Array.isArray(content)) return '';
  var out = '';
  for (var i = 0; i < content.length; i++) {
    var block = content[i];
    if (block && block.type === 'text' && block.text) out += block.text;
  }
  return out;
}

/* Extract the outermost JSON object from a possibly-fenced text blob. */
function parseJsonFromText(text) {
  if (!text) return null;
  var stripped = String(text).replace(/```json\s*/gi, '').replace(/```/g, '').trim();
  var start = stripped.indexOf('{');
  var end   = stripped.lastIndexOf('}');
  if (start === -1 || end <= start) return null;
  try { return JSON.parse(stripped.substring(start, end + 1)); }
  catch (_) { return null; }
}

function claudeEnrichSearch(query, sheetName) {
  var apiKey = getAnthropicKey();
  if (!apiKey) {
    return { error: 'Missing ANTHROPIC_API_KEY — set it in Apps Script → Project Settings → Script Properties' };
  }

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
      'x-api-key':         apiKey,
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
