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

  /* Ensure all expected columns exist on both sheets so next_airs / next_game
     are never silently dropped, even before the first write. */
  var contentSheet = ss.getSheetByName(CONTENT_MASTER);
  var liveSheet    = ss.getSheetByName(LIVE_TV_SHEET);
  if (contentSheet) ensureColumns(contentSheet, CONTENT_FIELDS);
  if (liveSheet)    ensureColumns(liveSheet,    LIVE_TV_FIELDS);

  /* Content_Master → split by content_type, dedup by (content_type + title)
     so a Movie and a TV Show with the same title are kept as separate entries. */
  if (contentSheet) {
    var rawData = contentSheet.getDataRange().getValues();
    if (rawData.length > 1) {
      var headers    = normalizeHeaders(rawData[0]);
      var seenMovies = {};
      var seenShows  = {};
      for (var i = 1; i < rawData.length; i++) {
        var row = buildObj(headers, rawData[i]);
        var key = String(row['title'] || '').toLowerCase().trim();
        if (!key) continue;            // skip blank rows
        var item = projectFields(row, CONTENT_FIELDS);
        item.rowIndex = i + 1;
        var ct = String(row['content_type'] || '').trim();
        if (ct === 'Movie') {
          if (seenMovies[key]) continue;
          seenMovies[key] = true;
          movies.push(item);
        } else if (ct === 'TV Show') {
          if (seenShows[key]) continue;
          seenShows[key] = true;
          shows.push(item);
        }
      }
    }
  }

  /* Live_TV_Channels — dedup by favorite_team_or_channel */
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

/* ── Ensure columns exist ────────────────────────────────── */
/* Adds any headers from requiredCols that are not yet present as columns
   in the sheet. Call this before every write so optional fields like
   next_airs and next_game are never silently dropped. */
function ensureColumns(sheet, requiredCols) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? normalizeHeaders(sheet.getRange(1, 1, 1, lastCol).getValues()[0])
    : [];
  var added = 0;
  for (var i = 0; i < requiredCols.length; i++) {
    if (headers.indexOf(requiredCols[i]) === -1) {
      sheet.getRange(1, lastCol + added + 1).setValue(requiredCols[i]);
      added++;
    }
  }
  if (added > 0) invalidateCache();
}

/* ── Add row ─────────────────────────────────────────────── */
function handleAddRow(sheetName, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (isLiveTVSheet(sheetName)) {
    var sheet = ss.getSheetByName(LIVE_TV_SHEET);
    if (!sheet) return { error: LIVE_TV_SHEET + ' sheet not found' };
    ensureColumns(sheet, LIVE_TV_FIELDS);
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
    ensureColumns(sheet, CONTENT_FIELDS);

    var kind       = isShowsSheet(sheetName) ? 'TV Show' : 'Movie';
    var contentRow = mapToSheetRow(rowData, kind);
    var titleVal   = (contentRow.title || '').toLowerCase().trim();
    if (titleVal && hasDuplicate(sheet, 'title', titleVal, kind)) {
      return { success: true, duplicate: true };
    }
    appendByHeaders(sheet, contentRow);
  }

  invalidateCache();
  return { success: true };
}

/* Returns true if the sheet already has a row whose titleHeader column
   matches newTitle (case-insensitive). When contentType is provided, also
   requires the content_type column to match — so a Movie and a TV Show with
   the same title are not considered duplicates of each other. */
function hasDuplicate(sheet, titleHeader, newTitle, contentType) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  var headers  = normalizeHeaders(data[0]);
  var titleIdx = headers.indexOf(titleHeader);
  if (titleIdx === -1) return false;
  var ctIdx = contentType ? headers.indexOf('content_type') : -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][titleIdx]).toLowerCase().trim() !== newTitle) continue;
    if (ctIdx !== -1 && String(data[i][ctIdx]).trim() !== contentType) continue;
    return true;
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

  /* For Content_Master, include content_type in the key so a Movie and a TV
     Show with identical titles are treated as separate entries, not duplicates. */
  var ctIdx = !isLiveTVSheet(sheetName) ? headers.indexOf('content_type') : -1;

  var seen         = {};
  var rowsToDelete = [];

  for (var i = 1; i < data.length; i++) {
    var title = String(data[i][keyIdx]).toLowerCase().trim();
    if (!title) continue;
    var key = ctIdx !== -1
      ? String(data[i][ctIdx]).trim().toLowerCase() + ':' + title
      : title;
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
      next_game:     firstOf(data, ['nextGame', 'next_game', 'nextAirs', 'next_airs']),
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
    next_airs:      firstOf(data, ['nextAirs', 'next_airs', 'nextAiring', 'airing', 'whenitairs'])
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

  /* Ensure all expected columns exist so next_airs / next_game are never
     silently dropped when the sheet is missing those headers. */
  ensureColumns(sheet, isLiveTVSheet(sheetName) ? LIVE_TV_FIELDS : CONTENT_FIELDS);

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

  var today = new Date();
  var yyyy  = today.getFullYear();
  var mm    = String(today.getMonth() + 1).padStart(2, '0');
  var dd    = String(today.getDate()).padStart(2, '0');
  var todayStr = yyyy + '-' + mm + '-' + dd;

  var prompt =
    'You are a media database assistant. Today\'s date is ' + todayStr + '. The user searched for: "' + query + '"\n\n' +
    'Use the web_search tool (up to 8 times) to look up current, accurate information from ' +
    'credible sources (IMDb, Rotten Tomatoes, Wikipedia, official network and streaming-service pages, ' +
    'TV Guide, Sports Reference). Search specifically for the next air date / next game if applicable.\n\n' +
    'Return ONLY a single raw JSON object — no markdown fences, no explanation, no extra text.\n\n' +
    'For a Movie use exactly these keys:\n' +
    '{"type":"Movie","title":"","year":"<4-digit year>","genre":"<primary genre>","rating":"<MPAA rating e.g. PG-13>","description":"<1-2 sentence plot summary>","director":"","cast":"<comma-separated top 3 actors>","streamingOn":"<platform name>","imdbScore":"<e.g. 8.2>","tone":"<e.g. Action, Comedy, Drama, Thriller>"}\n\n' +
    'For a TV Show use exactly these keys:\n' +
    '{"type":"Show","title":"","year":"<year show started>","genre":"<primary genre>","rating":"<TV rating e.g. TV-MA>","description":"<1-2 sentence premise>","network":"<broadcast network or streaming service>","seasons":"<number>","latestEpisode":"<S##E## Title if known>","status":"<Returning | Ended | Cancelled | On Hiatus>","nextAirs":"<YYYY-MM-DD HH:MM TZ or descriptive e.g. \'Tuesdays 9PM ET on NBC\'>","cast":"<comma-separated top 3 actors>","streamingOn":"<streaming platform if different from network>","imdbScore":"<e.g. 8.2>","tone":"<e.g. Drama, Comedy, Thriller>"}\n\n' +
    'For Live TV / Sports channel use exactly these keys:\n' +
    '{"type":"LiveTV","channel":"<channel or team name>","network":"<broadcast network>","league":"<e.g. NFL, NBA, EPL>","genre":"<Sports | News | Entertainment>","description":"<brief description>","streamingOn":"<streaming platform>","nextGame":"<YYYY-MM-DD HH:MM TZ or descriptive>","tvChannel":"<cable/satellite channel name>"}\n\n' +
    'Rules: real data only; leave a field empty string if truly unknown; dates MUST be in YYYY-MM-DD format when an exact date is known.';

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 2048,
    tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 8 }],
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
  return headerRow.map(function(h) {
    return String(h).trim().toLowerCase().replace(/\s+/g, '_');
  });
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
