/* ═══════════════════════════════════════════════════════════
   Family Media Tracker — Apps Script backend
   Sheet layout:
     Content_Master    — movies (content_type="Movie") and
                          shows  (content_type="TV Show")
     Live_TV_Channels  — live TV / sports team entries
     Episode_Schedule  — per-episode air dates, joined to shows by title
     Schedules         — per-game dates, joined to live TV by channel_id
                          (matched against favorite_team_or_channel)
════════════════════════════════════════════════════════════ */

var CONTENT_MASTER     = 'Content_Master';
var LIVE_TV_SHEET      = 'Live_TV_Channels';
var EPISODE_SCHEDULE   = 'Episode_Schedule';
var SCHEDULES_SHEET    = 'Schedules';

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

/* Spreadsheet access — works for both container-bound and standalone scripts.
   For a standalone web app, set the SPREADSHEET_ID Script Property to the
   ID from your Google Sheet URL:
     https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit
   For a container-bound script (created inside the sheet), leave it unset. */
function getSpreadsheet() {
  var ssId = '';
  try {
    ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
  } catch (_) {}
  if (ssId) return SpreadsheetApp.openById(ssId);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error(
    'No spreadsheet found. Set the SPREADSHEET_ID Script Property to your sheet\'s ID, ' +
    'or run this script from within your Google Sheet.'
  );
  return active;
}

/* Fields projected for each content type. Optional columns
   (streaming_on, imdb_score, etc.) are surfaced if the sheet has them; if
   not, projectFields just emits an empty string for that key. */
var CONTENT_FIELDS = [
  'title', 'content_type', 'genre_primary', 'age_rating',
  'description', 'year_started', 'seasons_count', 'tone', 'family_safe',
  'streaming_on', 'imdb_score', 'cast', 'director',
  'network', 'status', 'latest_episode', 'next_airs', 'favorites', 'profile',
  'watch_status'
];
var LIVE_TV_FIELDS = [
  'favorite_team_or_channel', 'live_tv_type', 'league',
  'default_channel_or_provider', 'profile_name',
  'network', 'genre', 'description', 'streaming_on',
  'next_game', 'tv_channel', 'favorites', 'profile',
  'watch_status'
];
/* Episode_Schedule: one row per upcoming/recent episode; joined to a show
   by lowercased title. air_date should be ISO YYYY-MM-DD when known. */
var EPISODE_FIELDS = [
  'title', 'season', 'episode', 'episode_title', 'air_date', 'airstamp', 'network'
];
/* Schedules: one row per game; joined to a live-TV team/channel by
   channel_id (matched against favorite_team_or_channel). date should be
   ISO YYYY-MM-DD; time is free-form (e.g. "7:10 PM PT"). */
var SCHEDULE_FIELDS = [
  'channel_id', 'team', 'league', 'date', 'time', 'opponent', 'tv_channel'
];

/* ── Settings sheet ──────────────────────────────────────── */
/* Reads the Settings tab (if it exists) and returns an object of
   { setting_name: value } pairs. Falls back to safe defaults when the
   sheet is absent or a row is missing so the app keeps working even
   before the Settings tab is created. */
function getSettings() {
  var defaults = {
    search_enabled:    'TRUE',
    writes_enabled:    'TRUE',
    sports_enabled:    'TRUE',
    movies_enabled:    'TRUE',
    shows_enabled:     'TRUE',
    default_timezone:  'America/Los_Angeles'
  };
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return defaults;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0] || '').trim().toLowerCase().replace(/\s+/g, '_');
      var val = String(data[i][1] || '').trim();
      if (key) defaults[key] = val;
    }
  } catch (_) {}
  return defaults;
}

function settingEnabled(settings, key) {
  var v = String(settings[key] || 'TRUE').toUpperCase();
  return v !== 'FALSE' && v !== '0' && v !== 'NO';
}

/* ── Cache ───────────────────────────────────────────────── */
var CACHE_KEY = 'allMedia_v1';
var CACHE_TTL = 1800; // seconds (30 minutes)

function invalidateCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (_) {}
}

/* ═══════════════════════════════════════════════════════════
   EXTERNAL SEARCH APIS
   Three reliable, purpose-built data sources replace the
   Claude web-search approach for initial content discovery:
     • TVmaze  — TV show search + episode schedule (free, no key)
     • TMDB    — Movie search + streaming providers (free API key)
     • TheSportsDB — Sports team search + upcoming events (free key "3")
   Claude is still used for card-level refreshes (richer narrative data).
════════════════════════════════════════════════════════════ */

function getTMDBKey() {
  try { return PropertiesService.getScriptProperties().getProperty('TMDB_API_KEY') || ''; } catch (_) { return ''; }
}
function getSportsDBKey() {
  try { return PropertiesService.getScriptProperties().getProperty('THESPORTSDB_API_KEY') || '3'; } catch (_) { return '3'; }
}

/* TMDB genre ID → name map (movie + TV IDs) */
var TMDB_GENRES = {
  28:'Action', 12:'Adventure', 16:'Animation', 35:'Comedy', 80:'Crime',
  99:'Documentary', 18:'Drama', 10751:'Family', 14:'Fantasy', 36:'History',
  27:'Horror', 10402:'Music', 9648:'Mystery', 10749:'Romance', 878:'Sci-Fi',
  53:'Thriller', 10752:'War', 37:'Western', 10759:'Action & Adventure',
  10762:'Kids', 10765:'Sci-Fi & Fantasy', 10768:'War & Politics'
};
function tmdbGenreNames(ids) {
  if (!Array.isArray(ids) || !ids.length) return '';
  return ids.slice(0, 2).map(function(id) { return TMDB_GENRES[id] || ''; }).filter(Boolean).join(', ');
}

/* ── Unified search dispatcher ───────────────────────────── */
/* Two-round parallel fetch: all primary searches fire together, then all
   secondary enrichment requests (episodes / providers / events) fire together,
   cutting 6 sequential HTTP round-trips down to 2. */
function handleSearch(query, searchType) {
  var settings = getSettings();
  if (!settingEnabled(settings, 'search_enabled')) {
    return { error: 'Search is disabled. Set search_enabled = TRUE in the Settings sheet.' };
  }
  if (!query || !String(query).trim()) return { error: 'Missing search query' };

  var type     = String(searchType || 'all').toLowerCase().replace(/s$/, '');
  var tmdbKey  = (type === 'movie' || type === 'all') ? getTMDBKey() : '';
  var sportsKey = getSportsDBKey();

  if (type === 'movie' && !tmdbKey) {
    return { error: 'Movie search requires a TMDB_API_KEY. Add it in Apps Script → Project Settings → Script Properties.' };
  }

  var q = encodeURIComponent(String(query));

  /* ── Round 1: all primary searches in parallel ─────────── */
  var r1Reqs = [], r1Tags = [];
  if (type === 'show' || type === 'all') {
    r1Reqs.push({ url: 'https://api.tvmaze.com/search/shows?q=' + q, muteHttpExceptions: true });
    r1Tags.push('tvmaze');
  }
  if ((type === 'movie' || type === 'all') && tmdbKey) {
    r1Reqs.push({ url: 'https://api.themoviedb.org/3/search/movie?api_key=' + tmdbKey + '&query=' + q + '&language=en-US&page=1', muteHttpExceptions: true });
    r1Tags.push('tmdb');
  }
  if (type === 'sport' || type === 'all') {
    r1Reqs.push({ url: 'https://www.thesportsdb.com/api/v1/json/' + sportsKey + '/searchteams.php?t=' + q, muteHttpExceptions: true });
    r1Tags.push('sports');
  }

  var r1Resps    = UrlFetchApp.fetchAll(r1Reqs);
  var tvShows    = [], tmdbMovies = [], sportsTeams = [];
  var errors     = [];
  var r2Reqs     = [], r2Tags = [];

  /* ── Parse primary responses, queue secondary requests ───── */
  r1Resps.forEach(function(resp, i) {
    var tag = r1Tags[i];
    if (resp.getResponseCode() !== 200) { errors.push(tag + ' HTTP ' + resp.getResponseCode()); return; }
    try {
      var body = JSON.parse(resp.getContentText());
      if (tag === 'tvmaze') {
        tvShows = (body || []).slice(0, 5).map(function(item) { return item.show || item; });
        tvShows.forEach(function(show) {
          if (!show.id) return;
          r2Reqs.push({ url: 'https://api.tvmaze.com/shows/' + show.id + '/episodes', muteHttpExceptions: true });
          r2Tags.push({ src: 'ep', id: show.id });
        });
      } else if (tag === 'tmdb') {
        tmdbMovies = (body.results || []).slice(0, 5);
        tmdbMovies.forEach(function(movie) {
          r2Reqs.push({ url: 'https://api.themoviedb.org/3/movie/' + movie.id + '/watch/providers?api_key=' + tmdbKey, muteHttpExceptions: true });
          r2Tags.push({ src: 'prov', id: movie.id });
        });
      } else if (tag === 'sports') {
        sportsTeams = (body.teams || []).slice(0, 5);
        sportsTeams.forEach(function(team) {
          r2Reqs.push({ url: 'https://www.thesportsdb.com/api/v1/json/' + sportsKey + '/eventsnext.php?id=' + team.idTeam, muteHttpExceptions: true });
          r2Tags.push({ src: 'ev', id: team.idTeam });
        });
      }
    } catch (e) { errors.push(tag + ': ' + e.message); }
  });

  /* ── Round 2: all secondary enrichment requests in parallel ─ */
  var r2Resps  = r2Reqs.length ? UrlFetchApp.fetchAll(r2Reqs) : [];
  var todayMs  = new Date().setHours(0, 0, 0, 0);
  var epMap    = {}, provMap = {}, evMap = {};

  r2Resps.forEach(function(resp, i) {
    var tag = r2Tags[i];
    if (resp.getResponseCode() !== 200) return;
    try {
      var body = JSON.parse(resp.getContentText());
      if (tag.src === 'ep') {
        epMap[tag.id] = (body || []).filter(function(ep) {
          return ep.airdate && new Date(ep.airdate).getTime() >= todayMs;
        }).slice(0, 5);
      } else if (tag.src === 'prov') {
        var us = body.results && body.results.US;
        if (us && us.flatrate && us.flatrate.length) provMap[tag.id] = us.flatrate[0].provider_name;
      } else if (tag.src === 'ev') {
        evMap[tag.id] = (body.events || []).slice(0, 15);
      }
    } catch (_) {}
  });

  /* ── Normalize all results ──────────────────────────────── */
  var results = [];

  tvShows.forEach(function(show) {
    results.push(normalizeTVMazeShow(show, epMap[show.id] || []));
  });

  tmdbMovies.forEach(function(item) {
    results.push({
      type: 'Movie',
      title: item.title || '',
      year:  (item.release_date || '').substring(0, 4),
      genre: tmdbGenreNames(item.genre_ids),
      description: (item.overview || '').substring(0, 220),
      streamingOn: provMap[item.id] || '',
      imdbScore:   item.vote_average ? String(Number(item.vote_average).toFixed(1)) : '',
      tmdbId:      item.id || ''
    });
  });

  sportsTeams.forEach(function(team) {
    var events = evMap[team.idTeam] || [];
    var games  = events.map(function(ev) {
      var opp      = ev.strHomeTeam === team.strTeam ? ev.strAwayTeam : ev.strHomeTeam;
      var gameDate = ev.dateEvent || '';
      var gameTime = ev.strTime   || '';
      /* TheSportsDB returns dateEvent and strTime in UTC.
         Combine them into a full ISO-8601 UTC timestamp and convert to
         America/Los_Angeles so stored values match the viewer's local time.
         A Mariners 7:10 PM PDT home game arrives as "02:10:00+00:00" UTC
         the next calendar day — without conversion it would show as 2 AM. */
      if (gameDate && gameTime) {
        try {
          var utcStr = gameDate + 'T' + gameTime;
          // Append Z if there is no timezone indicator at the end already
          if (!/Z$|[+-]\d{2}:\d{2}$/.test(utcStr)) utcStr += 'Z';
          var d = new Date(utcStr);
          if (!isNaN(d.getTime())) {
            gameDate = Utilities.formatDate(d, 'America/Los_Angeles', 'yyyy-MM-dd');
            gameTime = Utilities.formatDate(d, 'America/Los_Angeles', 'h:mm a z');
          }
        } catch (_) {}
      }
      return { date: gameDate, time: gameTime, opponent: opp, tv_channel: ev.strTVStation || '' };
    });
    var next = games.length
      ? (games[0].date + (games[0].time ? ' ' + games[0].time : '') + (games[0].opponent ? ' vs ' + games[0].opponent : '')).trim()
      : '';
    results.push({
      type:        'LiveTV',
      channel:     team.strTeam || '',
      league:      team.strLeague || '',
      description: (team.strDescriptionEN || '').replace(/<[^>]+>/g, '').substring(0, 220),
      genre:       'Sports',
      network:     team.strLeague || '',
      nextGame:    next,
      games:       games,
      sportsdbId:  team.idTeam || ''
    });
  });

  if (results.length === 0 && errors.length > 0) {
    return { error: errors.join('; ') };
  }

  logSearch(query, searchType || 'all', results.length);
  return { success: true, results: results };
}

function normalizeTVMazeShow(show, upcomingEps) {
  var network   = (show.network && show.network.name) ||
                  (show.webChannel && show.webChannel.name) || '';
  var genres    = Array.isArray(show.genres) ? show.genres.slice(0, 2).join(', ') : '';
  var summary   = (show.summary || '').replace(/<[^>]+>/g, '').trim().substring(0, 220);
  var status    = show.status === 'Running' ? 'Returning' : (show.status || '');
  var rating    = show.rating && show.rating.average ? String(show.rating.average) : '';
  var premiered = (show.premiered || '').substring(0, 4);
  var episodes  = (upcomingEps || []).map(function(ep) {
    return { season: ep.season, episode: ep.number, episode_title: ep.name || '', air_date: ep.airdate || '', airstamp: ep.airstamp || '', network: network };
  });
  return {
    type: 'Show',
    title: show.name || '',
    year: premiered,
    genre: genres,
    description: summary,
    network: network,
    streamingOn: network,
    status: status,
    nextAirs: episodes.length ? episodes[0].air_date : '',
    imdbScore: rating,
    episodes: episodes,
    tvmazeId: show.id || ''
  };
}

/* ── Log searches to SearchLogs sheet (best-effort) ─────── */
function logSearch(query, type, count) {
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('SearchLogs');
    if (!sheet) return;
    ensureColumns(sheet, ['timestamp', 'search_term', 'search_type', 'results_count']);
    appendByHeaders(sheet, {
      timestamp:     new Date().toISOString(),
      search_term:   query,
      search_type:   type,
      results_count: count
    });
  } catch (_) {}
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
    if (!e || !e.postData || !e.postData.contents) {
      return respondJson({ error: 'Empty request body' });
    }
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;

    if (action === 'search')            return respondJson(handleSearch(body.query, body.searchType));
    if (action === 'addRow')            return respondJson(handleAddRow(body.sheetName, body.rowData));
    if (action === 'updateRow')         return respondJson(handleUpdateRow(body.sheetName, body.rowIndex, body.rowData));
    if (action === 'deleteRow')         return respondJson(handleDeleteRow(body.sheetName, body.rowIndex));
    if (action === 'claudeSearch')      return respondJson(handleClaudeSearch(body.query, body.sheetName, body.clientDatetime));
    if (action === 'recommendForMe')    return respondJson(handleRecommendForMe(body));
    if (action === 'dislike')           return respondJson(handleDislike(body.title, body.type));
    if (action === 'removeDuplicates')  return respondJson(removeDuplicatesFromSheet(body.sheetName));
    if (action === 'saveEpisodes')      return respondJson(handleSaveEpisodes(body.title, body.episodes));
    if (action === 'saveGames')         return respondJson(handleSaveGames(body.channelId, body.games));

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
  var ss      = getSpreadsheet();
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
      var seenLiveC  = {};
      for (var i = 1; i < rawData.length; i++) {
        var row = buildObj(headers, rawData[i]);
        var key = String(row['title'] || '').toLowerCase().trim();
        if (!key) continue;            // skip blank rows
        var item = projectFields(row, CONTENT_FIELDS);
        item.rowIndex = i + 1;
        var ctRaw = String(row['content_type'] || '').trim();
        var fmt   = String(row['format']       || '').trim().toLowerCase();
        var ct    = ctRaw.toLowerCase();

        // Infer content_type when blank: check format column, then default to TV Show
        if (!ct) {
          ct = (fmt === 'movie' || fmt === 'film') ? 'movie' : 'tv show';
        }

        if (ct === 'movie' || ct === 'film') {
          if (seenMovies[key]) continue;
          seenMovies[key] = true;
          movies.push(item);
        } else if (ct === 'tv show' || ct === 'show' || ct === 'series' ||
                   ct === 'mini-series' || ct === 'limited series' || ct === 'miniseries') {
          if (seenShows[key]) continue;
          seenShows[key] = true;
          shows.push(item);
        } else if (ct === 'sports' || ct === 'sport' || ct === 'live event' ||
                   ct === 'live tv' || ct === 'live') {
          // Sports/Live-TV rows in Content_Master: map into liveTV list
          if (seenLiveC[key]) continue;
          seenLiveC[key] = true;
          liveTV.push({
            favorite_team_or_channel: String(row['title'] || '').trim(),
            live_tv_type:             ctRaw || fmt || 'Sports',
            league:                   String(row['genre_primary'] || '').trim(),
            description:              String(row['description']   || '').trim(),
            favorites:                String(row['favorites']     || '').trim(),
            rowIndex:                 i + 1,
          });
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

  /* Episode_Schedule — group rows by lowercased title and attach the
     upcoming/recent episodes to the matching show. The earliest upcoming
     episode also fills next_airs / latest_episode if those are blank. */
  var episodesByTitle = readScheduleRows(ss, EPISODE_SCHEDULE, EPISODE_FIELDS, 'title');
  attachSchedule(shows, 'title', episodesByTitle, function(item, list) {
    item.episodes = list;
    var upcoming = pickUpcoming(list, 'air_date');
    if (upcoming) {
      if (!item.next_airs) item.next_airs = upcoming.air_date || '';
      if (!item.latest_episode && (upcoming.season || upcoming.episode || upcoming.episode_title)) {
        var s = upcoming.season ? 'S' + String(upcoming.season).padStart(2, '0') : '';
        var e = upcoming.episode ? 'E' + String(upcoming.episode).padStart(2, '0') : '';
        var t = upcoming.episode_title ? ' ' + upcoming.episode_title : '';
        item.latest_episode = (s + e + t).trim();
      }
    }
  });

  /* Schedules — group rows by lowercased channel_id and attach upcoming
     games to the matching live-TV entry. The earliest upcoming game fills
     next_game / tv_channel if blank. */
  var gamesByChannel = readScheduleRows(ss, SCHEDULES_SHEET, SCHEDULE_FIELDS, 'channel_id');
  attachSchedule(liveTV, 'favorite_team_or_channel', gamesByChannel, function(item, list) {
    item.games = list;
    var upcoming = pickUpcoming(list, 'date');
    if (upcoming) {
      if (!item.next_game) {
        var when = upcoming.date || '';
        if (upcoming.time) when += (when ? ' ' : '') + upcoming.time;
        if (upcoming.opponent) when += ' vs ' + upcoming.opponent;
        item.next_game = when.trim();
      }
      if (!item.tv_channel && upcoming.tv_channel) item.tv_channel = upcoming.tv_channel;
    }
  });

  return { success: true, movies: movies, shows: shows, liveTV: liveTV };
}

/* Read a schedule-style sheet (Episode_Schedule / Schedules) and return
   { keyLower: [rows] }. Rows are projected to the supplied field list and
   keyed case-insensitively on the join column so a sheet with "Seahawks"
   matches a live-TV entry "seahawks". */
function readScheduleRows(ss, sheetName, fields, joinKey) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};
  ensureColumns(sheet, fields);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  var headers = normalizeHeaders(data[0]);
  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var row = buildObj(headers, data[i]);
    var key = String(row[joinKey] || '').toLowerCase().trim();
    if (!key) continue;
    if (!groups[key]) groups[key] = [];
    groups[key].push(projectFields(row, fields));
  }
  return groups;
}

/* For each item in the list, look up its grouped rows and run the writer
   callback to attach them. itemKey is the field on the item that matches
   the schedule sheet's join column. */
function attachSchedule(items, itemKey, groups, writer) {
  if (!items || !items.length) return;
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var key  = String(item[itemKey] || '').toLowerCase().trim();
    var list = groups[key] || [];
    writer(item, list);
  }
}

/* Pick the earliest row whose date column (parsed as YYYY-MM-DD) is today
   or later. Returns null when every row is in the past so callers never
   surface stale "last week's game" data on a card. */
function pickUpcoming(rows, dateField) {
  if (!rows || !rows.length) return null;
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var best = null, bestTime = Infinity;
  for (var i = 0; i < rows.length; i++) {
    var raw = String(rows[i][dateField] || '').trim();
    if (!raw) continue;
    var m = raw.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (!m) continue;
    var d = new Date(+m[1], +m[2] - 1, +m[3]);
    if (isNaN(d.getTime())) continue;
    if (d.getTime() < today.getTime()) continue;
    if (d.getTime() < bestTime) { best = rows[i]; bestTime = d.getTime(); }
  }
  return best;
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
  var settings = getSettings();
  if (!settingEnabled(settings, 'writes_enabled')) {
    return { error: 'Writes are currently disabled. Enable in the Settings sheet (writes_enabled = TRUE).' };
  }

  var ss = getSpreadsheet();

  if (isLiveTVSheet(sheetName)) {
    var sheet = ss.getSheetByName(LIVE_TV_SHEET);
    if (!sheet) return { error: LIVE_TV_SHEET + ' sheet not found' };
    ensureColumns(sheet, LIVE_TV_FIELDS);
    var liveRow  = mapToSheetRow(rowData, 'liveTV');
    var liveTitle = (liveRow.favorite_team_or_channel || '').toLowerCase().trim();
    if (liveTitle && hasDuplicate(sheet, 'favorite_team_or_channel', liveTitle, null, liveRow.profile)) {
      return { success: true, duplicate: true };
    }
    var rowIndex = appendByHeaders(sheet, liveRow);

  } else {
    /* Movies and Shows both go into Content_Master */
    var sheet = ss.getSheetByName(CONTENT_MASTER);
    if (!sheet) return { error: CONTENT_MASTER + ' sheet not found' };
    ensureColumns(sheet, CONTENT_FIELDS);

    var kind       = isShowsSheet(sheetName) ? 'TV Show' : 'Movie';
    var contentRow = mapToSheetRow(rowData, kind);
    var titleVal   = (contentRow.title || '').toLowerCase().trim();
    if (titleVal && hasDuplicate(sheet, 'title', titleVal, kind, contentRow.profile)) {
      return { success: true, duplicate: true };
    }
    var rowIndex = appendByHeaders(sheet, contentRow);
  }

  invalidateCache();
  return { success: true, rowIndex: rowIndex };
}

/* Returns true if the sheet already has a row whose titleHeader column
   matches newTitle (case-insensitive). When contentType is provided, also
   requires the content_type column to match — so a Movie and a TV Show with
   the same title are not considered duplicates of each other.
   When profile is provided, only rows with the same profile value are
   considered duplicates — items from different profiles can share a title. */
function hasDuplicate(sheet, titleHeader, newTitle, contentType, profile) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  var headers    = normalizeHeaders(data[0]);
  var titleIdx   = headers.indexOf(titleHeader);
  if (titleIdx === -1) return false;
  var ctIdx      = contentType ? headers.indexOf('content_type') : -1;
  var profileIdx = headers.indexOf('profile');
  var newProfileLc = (profile || '').toLowerCase().trim();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][titleIdx]).toLowerCase().trim() !== newTitle) continue;
    if (ctIdx !== -1 && String(data[i][ctIdx]).trim().toLowerCase() !== contentType.toLowerCase()) continue;
    /* Different profiles → not a duplicate; each profile owns its own library */
    var existingProfile = profileIdx !== -1 ? String(data[i][profileIdx] || '').toLowerCase().trim() : '';
    if (existingProfile !== newProfileLc) continue;
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
  var ss = getSpreadsheet();
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
      tv_channel:    firstOf(data, ['tvChannel', 'tv_channel']),
      favorites:     firstOf(data, ['favorites', 'favorite']),
      profile:       firstOf(data, ['profile']),
      watch_status:  firstOf(data, ['watch_status', 'watchStatus'])
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
    next_airs:      firstOf(data, ['nextAirs', 'next_airs', 'nextAiring', 'airing', 'whenitairs']),
    favorites:      firstOf(data, ['favorites', 'favorite']),
    profile:        firstOf(data, ['profile']),
    watch_status:   firstOf(data, ['watch_status', 'watchStatus'])
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
  var rowNum = parseInt(rowIndex, 10);
  if (isNaN(rowNum) || rowNum < 2) return { error: 'Invalid rowIndex: ' + rowIndex };

  var ss    = getSpreadsheet();
  var sheet = isLiveTVSheet(sheetName)
    ? ss.getSheetByName(LIVE_TV_SHEET)
    : ss.getSheetByName(CONTENT_MASTER);

  if (!sheet) return { error: 'Sheet not found for: ' + sheetName };
  if (rowNum > sheet.getLastRow()) return { error: 'Row out of range: ' + rowIndex };

  /* Ensure all expected columns exist so next_airs / next_game are never
     silently dropped when the sheet is missing those headers. */
  ensureColumns(sheet, isLiveTVSheet(sheetName) ? LIVE_TV_FIELDS : CONTENT_FIELDS);

  var lastCol  = sheet.getLastColumn();
  var headers  = normalizeHeaders(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  var existing = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];

  /* Normalize Claude-style keys (streamingOn → streaming_on, etc.)
     so a refresh always persists the freshest values from Claude. */
  var kind       = inferContentKind(sheetName);
  var mappedKind = kind === 'liveTV' ? 'liveTV' : (kind === 'TV Show' ? 'TV Show' : 'Movie');
  var normalized = mapToSheetRow(rowData, mappedKind);

  var row = headers.map(function(h, i) {
    /* Prefer normalized value (from Claude's camelCase keys mapped to sheet headers).
       Use the existing cell only when Claude didn't return this field at all. */
    var v = normalized[h];
    if (v !== undefined && v !== '') return v;
    /* Fall back to raw rowData key (direct match) */
    v = rowData[h];
    if (v !== undefined && v !== '') return v;
    /* Keep existing cell — Claude didn't touch this field */
    return existing[i] !== undefined ? existing[i] : '';
  });

  sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
  invalidateCache();
  return { success: true };
}

/* ── Delete row ──────────────────────────────────────────── */
function handleDeleteRow(sheetName, rowIndex) {
  var ss    = getSpreadsheet();
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

/* ── Save episode rows to Episode_Schedule ───────────────── */
/* Replaces all rows for the given show title, then appends the new ones.
   Title is matched case-insensitively. Returns { success, written }. */
function handleSaveEpisodes(title, episodes) {
  if (!title) return { error: 'Missing title' };
  if (!Array.isArray(episodes)) return { error: 'episodes must be an array' };
  return writeScheduleRows(EPISODE_SCHEDULE, EPISODE_FIELDS, 'title', title, episodes);
}

/* ── Save game rows to Schedules ─────────────────────────── */
/* Replaces all rows for the given channelId, then appends the new ones.
   channelId is matched against the channel_id column case-insensitively
   and stamped onto every appended row. Returns { success, written }. */
function handleSaveGames(channelId, games) {
  if (!channelId) return { error: 'Missing channelId' };
  if (!Array.isArray(games)) return { error: 'games must be an array' };
  return writeScheduleRows(SCHEDULES_SHEET, SCHEDULE_FIELDS, 'channel_id', channelId, games);
}

function writeScheduleRows(sheetName, fields, joinKey, joinValue, rows) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(fields);
  } else {
    ensureColumns(sheet, fields);
  }

  var data    = sheet.getDataRange().getValues();
  var headers = data.length > 0 ? normalizeHeaders(data[0]) : fields.slice();
  var keyIdx  = headers.indexOf(joinKey);
  var matchLc = String(joinValue).toLowerCase().trim();

  /* Delete existing rows for the current join value AND prune any rows from
     other channels/titles that are more than 30 days in the past (Schedules
     sheet only). Single bottom-up pass keeps row indices valid throughout. */
  var cutoffMs = Date.now() - 30 * 24 * 60 * 60 * 1000;
  var dateColIdx = headers.indexOf('date');
  if (keyIdx !== -1) {
    for (var i = data.length - 1; i >= 1; i--) {
      var rowKey = String(data[i][keyIdx]).toLowerCase().trim();
      if (rowKey === matchLc) {
        sheet.deleteRow(i + 1);
        continue;
      }
      /* Only auto-prune stale rows on the Schedules sheet (not episode lists) */
      if (sheetName === SCHEDULES_SHEET && dateColIdx !== -1) {
        var rawDate = String(data[i][dateColIdx] || '').trim();
        var dm = rawDate.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (dm) {
          var rowDate = new Date(+dm[1], +dm[2] - 1, +dm[3]);
          if (!isNaN(rowDate.getTime()) && rowDate.getTime() < cutoffMs) {
            sheet.deleteRow(i + 1);
          }
        }
      }
    }
  }

  /* Append the new rows. The join column is force-stamped so callers can
     omit it (they're already telling us the join value). */
  var written = 0;
  for (var r = 0; r < rows.length; r++) {
    var src = rows[r] || {};
    var obj = {};
    fields.forEach(function(f) { obj[f] = src[f] !== undefined ? src[f] : ''; });
    obj[joinKey] = joinValue;
    appendByHeaders(sheet, obj);
    written++;
  }

  invalidateCache();
  return { success: true, written: written };
}

/* ── Claude search (with web_search tool) ────────────────── */
function handleClaudeSearch(query, sheetName, clientDatetime) {
  var settings = getSettings();
  if (!settingEnabled(settings, 'search_enabled')) {
    return { error: 'Search is currently disabled. Enable it in the Settings sheet (search_enabled = TRUE).' };
  }

  var apiKey = getAnthropicKey();
  if (!apiKey) {
    return { error: 'Missing ANTHROPIC_API_KEY — set it in Apps Script → Project Settings → Script Properties' };
  }

  var localDate, localTime;
  if (clientDatetime) {
    var dtMatch = String(clientDatetime).match(/^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);
    if (dtMatch) { localDate = dtMatch[1]; localTime = dtMatch[2]; }
  }
  if (!localDate) {
    var today = new Date();
    localDate = today.getFullYear() + '-' +
      String(today.getMonth() + 1).padStart(2, '0') + '-' +
      String(today.getDate()).padStart(2, '0');
    localTime = '';
  }

  var nowContext = localTime
    ? 'Today\'s date is ' + localDate + ' and the user\'s current local time is ' + localTime + '.'
    : 'Today\'s date is ' + localDate + '.';

  var sportsLiveNote = localTime
    ? 'CRITICAL FOR SPORTS/LIVE TV: Use ' + localDate + ' as today\'s date when searching schedules. ' +
      'If a game or event for the searched team is currently in progress right now (its scheduled start time is at or before ' + localTime + ' and a typical game duration means it has not yet ended), ' +
      'you MUST include it first in the games array at its actual scheduled start time — do NOT skip or omit currently live games. ' +
      'Always verify today\'s full schedule before listing future dates.\n\n'
    : '';

  var prompt =
    'You are a media database assistant. ' + nowContext + ' The user searched for: "' + query + '"\n\n' +
    sportsLiveNote +
    'Use the web_search tool (up to 8 times) to look up current, accurate information from ' +
    'credible sources (IMDb, Rotten Tomatoes, Wikipedia, official network and streaming-service pages, ' +
    'TV Guide, Sports Reference). Search specifically for the next air date / next game if applicable.\n\n' +
    'Return ONLY a single raw JSON object — no markdown fences, no explanation, no extra text.\n\n' +
    'For a Movie use exactly these keys:\n' +
    '{"type":"Movie","title":"","year":"<4-digit year>","genre":"<primary genre>","rating":"<MPAA rating e.g. PG-13>","description":"<1-2 sentence plot summary>","director":"","cast":"<comma-separated top 3 actors>","streamingOn":"<platform name>","imdbScore":"<e.g. 8.2>","tone":"<e.g. Action, Comedy, Drama, Thriller>"}\n\n' +
    'For a TV Show use exactly these keys (and include an "episodes" array of the next 5 upcoming or most recent episodes when known):\n' +
    '{"type":"Show","title":"","year":"<year show started>","genre":"<primary genre>","rating":"<TV rating e.g. TV-MA>","description":"<1-2 sentence premise>","network":"<broadcast network or streaming service>","seasons":"<number>","latestEpisode":"<S##E## Title if known>","status":"<Returning | Ended | Cancelled | On Hiatus>","nextAirs":"<YYYY-MM-DD HH:MM TZ or descriptive e.g. \'Tuesdays 9PM ET on NBC\'>","cast":"<comma-separated top 3 actors>","streamingOn":"<streaming platform if different from network>","imdbScore":"<e.g. 8.2>","tone":"<e.g. Drama, Comedy, Thriller>","episodes":[{"season":"5","episode":"3","episode_title":"","air_date":"YYYY-MM-DD","network":""}]}\n\n' +
    'For Live TV / Sports channel use exactly these keys (and include a "games" array of the next 15 games when known):\n' +
    '{"type":"LiveTV","channel":"<channel or team name>","network":"<broadcast network>","league":"<e.g. NFL, NBA, EPL>","genre":"<Sports | News | Entertainment>","description":"<brief description>","streamingOn":"<streaming platform>","nextGame":"<YYYY-MM-DD HH:MM PT or descriptive>","tvChannel":"<cable/satellite channel name>","games":[{"date":"YYYY-MM-DD","time":"7:10 PM PDT","opponent":"","tv_channel":""}]}\n\n' +
    'Rules: real data only; leave a field empty string if truly unknown; dates MUST be in YYYY-MM-DD format when an exact date is known. Game times MUST be in Pacific Time (e.g. "7:10 PM PDT" or "1:05 PM PST"). Return at most 15 games and 5 episodes. Omit the array (or return []) if you cannot find scheduled dates.';

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 2048,
    tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 8 }],
    messages: [{ role: 'user', content: prompt }]
  };

  /* Call Anthropic with one automatic retry on transient 5xx / network errors. */
  var response, code, body;
  var fetchOptions = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  for (var attempt = 0; attempt < 2; attempt++) {
    if (attempt > 0) Utilities.sleep(2000);
    try {
      response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', fetchOptions);
    } catch (netErr) {
      if (attempt === 0) continue;
      return { error: 'Network error contacting Anthropic: ' + netErr.message };
    }
    code = response.getResponseCode();
    body = response.getContentText();
    /* Retry on 5xx server errors or 529 (overloaded) only */
    if (code >= 500 || code === 529) {
      if (attempt === 0) continue;
    }
    break;
  }

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

  /* Persist episodes / games arrays so the schedule sheets stay in sync
     with whatever Claude just returned. Failures here must not break the
     search response — the card data is still useful on its own. */
  try {
    var resultType = String(mediaData.type || '').toLowerCase();
    if (Array.isArray(mediaData.episodes) && mediaData.episodes.length &&
        (resultType.indexOf('show') !== -1 || resultType.indexOf('series') !== -1)) {
      var showTitle = mediaData.title || '';
      if (showTitle) handleSaveEpisodes(showTitle, mediaData.episodes);
    }
    if (Array.isArray(mediaData.games) && mediaData.games.length &&
        (resultType.indexOf('live') !== -1 || resultType.indexOf('sport') !== -1 ||
         resultType.indexOf('team') !== -1 || resultType.indexOf('channel') !== -1)) {
      var channelId = mediaData.channel || mediaData.title || '';
      if (channelId) {
        /* Stamp team/league onto each row so the Schedules sheet is
           self-describing — useful when a single sheet has rows from many teams. */
        var stamped = mediaData.games.map(function(g) {
          var copy = shallowCopy(g || {});
          if (!copy.team)   copy.team   = channelId;
          if (!copy.league) copy.league = mediaData.league || '';
          return copy;
        });
        handleSaveGames(channelId, stamped);
      }
    }
  } catch (scheduleErr) { /* surface nothing; schedules are best-effort */ }

  return { success: true, data: mediaData };
}

/* ── Personalized recommendations ─────────────────────────────
   Reads the user's saved Movies + Shows, asks Claude to follow the
   four-phase recommendation framework (build a taste profile, search the
   web in three parallel passes, score, and return 2 movies + 2 shows),
   then returns the parsed result for the front-end to render.

   Response shape:
     { profile: "<2-3 sentence taste fingerprint>",
       results: [
         { type, title, year, genre, description, streamingOn, imdbScore,
           tone, network?, status?, seasons?, nextAirs?,
           whyItFits, confidence },
         ...4 total
       ],
       note?: "<optional explanation, e.g. library too small>" }
*/
/* ── Dislike a recommendation ────────────────────────────── */
/* Appends a row to the "Disliked" sheet (created automatically on first use).
   These titles are read back by handleRecommendForMe() and injected into the
   prompt + post-filter so Claude never surfaces them again. */
function handleDislike(title, type) {
  if (!title || !String(title).trim()) return { error: 'Missing title' };
  try {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName('Disliked');
    if (!sheet) {
      sheet = ss.insertSheet('Disliked');
      sheet.getRange(1, 1, 1, 3).setValues([['title', 'type', 'disliked_at']]);
    }
    sheet.appendRow([String(title).trim(), String(type || '').trim(), new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

/* ── Read all disliked titles (for recommendation exclusion) ─ */
function readDislikedTitles() {
  try {
    var sheet = getSpreadsheet().getSheetByName('Disliked');
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .map(function(row) { return String(row[0] || '').toLowerCase().trim(); })
      .filter(Boolean);
  } catch (_) { return []; }
}

function handleRecommendForMe(body) {
  var settings = getSettings();
  if (!settingEnabled(settings, 'search_enabled')) {
    return { error: 'Search is currently disabled. Enable it in the Settings sheet (search_enabled = TRUE).' };
  }

  var apiKey = getAnthropicKey();
  if (!apiKey) {
    return { error: 'Missing ANTHROPIC_API_KEY — set it in Apps Script → Project Settings → Script Properties' };
  }

  /* Use the profile-filtered library sent by the client when available.
     This ensures recommendations are scoped to the active profile's content. */
  var movies, shows, livetv;
  var activeProfile = (body && body.profile) || '';
  var providedLib   = body && body.library;

  if (providedLib && Array.isArray(providedLib.movies)) {
    movies = providedLib.movies;
    shows  = providedLib.shows  || [];
    livetv = providedLib.livetv || [];
  } else {
    var media = readAllMedia();
    movies = (media && media.movies) || [];
    shows  = (media && media.shows)  || [];
    livetv = (media && media.liveTV) || [];
  }

  /* Backend-side profile filter (safety net when frontend filter may not have run) */
  function isPopNana(item) { return (item && (item.profile || '')).toLowerCase() === 'popnana'; }
  function isSeattleTeam(item) {
    var name = ((item.favorite_team_or_channel || item.title || item.team || '')).toLowerCase();
    return name.indexOf('seahawk') !== -1 || name.indexOf('mariner') !== -1;
  }
  if (activeProfile === 'popnana') {
    movies = movies.filter(isPopNana);
    shows  = shows.filter(isPopNana);
    livetv = livetv.filter(function(l) { return isPopNana(l) || isSeattleTeam(l); });
  } else {
    movies = movies.filter(function(m) { return !isPopNana(m); });
    shows  = shows.filter(function(s)  { return !isPopNana(s); });
    livetv = livetv.filter(function(l) { return !isPopNana(l); });
  }

  /* Build a digestible sports/live-TV list for the taste profile */
  var liveList = livetv.map(function(l) {
    var name   = String(l.favorite_team_or_channel || l.title || l.team || '').trim();
    if (!name) return '';
    var league = String(l.league || l.live_tv_type || '').trim();
    return league ? name + ' (' + league + ')' : name;
  }).filter(Boolean).slice(0, 20);

  if (movies.length + shows.length + liveList.length < 2) {
    return {
      profile: '',
      results: [],
      note: 'Add at least a couple of titles to your library so we can build a taste profile.'
    };
  }

  /* Compact each library entry to the fields most useful for taste
     analysis. Keeps the prompt small enough to fit comfortably. */
  function digest(item) {
    var n = {};
    Object.keys(item || {}).forEach(function(k) { n[k.toLowerCase().replace(/[_\s]/g, '')] = item[k]; });
    function p() { for (var i = 0; i < arguments.length; i++) { var v = n[arguments[i].toLowerCase().replace(/[_\s]/g, '')]; if (v != null && v !== '') return String(v); } return ''; }
    var parts = [];
    var title = p('title');
    if (!title) return '';
    parts.push(title);
    var year  = p('yearstarted', 'year');
    if (year)  parts.push('(' + year + ')');
    var genre = p('genreprimary', 'genre');
    if (genre) parts.push('— ' + genre);
    var tone  = p('tone');
    if (tone)  parts.push('[tone: ' + tone + ']');
    var rating = p('agerating', 'rating');
    if (rating) parts.push('[' + rating + ']');
    var imdb = p('imdbscore', 'imdb');
    if (imdb) parts.push('[imdb: ' + imdb + ']');
    var fav = p('favorites');
    if (fav && (fav.toLowerCase() === 'yes' || fav === '1' || fav.toLowerCase() === 'true')) parts.push('[FAVORITE]');
    return parts.join(' ');
  }

  var movieList = movies.map(digest).filter(Boolean).slice(0, 80);
  var showList  = shows.map(digest).filter(Boolean).slice(0, 80);

  /* Excluded titles — library content + anything the user explicitly disliked. */
  var dislikedTitles = readDislikedTitles();
  var excludedTitles = []
    .concat(movies.map(function(m) { return String(m.title || '').toLowerCase().trim(); }))
    .concat(shows.map(function(s)  { return String(s.title || '').toLowerCase().trim(); }))
    .concat(dislikedTitles)
    .filter(Boolean);

  var today = new Date();
  var todayStr = today.getFullYear() + '-' +
    String(today.getMonth() + 1).padStart(2, '0') + '-' +
    String(today.getDate()).padStart(2, '0');

  var prompt =
    'You are a personalized entertainment recommendation agent. Today\'s date is ' + todayStr + '.\n\n' +
    'The user\'s library is below — these are titles they have already watched and enjoyed. ' +
    'Do NOT recommend anything from this list.\n\n' +
    'MOVIES THE USER HAS SAVED (' + movieList.length + '):\n' +
    (movieList.length ? '- ' + movieList.join('\n- ') : '(none)') + '\n\n' +
    'TV SHOWS THE USER HAS SAVED (' + showList.length + '):\n' +
    (showList.length ? '- ' + showList.join('\n- ') : '(none)') + '\n\n' +
    'SPORTS & LIVE TV THE USER FOLLOWS (' + liveList.length + '):\n' +
    (liveList.length ? '- ' + liveList.join('\n- ') : '(none)') + '\n\n' +
    'TITLES THE USER HAS EXPLICITLY DISLIKED — do NOT recommend these under any circumstances:\n' +
    (dislikedTitles.length ? '- ' + dislikedTitles.join('\n- ') : '(none)') + '\n\n' +
    'Follow this four-phase framework strictly:\n\n' +
    'PHASE 1 — BUILD THE TASTE PROFILE BEFORE SEARCHING\n' +
    'Synthesize the lists above (do not search yet):\n' +
    '  1. GENRES: which genres dominate?\n' +
    '  2. TONES: which tones recur (slow-burn, satirical, emotionally heavy, witty, etc.)?\n' +
    '  3. THEMES: which subject matter repeats (moral ambiguity, class tension, found family, unreliable narrator, crime procedural, etc.)?\n' +
    '  4. ERA / FORMAT: prestige TV, indie film, blockbusters, foreign language, classics?\n' +
    '  5. SPORTS & LIVE INTERESTS: factor in any teams or leagues from the SPORTS & LIVE TV list — these reveal regional loyalties and live-event preferences that should inform recommendations (sports documentaries, team-related content, broadcast live events).\n' +
    '  6. AVOID SIGNALS: which genres/tones are completely absent? Treat these as soft avoids.\n' +
    '  7. TASTE FINGERPRINT: write a 2-3 sentence summary you will score every candidate against.\n\n' +
    'PHASE 2 — RUN THREE PARALLEL SEARCH PASSES (use the web_search tool, free public sources only)\n' +
    '  PASS A — Streaming catalog: target JustWatch, Letterboxd, IMDb lists, Rotten Tomatoes. ' +
    'Query like "[theme/genre from fingerprint] best movies/shows ' + today.getFullYear() + '" and ' +
    '"hidden gem [genre] streaming". Goal: 10–15 candidates currently streamable.\n' +
    '  PASS B — Social discovery: target Reddit (r/moviesuggestions, r/ifyoulikeblank, r/television), ' +
    'Letterboxd lists, fan wikis. Query like "if you liked [top 2-3 titles from their list] recommendations" ' +
    'and "fans of [title] also loved". Goal: 5–10 community-endorsed thematic adjacents.\n' +
    '  PASS C — Live & broadcast: target TV Guide, Reelgood, network sites (NBC, HBO, PBS), sports schedules. ' +
    'Query upcoming TV premieres and live events for ' + today.getFullYear() + '. Goal: 3–5 time-sensitive picks.\n\n' +
    'PHASE 3 — SCORE AND RANK\n' +
    'For every candidate, score against the taste fingerprint:\n' +
    '  THEME MATCH (0–3) · TONE MATCH (0–2) · AVOID PENALTY (–2 if it leans on a genre/tone they\'ve clearly avoided) · ' +
    'NOVELTY BONUS (+1 if underrepresented in their list) · RECENCY (+1 if released/airing in the last 18 months) · ' +
    'LIVE URGENCY (mark Pass C titles airing within 7 days as URGENT).\n' +
    'Discard anything below 3 points. Rank by score.\n\n' +
    'PHASE 4 — OUTPUT\n' +
    'Return EXACTLY 3 movies and 3 TV shows (6 total) — the highest-scoring picks across passes A and B. ' +
    'Skip Pass C unless one of the six picks is naturally a live/limited-series premiere.\n\n' +
    'ACCURACY RULES\n' +
    '- Never fabricate availability. If you cannot confirm where a title streams, set streamingOn to "Check JustWatch".\n' +
    '- Every "whyItFits" must reference 1–2 specific titles from the user\'s own list as comparison anchors.\n' +
    '- Titles marked [FAVORITE] in the library are loved most — weight them 2× when building the taste fingerprint and use them as primary anchors in whyItFits.\n' +
    '- Do not recommend any title already in the user\'s library (case-insensitive match on title).\n' +
    '- Prefer specificity over volume.\n\n' +
    'Return ONLY a single raw JSON object — no markdown fences, no explanation, no extra text — with this exact shape:\n' +
    '{\n' +
    '  "profile": "<2-3 sentence taste fingerprint>",\n' +
    '  "results": [\n' +
    '    {\n' +
    '      "type": "Movie",\n' +
    '      "title": "",\n' +
    '      "year": "<4-digit year>",\n' +
    '      "genre": "<primary genre>",\n' +
    '      "rating": "<MPAA rating>",\n' +
    '      "description": "<1-2 sentence plot summary>",\n' +
    '      "director": "",\n' +
    '      "cast": "<comma-separated top 3>",\n' +
    '      "streamingOn": "<platform or \'Check JustWatch\'>",\n' +
    '      "imdbScore": "<e.g. 8.2>",\n' +
    '      "tone": "<e.g. Drama, Thriller>",\n' +
    '      "whyItFits": "<2-3 sentences citing specific titles from their list>",\n' +
    '      "confidence": "High | Medium | Worth a shot"\n' +
    '    },\n' +
    '    { "type": "Movie", ... },\n' +
    '    {\n' +
    '      "type": "Show",\n' +
    '      "title": "",\n' +
    '      "year": "<year show started>",\n' +
    '      "genre": "<primary genre>",\n' +
    '      "rating": "<TV rating>",\n' +
    '      "description": "<1-2 sentence premise>",\n' +
    '      "network": "<broadcast network or streaming service>",\n' +
    '      "seasons": "<number>",\n' +
    '      "status": "<Returning | Ended | Cancelled | On Hiatus>",\n' +
    '      "nextAirs": "<YYYY-MM-DD HH:MM TZ or descriptive, empty string if unknown>",\n' +
    '      "cast": "<comma-separated top 3>",\n' +
    '      "streamingOn": "<streaming platform or \'Check JustWatch\'>",\n' +
    '      "imdbScore": "<e.g. 8.2>",\n' +
    '      "tone": "<e.g. Drama, Thriller>",\n' +
    '      "whyItFits": "<2-3 sentences citing specific titles from their list>",\n' +
    '      "confidence": "High | Medium | Worth a shot"\n' +
    '    },\n' +
    '    { "type": "Show", ... },\n' +
    '    { "type": "Show", ... }\n' +
    '  ]\n' +
    '}\n\n' +
    'Order: 3 movies first, then 3 shows. Use empty string for any field you cannot confirm. ' +
    'The "results" array MUST contain exactly 6 entries — 3 with type "Movie" and 3 with type "Show".';

  var payload = {
    model:      ANTHROPIC_MODEL,
    max_tokens: 16000,
    tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 12 }],
    messages: [{ role: 'user', content: prompt }]
  };

  var fetchOptions = {
    method:      'post',
    contentType: 'application/json',
    headers: {
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response, code, body;
  for (var attempt = 0; attempt < 2; attempt++) {
    if (attempt > 0) Utilities.sleep(2000);
    try {
      response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', fetchOptions);
    } catch (netErr) {
      if (attempt === 0) continue;
      return { error: 'Network error contacting Anthropic: ' + netErr.message };
    }
    code = response.getResponseCode();
    body = response.getContentText();
    if (code >= 500 || code === 529) {
      if (attempt === 0) continue;
    }
    break;
  }

  if (code < 200 || code >= 300) {
    var apiErr;
    try { apiErr = JSON.parse(body); } catch (_) {}
    var msg = (apiErr && apiErr.error && apiErr.error.message) || body;
    return { error: 'Anthropic API ' + code + ': ' + msg };
  }

  var apiResult;
  try { apiResult = JSON.parse(body); }
  catch (e) { return { error: 'Bad API response: ' + body.substring(0, 200) }; }

  if (apiResult.stop_reason === 'max_tokens') {
    return { error: 'Recommendation response was too long to process — please try again.' };
  }

  var text = extractTextFromContent(apiResult.content);
  if (!text) return { error: 'Empty response from Claude' };

  var parsed = parseJsonFromText(text);
  if (!parsed || !Array.isArray(parsed.results)) {
    return { error: 'Could not parse recommendations from response', raw: text.substring(0, 400) };
  }

  /* Filter out anything Claude accidentally returned that the user already
     owns. Case-insensitive title match. */
  var excludeSet = {};
  excludedTitles.forEach(function(t) { excludeSet[t] = true; });
  var filtered = parsed.results.filter(function(r) {
    var t = String((r && r.title) || '').toLowerCase().trim();
    return t && !excludeSet[t];
  });

  return {
    profile: String(parsed.profile || ''),
    results: filtered,
    note:    filtered.length ? '' : 'Couldn\'t find fresh picks that aren\'t already in your library — try again later.'
  };
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

/* Extract the outermost JSON object from a possibly-fenced text blob.
   Uses balanced-bracket counting so trailing text after the closing }
   (e.g. a model explanation) does not corrupt the extracted substring. */
function parseJsonFromText(text) {
  if (!text) return null;
  var stripped = String(text).replace(/```json\s*/gi, '').replace(/```/g, '').trim();

  /* Fast path: the whole string is already valid JSON */
  try { var direct = JSON.parse(stripped); if (direct && typeof direct === 'object') return direct; } catch (_) {}

  /* Balanced-bracket scan to find the first complete {...} object */
  var start = stripped.indexOf('{');
  if (start === -1) return null;
  var depth = 0, inStr = false, esc = false;
  for (var i = start; i < stripped.length; i++) {
    var c = stripped[i];
    if (esc)            { esc = false; continue; }
    if (c === '\\' && inStr) { esc = true;  continue; }
    if (c === '"')      { inStr = !inStr; continue; }
    if (inStr)          { continue; }
    if (c === '{')      { depth++; }
    else if (c === '}') {
      depth--;
      if (depth === 0) {
        try { return JSON.parse(stripped.substring(start, i + 1)); } catch (_) { return null; }
      }
    }
  }
  return null;
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
  return sheet.getLastRow(); // row index of the newly appended row
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
