var LISTINGS_SHEET_NAME = "listings";
var INTERESTS_SHEET_NAME = "interests";
var LISTINGS_PROJECTION_SHEET_NAME = "listings_projection";
var LISTING_INDEX_SHEET_NAME = "listing_index";
var INTEREST_INDEX_SHEET_NAME = "interest_index";
var RUNTIME_META_SHEET_NAME = "runtime_meta";
var LEGACY_BOOKS_SHEET_NAME = "books";

var LISTINGS_HEADERS = [
  "created_at",
  "listing_id",
  "item_type",
  "category",
  "title",
  "description",
  "image_urls_json",
  "jan",
  "subject_tags_json",
  "author",
  "publisher",
  "published_date",
  "seller_name",
  "seller_id",
  "status",
];

var INTERESTS_HEADERS = ["created_at", "interest_id", "listing_id", "viewer_id", "viewer_name", "is_active"];
var INTERESTS_HEADERS_V1 = ["created_at", "interest_id", "listing_id", "viewer_id", "viewer_name"];

var LISTINGS_PROJECTION_HEADERS = [
  "projection_row_id",
  "created_at",
  "listing_id",
  "item_type",
  "category",
  "title",
  "thumb_url",
  "detail_image_url",
  "jan",
  "subject_tags_json",
  "author",
  "publisher",
  "seller_name",
  "seller_id",
  "status",
  "is_active",
  "wanted_count",
  "wanted_viewer_ids_json",
  "updated_at",
];

var LISTING_INDEX_HEADERS = ["listing_id", "listings_row", "projection_row"];
var INTEREST_INDEX_HEADERS = ["interest_key", "interests_row", "is_active"];
var RUNTIME_META_HEADERS = ["key", "value"];

var LEGACY_BOOK_HEADERS = [
  "created_at",
  "book_id",
  "jan",
  "title",
  "summary",
  "cover_url",
  "author",
  "publisher",
  "published_date",
  "seller_name",
  "seller_id",
  "status",
];

var LEGACY_INTEREST_HEADERS = ["created_at", "interest_id", "book_id", "jan", "viewer_id", "viewer_name"];

var DEFAULT_STATUS = "AVAILABLE";
var ACTIVE_FLAG_ON = "1";
var ACTIVE_FLAG_OFF = "0";
var UPLOAD_FOLDER_ID = "";
var MARKET_SNAPSHOT_CACHE_KEY = "market_snapshot_v2";
var MARKET_SNAPSHOT_TTL_SEC = 60;
var SCHEMA_CACHE_KEY = "market_schema_ok_v3";
var SCHEMA_CACHE_TTL_SEC = 3600;
var MARKET_COMPUTE_VERSION_PROPERTY = "market_compute_version_v1";
var MARKET_COMPUTE_CACHE_TTL_SEC = 60;
var LISTINGS_SUMMARY_SNAPSHOT_CACHE_KEY = "listings_summary_snapshot_v1";
var LISTINGS_SUMMARY_SNAPSHOT_CACHE_TTL_SEC = 120;
var LISTINGS_V3_CACHE_KEY = "listings_v3";
var LISTINGS_V3_CHUNK_MULTIPLIER = 3;
var LISTINGS_V3_USE_FOR_V2 = true;
var MARKET_LOCK_WAIT_MS = 20000;
var PROJECTION_VERSION_KEY = "projection_version";
var PROJECTION_LAST_REBUILD_KEY = "last_rebuild_at";
var PROJECTION_DIRTY_KEY = "dirty_flag";

var RESIDENTS_BY_ROOM = {
  "105": "坪田　晴琉",
  "106": "杉本　博斗",
  "108": "前林　諒叡",
  "110": "高廣　翼",
  "204": "若林　永大",
  "205": "竹腰　啓人",
  "208": "氷見　要",
  "209": "佐藤　秀哉",
  "210": "岡上　拓海",
  "211": "松田　優真",
  "212": "和泉　侑汰",
  "215": "坂野　友哉",
  "218": "竹腰　陽樹",
  "219": "山原　準世",
  "221": "樋口　康介",
  "223": "菊　悠太",
  "224": "黒澤　望",
  "227": "東海　立",
  "228": "橋本　武蔵",
  "229": "山口　順生",
  "303": "大島　一輝",
  "304": "小林　佳太郎",
  "305": "清水　陽太",
  "306": "伊藤　哲也",
  "310": "石田　礼門",
  "311": "阿部　稔暁",
  "312": "平　明真",
  "313": "平　啓吾",
  "316": "森田　旺輔",
  "317": "安達　寛人",
  "318": "金川　明仁",
  "322": "平岡　桐弥",
  "324": "松井　俊介",
  "329": "江尻　凌太朗",
};

function doGet(e) {
  return handleRequest_("GET", e);
}

function doPost(e) {
  return handleRequest_("POST", e);
}

function handleRequest_(method, e) {
  try {
    ensureSchema_();

    var req = parseRequest_(method, e);
    var action = req.action;

    if (!action) {
      throw new Error("action is required");
    }

    var result;
    switch (action) {
      case "initSheets":
        result = {
          ok: true,
          listingsHeaders: LISTINGS_HEADERS,
          interestsHeaders: INTERESTS_HEADERS,
          residentCount: Object.keys(RESIDENTS_BY_ROOM).length,
        };
        break;
      case "loginByRoom":
        result = loginByRoom_(req.payload, req.params);
        break;
      case "listListings":
        result = listListings_(req.payload, req.params);
        break;
      case "listListingsV2":
        result = listListingsV2_(req.payload, req.params);
        break;
      case "listListingsV3":
        result = listListingsV3_(req.payload, req.params);
        break;
      case "getListingsVersion":
        result = getListingsVersion_(req.payload, req.params);
        break;
      case "rebuildProjection":
        result = rebuildProjection_(req.payload, req.params);
        break;
      case "getProjectionHealth":
        result = getProjectionHealth_();
        break;
      case "getListingDetailV2":
        result = getListingDetailV2_(req.payload, req.params);
        break;
      case "listWishersV2":
        result = listWishersV2_(req.payload, req.params);
        break;
      case "uploadImage":
        result = uploadImage_(req.payload);
        break;
      case "addListing":
        result = addListing_(req.payload);
        break;
      case "addListingsBatch":
        result = addListingsBatch_(req.payload);
        break;
      case "updateListing":
        result = updateListing_(req.payload);
        break;
      case "addInterest":
        result = addInterest_(req.payload);
        break;
      case "removeInterest":
        result = removeInterest_(req.payload);
        break;
      case "cancelListing":
        result = cancelListing_(req.payload);
        break;
      case "listMyPage":
        result = listMyPage_(req.payload, req.params);
        break;
      case "listConflicts":
        result = listConflicts_();
        break;
      default:
        throw new Error("unsupported action: " + action);
    }

    return json_(result);
  } catch (error) {
    return json_({
      ok: false,
      error: String(error && error.message ? error.message : error),
    });
  }
}

function parseRequest_(method, e) {
  var params = (e && e.parameter) || {};
  var payload = {};

  if (params.payload) {
    payload = safeJsonParse_(params.payload);
  } else if (method === "POST" && e && e.postData && e.postData.contents) {
    var raw = String(e.postData.contents || "");
    if (raw) {
      if (raw.charAt(0) === "{") {
        payload = safeJsonParse_(raw);
      } else {
        var parsed = parseUrlEncoded_(raw);
        if (parsed.payload) {
          payload = safeJsonParse_(parsed.payload);
        }
        if (!params.action && parsed.action) {
          params.action = parsed.action;
        }
      }
    }
  }

  var action = toText_(params.action || payload.action);
  return {
    action: action,
    payload: payload || {},
    params: params || {},
  };
}

function parseUrlEncoded_(query) {
  var out = {};
  if (!query) {
    return out;
  }

  var pairs = String(query).split("&");
  for (var i = 0; i < pairs.length; i += 1) {
    var item = pairs[i].split("=");
    var key = decodeURIComponent(item[0] || "").trim();
    if (!key) {
      continue;
    }
    var value = decodeURIComponent((item[1] || "").replace(/\+/g, " "));
    out[key] = value;
  }

  return out;
}

function safeJsonParse_(text) {
  if (!text) {
    return {};
  }
  try {
    var parsed = JSON.parse(text);
    if (parsed && typeof parsed === "object") {
      return parsed;
    }
    return {};
  } catch (error) {
    throw new Error("invalid payload JSON");
  }
}

function ensureSchema_() {
  if (readFromCache_(SCHEMA_CACHE_KEY) === "1") {
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
  if (!listingsSheet) {
    listingsSheet = ss.insertSheet(LISTINGS_SHEET_NAME);
  }
  ensureSheetWithHeaderMigration_(listingsSheet, LISTINGS_HEADERS, "listings");

  var interestsSheet = ss.getSheetByName(INTERESTS_SHEET_NAME);
  if (!interestsSheet) {
    interestsSheet = ss.insertSheet(INTERESTS_SHEET_NAME);
  }
  ensureSheetWithHeaderMigration_(interestsSheet, INTERESTS_HEADERS, "interests");

  var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  if (!projectionSheet) {
    projectionSheet = ss.insertSheet(LISTINGS_PROJECTION_SHEET_NAME);
  }
  ensureSheetWithHeaderMigration_(projectionSheet, LISTINGS_PROJECTION_HEADERS, "projection");

  var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
  if (!listingIndexSheet) {
    listingIndexSheet = ss.insertSheet(LISTING_INDEX_SHEET_NAME);
  }
  ensureSheetWithHeaderMigration_(listingIndexSheet, LISTING_INDEX_HEADERS, "listing_index");

  var interestIndexSheet = ss.getSheetByName(INTEREST_INDEX_SHEET_NAME);
  if (!interestIndexSheet) {
    interestIndexSheet = ss.insertSheet(INTEREST_INDEX_SHEET_NAME);
  }
  ensureSheetWithHeaderMigration_(interestIndexSheet, INTEREST_INDEX_HEADERS, "interest_index");

  var runtimeMetaSheet = ss.getSheetByName(RUNTIME_META_SHEET_NAME);
  if (!runtimeMetaSheet) {
    runtimeMetaSheet = ss.insertSheet(RUNTIME_META_SHEET_NAME);
  }
  ensureSheetWithHeaderMigration_(runtimeMetaSheet, RUNTIME_META_HEADERS, "runtime_meta");
  ensureRuntimeMetaDefaults_(runtimeMetaSheet);

  migrateLegacyBooksSheetIfNeeded_(ss, listingsSheet);
  migrateLegacyInterestsSheetIfNeeded_(interestsSheet);

  if (projectionSheet.getLastRow() <= 1 && listingsSheet.getLastRow() > 1) {
    rebuildProjectionInternal_(ss, true);
  }
  writeToCache_(SCHEMA_CACHE_KEY, "1", SCHEMA_CACHE_TTL_SEC);
}

function readMarketplaceSnapshot_() {
  var cached = readFromCache_(MARKET_SNAPSHOT_CACHE_KEY);
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      if (parsed && parsed.listings && parsed.interests) {
        return parsed;
      }
    } catch (error) {
      // ignore broken cache value
    }
  }

  var snapshot = {
    listings: readObjectsByName_(LISTINGS_SHEET_NAME, LISTINGS_HEADERS),
    interests: readObjectsByName_(INTERESTS_SHEET_NAME, INTERESTS_HEADERS),
  };

  writeToCache_(MARKET_SNAPSHOT_CACHE_KEY, JSON.stringify(snapshot), MARKET_SNAPSHOT_TTL_SEC);
  return snapshot;
}

function invalidateMarketplaceCache_() {
  removeFromCache_(MARKET_SNAPSHOT_CACHE_KEY);
  bumpMarketComputeVersion_();
  touchProjectionVersion_();
}

function readMarketComputeVersion_() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var value = toText_(properties.getProperty(MARKET_COMPUTE_VERSION_PROPERTY));
    if (value) {
      return value;
    }
    value = String(new Date().getTime());
    properties.setProperty(MARKET_COMPUTE_VERSION_PROPERTY, value);
    return value;
  } catch (error) {
    return "0";
  }
}

function bumpMarketComputeVersion_() {
  try {
    var properties = PropertiesService.getScriptProperties();
    properties.setProperty(MARKET_COMPUTE_VERSION_PROPERTY, String(new Date().getTime()));
  } catch (error) {
    // ignore properties failures
  }
}

function readComputedCache_(prefix, suffix) {
  var version = readMarketComputeVersion_();
  return readFromCache_(prefix + ":" + version + ":" + suffix);
}

function writeComputedCache_(prefix, suffix, value, ttlSec) {
  var version = readMarketComputeVersion_();
  writeToCache_(prefix + ":" + version + ":" + suffix, value, ttlSec);
}

function readFromCache_(key) {
  try {
    return CacheService.getScriptCache().get(key);
  } catch (error) {
    return "";
  }
}

function writeToCache_(key, value, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, value, ttlSec);
  } catch (error) {
    // ignore cache write failures (size/quota)
  }
}

function removeFromCache_(key) {
  try {
    CacheService.getScriptCache().remove(key);
  } catch (error) {
    // ignore cache failures
  }
}

function ensureSheetWithHeaderMigration_(sheet, targetHeaders, typeName) {
  var currentHeaders = getHeaderRow_(sheet);

  if (!currentHeaders.length) {
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    return;
  }

  if (headersMatch_(currentHeaders, targetHeaders)) {
    return;
  }

  if (typeName === "interests" && headersMatch_(currentHeaders, LEGACY_INTEREST_HEADERS)) {
    migrateLegacyInterestsSheetIfNeeded_(sheet);
    return;
  }

  if (typeName === "listings" && headersMatch_(currentHeaders, LEGACY_BOOK_HEADERS)) {
    migrateLegacyListingsSheet_(sheet);
    return;
  }

  backupAndResetSheet_(sheet, targetHeaders);
}

function migrateLegacyBooksSheetIfNeeded_(ss, listingsSheet) {
  if (listingsSheet.getLastRow() > 1) {
    return;
  }

  var legacyBooksSheet = ss.getSheetByName(LEGACY_BOOKS_SHEET_NAME);
  if (!legacyBooksSheet) {
    return;
  }

  var headers = getHeaderRow_(legacyBooksSheet);
  if (!headersMatch_(headers, LEGACY_BOOK_HEADERS)) {
    return;
  }

  var legacyRows = readObjectsBySheet_(legacyBooksSheet, LEGACY_BOOK_HEADERS);
  if (!legacyRows.length) {
    return;
  }

  var converted = [];
  for (var i = 0; i < legacyRows.length; i += 1) {
    converted.push(convertLegacyBookToListing_(legacyRows[i]));
  }

  appendObjectRows_(listingsSheet, LISTINGS_HEADERS, converted);
}

function migrateLegacyInterestsSheetIfNeeded_(interestsSheet) {
  var headers = getHeaderRow_(interestsSheet);
  if (!headers.length) {
    interestsSheet.getRange(1, 1, 1, INTERESTS_HEADERS.length).setValues([INTERESTS_HEADERS]);
    return;
  }

  if (headersMatch_(headers, INTERESTS_HEADERS)) {
    return;
  }

  if (headersMatch_(headers, INTERESTS_HEADERS_V1)) {
    var oldRows = readObjectsBySheet_(interestsSheet, INTERESTS_HEADERS_V1);
    interestsSheet.clearContents();
    interestsSheet.getRange(1, 1, 1, INTERESTS_HEADERS.length).setValues([INTERESTS_HEADERS]);
    var migratedRows = [];
    for (var o = 0; o < oldRows.length; o += 1) {
      var oldRow = oldRows[o];
      migratedRows.push({
        created_at: toText_(oldRow.created_at),
        interest_id: toText_(oldRow.interest_id) || createId_("interest"),
        listing_id: toText_(oldRow.listing_id),
        viewer_id: toText_(oldRow.viewer_id),
        viewer_name: toText_(oldRow.viewer_name),
        is_active: ACTIVE_FLAG_ON,
      });
    }
    appendObjectRows_(interestsSheet, INTERESTS_HEADERS, migratedRows);
    return;
  }

  if (!headersMatch_(headers, LEGACY_INTEREST_HEADERS)) {
    backupAndResetSheet_(interestsSheet, INTERESTS_HEADERS);
    return;
  }

  var legacyRows = readObjectsBySheet_(interestsSheet, LEGACY_INTEREST_HEADERS);
  interestsSheet.clearContents();
  interestsSheet.getRange(1, 1, 1, INTERESTS_HEADERS.length).setValues([INTERESTS_HEADERS]);

  var converted = [];
  for (var i = 0; i < legacyRows.length; i += 1) {
    var row = legacyRows[i];
    converted.push({
      created_at: toText_(row.created_at),
      interest_id: toText_(row.interest_id) || createId_("interest"),
      listing_id: toText_(row.book_id),
      viewer_id: toText_(row.viewer_id),
      viewer_name: toText_(row.viewer_name),
      is_active: ACTIVE_FLAG_ON,
    });
  }

  appendObjectRows_(interestsSheet, INTERESTS_HEADERS, converted);
}

function migrateLegacyListingsSheet_(listingsSheet) {
  var rows = readObjectsBySheet_(listingsSheet, LEGACY_BOOK_HEADERS);
  listingsSheet.clearContents();
  listingsSheet.getRange(1, 1, 1, LISTINGS_HEADERS.length).setValues([LISTINGS_HEADERS]);

  var converted = [];
  for (var i = 0; i < rows.length; i += 1) {
    converted.push(convertLegacyBookToListing_(rows[i]));
  }

  appendObjectRows_(listingsSheet, LISTINGS_HEADERS, converted);
}

function convertLegacyBookToListing_(row) {
  var coverUrl = toText_(row.cover_url);
  var imageUrls = coverUrl ? [coverUrl] : [];

  return {
    created_at: toText_(row.created_at) || new Date().toISOString(),
    listing_id: toText_(row.book_id) || createId_("listing"),
    item_type: "book",
    category: "書籍",
    title: toText_(row.title) || "タイトル不明",
    description: toText_(row.summary),
    image_urls_json: JSON.stringify(imageUrls),
    jan: normalizeJan_(row.jan),
    subject_tags_json: JSON.stringify([]),
    author: toText_(row.author),
    publisher: toText_(row.publisher),
    published_date: toText_(row.published_date),
    seller_name: toText_(row.seller_name),
    seller_id: toText_(row.seller_id),
    status: toText_(row.status) || DEFAULT_STATUS,
  };
}

function backupAndResetSheet_(sheet, headers) {
  var ss = sheet.getParent();
  var backupName = sheet.getName() + "_backup_" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmmss");
  var backup = ss.insertSheet(backupName);
  var range = sheet.getDataRange();
  if (range.getNumRows() > 0 && range.getNumColumns() > 0) {
    range.copyTo(backup.getRange(1, 1), { contentsOnly: true });
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getHeaderRow_(sheet) {
  var maxColumn = sheet.getLastColumn();
  if (maxColumn < 1) {
    return [];
  }

  var values = sheet.getRange(1, 1, 1, maxColumn).getValues()[0];
  var headers = [];
  for (var i = 0; i < values.length; i += 1) {
    var value = toText_(values[i]);
    if (!value) {
      continue;
    }
    headers.push(value);
  }
  return headers;
}

function headersMatch_(current, expected) {
  if (!current || !expected || current.length !== expected.length) {
    return false;
  }

  for (var i = 0; i < expected.length; i += 1) {
    if (toText_(current[i]) !== toText_(expected[i])) {
      return false;
    }
  }

  return true;
}

function loginByRoom_(payload, params) {
  var roomNumber = normalizeRoomNumber_(
    (payload && payload.roomNumber) || (params && params.roomNumber)
  );
  assert_(roomNumber, "roomNumber is required");

  var residentName = RESIDENTS_BY_ROOM[roomNumber];
  if (!residentName) {
    throw new Error("該当する部屋番号が見つかりません");
  }

  return {
    ok: true,
    resident: {
      userId: createResidentId_(roomNumber),
      roomNumber: roomNumber,
      name: residentName,
    },
  };
}

function listListings_(payload, params) {
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  var category = normalizeQuery_((payload && payload.category) || (params && params.category));
  var subject = normalizeQuery_((payload && payload.subject) || (params && params.subject));
  var q = normalizeQuery_((payload && payload.q) || (params && params.q));

  var snapshot = readMarketplaceSnapshot_();
  var listings = snapshot.listings || [];
  var interests = snapshot.interests || [];

  var responses = buildListingResponses_(listings, interests, viewerId);
  var filtered = [];

  for (var i = 0; i < responses.length; i += 1) {
    var listing = responses[i];
    if (!matchListingFilter_(listing, category, subject, q)) {
      continue;
    }
    filtered.push(listing);
  }

  filtered.sort(function (a, b) {
    var av = toText_(a.createdAt);
    var bv = toText_(b.createdAt);
    if (av < bv) return 1;
    if (av > bv) return -1;
    return 0;
  });

  return {
    ok: true,
    listings: filtered,
  };
}

function listListingsV2_(payload, params) {
  if (!LISTINGS_V3_USE_FOR_V2) {
    return listListingsV2Legacy_(payload, params);
  }
  return listListingsV3_(payload, params);
}

function listListingsV3_(payload, params) {
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  var limit = parsePositiveInt_((payload && payload.limit) || (params && params.limit), 80, 1, 200);
  var cursor = parsePositiveInt_((payload && payload.cursor) || (params && params.cursor), 0, 0, 1000000);
  var startedAtMs = new Date().getTime();

  if (isProjectionDirty_()) {
    return listListingsV2Legacy_(payload, params);
  }

  var projectionVersion = readProjectionVersion_();
  var cacheSuffix = [projectionVersion || "_", viewerId || "_", String(limit), String(cursor)].join("|");
  var cached = readFromCache_(LISTINGS_V3_CACHE_KEY + ":" + cacheSuffix);
  if (cached) {
    try {
      var parsedCached = JSON.parse(cached);
      if (parsedCached && parsedCached.ok && Array.isArray(parsedCached.listings) && toText_(parsedCached.version)) {
        return parsedCached;
      }
    } catch (error) {
      // ignore broken cache value
    }
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    assert_(sheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        ok: true,
        version: projectionVersion || readMarketComputeVersion_(),
        listings: [],
        nextCursor: "",
        totalCount: 0,
      };
    }

    var chunkSize = Math.max(limit * LISTINGS_V3_CHUNK_MULTIPLIER, limit);
    var pageChunk = readProjectionChunkFromTail_(sheet, cursor, chunkSize);
    var activeCount = pageChunk.totalActive;
    var rowsRead = pageChunk.rows.length;
    var page = [];
    var maxRows = Math.min(pageChunk.rows.length, limit);
    for (var i = 0; i < maxRows; i += 1) {
      page.push(mapProjectionRowToSummaryResponse_(pageChunk.rows[i], viewerId));
    }
    var nextOffset = cursor + page.length;

    var response = {
      ok: true,
      version: projectionVersion || readMarketComputeVersion_(),
      listings: page,
      nextCursor: nextOffset < activeCount ? String(nextOffset) : "",
      totalCount: activeCount,
    };

    writeToCache_(LISTINGS_V3_CACHE_KEY + ":" + cacheSuffix, JSON.stringify(response), MARKET_COMPUTE_CACHE_TTL_SEC);
    console.log(
      "[listListingsV3] rowsRead=%s activeTotal=%s elapsedMs=%s",
      String(rowsRead),
      String(activeCount),
      String(new Date().getTime() - startedAtMs)
    );
    return response;
  } catch (error) {
    markProjectionDirty_(true);
    return listListingsV2Legacy_(payload, params);
  }
}

function listListingsV2Legacy_(payload, params) {
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  var limit = parsePositiveInt_((payload && payload.limit) || (params && params.limit), 80, 1, 200);
  var cursor = parsePositiveInt_((payload && payload.cursor) || (params && params.cursor), 0, 0, 1000000);

  var cacheSuffix = [viewerId || "_", String(limit), String(cursor)].join("|");
  var cached = readComputedCache_("listings_v2", cacheSuffix);
  if (cached) {
    try {
      var parsedLegacy = JSON.parse(cached);
      if (parsedLegacy && parsedLegacy.ok && Array.isArray(parsedLegacy.listings) && toText_(parsedLegacy.version)) {
        return parsedLegacy;
      }
    } catch (error) {
      // ignore broken cache value
    }
  }

  var summarySnapshot = readListingsSummarySnapshot_();
  var summaries = Array.isArray(summarySnapshot.listings) ? summarySnapshot.listings : [];
  var start = Math.min(cursor, summaries.length);
  var end = Math.min(start + limit, summaries.length);
  var page = [];

  for (var i = start; i < end; i += 1) {
    page.push(mapSummarySnapshotForViewer_(summaries[i], viewerId));
  }

  var response = {
    ok: true,
    version: toText_(summarySnapshot.version) || readMarketComputeVersion_(),
    listings: page,
    nextCursor: end < summaries.length ? String(end) : "",
    totalCount: summaries.length,
  };

  writeComputedCache_("listings_v2", cacheSuffix, JSON.stringify(response), MARKET_COMPUTE_CACHE_TTL_SEC);
  return response;
}

function rebuildProjection_(payload, params) {
  var startedAtMs = new Date().getTime();
  var result = withMarketplaceLock_(function () {
    var output = rebuildProjectionInternal_(SpreadsheetApp.getActiveSpreadsheet(), true);
    markProjectionDirty_(false);
    return output;
  });

  console.log(
    "[rebuildProjection] listings=%s interests=%s elapsedMs=%s",
    String(result.rowCounts.listings),
    String(result.rowCounts.interests),
    String(new Date().getTime() - startedAtMs)
  );

  return {
    ok: true,
    rebuilt: true,
    version: toText_(result.version),
    lastRebuildAt: toText_(result.lastRebuildAt),
    rowCounts: result.rowCounts,
    usedBatchGet: Boolean(result.usedBatchGet),
  };
}

function getProjectionHealth_() {
  var meta = readRuntimeMetaMap_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
  var interestsSheet = ss.getSheetByName(INTERESTS_SHEET_NAME);
  var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
  var interestIndexSheet = ss.getSheetByName(INTEREST_INDEX_SHEET_NAME);

  return {
    ok: true,
    version: toText_(meta[PROJECTION_VERSION_KEY]) || "0",
    lastRebuildAt: toText_(meta[PROJECTION_LAST_REBUILD_KEY]),
    dirty: toText_(meta[PROJECTION_DIRTY_KEY]) === ACTIVE_FLAG_ON,
    rowCounts: {
      listings: listingsSheet ? Math.max(listingsSheet.getLastRow() - 1, 0) : 0,
      interests: interestsSheet ? Math.max(interestsSheet.getLastRow() - 1, 0) : 0,
      projection: projectionSheet ? Math.max(projectionSheet.getLastRow() - 1, 0) : 0,
      listingIndex: listingIndexSheet ? Math.max(listingIndexSheet.getLastRow() - 1, 0) : 0,
      interestIndex: interestIndexSheet ? Math.max(interestIndexSheet.getLastRow() - 1, 0) : 0,
    },
  };
}

function getListingsVersion_(payload, params) {
  if (!isProjectionDirty_()) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    var meta = readRuntimeMetaMap_();
    var totalCount = projectionSheet ? countActiveProjectionRows_(projectionSheet) : 0;
    return {
      ok: true,
      version: toText_(meta[PROJECTION_VERSION_KEY]) || readMarketComputeVersion_(),
      totalCount: totalCount,
      generatedAt: toText_(meta[PROJECTION_LAST_REBUILD_KEY]),
    };
  }

  var summarySnapshot = readListingsSummarySnapshot_();
  var legacyTotalCount = Array.isArray(summarySnapshot.listings) ? summarySnapshot.listings.length : 0;
  return {
    ok: true,
    version: toText_(summarySnapshot.version) || readMarketComputeVersion_(),
    totalCount: legacyTotalCount,
    generatedAt: toText_(summarySnapshot.generatedAt),
  };
}

function readListingsSummarySnapshot_() {
  var cached = readComputedCache_(LISTINGS_SUMMARY_SNAPSHOT_CACHE_KEY, "all");
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      if (parsed && parsed.ok && Array.isArray(parsed.listings) && toText_(parsed.version)) {
        return parsed;
      }
    } catch (error) {
      // ignore broken cache value
    }
  }

  var snapshot = readMarketplaceSnapshot_();
  var listings = snapshot.listings || [];
  var interests = snapshot.interests || [];
  var groupedInterests = buildGroupedInterests_(interests);
  var availableRows = [];

  for (var i = 0; i < listings.length; i += 1) {
    var row = listings[i];
    var status = toText_(row.status) || DEFAULT_STATUS;
    if (status !== DEFAULT_STATUS) {
      continue;
    }
    availableRows.push(row);
  }

  availableRows.sort(function (a, b) {
    var av = toText_(a.created_at);
    var bv = toText_(b.created_at);
    if (av < bv) return 1;
    if (av > bv) return -1;
    return 0;
  });

  var summaries = [];
  for (var j = 0; j < availableRows.length; j += 1) {
    var summaryRow = availableRows[j];
    var listingId = toText_(summaryRow.listing_id);
    var interestMap = groupedInterests[listingId] || {};
    var summary = mapListingRowToSummaryResponse_(summaryRow, interestMap, "");
    summary.wantedViewerIds = objectKeys_(interestMap);
    delete summary.alreadyWanted;
    summaries.push(summary);
  }

  var response = {
    ok: true,
    version: readMarketComputeVersion_(),
    generatedAt: new Date().toISOString(),
    listings: summaries,
  };

  writeComputedCache_(
    LISTINGS_SUMMARY_SNAPSHOT_CACHE_KEY,
    "all",
    JSON.stringify(response),
    LISTINGS_SUMMARY_SNAPSHOT_CACHE_TTL_SEC
  );
  return response;
}

function mapSummarySnapshotForViewer_(summaryRow, viewerId) {
  var row = summaryRow || {};
  var wantedViewerIds = Array.isArray(row.wantedViewerIds) ? row.wantedViewerIds : [];
  var wantedCount = parsePositiveInt_(row.wantedCount, 0, 0, 1000000);
  var ownListing = Boolean(viewerId && toText_(row.sellerId) === viewerId);
  var alreadyWanted = Boolean(!ownListing && viewerId && wantedViewerIds.indexOf(viewerId) !== -1);

  return {
    createdAt: toText_(row.createdAt),
    listingId: toText_(row.listingId),
    itemType: normalizeItemType_(row.itemType),
    category: toText_(row.category),
    title: toText_(row.title),
    thumbUrl: normalizeImageUrl_(row.thumbUrl),
    detailImageUrl: normalizeImageUrl_(row.detailImageUrl),
    jan: normalizeJan_(row.jan),
    subjectTags: sanitizeTextArray_(row.subjectTags, 50),
    author: toText_(row.author),
    publisher: toText_(row.publisher),
    sellerName: toText_(row.sellerName),
    sellerId: toText_(row.sellerId),
    status: toText_(row.status) || DEFAULT_STATUS,
    wantedCount: wantedCount,
    alreadyWanted: alreadyWanted,
  };
}

function getListingDetailV2_(payload, params) {
  var listingId = toText_((payload && payload.listingId) || (params && params.listingId));
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  assert_(listingId, "listingId is required");

  var cacheSuffix = [listingId, viewerId || "_"].join("|");
  var cached = readComputedCache_("detail_v2", cacheSuffix);
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      if (parsed && parsed.ok && parsed.listing) {
        return parsed;
      }
    } catch (error) {
      // ignore broken cache value
    }
  }

  var snapshot = readMarketplaceSnapshot_();
  var listings = snapshot.listings || [];
  var interests = snapshot.interests || [];
  var target = null;

  for (var i = 0; i < listings.length; i += 1) {
    if (toText_(listings[i].listing_id) !== listingId) {
      continue;
    }
    var status = toText_(listings[i].status) || DEFAULT_STATUS;
    if (status !== DEFAULT_STATUS) {
      throw new Error("listing not available");
    }
    target = listings[i];
    break;
  }
  assert_(target, "listing not found");

  var groupedInterests = buildGroupedInterests_(interests);
  var interestMap = groupedInterests[listingId] || {};
  var listing = mapListingRowToDetailResponse_(target, interestMap, viewerId);

  var response = {
    ok: true,
    listing: listing,
  };
  writeComputedCache_("detail_v2", cacheSuffix, JSON.stringify(response), MARKET_COMPUTE_CACHE_TTL_SEC);
  return response;
}

function listWishersV2_(payload, params) {
  var listingId = toText_((payload && payload.listingId) || (params && params.listingId));
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  assert_(listingId, "listingId is required");

  var cacheSuffix = [listingId, viewerId ? "login" : "guest"].join("|");
  var cached = readComputedCache_("wishers_v2", cacheSuffix);
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      if (parsed && parsed.ok && parsed.listingId === listingId) {
        return parsed;
      }
    } catch (error) {
      // ignore broken cache value
    }
  }

  var snapshot = readMarketplaceSnapshot_();
  var listings = snapshot.listings || [];
  var interests = snapshot.interests || [];

  var exists = false;
  for (var i = 0; i < listings.length; i += 1) {
    if (toText_(listings[i].listing_id) !== listingId) {
      continue;
    }
    var status = toText_(listings[i].status) || DEFAULT_STATUS;
    if (status === DEFAULT_STATUS) {
      exists = true;
    }
    break;
  }
  assert_(exists, "listing not found");

  var groupedInterests = buildGroupedInterests_(interests);
  var wishers = objectValues_(groupedInterests[listingId] || {});
  var response = {
    ok: true,
    listingId: listingId,
    wantedCount: wishers.length,
    wantedBy: viewerId ? wishers : [],
  };
  writeComputedCache_("wishers_v2", cacheSuffix, JSON.stringify(response), MARKET_COMPUTE_CACHE_TTL_SEC);
  return response;
}

function addListing_(payload) {
  var payloadObj = payload || {};
  var singleListing = payloadObj.listing || {};
  var created = addListingsBatchCore_(payloadObj.sellerId, payloadObj.sellerName, [singleListing]);
  var first = created.records[0];

  return {
    ok: true,
    listingId: first ? toText_(first.listing_id) : "",
    record: first ? mapListingRowToResponse_(first, {}, "") : {},
  };
}

function addListingsBatch_(payload) {
  var payloadObj = payload || {};
  var listings = Array.isArray(payloadObj.listings) ? payloadObj.listings : [];
  assert_(listings.length > 0, "listings is required");
  if (listings.length > 20) {
    throw new Error("listings must be 20 or less");
  }

  var created = addListingsBatchCore_(payloadObj.sellerId, payloadObj.sellerName, listings);
  return {
    ok: true,
    createdCount: created.records.length,
    listingIds: created.listingIds,
  };
}

function addListingsBatchCore_(sellerIdRaw, sellerNameRaw, listings) {
  var sellerId = toText_(sellerIdRaw);
  var sellerName = toText_(sellerNameRaw);
  assert_(sellerId, "sellerId is required");
  assert_(sellerName, "sellerName is required");

  var records = [];
  for (var i = 0; i < listings.length; i += 1) {
    records.push(buildListingRecordFromPayload_(listings[i] || {}, sellerId, sellerName));
  }

  return withMarketplaceLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
    var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);

    assert_(listingsSheet, "sheet not found: " + LISTINGS_SHEET_NAME);
    assert_(projectionSheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
    assert_(listingIndexSheet, "sheet not found: " + LISTING_INDEX_SHEET_NAME);

    var listingStartRow = appendObjectRowsAndGetStartRow_(listingsSheet, LISTINGS_HEADERS, records);
    var projectionRecords = [];
    var listingIndexRecords = [];

    for (var j = 0; j < records.length; j += 1) {
      var record = records[j];
      projectionRecords.push(buildProjectionRecordFromListing_(record, [], 0));
      listingIndexRecords.push({
        listing_id: toText_(record.listing_id),
        listings_row: listingStartRow + j,
        projection_row: 0,
      });
    }

    var projectionStartRow = appendObjectRowsAndGetStartRow_(
      projectionSheet,
      LISTINGS_PROJECTION_HEADERS,
      projectionRecords
    );

    for (var k = 0; k < listingIndexRecords.length; k += 1) {
      listingIndexRecords[k].projection_row = projectionStartRow + k;
    }
    appendObjectRows_(listingIndexSheet, LISTING_INDEX_HEADERS, listingIndexRecords);

    invalidateMarketplaceCache_();
    markProjectionDirty_(false);

    var listingIds = [];
    for (var m = 0; m < records.length; m += 1) {
      listingIds.push(toText_(records[m].listing_id));
    }

    return {
      records: records,
      listingIds: listingIds,
    };
  });
}

function buildListingRecordFromPayload_(listing, sellerId, sellerName) {
  var itemType = normalizeItemType_(listing.itemType);
  var category = toText_(listing.category) || defaultCategoryByItemType_(itemType);
  var title = toText_(listing.title);
  var description = toText_(listing.description);
  var jan = normalizeJan_(listing.jan);

  assert_(title, "title is required");
  assert_(category, "category is required");
  if (jan && !isValidBookJan_(jan)) {
    throw new Error("invalid jan");
  }

  var imageUrls = sanitizeUrlArray_(listing.imageUrls, 5);
  var subjectTags = itemType === "book" ? sanitizeTextArray_(listing.subjectTags, 20) : [];

  return {
    created_at: new Date().toISOString(),
    listing_id: createId_("listing"),
    item_type: itemType,
    category: category,
    title: title,
    description: description,
    image_urls_json: JSON.stringify(imageUrls),
    jan: jan,
    subject_tags_json: JSON.stringify(subjectTags),
    author: itemType === "book" ? toText_(listing.author) : "",
    publisher: itemType === "book" ? toText_(listing.publisher) : "",
    published_date: itemType === "book" ? toText_(listing.publishedDate) : "",
    seller_name: sellerName,
    seller_id: sellerId,
    status: DEFAULT_STATUS,
  };
}

function buildListingRecordForUpdate_(currentRow, listing) {
  var itemType = normalizeItemType_(listing.itemType);
  var category = toText_(listing.category) || defaultCategoryByItemType_(itemType);
  var title = toText_(listing.title);
  var description = toText_(listing.description);
  var jan = normalizeJan_(listing.jan);

  assert_(title, "title is required");
  assert_(category, "category is required");

  if (itemType === "book" && jan && !isValidBookJan_(jan)) {
    throw new Error("invalid jan");
  }

  var imageUrls = sanitizeUrlArray_(listing.imageUrls, 5);
  var subjectTags = itemType === "book" ? sanitizeTextArray_(listing.subjectTags, 20) : [];

  return {
    created_at: toText_(currentRow.created_at),
    listing_id: toText_(currentRow.listing_id),
    item_type: itemType,
    category: category,
    title: title,
    description: description,
    image_urls_json: JSON.stringify(imageUrls),
    jan: itemType === "book" ? jan : "",
    subject_tags_json: JSON.stringify(subjectTags),
    author: itemType === "book" ? toText_(listing.author) : "",
    publisher: itemType === "book" ? toText_(listing.publisher) : "",
    published_date: itemType === "book" ? toText_(listing.publishedDate) : "",
    seller_name: toText_(currentRow.seller_name),
    seller_id: toText_(currentRow.seller_id),
    status: DEFAULT_STATUS,
  };
}

function updateListing_(payload) {
  var payloadObj = payload || {};
  var sellerId = toText_(payloadObj.sellerId);
  var listingId = toText_(payloadObj.listingId);
  var listing = payloadObj.listing || {};

  assert_(sellerId, "sellerId is required");
  assert_(listingId, "listingId is required");

  return withMarketplaceLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
    var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
    assert_(listingsSheet, "sheet not found: " + LISTINGS_SHEET_NAME);
    assert_(projectionSheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
    assert_(listingIndexSheet, "sheet not found: " + LISTING_INDEX_SHEET_NAME);

    var refs = resolveListingRows_(listingId, listingsSheet, projectionSheet, listingIndexSheet);
    if (!refs.listingsRow) {
      throw new Error("listing not found");
    }

    var currentValues = listingsSheet.getRange(refs.listingsRow, 1, 1, LISTINGS_HEADERS.length).getValues()[0];
    var currentRow = rowValuesToObject_(currentValues, LISTINGS_HEADERS);

    if (toText_(currentRow.seller_id) !== sellerId) {
      throw new Error("only seller can update listing");
    }
    var rowStatus = toText_(currentRow.status) || DEFAULT_STATUS;
    if (rowStatus !== DEFAULT_STATUS) {
      throw new Error("only available listing can be updated");
    }

    var nextRow = buildListingRecordForUpdate_(currentRow, listing);
    listingsSheet
      .getRange(refs.listingsRow, 1, 1, LISTINGS_HEADERS.length)
      .setValues([objectToOrderedRow_(nextRow, LISTINGS_HEADERS)]);

    var projectionCurrent = refs.projectionRow
      ? rowValuesToObject_(
          projectionSheet.getRange(refs.projectionRow, 1, 1, LISTINGS_PROJECTION_HEADERS.length).getValues()[0],
          LISTINGS_PROJECTION_HEADERS
        )
      : null;
    var wantedViewerIds = parseProjectionViewerIds_(projectionCurrent && projectionCurrent.wanted_viewer_ids_json);
    var wantedCount = parsePositiveInt_(
      projectionCurrent ? projectionCurrent.wanted_count : wantedViewerIds.length,
      wantedViewerIds.length,
      0,
      1000000
    );
    var projectionRecord = buildProjectionRecordFromListing_(nextRow, wantedViewerIds, wantedCount);
    if (projectionCurrent && toText_(projectionCurrent.projection_row_id)) {
      projectionRecord.projection_row_id = toText_(projectionCurrent.projection_row_id);
    }

    var nextProjectionRow = refs.projectionRow;
    if (nextProjectionRow) {
      projectionSheet
        .getRange(nextProjectionRow, 1, 1, LISTINGS_PROJECTION_HEADERS.length)
        .setValues([objectToOrderedRow_(projectionRecord, LISTINGS_PROJECTION_HEADERS)]);
    } else {
      nextProjectionRow = appendObjectRowsAndGetStartRow_(
        projectionSheet,
        LISTINGS_PROJECTION_HEADERS,
        [projectionRecord]
      );
    }

    upsertListingIndexEntry_(listingIndexSheet, listingId, refs.listingsRow, nextProjectionRow);
    invalidateMarketplaceCache_();
    markProjectionDirty_(false);

    return {
      ok: true,
      listingId: listingId,
      listing: mapListingRowToResponse_(nextRow, {}, sellerId),
    };
  });
}

function addInterest_(payload) {
  var viewerId = toText_(payload.viewerId);
  var viewerName = toText_(payload.viewerName);
  var listingId = toText_(payload.listingId);

  assert_(viewerId, "viewerId is required");
  assert_(viewerName, "viewerName is required");
  assert_(listingId, "listingId is required");

  return withMarketplaceLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
    var interestsSheet = ss.getSheetByName(INTERESTS_SHEET_NAME);
    var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
    var interestIndexSheet = ss.getSheetByName(INTEREST_INDEX_SHEET_NAME);
    assert_(listingsSheet, "sheet not found: " + LISTINGS_SHEET_NAME);
    assert_(interestsSheet, "sheet not found: " + INTERESTS_SHEET_NAME);
    assert_(projectionSheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
    assert_(listingIndexSheet, "sheet not found: " + LISTING_INDEX_SHEET_NAME);
    assert_(interestIndexSheet, "sheet not found: " + INTEREST_INDEX_SHEET_NAME);

    var listingRefs = resolveListingRows_(listingId, listingsSheet, projectionSheet, listingIndexSheet);
    assert_(listingRefs.listingsRow, "listing not found");
    var listingRowValues = listingsSheet.getRange(listingRefs.listingsRow, 1, 1, LISTINGS_HEADERS.length).getValues()[0];
    var listing = rowValuesToObject_(listingRowValues, LISTINGS_HEADERS);
    var listingStatus = toText_(listing.status) || DEFAULT_STATUS;
    if (listingStatus !== DEFAULT_STATUS) {
      throw new Error("listing not available");
    }
    if (toText_(listing.seller_id) === viewerId) {
      throw new Error("cannot add interest to own listing");
    }

    var interestKey = buildInterestKey_(listingId, viewerId);
    var interestIndexMap = readInterestIndexMap_(interestIndexSheet);
    var interestRef = interestIndexMap[interestKey];
    var existingActiveRows = findInterestRowsByListingAndViewer_(interestsSheet, listingId, viewerId);
    if (existingActiveRows.length > 0) {
      upsertInterestIndexEntry_(interestIndexSheet, interestKey, existingActiveRows[0], true);
      return {
        ok: true,
        created: false,
      };
    }

    if (
      interestRef &&
      interestRef.active &&
      interestRef.interestsRow > 1 &&
      interestRef.interestsRow <= interestsSheet.getLastRow()
    ) {
      var existingInterestValues = interestsSheet
        .getRange(interestRef.interestsRow, 1, 1, INTERESTS_HEADERS.length)
        .getValues()[0];
      if (
        toText_(existingInterestValues[INTERESTS_HEADERS.indexOf("listing_id")]) === listingId &&
        toText_(existingInterestValues[INTERESTS_HEADERS.indexOf("viewer_id")]) === viewerId &&
        isActiveFlag_(existingInterestValues[INTERESTS_HEADERS.indexOf("is_active")], true)
      ) {
        return {
          ok: true,
          created: false,
        };
      }
    }

    var record = {
      created_at: new Date().toISOString(),
      interest_id: createId_("interest"),
      listing_id: listingId,
      viewer_id: viewerId,
      viewer_name: viewerName,
      is_active: ACTIVE_FLAG_ON,
    };
    var interestRow = 0;
    if (interestRef && interestRef.interestsRow > 1) {
      interestRow = interestRef.interestsRow;
      interestsSheet
        .getRange(interestRow, 1, 1, INTERESTS_HEADERS.length)
        .setValues([objectToOrderedRow_(record, INTERESTS_HEADERS)]);
      upsertInterestIndexEntry_(interestIndexSheet, interestKey, interestRow, true);
    } else {
      interestRow = appendObjectRowsAndGetStartRow_(interestsSheet, INTERESTS_HEADERS, [record]);
      upsertInterestIndexEntry_(interestIndexSheet, interestKey, interestRow, true);
    }

    var nextProjectionRow = updateProjectionWantedState_(
      projectionSheet,
      listingRefs.projectionRow,
      listing,
      viewerId,
      true
    );
    upsertListingIndexEntry_(listingIndexSheet, listingId, listingRefs.listingsRow, nextProjectionRow);

    invalidateMarketplaceCache_();
    markProjectionDirty_(false);

    return {
      ok: true,
      created: true,
      interestId: record.interest_id,
    };
  });
}

function removeInterest_(payload) {
  var viewerId = toText_(payload.viewerId);
  var listingId = toText_(payload.listingId);

  assert_(viewerId, "viewerId is required");
  assert_(listingId, "listingId is required");

  return withMarketplaceLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var interestsSheet = ss.getSheetByName(INTERESTS_SHEET_NAME);
    var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
    var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
    var interestIndexSheet = ss.getSheetByName(INTEREST_INDEX_SHEET_NAME);
    if (!interestsSheet || interestsSheet.getLastRow() <= 1) {
      return {
        ok: true,
        removed: false,
        removedCount: 0,
      };
    }

    var listingRefs = resolveListingRows_(listingId, listingsSheet, projectionSheet, listingIndexSheet);
    var listingRecord = null;
    if (listingRefs.listingsRow) {
      listingRecord = rowValuesToObject_(
        listingsSheet.getRange(listingRefs.listingsRow, 1, 1, LISTINGS_HEADERS.length).getValues()[0],
        LISTINGS_HEADERS
      );
    }

    var interestKey = buildInterestKey_(listingId, viewerId);
    var interestIndexMap = readInterestIndexMap_(interestIndexSheet);
    var interestRef = interestIndexMap[interestKey];
    var removedRows = [];

    if (interestRef && interestRef.interestsRow > 1 && interestRef.active) {
      interestsSheet
        .getRange(interestRef.interestsRow, INTERESTS_HEADERS.indexOf("is_active") + 1)
        .setValue(ACTIVE_FLAG_OFF);
      upsertInterestIndexEntry_(interestIndexSheet, interestKey, interestRef.interestsRow, false);
      removedRows.push(interestRef.interestsRow);
    } else {
      var fallbackRows = findInterestRowsByListingAndViewer_(interestsSheet, listingId, viewerId);
      for (var i = 0; i < fallbackRows.length; i += 1) {
        interestsSheet
          .getRange(fallbackRows[i], INTERESTS_HEADERS.indexOf("is_active") + 1)
          .setValue(ACTIVE_FLAG_OFF);
        removedRows.push(fallbackRows[i]);
      }
      if (removedRows.length > 0) {
        upsertInterestIndexEntry_(interestIndexSheet, interestKey, removedRows[0], false);
      }
    }

    if (removedRows.length > 0) {
      if (listingRefs.projectionRow && listingRecord) {
        updateProjectionWantedState_(projectionSheet, listingRefs.projectionRow, listingRecord, viewerId, false);
        markProjectionDirty_(false);
      } else {
        markProjectionDirty_(true);
      }
      invalidateMarketplaceCache_();
    }

    return {
      ok: true,
      removed: removedRows.length > 0,
      removedCount: removedRows.length,
    };
  });
}

function cancelListing_(payload) {
  var payloadObj = payload || {};
  var sellerId = toText_(payloadObj.sellerId);
  var listingId = toText_(payloadObj.listingId);
  assert_(sellerId, "sellerId is required");
  assert_(listingId, "listingId is required");

  return withMarketplaceLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
    var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
    var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
    var interestsSheet = ss.getSheetByName(INTERESTS_SHEET_NAME);
    var interestIndexSheet = ss.getSheetByName(INTEREST_INDEX_SHEET_NAME);

    assert_(listingsSheet, "sheet not found: " + LISTINGS_SHEET_NAME);
    assert_(projectionSheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
    assert_(listingIndexSheet, "sheet not found: " + LISTING_INDEX_SHEET_NAME);

    var refs = resolveListingRows_(listingId, listingsSheet, projectionSheet, listingIndexSheet);
    if (!refs.listingsRow) {
      throw new Error("listing not found");
    }

    var currentValues = listingsSheet.getRange(refs.listingsRow, 1, 1, LISTINGS_HEADERS.length).getValues()[0];
    var currentRow = rowValuesToObject_(currentValues, LISTINGS_HEADERS);
    if (toText_(currentRow.seller_id) !== sellerId) {
      throw new Error("only seller can cancel listing");
    }

    var rowStatus = toText_(currentRow.status) || DEFAULT_STATUS;
    if (rowStatus !== DEFAULT_STATUS) {
      return {
        ok: true,
        cancelled: false,
        listingId: listingId,
      };
    }

    currentRow.status = "CANCELLED";
    listingsSheet
      .getRange(refs.listingsRow, 1, 1, LISTINGS_HEADERS.length)
      .setValues([objectToOrderedRow_(currentRow, LISTINGS_HEADERS)]);

    var removedCount = removeInterestsByListingId_(listingId, interestsSheet, interestIndexSheet);

    var nextProjectionRow = updateProjectionStatusOnly_(
      projectionSheet,
      refs.projectionRow,
      currentRow,
      "CANCELLED",
      ACTIVE_FLAG_OFF,
      []
    );
    upsertListingIndexEntry_(listingIndexSheet, listingId, refs.listingsRow, nextProjectionRow);

    invalidateMarketplaceCache_();
    markProjectionDirty_(false);

    return {
      ok: true,
      cancelled: true,
      listingId: listingId,
      removedInterests: removedCount,
    };
  });
}

function removeInterestsByListingId_(listingId, interestsSheetArg, interestIndexSheetArg) {
  var sheet = interestsSheetArg || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTERESTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  var interestIndexSheet =
    interestIndexSheetArg || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTEREST_INDEX_SHEET_NAME);
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, INTERESTS_HEADERS.length).getValues();
  var rowsToDeactivate = [];

  for (var i = 0; i < values.length; i += 1) {
    var rowListingId = toText_(values[i][2]);
    var activeRaw = values[i][INTERESTS_HEADERS.indexOf("is_active")];
    if (rowListingId === listingId && isActiveFlag_(activeRaw, true)) {
      rowsToDeactivate.push(i + 2);
    }
  }

  for (var j = 0; j < rowsToDeactivate.length; j += 1) {
    var rowNumber = rowsToDeactivate[j];
    sheet.getRange(rowNumber, INTERESTS_HEADERS.indexOf("is_active") + 1).setValue(ACTIVE_FLAG_OFF);
    var rowValues = sheet.getRange(rowNumber, 1, 1, INTERESTS_HEADERS.length).getValues()[0];
    var key = buildInterestKey_(rowValues[2], rowValues[3]);
    if (interestIndexSheet && toText_(rowValues[2]) && toText_(rowValues[3])) {
      upsertInterestIndexEntry_(interestIndexSheet, key, rowNumber, false);
    }
  }
  return rowsToDeactivate.length;
}

function listMyPage_(payload, params) {
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  assert_(viewerId, "viewerId is required");

  var snapshot = readMarketplaceSnapshot_();
  var listings = snapshot.listings || [];
  var interests = snapshot.interests || [];
  var responses = buildListingResponses_(listings, interests, viewerId);

  var wantedListings = [];
  var myListings = [];

  for (var i = 0; i < responses.length; i += 1) {
    var listing = responses[i];
    var ownListing = toText_(listing.sellerId) === viewerId;
    if (listing.alreadyWanted && !ownListing) {
      wantedListings.push(listing);
    }
    if (ownListing) {
      myListings.push(listing);
    }
  }

  return {
    ok: true,
    wantedListings: wantedListings,
    myListings: myListings,
  };
}

function listConflicts_() {
  var snapshot = readMarketplaceSnapshot_();
  var listings = snapshot.listings || [];
  var interests = snapshot.interests || [];
  var responses = buildListingResponses_(listings, interests, "");

  var conflicts = [];
  for (var i = 0; i < responses.length; i += 1) {
    var listing = responses[i];
    if (Number(listing.wantedCount || 0) < 2) {
      continue;
    }
    conflicts.push({
      book: listing,
      interestedUsers: listing.wantedBy,
      count: listing.wantedCount,
    });
  }

  conflicts.sort(function (a, b) {
    if (a.count < b.count) return 1;
    if (a.count > b.count) return -1;
    return 0;
  });

  return {
    ok: true,
    conflicts: conflicts,
  };
}

function uploadImage_(payload) {
  var fileName = toText_(payload.fileName) || "listing-image.jpg";
  var mimeType = toText_(payload.mimeType) || "image/jpeg";
  var dataUrl = toText_(payload.dataUrl);

  assert_(dataUrl, "dataUrl is required");

  var parsed = parseDataUrl_(dataUrl, mimeType);
  assert_(parsed && parsed.base64, "invalid dataUrl");

  var bytes = Utilities.base64Decode(parsed.base64);
  var blob = Utilities.newBlob(bytes, parsed.mimeType, fileName);

  var folder = getUploadFolder_();
  var file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    // Domain policy may block changing sharing; ignore and still return file URL.
  }

  return {
    ok: true,
    imageId: file.getId(),
    imageUrl: buildDriveImageUrl_(file.getId()),
  };
}

function getUploadFolder_() {
  if (UPLOAD_FOLDER_ID) {
    return DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  }
  return DriveApp.getRootFolder();
}

function parseDataUrl_(dataUrl, fallbackMimeType) {
  var value = toText_(dataUrl);
  var match = value.match(/^data:([^;]+);base64,(.+)$/);

  if (match) {
    return {
      mimeType: toText_(match[1]) || fallbackMimeType,
      base64: match[2],
    };
  }

  if (/^[A-Za-z0-9+/=\s]+$/.test(value)) {
    return {
      mimeType: fallbackMimeType,
      base64: value.replace(/\s+/g, ""),
    };
  }

  return null;
}

function buildDriveImageUrl_(fileId, edgeSize) {
  var size = parsePositiveInt_(edgeSize, 1200, 120, 2000);
  return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(fileId) + "&sz=w" + size;
}

function buildListingResponses_(listingRows, interests, viewerId) {
  var groupedInterests = buildGroupedInterests_(interests);
  var responses = [];

  for (var j = 0; j < listingRows.length; j += 1) {
    var listingRow = listingRows[j];
    var status = toText_(listingRow.status) || DEFAULT_STATUS;
    if (status !== DEFAULT_STATUS) {
      continue;
    }

    var map = groupedInterests[toText_(listingRow.listing_id)] || {};
    responses.push(mapListingRowToResponse_(listingRow, map, viewerId));
  }

  return responses;
}

function buildGroupedInterests_(interests) {
  var groupedInterests = {};
  for (var i = 0; i < interests.length; i += 1) {
    var interest = interests[i];
    if (!isActiveFlag_(interest.is_active, true)) {
      continue;
    }
    var listingId = toText_(interest.listing_id);
    var currentViewerId = toText_(interest.viewer_id);
    if (!listingId || !currentViewerId) {
      continue;
    }

    if (!groupedInterests[listingId]) {
      groupedInterests[listingId] = {};
    }

    groupedInterests[listingId][currentViewerId] = {
      viewerId: currentViewerId,
      viewerName: toText_(interest.viewer_name),
    };
  }
  return groupedInterests;
}

function resolveListingImageVariants_(row) {
  var imageUrls = parseJsonArray_(row.image_urls_json);
  var firstImageUrl = imageUrls.length ? normalizeImageUrl_(imageUrls[0]) : "";
  var driveId = extractDriveFileId_(firstImageUrl);
  if (driveId) {
    return {
      thumbUrl: buildDriveImageUrl_(driveId, 480),
      detailImageUrl: buildDriveImageUrl_(driveId, 1200),
    };
  }
  return {
    thumbUrl: firstImageUrl,
    detailImageUrl: firstImageUrl,
  };
}

function mapListingRowToSummaryResponse_(row, interestMap, viewerId) {
  var listingId = toText_(row.listing_id);
  var subjectTags = parseJsonArray_(row.subject_tags_json);
  var variants = resolveListingImageVariants_(row);
  var wantedCount = objectKeys_(interestMap).length;
  var ownListing = Boolean(viewerId && toText_(row.seller_id) === viewerId);

  return {
    createdAt: toText_(row.created_at),
    listingId: listingId,
    itemType: normalizeItemType_(row.item_type),
    category: toText_(row.category),
    title: toText_(row.title),
    thumbUrl: variants.thumbUrl,
    detailImageUrl: variants.detailImageUrl,
    jan: normalizeJan_(row.jan),
    subjectTags: subjectTags,
    author: toText_(row.author),
    publisher: toText_(row.publisher),
    sellerName: toText_(row.seller_name),
    sellerId: toText_(row.seller_id),
    status: toText_(row.status) || DEFAULT_STATUS,
    wantedCount: wantedCount,
    alreadyWanted: Boolean(!ownListing && viewerId && interestMap && interestMap[viewerId]),
  };
}

function mapListingRowToDetailResponse_(row, interestMap, viewerId) {
  var detail = mapListingRowToResponse_(row, interestMap, viewerId);
  delete detail.wantedBy;
  return detail;
}

function matchListingFilter_(listing, category, subject, q) {
  var categoryLower = normalizeQuery_(listing.category);
  var subjectTags = sanitizeTextArray_(listing.subjectTags, 50);

  if (category && categoryLower !== category) {
    return false;
  }

  if (subject) {
    var foundSubject = false;
    for (var i = 0; i < subjectTags.length; i += 1) {
      if (normalizeQuery_(subjectTags[i]) === subject) {
        foundSubject = true;
        break;
      }
    }
    if (!foundSubject) {
      return false;
    }
  }

  if (q) {
    var haystack = [
      listing.title,
      listing.description,
      listing.category,
      listing.sellerName,
      listing.author,
      listing.publisher,
      listing.jan,
      subjectTags.join(" "),
    ]
      .join(" ")
      .toLowerCase();

    if (haystack.indexOf(q) === -1) {
      return false;
    }
  }

  return true;
}

function mapListingRowToResponse_(row, interestMap, viewerId) {
  var wantedBy = objectValues_(interestMap);
  var listingId = toText_(row.listing_id);
  var subjectTags = parseJsonArray_(row.subject_tags_json);
  var variants = resolveListingImageVariants_(row);
  var ownListing = Boolean(viewerId && toText_(row.seller_id) === viewerId);

  return {
    createdAt: toText_(row.created_at),
    listingId: listingId,
    itemType: normalizeItemType_(row.item_type),
    category: toText_(row.category),
    title: toText_(row.title),
    description: toText_(row.description),
    imageUrls: parseJsonArray_(row.image_urls_json),
    thumbUrl: variants.thumbUrl,
    detailImageUrl: variants.detailImageUrl,
    jan: normalizeJan_(row.jan),
    subjectTags: subjectTags,
    author: toText_(row.author),
    publisher: toText_(row.publisher),
    publishedDate: toText_(row.published_date),
    sellerName: toText_(row.seller_name),
    sellerId: toText_(row.seller_id),
    status: toText_(row.status) || DEFAULT_STATUS,
    wantedCount: wantedBy.length,
    alreadyWanted: Boolean(!ownListing && viewerId && interestMap && interestMap[viewerId]),
    wantedBy: wantedBy,
  };
}

function withMarketplaceLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(MARKET_LOCK_WAIT_MS);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function isActiveFlag_(value, fallback) {
  var text = toText_(value).toLowerCase();
  if (!text) {
    return Boolean(fallback);
  }
  if (text === ACTIVE_FLAG_ON || text === "true" || text === "active" || text === "yes") {
    return true;
  }
  if (text === ACTIVE_FLAG_OFF || text === "false" || text === "inactive" || text === "no") {
    return false;
  }
  return Boolean(fallback);
}

function rowValuesToObject_(values, headers) {
  var rowValues = Array.isArray(values) ? values : [];
  var row = {};
  for (var i = 0; i < headers.length; i += 1) {
    row[headers[i]] = i < rowValues.length ? rowValues[i] : "";
  }
  return row;
}

function objectToOrderedRow_(obj, headers) {
  var record = obj || {};
  var row = [];
  for (var i = 0; i < headers.length; i += 1) {
    var key = headers[i];
    row.push(record[key] !== undefined ? record[key] : "");
  }
  return row;
}

function appendObjectRowsAndGetStartRow_(sheet, headers, records) {
  if (!records || !records.length) {
    return 0;
  }
  var startRow = sheet.getLastRow() + 1;
  appendObjectRows_(sheet, headers, records);
  return startRow;
}

function buildInterestKey_(listingId, viewerId) {
  return toText_(listingId) + "|" + toText_(viewerId);
}

function parseProjectionViewerIds_(value) {
  if (!value) {
    return [];
  }
  var ids = parseJsonArray_(value);
  return sanitizeTextArray_(ids, 1000);
}

function buildProjectionRecordFromListing_(listingRow, wantedViewerIds, wantedCount) {
  var row = listingRow || {};
  var status = toText_(row.status) || DEFAULT_STATUS;
  var variants = resolveListingImageVariants_(row);
  var viewerIds = sanitizeTextArray_(wantedViewerIds, 1000);
  var count = parsePositiveInt_(wantedCount, viewerIds.length, 0, 1000000);
  if (count < viewerIds.length) {
    count = viewerIds.length;
  }
  var subjectTags = parseJsonArray_(row.subject_tags_json);

  return {
    projection_row_id: createId_("projection"),
    created_at: toText_(row.created_at),
    listing_id: toText_(row.listing_id),
    item_type: normalizeItemType_(row.item_type),
    category: toText_(row.category),
    title: toText_(row.title),
    thumb_url: variants.thumbUrl,
    detail_image_url: variants.detailImageUrl,
    jan: normalizeJan_(row.jan),
    subject_tags_json: JSON.stringify(subjectTags),
    author: toText_(row.author),
    publisher: toText_(row.publisher),
    seller_name: toText_(row.seller_name),
    seller_id: toText_(row.seller_id),
    status: status,
    is_active: status === DEFAULT_STATUS ? ACTIVE_FLAG_ON : ACTIVE_FLAG_OFF,
    wanted_count: count,
    wanted_viewer_ids_json: JSON.stringify(viewerIds),
    updated_at: new Date().toISOString(),
  };
}

function mapProjectionRowToSummaryResponse_(row, viewerId) {
  var projectionRow = row || {};
  var wantedViewerIds = parseProjectionViewerIds_(projectionRow.wanted_viewer_ids_json);
  var wantedCount = parsePositiveInt_(projectionRow.wanted_count, wantedViewerIds.length, 0, 1000000);
  var sellerId = toText_(projectionRow.seller_id);
  var ownListing = Boolean(viewerId && sellerId === viewerId);
  var alreadyWanted = Boolean(!ownListing && viewerId && wantedViewerIds.indexOf(viewerId) !== -1);

  return {
    createdAt: toText_(projectionRow.created_at),
    listingId: toText_(projectionRow.listing_id),
    itemType: normalizeItemType_(projectionRow.item_type),
    category: toText_(projectionRow.category),
    title: toText_(projectionRow.title),
    thumbUrl: normalizeImageUrl_(projectionRow.thumb_url),
    detailImageUrl: normalizeImageUrl_(projectionRow.detail_image_url),
    jan: normalizeJan_(projectionRow.jan),
    subjectTags: parseJsonArray_(projectionRow.subject_tags_json),
    author: toText_(projectionRow.author),
    publisher: toText_(projectionRow.publisher),
    sellerName: toText_(projectionRow.seller_name),
    sellerId: sellerId,
    status: toText_(projectionRow.status) || DEFAULT_STATUS,
    wantedCount: wantedCount,
    alreadyWanted: alreadyWanted,
  };
}

function countActiveProjectionRows_(sheet) {
  var projectionSheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  if (!projectionSheet || projectionSheet.getLastRow() <= 1) {
    return 0;
  }

  var activeColumn = LISTINGS_PROJECTION_HEADERS.indexOf("is_active") + 1;
  var values = projectionSheet.getRange(2, activeColumn, projectionSheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < values.length; i += 1) {
    if (isActiveFlag_(values[i][0], false)) {
      count += 1;
    }
  }
  return count;
}

function readProjectionChunkFromTail_(sheet, activeOffset, chunkSize) {
  var projectionSheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  if (!projectionSheet || projectionSheet.getLastRow() <= 1) {
    return {
      rows: [],
      totalActive: 0,
    };
  }

  var values = projectionSheet
    .getRange(2, 1, projectionSheet.getLastRow() - 1, LISTINGS_PROJECTION_HEADERS.length)
    .getValues();
  var offset = parsePositiveInt_(activeOffset, 0, 0, 10000000);
  var maxRows = parsePositiveInt_(chunkSize, 80, 1, 2000);
  var skipped = 0;
  var rows = [];
  var totalActive = 0;

  for (var i = values.length - 1; i >= 0; i -= 1) {
    var row = rowValuesToObject_(values[i], LISTINGS_PROJECTION_HEADERS);
    if (!isActiveFlag_(row.is_active, false)) {
      continue;
    }

    totalActive += 1;
    if (skipped < offset) {
      skipped += 1;
      continue;
    }

    if (rows.length < maxRows) {
      rows.push(row);
    }
  }

  return {
    rows: rows,
    totalActive: totalActive,
  };
}

function ensureRuntimeMetaDefaults_(runtimeMetaSheet) {
  var sheet = runtimeMetaSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RUNTIME_META_SHEET_NAME);
  if (!sheet) {
    return;
  }
  var meta = readRuntimeMetaMap_(sheet);
  var updates = {};

  if (!toText_(meta[PROJECTION_VERSION_KEY])) {
    updates[PROJECTION_VERSION_KEY] = String(new Date().getTime());
  }
  if (!toText_(meta[PROJECTION_LAST_REBUILD_KEY])) {
    updates[PROJECTION_LAST_REBUILD_KEY] = "";
  }
  if (!toText_(meta[PROJECTION_DIRTY_KEY])) {
    updates[PROJECTION_DIRTY_KEY] = ACTIVE_FLAG_OFF;
  }

  if (objectKeys_(updates).length > 0) {
    writeRuntimeMetaValues_(updates, sheet);
  }
}

function readRuntimeMetaMap_(runtimeMetaSheet) {
  var sheet = runtimeMetaSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RUNTIME_META_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return {};
  }

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var meta = {};
  for (var i = 0; i < values.length; i += 1) {
    var key = toText_(values[i][0]);
    if (!key) {
      continue;
    }
    meta[key] = toText_(values[i][1]);
  }
  return meta;
}

function writeRuntimeMetaValues_(updates, runtimeMetaSheet) {
  var sheet = runtimeMetaSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RUNTIME_META_SHEET_NAME);
  if (!sheet) {
    return;
  }

  var keys = objectKeys_(updates);
  if (!keys.length) {
    return;
  }

  var keyToRow = {};
  if (sheet.getLastRow() > 1) {
    var currentValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < currentValues.length; i += 1) {
      var currentKey = toText_(currentValues[i][0]);
      if (!currentKey) {
        continue;
      }
      keyToRow[currentKey] = i + 2;
    }
  }

  var rowsToAppend = [];
  for (var j = 0; j < keys.length; j += 1) {
    var key = toText_(keys[j]);
    if (!key) {
      continue;
    }
    var value = toText_(updates[key]);
    var targetRow = keyToRow[key];
    if (targetRow) {
      sheet.getRange(targetRow, 2).setValue(value);
    } else {
      rowsToAppend.push([key, value]);
    }
  }

  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
  }
}

function readProjectionVersion_() {
  var meta = readRuntimeMetaMap_();
  return toText_(meta[PROJECTION_VERSION_KEY]) || "0";
}

function touchProjectionVersion_() {
  writeRuntimeMetaValues_(
    (function () {
      var out = {};
      out[PROJECTION_VERSION_KEY] = String(new Date().getTime());
      return out;
    })()
  );
}

function markProjectionDirty_(dirty) {
  writeRuntimeMetaValues_(
    (function () {
      var out = {};
      out[PROJECTION_DIRTY_KEY] = dirty ? ACTIVE_FLAG_ON : ACTIVE_FLAG_OFF;
      return out;
    })()
  );
}

function isProjectionDirty_() {
  var meta = readRuntimeMetaMap_();
  return toText_(meta[PROJECTION_DIRTY_KEY]) === ACTIVE_FLAG_ON;
}

function findRowByColumnValue_(sheet, columnIndex, targetValue) {
  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  var expected = toText_(targetValue);
  if (!expected) {
    return 0;
  }

  var values = sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i += 1) {
    if (toText_(values[i][0]) === expected) {
      return i + 2;
    }
  }
  return 0;
}

function resolveListingRows_(listingId, listingsSheet, projectionSheet, listingIndexSheet) {
  var listSheet = listingsSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_SHEET_NAME);
  var projSheet =
    projectionSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  var indexSheet = listingIndexSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTING_INDEX_SHEET_NAME);
  assert_(listSheet, "sheet not found: " + LISTINGS_SHEET_NAME);
  assert_(projSheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);

  var targetListingId = toText_(listingId);
  if (!targetListingId) {
    return {
      listingsRow: 0,
      projectionRow: 0,
    };
  }

  var indexRow = findRowByColumnValue_(indexSheet, 1, targetListingId);
  var listingsRow = 0;
  var projectionRow = 0;
  if (indexRow > 1) {
    var indexValues = indexSheet.getRange(indexRow, 1, 1, LISTING_INDEX_HEADERS.length).getValues()[0];
    listingsRow = parsePositiveInt_(indexValues[1], 0, 0, 100000000);
    projectionRow = parsePositiveInt_(indexValues[2], 0, 0, 100000000);
  }

  var listingsIdColumn = LISTINGS_HEADERS.indexOf("listing_id") + 1;
  if (
    !listingsRow ||
    listingsRow < 2 ||
    listingsRow > listSheet.getLastRow() ||
    toText_(listSheet.getRange(listingsRow, listingsIdColumn).getValue()) !== targetListingId
  ) {
    listingsRow = findRowByColumnValue_(listSheet, listingsIdColumn, targetListingId);
  }

  var projectionIdColumn = LISTINGS_PROJECTION_HEADERS.indexOf("listing_id") + 1;
  if (
    !projectionRow ||
    projectionRow < 2 ||
    projectionRow > projSheet.getLastRow() ||
    toText_(projSheet.getRange(projectionRow, projectionIdColumn).getValue()) !== targetListingId
  ) {
    projectionRow = findRowByColumnValue_(projSheet, projectionIdColumn, targetListingId);
  }

  if (listingsRow || projectionRow) {
    upsertListingIndexEntry_(indexSheet, targetListingId, listingsRow, projectionRow);
  }

  return {
    listingsRow: listingsRow,
    projectionRow: projectionRow,
  };
}

function upsertListingIndexEntry_(listingIndexSheet, listingId, listingsRow, projectionRow) {
  var sheet =
    listingIndexSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTING_INDEX_SHEET_NAME);
  if (!sheet) {
    return;
  }

  var key = toText_(listingId);
  if (!key) {
    return;
  }

  var rowToWrite = findRowByColumnValue_(sheet, 1, key);
  var record = {
    listing_id: key,
    listings_row: parsePositiveInt_(listingsRow, 0, 0, 100000000),
    projection_row: parsePositiveInt_(projectionRow, 0, 0, 100000000),
  };

  if (rowToWrite > 1) {
    sheet
      .getRange(rowToWrite, 1, 1, LISTING_INDEX_HEADERS.length)
      .setValues([objectToOrderedRow_(record, LISTING_INDEX_HEADERS)]);
  } else {
    appendObjectRows_(sheet, LISTING_INDEX_HEADERS, [record]);
  }
}

function readInterestIndexMap_(interestIndexSheet) {
  var sheet =
    interestIndexSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTEREST_INDEX_SHEET_NAME);
  var map = {};
  if (!sheet || sheet.getLastRow() <= 1) {
    return map;
  }

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, INTEREST_INDEX_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i += 1) {
    var key = toText_(values[i][0]);
    if (!key) {
      continue;
    }
    map[key] = {
      indexRow: i + 2,
      interestsRow: parsePositiveInt_(values[i][1], 0, 0, 100000000),
      active: isActiveFlag_(values[i][2], false),
    };
  }
  return map;
}

function upsertInterestIndexEntry_(interestIndexSheet, interestKey, interestsRow, active) {
  var sheet =
    interestIndexSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTEREST_INDEX_SHEET_NAME);
  if (!sheet) {
    return;
  }

  var key = toText_(interestKey);
  if (!key) {
    return;
  }

  var rowToWrite = findRowByColumnValue_(sheet, 1, key);
  var record = {
    interest_key: key,
    interests_row: parsePositiveInt_(interestsRow, 0, 0, 100000000),
    is_active: active ? ACTIVE_FLAG_ON : ACTIVE_FLAG_OFF,
  };

  if (rowToWrite > 1) {
    sheet
      .getRange(rowToWrite, 1, 1, INTEREST_INDEX_HEADERS.length)
      .setValues([objectToOrderedRow_(record, INTEREST_INDEX_HEADERS)]);
  } else {
    appendObjectRows_(sheet, INTEREST_INDEX_HEADERS, [record]);
  }
}

function findInterestRowsByListingAndViewer_(interestsSheet, listingId, viewerId) {
  if (!interestsSheet || interestsSheet.getLastRow() <= 1) {
    return [];
  }

  var out = [];
  var values = interestsSheet.getRange(2, 1, interestsSheet.getLastRow() - 1, INTERESTS_HEADERS.length).getValues();
  var listingIndex = INTERESTS_HEADERS.indexOf("listing_id");
  var viewerIndex = INTERESTS_HEADERS.indexOf("viewer_id");
  var activeIndex = INTERESTS_HEADERS.indexOf("is_active");

  for (var i = 0; i < values.length; i += 1) {
    var rowListingId = toText_(values[i][listingIndex]);
    var rowViewerId = toText_(values[i][viewerIndex]);
    var rowActive = values[i][activeIndex];
    if (rowListingId !== toText_(listingId) || rowViewerId !== toText_(viewerId)) {
      continue;
    }
    if (!isActiveFlag_(rowActive, true)) {
      continue;
    }
    out.push(i + 2);
  }

  return out;
}

function updateProjectionWantedState_(projectionSheet, projectionRow, listingRecord, viewerId, shouldAdd) {
  var sheet =
    projectionSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  if (!sheet) {
    throw new Error("sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
  }

  var nextProjectionRow = parsePositiveInt_(projectionRow, 0, 0, 100000000);
  var current = null;
  if (nextProjectionRow > 1 && nextProjectionRow <= sheet.getLastRow()) {
    current = rowValuesToObject_(
      sheet.getRange(nextProjectionRow, 1, 1, LISTINGS_PROJECTION_HEADERS.length).getValues()[0],
      LISTINGS_PROJECTION_HEADERS
    );
  }

  var viewerIds = parseProjectionViewerIds_(current ? current.wanted_viewer_ids_json : "");
  var targetViewerId = toText_(viewerId);
  var pos = viewerIds.indexOf(targetViewerId);
  if (shouldAdd) {
    if (targetViewerId && pos === -1) {
      viewerIds.push(targetViewerId);
    }
  } else if (pos !== -1) {
    viewerIds.splice(pos, 1);
  }

  if (current) {
    current.wanted_count = viewerIds.length;
    current.wanted_viewer_ids_json = JSON.stringify(viewerIds);
    current.updated_at = new Date().toISOString();
    sheet
      .getRange(nextProjectionRow, 1, 1, LISTINGS_PROJECTION_HEADERS.length)
      .setValues([objectToOrderedRow_(current, LISTINGS_PROJECTION_HEADERS)]);
    return nextProjectionRow;
  }

  var listing = listingRecord || {};
  var nextRecord = buildProjectionRecordFromListing_(listing, viewerIds, viewerIds.length);
  nextProjectionRow = appendObjectRowsAndGetStartRow_(sheet, LISTINGS_PROJECTION_HEADERS, [nextRecord]);
  return nextProjectionRow;
}

function updateProjectionStatusOnly_(
  projectionSheet,
  projectionRow,
  listingRecord,
  nextStatus,
  nextIsActive,
  nextWantedViewerIds
) {
  var sheet =
    projectionSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  if (!sheet) {
    throw new Error("sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
  }

  var current = null;
  if (projectionRow > 1 && projectionRow <= sheet.getLastRow()) {
    current = rowValuesToObject_(
      sheet.getRange(projectionRow, 1, 1, LISTINGS_PROJECTION_HEADERS.length).getValues()[0],
      LISTINGS_PROJECTION_HEADERS
    );
  }

  var viewerIds = Array.isArray(nextWantedViewerIds)
    ? sanitizeTextArray_(nextWantedViewerIds, 1000)
    : parseProjectionViewerIds_(current ? current.wanted_viewer_ids_json : "");
  var nextRecord = buildProjectionRecordFromListing_(listingRecord || {}, viewerIds, viewerIds.length);
  nextRecord.status = toText_(nextStatus) || nextRecord.status;
  nextRecord.is_active = isActiveFlag_(nextIsActive, false) ? ACTIVE_FLAG_ON : ACTIVE_FLAG_OFF;
  if (current && toText_(current.projection_row_id)) {
    nextRecord.projection_row_id = toText_(current.projection_row_id);
  }

  if (projectionRow > 1 && projectionRow <= sheet.getLastRow()) {
    sheet
      .getRange(projectionRow, 1, 1, LISTINGS_PROJECTION_HEADERS.length)
      .setValues([objectToOrderedRow_(nextRecord, LISTINGS_PROJECTION_HEADERS)]);
    return projectionRow;
  }

  return appendObjectRowsAndGetStartRow_(sheet, LISTINGS_PROJECTION_HEADERS, [nextRecord]);
}

function readSheetRowsWithRowNumber_(sheet, headers) {
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i += 1) {
    rows.push({
      rowNumber: i + 2,
      row: rowValuesToObject_(values[i], headers),
    });
  }
  return rows;
}

function valueRangeToRowsWithRowNumber_(values, headers) {
  var matrix = Array.isArray(values) ? values : [];
  if (matrix.length <= 1) {
    return [];
  }

  var headerRow = matrix[0] || [];
  var headerIndex = {};
  for (var i = 0; i < headers.length; i += 1) {
    headerIndex[headers[i]] = -1;
    for (var j = 0; j < headerRow.length; j += 1) {
      if (toText_(headerRow[j]) === headers[i]) {
        headerIndex[headers[i]] = j;
        break;
      }
    }
  }

  var out = [];
  for (var r = 1; r < matrix.length; r += 1) {
    var rowValues = matrix[r] || [];
    var row = {};
    for (var h = 0; h < headers.length; h += 1) {
      var key = headers[h];
      var idx = headerIndex[key];
      row[key] = idx >= 0 && idx < rowValues.length ? rowValues[idx] : "";
    }
    out.push({
      rowNumber: r + 1,
      row: row,
    });
  }
  return out;
}

function readRowsViaBatchGet_(spreadsheetId) {
  if (
    typeof Sheets === "undefined" ||
    !Sheets.Spreadsheets ||
    !Sheets.Spreadsheets.Values ||
    !Sheets.Spreadsheets.Values.batchGet
  ) {
    return {
      ok: false,
      usedBatchGet: false,
      listings: [],
      interests: [],
    };
  }

  try {
    var response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
      ranges: [LISTINGS_SHEET_NAME + "!A:Z", INTERESTS_SHEET_NAME + "!A:Z"],
      majorDimension: "ROWS",
    });
    var valueRanges = (response && response.valueRanges) || [];
    return {
      ok: true,
      usedBatchGet: true,
      listings: valueRangeToRowsWithRowNumber_(
        valueRanges[0] ? valueRanges[0].values : [],
        LISTINGS_HEADERS
      ),
      interests: valueRangeToRowsWithRowNumber_(
        valueRanges[1] ? valueRanges[1].values : [],
        INTERESTS_HEADERS
      ),
    };
  } catch (error) {
    return {
      ok: false,
      usedBatchGet: false,
      listings: [],
      interests: [],
    };
  }
}

function rewriteSheetWithObjects_(sheet, headers, records) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (!records || !records.length) {
    return;
  }

  var chunkSize = 500;
  for (var i = 0; i < records.length; i += chunkSize) {
    var chunk = records.slice(i, i + chunkSize);
    var rows = [];
    for (var j = 0; j < chunk.length; j += 1) {
      rows.push(objectToOrderedRow_(chunk[j], headers));
    }
    sheet.getRange(i + 2, 1, rows.length, headers.length).setValues(rows);
  }
}

function rebuildProjectionInternal_(spreadsheet, useBatchGet) {
  var ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var listingsSheet = ss.getSheetByName(LISTINGS_SHEET_NAME);
  var interestsSheet = ss.getSheetByName(INTERESTS_SHEET_NAME);
  var projectionSheet = ss.getSheetByName(LISTINGS_PROJECTION_SHEET_NAME);
  var listingIndexSheet = ss.getSheetByName(LISTING_INDEX_SHEET_NAME);
  var interestIndexSheet = ss.getSheetByName(INTEREST_INDEX_SHEET_NAME);
  var runtimeMetaSheet = ss.getSheetByName(RUNTIME_META_SHEET_NAME);

  assert_(listingsSheet, "sheet not found: " + LISTINGS_SHEET_NAME);
  assert_(interestsSheet, "sheet not found: " + INTERESTS_SHEET_NAME);
  assert_(projectionSheet, "sheet not found: " + LISTINGS_PROJECTION_SHEET_NAME);
  assert_(listingIndexSheet, "sheet not found: " + LISTING_INDEX_SHEET_NAME);
  assert_(interestIndexSheet, "sheet not found: " + INTEREST_INDEX_SHEET_NAME);
  assert_(runtimeMetaSheet, "sheet not found: " + RUNTIME_META_SHEET_NAME);

  var startedAtMs = new Date().getTime();
  var batchRead = {
    ok: false,
    usedBatchGet: false,
    listings: [],
    interests: [],
  };
  if (useBatchGet) {
    batchRead = readRowsViaBatchGet_(ss.getId());
  }

  var listingRows = batchRead.ok ? batchRead.listings : readSheetRowsWithRowNumber_(listingsSheet, LISTINGS_HEADERS);
  var interestRows = batchRead.ok
    ? batchRead.interests
    : readSheetRowsWithRowNumber_(interestsSheet, INTERESTS_HEADERS);

  var listingMap = {};
  var listingOrder = [];
  for (var i = 0; i < listingRows.length; i += 1) {
    var listingRow = listingRows[i].row || {};
    var listingId = toText_(listingRow.listing_id);
    if (!listingId) {
      continue;
    }
    if (!listingMap[listingId]) {
      listingOrder.push(listingId);
    }
    listingMap[listingId] = {
      rowNumber: listingRows[i].rowNumber,
      row: listingRow,
    };
  }

  var interestIndexMap = {};
  var groupedByListing = {};
  for (var j = 0; j < interestRows.length; j += 1) {
    var interestRow = interestRows[j].row || {};
    var interestListingId = toText_(interestRow.listing_id);
    var interestViewerId = toText_(interestRow.viewer_id);
    if (!interestListingId || !interestViewerId) {
      continue;
    }

    var interestKey = buildInterestKey_(interestListingId, interestViewerId);
    var interestActive = isActiveFlag_(interestRow.is_active, true);
    interestIndexMap[interestKey] = {
      interest_key: interestKey,
      interests_row: interestRows[j].rowNumber,
      is_active: interestActive ? ACTIVE_FLAG_ON : ACTIVE_FLAG_OFF,
    };

    if (!groupedByListing[interestListingId]) {
      groupedByListing[interestListingId] = {};
    }
    if (interestActive) {
      groupedByListing[interestListingId][interestViewerId] = {
        viewerId: interestViewerId,
        viewerName: toText_(interestRow.viewer_name),
      };
    } else {
      delete groupedByListing[interestListingId][interestViewerId];
    }
  }

  var projectionRecords = [];
  var listingIndexRecords = [];
  for (var k = 0; k < listingOrder.length; k += 1) {
    var id = listingOrder[k];
    var listingEntry = listingMap[id];
    if (!listingEntry) {
      continue;
    }
    var grouped = groupedByListing[id] || {};
    var viewerIds = objectKeys_(grouped);
    var projectionRecord = buildProjectionRecordFromListing_(listingEntry.row, viewerIds, viewerIds.length);
    projectionRecords.push(projectionRecord);
    listingIndexRecords.push({
      listing_id: id,
      listings_row: listingEntry.rowNumber,
      projection_row: projectionRecords.length + 1,
    });
  }

  var interestIndexRecords = objectValues_(interestIndexMap);

  rewriteSheetWithObjects_(projectionSheet, LISTINGS_PROJECTION_HEADERS, projectionRecords);
  rewriteSheetWithObjects_(listingIndexSheet, LISTING_INDEX_HEADERS, listingIndexRecords);
  rewriteSheetWithObjects_(interestIndexSheet, INTEREST_INDEX_HEADERS, interestIndexRecords);

  var nowIso = new Date().toISOString();
  var nextVersion = String(new Date().getTime());
  writeRuntimeMetaValues_(
    (function () {
      var updates = {};
      updates[PROJECTION_VERSION_KEY] = nextVersion;
      updates[PROJECTION_LAST_REBUILD_KEY] = nowIso;
      updates[PROJECTION_DIRTY_KEY] = ACTIVE_FLAG_OFF;
      return updates;
    })(),
    runtimeMetaSheet
  );

  removeFromCache_(MARKET_SNAPSHOT_CACHE_KEY);
  bumpMarketComputeVersion_();

  console.log(
    "[projection] rebuilt listings=%s interests=%s elapsedMs=%s",
    String(projectionRecords.length),
    String(interestRows.length),
    String(new Date().getTime() - startedAtMs)
  );

  return {
    version: nextVersion,
    lastRebuildAt: nowIso,
    usedBatchGet: batchRead.ok && batchRead.usedBatchGet,
    rowCounts: {
      listings: listingRows.length,
      interests: interestRows.length,
      projection: projectionRecords.length,
      listingIndex: listingIndexRecords.length,
      interestIndex: interestIndexRecords.length,
    },
  };
}

function findListingById_(listingId) {
  var snapshot = readMarketplaceSnapshot_();
  var rows = snapshot.listings || [];
  for (var i = 0; i < rows.length; i += 1) {
    if (toText_(rows[i].listing_id) === listingId) {
      return rows[i];
    }
  }
  return null;
}

function readObjectsByName_(sheetName, headers) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return [];
  }
  return readObjectsBySheet_(sheet, headers);
}

function readObjectsBySheet_(sheet, headers) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var rows = [];

  for (var i = 0; i < values.length; i += 1) {
    var row = {};
    for (var j = 0; j < headers.length; j += 1) {
      row[headers[j]] = values[i][j];
    }
    rows.push(row);
  }

  return rows;
}

function appendObjectRowByName_(sheetName, headers, record) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("sheet not found: " + sheetName);
  }
  appendObjectRows_(sheet, headers, [record]);
}

function appendObjectRows_(sheet, headers, records) {
  if (!records || !records.length) {
    return;
  }

  var rows = [];
  for (var i = 0; i < records.length; i += 1) {
    var record = records[i];
    var row = [];
    for (var j = 0; j < headers.length; j += 1) {
      row.push(record[headers[j]] !== undefined ? record[headers[j]] : "");
    }
    rows.push(row);
  }

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
    .setValues(rows);
}

function parseJsonArray_(value) {
  if (Array.isArray(value)) {
    return sanitizeTextArray_(value, 100);
  }

  var text = toText_(value);
  if (!text) {
    return [];
  }

  try {
    var parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      return sanitizeTextArray_(parsed, 100);
    }
  } catch (error) {
    return [];
  }

  return [];
}

function sanitizeUrlArray_(values, maxLen) {
  var array = Array.isArray(values) ? values : [];
  var seen = {};
  var output = [];

  for (var i = 0; i < array.length; i += 1) {
    var value = normalizeImageUrl_(array[i]);
    if (!value) {
      continue;
    }
    if (seen[value]) {
      continue;
    }
    seen[value] = true;
    output.push(value);
    if (output.length >= maxLen) {
      break;
    }
  }

  return output;
}

function normalizeImageUrl_(value) {
  var text = toText_(value);
  if (!text) {
    return "";
  }

  text = text.replace(/^http:\/\//i, "https://");
  var driveFileId = extractDriveFileId_(text);
  if (driveFileId) {
    return buildDriveImageUrl_(driveFileId, 1200);
  }
  return text;
}

function extractDriveFileId_(url) {
  var text = toText_(url);
  if (!text) {
    return "";
  }

  var patterns = [
    /[?&]id=([a-zA-Z0-9_-]{20,})/,
    /\/d\/([a-zA-Z0-9_-]{20,})/,
    /\/d\/([a-zA-Z0-9_-]{20,})=/,
  ];

  for (var i = 0; i < patterns.length; i += 1) {
    var match = text.match(patterns[i]);
    if (match && match[1]) {
      return match[1];
    }
  }
  return "";
}

function sanitizeTextArray_(values, maxLen) {
  var array = Array.isArray(values) ? values : [];
  var seen = {};
  var output = [];

  for (var i = 0; i < array.length; i += 1) {
    var value = toText_(array[i]);
    if (!value) {
      continue;
    }
    if (seen[value]) {
      continue;
    }
    seen[value] = true;
    output.push(value);
    if (output.length >= maxLen) {
      break;
    }
  }

  return output;
}

function normalizeItemType_(value) {
  var itemType = toText_(value).toLowerCase();
  if (itemType === "book" || itemType === "goods" || itemType === "bicycle" || itemType === "other") {
    return itemType;
  }
  return "other";
}

function defaultCategoryByItemType_(itemType) {
  switch (itemType) {
    case "book":
      return "書籍";
    case "goods":
      return "小物";
    case "bicycle":
      return "自転車";
    default:
      return "その他";
  }
}

function createResidentId_(roomNumber) {
  return "resident_" + roomNumber;
}

function normalizeRoomNumber_(value) {
  return String(value || "").replace(/[^\d]/g, "").slice(0, 3);
}

function normalizeJan_(value) {
  return String(value || "").replace(/[^\d]/g, "");
}

function isValidBookJan_(jan) {
  var code = normalizeJan_(jan);
  if (!/^\d{13}$/.test(code)) {
    return false;
  }
  if (code.indexOf("978") !== 0 && code.indexOf("979") !== 0) {
    return false;
  }

  var sum = 0;
  for (var i = 0; i < 12; i += 1) {
    var digit = Number(code.charAt(i));
    sum += digit * (i % 2 === 0 ? 1 : 3);
  }

  var check = (10 - (sum % 10)) % 10;
  return check === Number(code.charAt(12));
}

function normalizeQuery_(value) {
  return toText_(value).toLowerCase();
}

function parsePositiveInt_(value, fallback, min, max) {
  if (value === undefined || value === null || String(value).trim() === "") {
    return fallback;
  }
  var parsed = Number(value);
  if (!isFinite(parsed)) {
    return fallback;
  }
  var intValue = Math.floor(parsed);
  if (intValue < min) {
    return min;
  }
  if (intValue > max) {
    return max;
  }
  return intValue;
}

function toText_(value) {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value).trim();
}

function objectValues_(obj) {
  var keys = Object.keys(obj || {});
  var values = [];
  for (var i = 0; i < keys.length; i += 1) {
    values.push(obj[keys[i]]);
  }
  return values;
}

function objectKeys_(obj) {
  return Object.keys(obj || {});
}

function createId_(prefix) {
  return prefix + "_" + Utilities.getUuid().slice(0, 8) + "_" + new Date().getTime();
}

function assert_(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
