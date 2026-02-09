var LISTINGS_SHEET_NAME = "listings";
var INTERESTS_SHEET_NAME = "interests";
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

var INTERESTS_HEADERS = ["created_at", "interest_id", "listing_id", "viewer_id", "viewer_name"];

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
var UPLOAD_FOLDER_ID = "";
var MARKET_SNAPSHOT_CACHE_KEY = "market_snapshot_v2";
var MARKET_SNAPSHOT_TTL_SEC = 60;
var SCHEMA_CACHE_KEY = "market_schema_ok_v2";
var SCHEMA_CACHE_TTL_SEC = 3600;
var MARKET_COMPUTE_VERSION_PROPERTY = "market_compute_version_v1";
var MARKET_COMPUTE_CACHE_TTL_SEC = 60;
var LISTINGS_SUMMARY_SNAPSHOT_CACHE_KEY = "listings_summary_snapshot_v1";
var LISTINGS_SUMMARY_SNAPSHOT_CACHE_TTL_SEC = 120;

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
      case "getListingsVersion":
        result = getListingsVersion_(req.payload, req.params);
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

  migrateLegacyBooksSheetIfNeeded_(ss, listingsSheet);
  migrateLegacyInterestsSheetIfNeeded_(interestsSheet);
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
  var viewerId = toText_((payload && payload.viewerId) || (params && params.viewerId));
  var limit = parsePositiveInt_((payload && payload.limit) || (params && params.limit), 80, 1, 200);
  var cursor = parsePositiveInt_((payload && payload.cursor) || (params && params.cursor), 0, 0, 1000000);

  var cacheSuffix = [viewerId || "_", String(limit), String(cursor)].join("|");
  var cached = readComputedCache_("listings_v2", cacheSuffix);
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

function getListingsVersion_(payload, params) {
  var summarySnapshot = readListingsSummarySnapshot_();
  var totalCount = Array.isArray(summarySnapshot.listings) ? summarySnapshot.listings.length : 0;
  return {
    ok: true,
    version: toText_(summarySnapshot.version) || readMarketComputeVersion_(),
    totalCount: totalCount,
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
  var alreadyWanted = false;
  if (viewerId) {
    alreadyWanted = wantedViewerIds.indexOf(viewerId) !== -1;
  }

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

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_SHEET_NAME);
  assert_(sheet, "sheet not found: " + LISTINGS_SHEET_NAME);
  appendObjectRows_(sheet, LISTINGS_HEADERS, records);
  invalidateMarketplaceCache_();

  var listingIds = [];
  for (var j = 0; j < records.length; j += 1) {
    listingIds.push(toText_(records[j].listing_id));
  }

  return {
    records: records,
    listingIds: listingIds,
  };
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

function addInterest_(payload) {
  var viewerId = toText_(payload.viewerId);
  var viewerName = toText_(payload.viewerName);
  var listingId = toText_(payload.listingId);

  assert_(viewerId, "viewerId is required");
  assert_(viewerName, "viewerName is required");
  assert_(listingId, "listingId is required");

  var listing = findListingById_(listingId);
  assert_(listing, "listing not found");

  var snapshot = readMarketplaceSnapshot_();
  var interests = snapshot.interests || [];
  for (var i = 0; i < interests.length; i += 1) {
    var row = interests[i];
    if (toText_(row.listing_id) === listingId && toText_(row.viewer_id) === viewerId) {
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
  };

  appendObjectRowByName_(INTERESTS_SHEET_NAME, INTERESTS_HEADERS, record);
  invalidateMarketplaceCache_();

  return {
    ok: true,
    created: true,
    interestId: record.interest_id,
  };
}

function removeInterest_(payload) {
  var viewerId = toText_(payload.viewerId);
  var listingId = toText_(payload.listingId);

  assert_(viewerId, "viewerId is required");
  assert_(listingId, "listingId is required");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTERESTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      ok: true,
      removed: false,
      removedCount: 0,
    };
  }

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, INTERESTS_HEADERS.length).getValues();
  var rowsToDelete = [];

  for (var i = 0; i < values.length; i += 1) {
    var rowListingId = toText_(values[i][2]);
    var rowViewerId = toText_(values[i][3]);
    if (rowListingId === listingId && rowViewerId === viewerId) {
      rowsToDelete.push(i + 2);
    }
  }

  for (var j = rowsToDelete.length - 1; j >= 0; j -= 1) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  if (rowsToDelete.length > 0) {
    invalidateMarketplaceCache_();
  }

  return {
    ok: true,
    removed: rowsToDelete.length > 0,
    removedCount: rowsToDelete.length,
  };
}

function cancelListing_(payload) {
  var payloadObj = payload || {};
  var sellerId = toText_(payloadObj.sellerId);
  var listingId = toText_(payloadObj.listingId);
  assert_(sellerId, "sellerId is required");
  assert_(listingId, "listingId is required");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LISTINGS_SHEET_NAME);
  assert_(sheet, "sheet not found: " + LISTINGS_SHEET_NAME);
  if (sheet.getLastRow() <= 1) {
    throw new Error("listing not found");
  }

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, LISTINGS_HEADERS.length).getValues();
  var rowToUpdate = 0;
  var statusIndex = LISTINGS_HEADERS.indexOf("status");
  var sellerIndex = LISTINGS_HEADERS.indexOf("seller_id");
  var listingIndex = LISTINGS_HEADERS.indexOf("listing_id");

  for (var i = 0; i < values.length; i += 1) {
    var rowListingId = toText_(values[i][listingIndex]);
    if (rowListingId !== listingId) {
      continue;
    }

    var rowSellerId = toText_(values[i][sellerIndex]);
    if (rowSellerId !== sellerId) {
      throw new Error("only seller can cancel listing");
    }

    var rowStatus = toText_(values[i][statusIndex]) || DEFAULT_STATUS;
    if (rowStatus !== DEFAULT_STATUS) {
      return {
        ok: true,
        cancelled: false,
        listingId: listingId,
      };
    }

    rowToUpdate = i + 2;
    break;
  }

  if (!rowToUpdate) {
    throw new Error("listing not found");
  }

  sheet.getRange(rowToUpdate, statusIndex + 1).setValue("CANCELLED");
  removeInterestsByListingId_(listingId);
  invalidateMarketplaceCache_();

  return {
    ok: true,
    cancelled: true,
    listingId: listingId,
  };
}

function removeInterestsByListingId_(listingId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INTERESTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, INTERESTS_HEADERS.length).getValues();
  var rowsToDelete = [];

  for (var i = 0; i < values.length; i += 1) {
    var rowListingId = toText_(values[i][2]);
    if (rowListingId === listingId) {
      rowsToDelete.push(i + 2);
    }
  }

  for (var j = rowsToDelete.length - 1; j >= 0; j -= 1) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  return rowsToDelete.length;
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
    if (listing.alreadyWanted) {
      wantedListings.push(listing);
    }
    if (toText_(listing.sellerId) === viewerId) {
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
    alreadyWanted: Boolean(viewerId && interestMap && interestMap[viewerId]),
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
    alreadyWanted: Boolean(viewerId && interestMap && interestMap[viewerId]),
    wantedBy: wantedBy,
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
