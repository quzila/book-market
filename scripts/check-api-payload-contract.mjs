import fs from "node:fs";
import path from "node:path";

const CONFIG_PATH = path.join(process.cwd(), "config.js");
const REQUEST_TIMEOUT_MS = 30000;
const REQUEST_RETRIES = 2;

const requiredSummaryKeys = ["listingId", "title", "category", "itemType", "wantedCount", "createdAt"];

function readApiBaseUrl() {
  if (process.env.API_CONTRACT_BASE_URL) {
    return process.env.API_CONTRACT_BASE_URL;
  }
  const text = fs.readFileSync(CONFIG_PATH, "utf8");
  const match = text.match(/APPS_SCRIPT_BASE_URL\s*:\s*["']([^"']+)["']/);
  if (!match || !match[1]) {
    throw new Error("config.js から APPS_SCRIPT_BASE_URL を読み取れませんでした。");
  }
  return match[1];
}

async function fetchJsonWithRetry(url) {
  let lastError = null;
  for (let attempt = 0; attempt <= REQUEST_RETRIES; attempt += 1) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    try {
      const response = await fetch(url, { signal: controller.signal });
      const json = await response.json();
      if (!response.ok) {
        throw new Error(json?.error || `HTTP ${response.status}`);
      }
      return json;
    } catch (error) {
      lastError = error;
      if (attempt >= REQUEST_RETRIES) {
        break;
      }
      await new Promise((resolve) => setTimeout(resolve, 700 * (attempt + 1)));
    } finally {
      clearTimeout(timeoutId);
    }
  }
  throw lastError || new Error("request failed");
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function byteLength(value) {
  return Buffer.byteLength(JSON.stringify(value), "utf8");
}

async function main() {
  const baseUrl = readApiBaseUrl();
  const listingsUrl = new URL(baseUrl);
  listingsUrl.searchParams.set("action", "listListingsV2");
  listingsUrl.searchParams.set("limit", "50");

  const listResult = await fetchJsonWithRetry(listingsUrl.toString());
  assert(listResult?.ok === true, "listListingsV2: ok が true ではありません。");
  assert(Array.isArray(listResult.listings), "listListingsV2: listings が配列ではありません。");
  assert("nextCursor" in listResult, "listListingsV2: nextCursor がありません。");

  const payloadBytes = byteLength(listResult);
  console.log(`listListingsV2 payload size: ${payloadBytes} bytes`);
  assert(payloadBytes <= 160000, "listListingsV2 payload が想定より大きすぎます。");

  listResult.listings.forEach((item, index) => {
    requiredSummaryKeys.forEach((key) => {
      assert(item[key] !== undefined, `listings[${index}].${key} がありません。`);
    });
    assert(item.wantedBy === undefined, `listings[${index}] に wantedBy が含まれています。`);
  });

  const first = listResult.listings[0];
  if (!first || !first.listingId) {
    console.log("listings が空のため detail/wishers contract check はスキップしました。");
    return;
  }

  const detailUrl = new URL(baseUrl);
  detailUrl.searchParams.set("action", "getListingDetailV2");
  detailUrl.searchParams.set("listingId", String(first.listingId));
  const detailResult = await fetchJsonWithRetry(detailUrl.toString());
  assert(detailResult?.ok === true, "getListingDetailV2: ok が true ではありません。");
  assert(detailResult?.listing?.listingId === first.listingId, "getListingDetailV2: listingId が一致しません。");
  assert(detailResult?.listing?.description !== undefined, "getListingDetailV2: description がありません。");

  const wishersUrl = new URL(baseUrl);
  wishersUrl.searchParams.set("action", "listWishersV2");
  wishersUrl.searchParams.set("listingId", String(first.listingId));
  const wishersResult = await fetchJsonWithRetry(wishersUrl.toString());
  assert(wishersResult?.ok === true, "listWishersV2: ok が true ではありません。");
  assert(wishersResult?.listingId === first.listingId, "listWishersV2: listingId が一致しません。");
  assert(wishersResult?.wantedCount !== undefined, "listWishersV2: wantedCount がありません。");
  assert(Array.isArray(wishersResult?.wantedBy), "listWishersV2: wantedBy が配列ではありません。");
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
