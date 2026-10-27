// CA 404 finder — the whole weekly check in one file.
// The page shows these same lines while a workbook is processed.
//
// Step 1: We open the uploaded Excel file and read every sheet.
// Step 2: We keep changiairport page views and drop paths that contain a space.
// Step 3: We read the current sitemap and index each public URL.
// Step 4: We read last week's sitemap when the workbook has one.
// Step 5: We read CACONTENT and index each page path.
// Step 6: We turn every URL into one key so the three lists can be compared.
// Step 7: We classify each page that has at least 7 views.
// Step 8: We check pages missing from the sitemap. HEAD first, then GET when HEAD is 404.
// Step 9: We choose a recommended action from the page URL.
// Step 10: We keep only broken pages and pages whose source we cannot name.
// Step 11: We build the Excel file for download.

import https from "node:https";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { dirname } from "node:path";
import * as XLSX from "xlsx";

const MIN_VIEWS = 7;
const REPORT_SUITE = "changiairport";
const TRUNCATION_LIMIT = 86;
const BASE_URL = "https://www.changiairport.com";
const HEAD_NEEDS_GET = new Set([404, 410, 405, 501]);
const RETRY = new Set([429, 500, 502, 503, 504]);

// ---------------------------------------------------------------------------
// The run. Each show() line is one step on the page.
// ---------------------------------------------------------------------------

export async function processWorkbook(file, onStep, options = {}) {
  // Step 1: We open the uploaded Excel file and read every sheet.
  await show(onStep, "Step 1: We open the uploaded Excel file and read every sheet.");
  const book = await readWorkbook(file);
  const names = sheetNames(book);

  // Step 2: We keep changiairport page views and drop paths that contain a space.
  await show(onStep, "Step 2: We keep changiairport page views and drop paths that contain a space.");
  const pages = readPageViews(book, names);

  // Step 3: We read the current sitemap and index each public URL.
  await show(onStep, "Step 3: We read the current sitemap and index each public URL.");
  const sitemap = indexSheet(sheetRows(book, "sitemap"), "sitemap");

  // Step 4: We read last week's sitemap when the workbook has one.
  await show(onStep, "Step 4: We read last week's sitemap when the workbook has one.");
  const olderName = names.find((name) => name.toLowerCase().startsWith("sitemap ") && name !== "sitemap");
  const older = olderName ? indexSheet(sheetRows(book, olderName), "older sitemap") : buildIndex([]);

  // Step 5: We read CACONTENT and index each page path.
  await show(onStep, "Step 5: We read CACONTENT and index each page path.");
  const contentName = names.find((name) => name.toLowerCase().startsWith("cacontent")) || "";
  const content = contentName ? indexSheet(sheetRows(book, contentName), "cacontent") : buildIndex([]);

  // Step 6: We turn every URL into one key so the three lists can be compared.
  await show(onStep, "Step 6: We turn every URL into one key so the three lists can be compared.");

  // Step 7: We classify each page that has at least 7 views.
  await show(onStep, "Step 7: We classify each page that has at least 7 views.");
  const findings = await classifyPages(pages, { sitemap, older, content }, options.checkUrl, onStep);

  // Step 9: We choose a recommended action from the page URL.
  await show(onStep, "Step 9: We choose a recommended action from the page URL.");
  applyActions(findings, sitemap);

  // Step 10: We keep only broken pages and pages whose source we cannot name.
  await show(onStep, "Step 10: We keep only broken pages and pages whose source we cannot name.");
  const report = findings.filter(isBrokenOrUnknownSource);

  // Step 11: We build the Excel file for download.
  await show(onStep, "Step 11: We build the Excel file for download.");
  const blob = findingsBlob(report);

  await show(onStep, "Done");
  return { findings: report, blob, pageCount: pages.length };
}

function isBrokenOrUnknownSource(row) {
  return row.status === "BROKEN" || row.recommended_action === "UNKNOWN_SOURCE";
}

// Step 7. First match wins. Pages under 7 views are left out.
async function classifyPages(pages, lists, checkUrl, onStep) {
  const rows = [];
  const waiting = [];

  pages.forEach((page) => {
    if (page.pageView < MIN_VIEWS) {
      return;
    }

    const key = normalizeUrl(page.prop1);
    const row = {
      url: key ? buildPublicUrl(key) : "",
      page_views: page.pageView,
      status: "",
      evidence: "",
      prop1: page.prop1,
      item_id: page.itemId,
      recommended_action: "", 
    };

    // Step 7a: We mark a blank path Unknown. There is no URL to open.
    if (!key) {
      row.status = "UNKNOWN";
      row.evidence = "blank prop1";
      rows.push(row);
      return;
    }

    // Step 7b: We mark one current sitemap match Active. The page is still published.
    const onSitemap = lookup(lists.sitemap, key);
    if (onSitemap.length === 1) {
      row.url = onSitemap[0].url;
      row.status = "ACTIVE";
      row.evidence = "in the current sitemap";
      rows.push(row);
      return;
    }

    // Step 7c: We mark a page that was on last week's sitemap, and is gone now, Recently unpublished.
    const onOlder = lookup(lists.older, key);
    if (onOlder.length === 1 && onSitemap.length === 0) {
      row.url = onOlder[0].url;
      row.status = "RECENTLY_UNPUBLISHED";
      row.evidence = "in the older sitemap only";
      rows.push(row);
      return;
    }

    // Step 7d: We leave an 86-character key that fits two real pages as Unknown.
    if (onSitemap.length > 1) {
      row.status = "UNKNOWN";
      row.evidence = "86-character prefix matches more than one sitemap URL";
      rows.push(row);
      return;
    }

    const inContent = lookup(lists.content, key);
    if (inContent.length === 1) {
      row.url = inContent[0].url;
    }
    waiting.push({ row, inContent: inContent.length === 1 });
  });

  // Step 8: We request each page that is missing from the sitemap. HEAD first, then GET when HEAD is 404.
  for (let index = 0; index < waiting.length; index += 1) {
    const item = waiting[index];
    if (checkUrl) {
      onStep(`Step 8: Checking page ${index + 1} of ${waiting.length}. We send HEAD, then GET when HEAD is 404.`);
      const result = await checkUrl(item.row.url);
      applyHttp(item.row, result, item.inContent);
    } else if (item.inContent) {
      item.row.status = "UNKNOWN";
      item.row.evidence = "in CACONTENT, not on the sitemap";
    } else {
      item.row.status = "UNKNOWN";
      item.row.evidence = "not in the sitemap or CACONTENT";
    }
    rows.push(item.row);
  }

  rows.sort((left, right) => right.page_views - left.page_views || left.url.localeCompare(right.url));
  return rows;
}

// Step 8 result: only HTTP 404 or 410 is Broken. A HEAD 404 is not the final answer.
function applyHttp(row, result, inContent) {
  const code = result.status;
  const where = inContent ? "in CACONTENT, not on the sitemap" : "not in the sitemap or CACONTENT";
  if (code === 404 || code === 410) {
    row.status = "BROKEN";
    row.evidence = `${where}; HTTP ${code} via ${result.method}`;
    return;
  }
  if (inContent && code === 200) {
    row.status = "ACTIVE_EXCLUDED_FROM_SEARCH";
    row.evidence = `${where}; HTTP 200 via ${result.method}`;
    return;
  }
  row.status = "UNKNOWN";
  row.evidence = code ? `${where}; HTTP ${code} via ${result.method}` : `${where}; HTTP error: ${result.error}`;
}

// Step 9: We pick the first matching action. The order below is the priority.
function applyActions(findings, sitemap) {
  findings.forEach((row) => {
    if (row.status !== "UNKNOWN" && row.status !== "BROKEN") {
      return;
    }

    const text = row.url + " " + row.prop1;

    // Step 9a: We flag language-masters as a header or footer link to fix.
    if (text.includes("language-masters")) {
      row.recommended_action = "FIX_HEADER_FOOTER_LINK";
      return;
    }

    // Step 9b: We flag a shop, dine, attraction, event, facility, or promotion URL for a config check.
    if (/shop-detail\.|dine-detail\.|attractions-directory|events-directory|facilities-and-services-directory|(^|[:/])promotions([:/]|$)/i.test(text)) {
      row.recommended_action = "CHECK_DETAIL_PAGE_URL_CONFIG";
      return;
    }

    // Step 9c: We flag a missing child whose parent page is still on the sitemap.
    const parent = row.prop1.split(":").slice(0, -1).join(":");
    if (parent && lookup(sitemap, parent).length === 1) {
      row.recommended_action = "PUBLISH_SUBPAGES";
      return;
    }

    // Step 9d: We label everything else Unknown source. This row stays in the report.
    row.recommended_action = "UNKNOWN_SOURCE";
  });
}

// ---------------------------------------------------------------------------
// Step 2 helper. Read the pageView sheet.
// ---------------------------------------------------------------------------

function readPageViews(book, names) {
  const name = names.includes("pageView") ? "pageView" : names[0];
  const raw = sheetRows(book, name);

  return raw
    .map((row) => {
      let suite = text(row["value.1"]);
      let prop1 = text(row.prop1);

      // A raw Adobe row looks like changiairport:en:search. The first colon splits it.
      if (!prop1 && row.value) {
        const value = text(row.value);
        const splitAt = value.indexOf(":");
        if (splitAt >= 0) {
          suite = value.slice(0, splitAt);
          prop1 = value.slice(splitAt + 1);
        }
      }

      return {
        itemId: text(row.itemId).replace(/\.0$/, ""),
        suite,
        prop1,
        pageView: number(row.pageView ?? row.data),
      };
    })
    .filter((row) => row.suite === REPORT_SUITE && !row.prop1.includes(" "));
}

function indexSheet(rows, source) {
  const entries = [];
  rows.forEach((row) => {
    const loc = text(row.loc);
    const path = text(row.path);
    const prop1 = text(row.prop1);
    const raw = loc || path || prop1;
    const key = normalizeUrl(raw);
    if (!key) {
      return;
    }
    entries.push({
      key,
      url: loc || buildPublicUrl(key),
      source,
    });
  });
  return buildIndex(entries);
}

// ---------------------------------------------------------------------------
// Step 6 helper. One key for a prop1, a sitemap URL, or a CACONTENT path.
// Example: https://www.changiairport.com/en/search.html -> en:search
// ---------------------------------------------------------------------------

export function normalizeUrl(value) {
  if (value === null || value === undefined) {
    return "";
  }
  let raw = String(value).trim().toLowerCase();
  if (!raw || raw === "nan" || raw === "none") {
    return "";
  }
  if (raw.startsWith("changiairport:")) {
    raw = raw.slice("changiairport:".length);
  }

  let path = extractPath(raw);
  path = decodeRepeated(path);
  path = path.split("?")[0].split("#")[0];
  if (path.endsWith(".html")) {
    path = path.slice(0, -".html".length);
  }
  while (path.length > 1 && path.endsWith("/")) {
    path = path.slice(0, -1);
  }
  // /au/en/airport/at-changi -> au:en:at-changi
  // /au/en/airport stays au:en:airport
  path = path.replaceAll("/airport/", "/");
  return path.split("/").filter(Boolean).join(":");
}

function buildPublicUrl(key) {
  const cleaned = normalizeUrl(key);
  if (!cleaned) {
    return BASE_URL + "/.html";
  }
  // A space or "&" in a DAM name must be encoded or the URL is illegal.
  const path = cleaned.split(":").map((part) => encodeURIComponent(part)).join("/");
  return BASE_URL + "/" + path + ".html";
}

function buildIndex(rows) {
  const exact = new Map();
  const all = [];
  rows.forEach((row) => {
    if (!row.key || !row.url) {
      return;
    }
    const current = exact.get(row.key) || [];
    if (!current.some((item) => item.url === row.url)) {
      current.push(row);
      exact.set(row.key, current);
      all.push(row);
    }
  });
  return { exact, all };
}

function lookup(index, raw) {
  if (!index) {
    return [];
  }
  const key = normalizeUrl(raw);
  if (!key) {
    return [];
  }
  const exact = index.exact.get(key) || [];
  if (exact.length) {
    return exact;
  }
  // "en" is not "en:search". Only a key cut at exactly 86 characters may match one longer URL.
  if (key.length !== TRUNCATION_LIMIT) {
    return [];
  }
  return index.all.filter((row) => row.key.length > TRUNCATION_LIMIT && row.key.startsWith(key));
}

function extractPath(value) {
  if (value.includes("://") || value.startsWith("www.") || value.includes("changiairport.com/")) {
    const candidate = value.includes("://") ? value : "https://" + value;
    try {
      return new URL(candidate).pathname || "";
    } catch {
      return value;
    }
  }
  if (value.startsWith("/")) {
    return value;
  }
  return "/" + value.replaceAll(":", "/");
}

function decodeRepeated(value) {
  let current = value;
  for (let i = 0; i < 5; i += 1) {
    try {
      const decoded = decodeURIComponent(current);
      if (decoded === current) {
        break;
      }
      current = decoded;
    } catch {
      break;
    }
  }
  return current;
}

// ---------------------------------------------------------------------------
// Step 1 and Step 11 helpers. Read the workbook and build the download.
// ---------------------------------------------------------------------------

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array" });
}

function sheetRows(book, name) {
  const sheet = book.Sheets[name];
  if (!sheet) {
    return [];
  }
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function sheetNames(book) {
  return book.SheetNames || [];
}

export function findingsBlob(rows) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Findings");
  const bytes = XLSX.write(book, { bookType: "xlsx", type: "array" });
  return new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

// ---------------------------------------------------------------------------
// Step 8 helper. Ask the website. HEAD first. A HEAD 404 is checked again with GET.
// ---------------------------------------------------------------------------

export function createChecker(cachePath) {
  const cache = new Map();
  let lastRequest = 0;

  return {
    async load() {
      try {
        const saved = JSON.parse(await readFile(cachePath, "utf8"));
        Object.entries(saved).forEach(([url, row]) => cache.set(url, row));
      } catch {
        // no cache yet
      }
    },

    async check(url) {
      const cached = cache.get(url);
      if (cached && !needsGet(cached)) {
        return cached;
      }
      let result = await requestWithRetry("HEAD", url, wait);
      if (needsGet(result)) {
        result = await requestWithRetry("GET", url, wait);
      }
      cache.set(url, result);
      return result;
    },

    async save() {
      await mkdir(dirname(cachePath), { recursive: true });
      await writeFile(cachePath, JSON.stringify(Object.fromEntries(cache), null, 2));
    },
  };

  function wait() {
    const elapsed = Date.now() - lastRequest;
    const delay = Math.max(0, 500 - elapsed);
    lastRequest = Date.now() + delay;
    return new Promise((resolve) => setTimeout(resolve, delay));
  }
}

function needsGet(result) {
  return result.method === "HEAD" && (result.status === null || HEAD_NEEDS_GET.has(result.status));
}

async function requestWithRetry(method, url, wait) {
  let last = { url, status: null, method, error: "request failed" };
  for (let attempt = 0; attempt < 3; attempt += 1) {
    await wait();
    last = await requestOnce(method, url);
    if (last.status !== null && !RETRY.has(last.status)) {
      return last;
    }
  }
  return last;
}

function requestOnce(method, url) {
  let safeUrl = url;
  try {
    safeUrl = encodeUrl(url);
  } catch (err) {
    return Promise.resolve({ url, status: null, method, error: err.message });
  }
  return new Promise((resolve) => {
    const req = https.request(
      safeUrl,
      {
        method,
        timeout: 10000,
        headers: { "User-Agent": "ca-404-finder/1.0 (weekly support check)", Accept: "text/html" },
      },
      (res) => {
        res.resume();
        resolve({ url, status: res.statusCode, method, error: null });
      },
    );
    req.on("timeout", () => {
      req.destroy();
      resolve({ url, status: null, method, error: "timeout" });
    });
    req.on("error", (err) => {
      resolve({ url, status: null, method, error: err.message });
    });
    req.end();
  });
}

function encodeUrl(url) {
  const parts = new URL(url);
  parts.pathname = parts.pathname
    .split("/")
    .map((part) => {
      try {
        return encodeURIComponent(decodeURIComponent(part));
      } catch {
        return encodeURIComponent(part);
      }
    })
    .join("/");
  return parts.toString();
}

function text(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function number(value) {
  const parsed = Number(String(value).replace(/[^0-9.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function show(onStep, message) {
  onStep(message);
  return new Promise((resolve) => setTimeout(resolve, 350));
}
