const fs = require("node:fs");
const path = require("node:path");

const ISIN_PATTERN = /^[A-Z]{2}[A-Z0-9]{9}\d$/;

function loadIsinCache(dataDir) {
  const filePath = path.join(dataDir, "isin-cache.json");
  const cache = JSON.parse(fs.readFileSync(filePath, "utf8"));
  if (!cache || Array.isArray(cache) || typeof cache !== "object") {
    throw new Error("Soubor isin-cache.json musí obsahovat objekt ticker → ISIN.");
  }

  return Object.fromEntries(Object.entries(cache).map(([ticker, isin]) => {
    const normalizedTicker = ticker.trim().toUpperCase();
    const normalizedIsin = typeof isin === "string" ? isin.trim().toUpperCase() : "";
    if (!normalizedTicker || !ISIN_PATTERN.test(normalizedIsin)) {
      throw new Error(`Neplatná položka ticker → ISIN v souboru isin-cache.json: ${ticker}.`);
    }
    return [normalizedTicker, normalizedIsin];
  }));
}

function saveIsinCache(dataDir, updates) {
  const filePath = path.join(dataDir, "isin-cache.json");
  const cache = { ...loadIsinCache(dataDir), ...updates };
  const temporaryPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(temporaryPath, `${JSON.stringify(cache, null, 2)}\n`, "utf8");
  fs.renameSync(temporaryPath, filePath);
  return cache;
}

async function requestJson(url, endpointName, fetchImpl) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15_000);
  let response;
  try {
    response = await fetchImpl(url, { signal: controller.signal });
  } catch {
    return { error: `${endpointName} request failed; check the network and API key.` };
  } finally {
    clearTimeout(timeout);
  }

  if (!response.ok) {
    return { error: `${endpointName} returned HTTP ${response.status}.` };
  }

  try {
    return { result: await response.json() };
  } catch {
    return { error: `${endpointName} returned invalid JSON.` };
  }
}

function recordsFrom(result) {
  if (Array.isArray(result)) return result;
  return Array.isArray(result?.data) ? result.data : [];
}

function recordMatchesTicker(record, ticker) {
  const separator = ticker.lastIndexOf(".");
  const expectedCode = separator > 0 ? ticker.slice(0, separator) : ticker;
  const expectedExchange = separator > 0 ? ticker.slice(separator + 1) : "";
  const symbol = record.symbol || record.Symbol;
  if (symbol && String(symbol).trim().toUpperCase() === ticker) return true;
  const code = String(record.Code || record.code || symbol || "").trim().toUpperCase();
  const exchange = String(record.Exchange || record.exchange || "").trim().toUpperCase();
  return code === expectedCode
    && (!expectedExchange || !exchange || exchange === expectedExchange);
}

function isinFrom(record) {
  return String(record.isin || record.ISIN || "").trim().toUpperCase();
}

async function lookupIsin(ticker, apiKey, fetchImpl) {
  const separator = ticker.lastIndexOf(".");
  const code = separator > 0 ? ticker.slice(0, separator) : ticker;
  const exchange = separator > 0 ? ticker.slice(separator + 1) : "";
  const errors = [];

  const searchUrl = new URL(`https://eodhd.com/api/search/${encodeURIComponent(code)}`);
  searchUrl.searchParams.set("api_token", apiKey);
  searchUrl.searchParams.set("fmt", "json");
  if (exchange) searchUrl.searchParams.set("exchange", exchange);
  const searchResponse = await requestJson(searchUrl, "EODHD Search API", fetchImpl);
  if (searchResponse.error) {
    errors.push(searchResponse.error);
  } else {
    const records = recordsFrom(searchResponse.result);
    const record = records.find((item) => recordMatchesTicker(item, ticker));
    const isin = record && isinFrom(record);
    if (isin && ISIN_PATTERN.test(isin)) return isin;
    errors.push(records.length
      ? "Search API returned no valid ISIN for the exact symbol."
      : "Search API returned no matching records.");
  }

  const mappingUrl = new URL("https://eodhd.com/api/id-mapping");
  mappingUrl.searchParams.set("filter[symbol]", ticker);
  mappingUrl.searchParams.set("api_token", apiKey);
  mappingUrl.searchParams.set("fmt", "json");
  const mappingResponse = await requestJson(mappingUrl, "EODHD ID Mapping API", fetchImpl);
  if (mappingResponse.error) {
    errors.push(mappingResponse.error);
  } else {
    const records = recordsFrom(mappingResponse.result);
    const record = records.find((item) => recordMatchesTicker(item, ticker));
    const isin = record && isinFrom(record);
    if (isin && ISIN_PATTERN.test(isin)) return isin;
    errors.push(records.length
      ? "ID Mapping API returned no valid ISIN for the exact symbol."
      : "ID Mapping API returned no matching records.");
  }

  throw new Error(`EODHD returned no valid ISIN for ${ticker}. ${errors.join(" ")}`);
}

async function resolveIsinCache(tickers, {
  dataDir,
  apiKey = process.env.EODHD_API_KEY,
  fetchImpl = globalThis.fetch,
} = {}) {
  let cache = loadIsinCache(dataDir);
  const token = String(apiKey || "").trim();
  if (!token) return cache;
  if (typeof fetchImpl !== "function") {
    throw new Error("EODHD ISIN lookup requires a Node.js version with fetch support.");
  }

  const missingTickers = [...new Set(tickers
    .map((ticker) => String(ticker || "").trim().toUpperCase())
    .filter((ticker) => ticker && !cache[ticker]))];

  for (const ticker of missingTickers) {
    const isin = await lookupIsin(ticker, token, fetchImpl);
    cache = saveIsinCache(dataDir, { [ticker]: isin });
  }
  return cache;
}

module.exports = {
  loadIsinCache,
  resolveIsinCache,
};
