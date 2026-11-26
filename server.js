const fs = require("node:fs");
const http = require("node:http");
const path = require("node:path");
const Busboy = require("busboy");
const { buildReport } = require("./src/report");

const root = __dirname;
const dataDir = path.join(root, "data");
const publicDir = path.join(root, "public");
const ratesPath = path.join(dataDir, "cnb-rates.json");
const port = Number(process.env.PORT || 3000);
const MAX_UPLOAD_FILES = 8;
const MAX_UPLOAD_FILE_SIZE = 20 * 1024 * 1024;
const MAX_UPLOAD_TOTAL_SIZE = 64 * 1024 * 1024;
let eodhdApiKey = process.env.EODHD_API_KEY || "";

function readRates() {
  return JSON.parse(fs.readFileSync(ratesPath, "utf8"));
}

function writeRates(rates) {
  const temporaryPath = `${ratesPath}.${process.pid}.tmp`;
  fs.writeFileSync(temporaryPath, `${JSON.stringify(rates, null, 2)}\n`, "utf8");
  fs.renameSync(temporaryPath, ratesPath);
}

function sendJson(response, status, value) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(value));
}

function readJson(request) {
  return new Promise((resolve, reject) => {
    let body = "";
    request.setEncoding("utf8");
    request.on("data", (chunk) => {
      body += chunk;
      if (body.length > 64 * 1024) {
        reject(new Error("Požadavek je příliš velký."));
        request.destroy();
      }
    });
    request.on("end", () => {
      try {
        resolve(JSON.parse(body));
      } catch {
        reject(new Error("Neplatná data požadavku."));
      }
    });
    request.on("error", reject);
  });
}

function readMultipart(request) {
  return new Promise((resolve, reject) => {
    let parser;
    try {
      parser = Busboy({
        headers: request.headers,
        limits: {
          files: MAX_UPLOAD_FILES,
          fileSize: MAX_UPLOAD_FILE_SIZE,
          fields: 8,
          fieldSize: 64 * 1024,
        },
      });
    } catch {
      reject(new Error("Očekávám multipart/form-data požadavek."));
      return;
    }

    const fields = Object.create(null);
    const uploads = [];
    let totalSize = 0;
    let failure;
    const fail = (message) => {
      if (!failure) failure = new Error(message);
    };

    parser.on("field", (name, value, info) => {
      if (!["year", "rateType", "clientName", "storedFiles"].includes(name)) {
        fail(`Neočekávané pole formuláře: ${name}.`);
        return;
      }
      if (Object.hasOwn(fields, name) || info.valueTruncated || info.nameTruncated) {
        fail("Formulář obsahuje duplicitní nebo příliš dlouhé pole.");
        return;
      }
      fields[name] = value;
    });

    parser.on("file", (fieldName, stream, info) => {
      if (fieldName !== "uploads") {
        fail("Neočekávaný souborový vstup.");
        stream.resume();
        return;
      }
      const name = path.basename(String(info.filename || "").replace(/\\/g, "/"));
      if (!name) {
        fail("Nahraný soubor nemá název.");
        stream.resume();
        return;
      }

      const chunks = [];
      let fileSize = 0;
      stream.on("data", (chunk) => {
        fileSize += chunk.length;
        totalSize += chunk.length;
        if (fileSize > MAX_UPLOAD_FILE_SIZE) {
          fail(`Soubor ${name} překročil limit 20 MB.`);
        } else if (totalSize > MAX_UPLOAD_TOTAL_SIZE) {
          fail("Souhrnná velikost nahraných souborů překročila limit 64 MB.");
        } else if (!failure) {
          chunks.push(chunk);
        }
      });
      stream.on("limit", () => fail(`Soubor ${name} překročil limit 20 MB.`));
      stream.on("end", () => {
        if (fileSize <= MAX_UPLOAD_FILE_SIZE) {
          uploads.push({ name, buffer: Buffer.concat(chunks) });
        }
      });
    });

    parser.on("filesLimit", () => fail(`Lze nahrát nejvýše ${MAX_UPLOAD_FILES} souborů.`));
    parser.on("fieldsLimit", () => fail("Formulář obsahuje příliš mnoho polí."));
    parser.on("error", () => reject(new Error("Nahrávání souborů se nepodařilo zpracovat.")));
    parser.on("finish", () => {
      if (failure) reject(failure);
      else resolve({ fields, uploads });
    });
    request.on("aborted", () => reject(new Error("Nahrávání bylo přerušeno.")));
    request.pipe(parser);
  });
}

function assertSameOrigin(request) {
  const originHeader = request.headers.origin;
  if (!originHeader) return;

  let origin;
  try {
    origin = new URL(originHeader);
  } catch {
    throw new Error("Neplatný původ požadavku.");
  }
  const allowedHosts = [
    request.headers.host,
    request.headers["x-forwarded-host"],
  ].filter(Boolean).map((host) => String(host).toLowerCase());
  if (!allowedHosts.includes(origin.host.toLowerCase())) {
    throw new Error("Požadavek musí pocházet ze stejného webu.");
  }
}

function listSourceFiles() {
  return fs.readdirSync(dataDir)
    .filter((name) => name.toLowerCase().endsWith(".xlsx") && name !== "Fiscal Report.xlsx")
    .map((name) => {
      const match = name.match(/^([A-Za-z0-9]+)_(\d+)_.*\.xlsx$/);
      return {
        name,
        currency: match ? match[1].toUpperCase() : name.split("_")[0].toUpperCase(),
        account: match ? match[2] : "",
      };
    });
}

function sendStatic(response, pathname) {
  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.resolve(publicDir, `.${requestedPath}`);
  if (!filePath.startsWith(`${publicDir}${path.sep}`) || !fs.existsSync(filePath)) {
    response.writeHead(404);
    response.end("Not found");
    return;
  }

  const extension = path.extname(filePath).toLowerCase();
  const contentTypes = {
    ".css": "text/css; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
  };
  response.writeHead(200, {
    "Content-Type": contentTypes[extension] || "application/octet-stream",
  });
  fs.createReadStream(filePath).pipe(response);
}

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url, `http://${request.headers.host || "localhost"}`);

  try {
    if (request.method === "POST") assertSameOrigin(request);

    if (request.method === "GET" && url.pathname === "/api/metadata") {
      const rates = readRates();
      const years = new Set([
        ...Object.keys(rates.years || {}),
        ...Object.keys(rates.annualRates || {}),
      ]);
      sendJson(response, 200, {
        files: listSourceFiles(),
        years: [...years].sort(),
        annualRates: rates.annualRates || {},
        currencies: [...new Set(Object.values(rates.years || {})
          .flatMap((year) => Object.keys(year.rates || {})))].sort(),
        eodhdConfigured: Boolean(eodhdApiKey),
      });
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/settings/eodhd") {
      const input = await readJson(request);
      if (typeof input.apiKey !== "string" || input.apiKey.length > 512) {
        sendJson(response, 400, { error: "Zadejte platný EODHD API klíč." });
        return;
      }
      eodhdApiKey = input.apiKey.trim();
      sendJson(response, 200, { configured: Boolean(eodhdApiKey) });
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/rates") {
      const input = await readJson(request);
      const year = String(input.year || "");
      const currency = String(input.currency || "").trim().toUpperCase();
      const type = input.type;
      const value = Number(input.rate);
      if (!/^\d{4}$/.test(year) || !/^[A-Z0-9]{3,8}$/.test(currency)
          || !Number.isFinite(value) || value <= 0 || value > 1_000_000) {
        sendJson(response, 400, { error: "Zadejte platný rok, měnu a kladný kurz." });
        return;
      }

      const rates = readRates();
      if (type === "MF") {
        rates.annualRates ||= {};
        rates.annualRates[year] ||= {};
        rates.annualRates[year][currency] = value;
      } else if (type === "CNB") {
        const date = String(input.date || "");
        if (!new RegExp(`^${year}-\\d{2}-\\d{2}$`).test(date)
            || Number.isNaN(Date.parse(`${date}T00:00:00Z`))) {
          sendJson(response, 400, { error: "U kurzu ČNB zadejte platné datum v daném roce." });
          return;
        }
        rates.years ||= {};
        rates.years[year] ||= { checkedAt: new Date().toISOString(), rates: {} };
        rates.years[year].rates ||= {};
        rates.years[year].rates[currency] ||= {};
        rates.years[year].rates[currency][date] = value;
      } else {
        sendJson(response, 400, { error: "Neznámý typ kurzu." });
        return;
      }
      writeRates(rates);
      sendJson(response, 200, { ok: true });
      return;
    }

    if (request.method === "POST" && url.pathname === "/api/report") {
      const { fields, uploads } = await readMultipart(request);
      const availableFiles = new Set(listSourceFiles().map((file) => file.name));
      let storedFiles;
      try {
        storedFiles = JSON.parse(fields.storedFiles || "[]");
      } catch {
        throw new Error("Seznam uložených souborů je neplatný.");
      }
      if (!Array.isArray(storedFiles) || storedFiles.some((file) => typeof file !== "string")) {
        throw new Error("Seznam uložených souborů je neplatný.");
      }
      if (storedFiles.some((file) => !availableFiles.has(file))) {
        sendJson(response, 400, { error: "Vybraný uložený soubor neexistuje." });
        return;
      }
      if (!storedFiles.length && !uploads.length) {
        sendJson(response, 400, { error: "Vyberte nebo nahrajte alespoň jeden XLSX zdrojový soubor." });
        return;
      }
      if (storedFiles.length + uploads.length > MAX_UPLOAD_FILES) {
        sendJson(response, 400, { error: `Lze zpracovat nejvýše ${MAX_UPLOAD_FILES} souborů najednou.` });
        return;
      }
      const selectedNames = new Set(storedFiles);
      for (const upload of uploads) {
        if (!/^[A-Za-z0-9]+_.+\.xlsx$/i.test(upload.name)) {
          sendJson(response, 400, {
            error: `Název souboru ${upload.name} musí začínat měnou a obsahovat podtržítko (např. USD_...).`,
          });
          return;
        }
        if (selectedNames.has(upload.name)) {
          sendJson(response, 400, { error: `Soubor ${upload.name} je vybraný vícekrát.` });
          return;
        }
        selectedNames.add(upload.name);
      }
      const year = String(fields.year || "");
      if (!/^\d{4}$/.test(year) || !["CNB", "MF"].includes(fields.rateType)) {
        sendJson(response, 400, { error: "Zadejte platný rok a typ kurzu." });
        return;
      }
      const clientName = String(fields.clientName || "").trim();
      if (clientName.length > 100) {
        sendJson(response, 400, { error: "Jméno může mít nejvýše 100 znaků." });
        return;
      }

      const output = await buildReport({
        dataDir,
        files: [...storedFiles, ...uploads],
        year,
        rateType: fields.rateType,
        clientName,
        rates: readRates(),
        isinApiKey: eodhdApiKey,
      });
      const filename = `Fiscal-Report-${year}.xlsx`;
      response.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Content-Length": output.length,
      });
      response.end(output);
      return;
    }

    if (request.method === "GET" && url.pathname.startsWith("/")) {
      sendStatic(response, decodeURIComponent(url.pathname));
      return;
    }

    response.writeHead(405);
    response.end("Method not allowed");
  } catch (error) {
    sendJson(response, 400, { error: error.message || "Zpracování se nezdařilo." });
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Fiscal Report Builder běží na http://127.0.0.1:${port}`);
});
