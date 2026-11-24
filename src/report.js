const path = require("node:path");
const ExcelJS = require("exceljs");
const { resolveIsinCache } = require("./isin");

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DATE_FORMAT = "dd.mm.yyyy hh:mm";
const RATE_FILL = "FFDDEBF7";
const RESULT_FILL = "FFE2F0D9";
const RESULT_FONT = "FF215E21";
const SECTION_GAP_ROWS = 3;
const SECTION_TITLES = [
  "Přehled zisků a ztrát na CFD (Profit Loss Summary CFD)",
  "Přehled zisků a ztrát na STC/ETF (Profit Loss Summary STC)",
  "Přehled zisků a ztrát na STC (Profit Loss Summary STC)",
  "Přehled zisků a ztrát na Frakčních právech (Profit Loss Summary Fractional Rights)",
  "Přehled zisků a ztrát na CFD (Profit Loss Summary CFD)",
  "Souhrn zisku a ztráty s úpravami (Profit Loss Summary with adjustments)",
  "Seznam transakcí na STC/ETF (List of Stock Transactions)",
  "Seznam transakcí s Frakčními právy (List of Fractional Shares Transactions)",
  "Seznam CFD transakcí (List of CFD Transactions)",
  "Seznam dividend a dalších peněžních operací (List of Dividends and other cash operations)",
  "Seznam poplatků, bonusů, úrokových sazeb peněžních toků (List of Fees, Bonuses, Interest Rates cashflows)",
  "Úpravy transakcí (Transaction adjustments)",
  "Souhrn peněžních toků v rámci poplatků, prémií, úroků (Fees, Bonuses, Interest Rates cashflows Summary)",
  "Seznam otevřených obchodů (List of Open Trades)",
];
const ISIN_COUNTRY_NAMES = {
  AU: "Austrálie",
  BE: "Belgie",
  CA: "Kanada",
  CH: "Švýcarsko",
  DE: "Německo",
  DK: "Dánsko",
  ES: "Španělsko",
  FI: "Finsko",
  FR: "Francie",
  GB: "Spojené království",
  HK: "Hongkong",
  IE: "Irsko",
  IT: "Itálie",
  JP: "Japonsko",
  LU: "Lucembursko",
  NL: "Nizozemsko",
  NO: "Norsko",
  SE: "Švédsko",
  SG: "Singapur",
  US: "Spojené státy",
};
const COLUMN_CURRENCY = {
  AU: "AUD", US: "USD", UK: "GBP", DE: "EUR", FR: "EUR", ES: "EUR",
  IT: "EUR", NL: "EUR", BE: "EUR", PT: "EUR", IE: "EUR", CZ: "CZK",
  PL: "PLN", CH: "CHF", SE: "SEK", NO: "NOK", DK: "DKK", CA: "CAD",
  HK: "HKD", JP: "JPY", SG: "SGD", ZA: "ZAR", HU: "HUF", TR: "TRY",
};

function number(value) {
  if (value === null || value === undefined || value === "") return 0;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function roundMoney(value) {
  const scaled = Math.abs(value) * 100;
  const rounded = Math.round(scaled + Number.EPSILON * Math.max(1, scaled));
  if (rounded === 0) return 0;
  return Math.sign(value) * rounded / 100;
}

function dateValue(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number") return new Date(EXCEL_EPOCH + value * 86_400_000);
  if (typeof value === "string" && value) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function dateKey(value) {
  const date = dateValue(value);
  return date ? date.toISOString().slice(0, 10) : "";
}

function rowObject(row, headerRow) {
  const result = {};
  headerRow.eachCell((cell, column) => {
    const header = String(cell.value || "").trim();
    if (header) result[header] = row.getCell(column).value;
  });
  return result;
}

async function loadSource(file, dataDir) {
  const workbook = new ExcelJS.Workbook();
  const fileName = typeof file === "string"
    ? path.basename(file)
    : path.basename(String(file.name || "").replace(/\\/g, "/"));
  if (typeof file === "string") {
    await workbook.xlsx.readFile(path.join(dataDir, file));
  } else if (Buffer.isBuffer(file.buffer)) {
    await workbook.xlsx.load(file.buffer);
  } else {
    throw new Error(`Soubor ${fileName || "bez názvu"} nemá platná data XLSX.`);
  }
  const closedSheet = workbook.getWorksheet("Closed Positions");
  const cashSheet = workbook.getWorksheet("Cash Operations");
  const openSheet = workbook.getWorksheet("Open Positions");
  if (!closedSheet || !cashSheet || !openSheet) {
    throw new Error(`Soubor ${fileName} neobsahuje očekávané listy XTB.`);
  }

  const accountCurrency = fileName.match(/^([A-Za-z0-9]+)_/)?.[1]?.toUpperCase();
  if (!accountCurrency) {
    throw new Error(`Z názvu souboru ${fileName} nelze určit měnu účtu.`);
  }
  const account = String(closedSheet.getCell("B1").value || "");
  const closedHeaders = closedSheet.getRow(5);
  const cashHeaders = cashSheet.getRow(5);
  const openHeaders = openSheet.getRow(9);

  const valuesFrom = (sheet, headerRow, firstDataRow) => {
    const values = [];
    for (let rowNumber = firstDataRow; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const record = rowObject(row, headerRow);
      if (Object.values(record).every((value) => value === null || value === undefined || value === "")) continue;
      if (String(Object.values(record)[0]).toLowerCase() === "total") continue;
      values.push(record);
    }
    return values;
  };

  return {
    account,
    accountCurrency,
    closed: valuesFrom(closedSheet, closedHeaders, 6),
    cash: valuesFrom(cashSheet, cashHeaders, 6),
    open: valuesFrom(openSheet, openHeaders, 10),
  };
}

function currencyForTicker(ticker, fallback) {
  const suffix = String(ticker || "").match(/\.([A-Z]{2,3})$/)?.[1];
  return (suffix && COLUMN_CURRENCY[suffix]) || fallback;
}

function dividendPerShare(comment) {
  const match = String(comment || "").match(/\b([A-Z]{3})\s+(-?\d+(?:[.,]\d+)?)\s*\/\s*SHR\b/i);
  return match
    ? { currency: match[1].toUpperCase(), amount: Number(match[2].replace(",", ".")) }
    : null;
}

function sharesHeldAt(source, ticker, timestamp, positionId) {
  const normalizedTicker = String(ticker || "").toUpperCase();
  const eventDate = dateValue(timestamp);
  if (!eventDate || !Number.isFinite(eventDate.getTime()) || !normalizedTicker) return 0;
  const normalizedPositionId = String(positionId || "");
  const eventTime = eventDate.getTime();
  let shares = 0;

  for (const row of [...source.closed, ...source.open]) {
    if (String(row.Ticker || "").toUpperCase() !== normalizedTicker) continue;
    if (normalizedPositionId && String(row["Position ID"] || "") !== normalizedPositionId) continue;
    const openTime = dateValue(row["Open Time (UTC)"] || row["Open time (UTC)"])?.getTime();
    if (!openTime || openTime > eventTime) continue;
    const closeTime = dateValue(row["Close Time (UTC)"])?.getTime();
    if (closeTime && closeTime <= eventTime) continue;
    const direction = String(row.Type || "").toUpperCase() === "SELL" ? -1 : 1;
    shares += direction * number(row.Volume);
  }
  return shares;
}

function sharesHeldNearDate(source, ticker, timestamp, positionId, maxDays = 45) {
  const eventTime = dateValue(timestamp)?.getTime();
  if (!Number.isFinite(eventTime)) return null;
  for (let daysBefore = 0; daysBefore <= maxDays; daysBefore += 1) {
    const date = new Date(eventTime - daysBefore * 86_400_000);
    const shares = sharesHeldAt(source, ticker, date, positionId);
    if (shares !== 0) return { shares, daysBefore };
  }
  return null;
}

function deriveDividendAmounts(dividends, taxes, year, rates) {
  for (const row of dividends) {
    const perShare = dividendPerShare(row.Comment);
    if (!perShare) continue;
    row.originalCurrency = perShare.currency;
    const holding = sharesHeldNearDate(row.source, row.Ticker, row.Time, row["Position ID"]);
    let shares = holding ? Math.abs(holding.shares) : 0;
    if (!holding) {
      const date = dateValue(row.Time);
      const accountRate = fxFor(row.source.accountCurrency, date, "CNB", year, rates);
      const currencyRate = fxFor(perShare.currency, date, "CNB", year, rates);
      const accountAmountInCurrency = number(row.Amount) * accountRate.value / currencyRate.value;
      shares = Math.abs(accountAmountInCurrency / perShare.amount);
    }
    row.originalAmountOverride = Math.sign(number(row.Amount) || 1) * shares * perShare.amount;
    row.originalAmountNote = holding
      ? `${perShare.amount} ${perShare.currency} × ${shares} ks${holding.daysBefore ? ` (${holding.daysBefore} dní před výplatou)` : ""}`
      : `${perShare.amount} ${perShare.currency} × ${shares.toFixed(4)} ks (odhad z XTB)`;
  }

  for (const row of taxes) {
    const percent = String(row.Comment || "").match(/\bWHT\s+(\d+(?:[.,]\d+)?)\s*%/i);
    if (!percent) continue;
    const eventDate = dateKey(row.Time);
    const candidates = dividends.filter((dividend) =>
      dividend.originalAmountOverride !== undefined
      && dividend.source === row.source
      && String(dividend.Ticker || "").toUpperCase() === String(row.Ticker || "").toUpperCase()
      && String(dividend["Position ID"] || "") === String(row["Position ID"] || "")
      && dateKey(dividend.Time) === eventDate);
    if (!candidates.length) continue;
    const timestamp = dateValue(row.Time)?.getTime() || 0;
    const pairedDividend = candidates
      .sort((left, right) => Math.abs((dateValue(left.Time)?.getTime() || 0) - timestamp)
        - Math.abs((dateValue(right.Time)?.getTime() || 0) - timestamp))[0];
    const percentage = Number(percent[1].replace(",", ".")) / 100;
    row.originalCurrency = pairedDividend.originalCurrency;
    row.originalAmountOverride = -Math.abs(pairedDividend.originalAmountOverride) * percentage;
    row.originalAmountNote = `WHT ${percentage * 100}% × (${pairedDividend.originalAmountNote})`;
  }
}

function fxFor(currency, date, rateType, reportYear, rates) {
  if (currency === "CZK") return { value: 1, date: dateKey(date) };
  if (rateType === "MF") {
    const value = number(rates.annualRates?.[reportYear]?.[currency]);
    if (!value) {
      throw new Error(`Chybí roční kurz MF ${currency} pro ${reportYear}. Doplňte jej v části Kurzy.`);
    }
    return { value, date: reportYear };
  }

  const key = dateKey(date);
  const eventYear = key.slice(0, 4);
  const dailyRates = rates.years?.[eventYear]?.rates?.[currency];
  if (!dailyRates) {
    throw new Error(`Chybí denní kurz ČNB ${currency} pro rok ${eventYear}. Doplňte jej v části Kurzy.`);
  }
  const applicableDate = Object.keys(dailyRates).filter((day) => day <= key).sort().pop();
  if (!applicableDate) {
    throw new Error(`Chybí kurz ČNB ${currency} k datu ${key} i předchozímu pracovnímu dni.`);
  }
  return {
    value: number(dailyRates[applicableDate]),
    date: applicableDate,
  };
}

function findRow(sheet, text) {
  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    for (let column = 1; column <= row.cellCount; column += 1) {
      if (String(row.getCell(column).value || "").trim() === text) return rowNumber;
    }
  }
  return -1;
}

function clearSectionTailColumn(sheet, title, headerOffset, column) {
  const titleRow = findRow(sheet, title);
  if (titleRow < 0) return;
  const nextSectionRow = SECTION_TITLES
    .map((sectionTitle) => findRow(sheet, sectionTitle))
    .filter((rowNumber) => rowNumber > titleRow)
    .reduce((nearest, rowNumber) => Math.min(nearest, rowNumber), sheet.rowCount + 1);
  for (let rowNumber = titleRow + headerOffset; rowNumber < nextSectionRow; rowNumber += 1) {
    const cell = sheet.getRow(rowNumber).getCell(column);
    cell.value = null;
    cell.border = {};
    setFill(cell, "FFFFFFFF");
  }
}

function ensureSectionSpacing(sheet) {
  const sectionRows = [...new Set(SECTION_TITLES)]
    .map((title) => ({ title, rowNumber: findRow(sheet, title) }))
    .filter((section) => section.rowNumber > 0)
    .sort((left, right) => right.rowNumber - left.rowNumber);

  for (const section of sectionRows) {
    let titleRow = section.rowNumber;
    let blankRows = 0;
    for (let rowNumber = titleRow - 1; rowNumber > 0; rowNumber -= 1) {
      const isBlank = sheet.getRow(rowNumber).values.slice(1)
        .every((value) => value === null || value === undefined || value === "");
      if (!isBlank) break;
      blankRows += 1;
    }
    const rowsToInsert = Math.max(0, SECTION_GAP_ROWS - blankRows);
    if (rowsToInsert) {
      sheet.spliceRows(titleRow, 0, ...Array.from({ length: rowsToInsert }, () => []));
      titleRow += rowsToInsert;
    } else if (blankRows > SECTION_GAP_ROWS) {
      const rowsToRemove = blankRows - SECTION_GAP_ROWS;
      sheet.spliceRows(titleRow - blankRows, rowsToRemove);
      titleRow -= rowsToRemove;
    }

    for (let rowNumber = titleRow - SECTION_GAP_ROWS; rowNumber < titleRow; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      for (let column = 1; column <= Math.max(row.cellCount, sheet.columnCount, 23); column += 1) {
        row.getCell(column).border = {};
      }
    }
  }
}

function isinForTicker(ticker, isinCache) {
  return isinCache[String(ticker || "").trim().toUpperCase()] || "US";
}

function countryFromIsin(isin) {
  const code = String(isin || "").slice(0, 2).toUpperCase() || "US";
  return `${ISIN_COUNTRY_NAMES[code] || code} (${code})`;
}

function cloneStyle(source, target) {
  if (source.style) target.style = JSON.parse(JSON.stringify(source.style));
  if (source.numFmt) target.numFmt = source.numFmt;
}

function setNumberFormat(cell, format) {
  const style = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
  style.numFmt = format;
  cell.style = style;
}

function setFill(cell, color, emphasize = false) {
  const style = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
  style.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
  if (emphasize) {
    style.font = { ...(style.font || {}), bold: true, color: { argb: RESULT_FONT } };
  }
  cell.style = style;
}

function setFontColor(cell, color) {
  const style = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
  style.font = { ...(style.font || {}), color: { argb: color } };
  cell.style = style;
}

function clearFill(cell, color) {
  if (cell.fill?.fgColor?.argb !== color) return;
  const style = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
  delete style.fill;
  cell.style = style;
}

function isSummaryTotalRow(values) {
  return values.some((value) => typeof value === "string"
    && /^(Celkem|Součet|Total)\b/i.test(value.trim()));
}

function moveSummaryTotalsLast(rows) {
  const totalRows = rows.filter(isSummaryTotalRow);
  if (!totalRows.length) return;
  rows.splice(0, rows.length, ...rows.filter((row) => !isSummaryTotalRow(row)), ...totalRows);
}

function writeSection(
  sheet,
  title,
  headerOffset,
  values,
  dateColumns = [],
  numberFormats = {},
  rateColumns = [],
  resultColumns = [],
  normalizeRowStyles = false,
) {
  const titleRow = findRow(sheet, title);
  if (titleRow < 0) throw new Error(`V šabloně chybí sekce „${title}“.`);
  const headerRowNumber = titleRow + headerOffset;
  const headerRow = sheet.getRow(headerRowNumber);
  const firstDataRow = headerRowNumber + 1;
  let nextSectionRow = sheet.rowCount + 1;
  for (const candidate of SECTION_TITLES) {
    const found = findRow(sheet, candidate);
    if (found > titleRow && found < nextSectionRow) nextSectionRow = found;
  }

  const capacity = Math.max(0, nextSectionRow - firstDataRow - SECTION_GAP_ROWS);
  const extraRows = Math.max(0, values.length - capacity);
  if (extraRows) {
    const insertAt = nextSectionRow;
    sheet.spliceRows(insertAt, 0, ...Array.from({ length: extraRows }, () => []));
    const styleSource = sheet.getRow(firstDataRow);
    for (let rowNumber = insertAt; rowNumber < insertAt + extraRows; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      for (let column = 1; column <= Math.max(headerRow.cellCount, 22); column += 1) {
        cloneStyle(styleSource.getCell(column), row.getCell(column));
      }
      row.height = styleSource.height;
    }
  }

  const styleSource = sheet.getRow(firstDataRow);
  if (normalizeRowStyles) {
    const hasTableStyle = (cell) => Boolean(
      Object.keys(cell.border || {}).length || cell.alignment,
    );
    const addedColumns = [...new Set([...rateColumns, ...resultColumns])];
    for (let rowNumber = firstDataRow; rowNumber < firstDataRow + values.length; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      for (let column = 1; column <= Math.max(headerRow.cellCount, 22); column += 1) {
        cloneStyle(styleSource.getCell(column), row.getCell(column));
      }
      for (const column of addedColumns) {
        if (hasTableStyle(styleSource.getCell(column))) continue;
        let sourceColumn = column - 1;
        while (sourceColumn > 0 && !hasTableStyle(styleSource.getCell(sourceColumn))) {
          sourceColumn -= 1;
        }
        if (sourceColumn > 0) cloneStyle(styleSource.getCell(sourceColumn), row.getCell(column));
      }
    }
  }

  const endRow = nextSectionRow + extraRows - 1;
  for (let rowNumber = firstDataRow; rowNumber <= endRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    for (let column = 1; column <= Math.max(row.cellCount, 23); column += 1) {
      row.getCell(column).value = null;
    }
  }
  values.forEach((value, index) => {
    const row = sheet.getRow(firstDataRow + index);
    value.forEach((cellValue, cellIndex) => {
      row.getCell(cellIndex + 2).value = cellValue;
    });
    dateColumns.forEach((column) => {
      setNumberFormat(row.getCell(column), DATE_FORMAT);
    });
    Object.entries(numberFormats).forEach(([column, format]) => {
      setNumberFormat(row.getCell(Number(column)), format);
    });
    rateColumns.forEach((column) => setFill(row.getCell(column), RATE_FILL));
    resultColumns.forEach((column) => setFill(row.getCell(column), RESULT_FILL, true));
  });
  return firstDataRow;
}

async function buildReport({
  dataDir,
  files,
  year,
  rateType,
  clientName = "",
  rates,
  isinApiKey = process.env.EODHD_API_KEY,
}) {
  const sources = await Promise.all(files.map((file) => loadSource(file, dataDir)));
  const closed = sources.flatMap((source) => source.closed.map((row) => ({ ...row, source })));
  const cash = sources.flatMap((source) => source.cash.map((row) => ({ ...row, source })));
  const open = sources.flatMap((source) => source.open.map((row) => ({ ...row, source })));
  const inYear = (date) => dateKey(date).startsWith(`${year}-`);
  const trades = closed.filter((row) => inYear(row["Close Time (UTC)"]));
  const cashRows = cash.filter((row) => inYear(row.Time));
  const openRows = open.filter((row) => dateValue(row["Open time (UTC)"]));
  const cfdRows = [];
  const stockRows = [];
  const fractionalRows = [];

  for (const row of trades) {
    const category = String(row.Category || "").toUpperCase();
    const product = String(row.Product || "").toUpperCase();
    if (category.includes("CFD") || product.includes("CFD")) cfdRows.push(row);
    else if (category.includes("FRACTION") || category.includes("RIGHT")) fractionalRows.push(row);
    else stockRows.push(row);
  }

  const decorateTrade = (row) => {
    const accountCurrency = row.source.accountCurrency;
    const closeDate = dateValue(row["Close Time (UTC)"]);
    const openDate = dateValue(row["Open Time (UTC)"]);
    const purchase = number(row["Purchase Value"]);
    const sale = number(row["Sale Value"]);
    const commission = number(row.Commission);
    const ticker = String(row.Ticker || "");
    const instrumentCurrency = currencyForTicker(ticker, accountCurrency);
    const instrumentPurchase = number(row["Open Conversion Rate"])
      ? purchase / number(row["Open Conversion Rate"])
      : purchase;
    const instrumentSale = number(row["Close Conversion Rate"])
      ? sale / number(row["Close Conversion Rate"])
      : sale;
    const purchaseRate = fxFor(instrumentCurrency, openDate, rateType, year, rates);
    const saleRate = fxFor(instrumentCurrency, closeDate, rateType, year, rates);
    const accountSaleRate = fxFor(accountCurrency, closeDate, rateType, year, rates);
    const purchaseCzk = roundMoney(instrumentPurchase * purchaseRate.value);
    const saleCzk = roundMoney(instrumentSale * saleRate.value);
    const commissionCzk = roundMoney(commission * accountSaleRate.value);
    const netCzk = roundMoney(saleCzk - purchaseCzk + commissionCzk);
    return {
      row,
      purchaseRate,
      saleRate,
      accountSaleRate,
      purchaseCzk,
      saleCzk,
      commissionCzk,
      netCzk,
      instrumentCurrency,
      instrumentPurchase,
      instrumentSale,
      commission,
      purchase,
      sale,
    };
  };
  const preparedStocks = stockRows.map(decorateTrade);
  const preparedFractionals = fractionalRows.map(decorateTrade);

  const getCashType = (row) => String(row.Type || "").toLowerCase();
  const dividends = cashRows.filter((row) => getCashType(row).includes("dividend"));
  const taxes = cashRows.filter((row) => getCashType(row).includes("withholding tax"));
  deriveDividendAmounts(dividends, taxes, year, rates);
  const isinCache = await resolveIsinCache(
    [...dividends, ...taxes].map((row) => row.Ticker),
    { dataDir, apiKey: isinApiKey },
  );
  const transfers = cashRows.filter((row) => getCashType(row).includes("transfer")
    || getCashType(row).includes("adjustment")
    || getCashType(row).includes("deposit"));
  const closedPositionIds = new Set(trades.map((row) => String(row["Position ID"] || "")));
  const excludedTypes = [
    "stock purchase", "stock sell", "dividend", "withholding tax",
    "transfer", "adjustment", "deposit", "close trade",
  ];
  const fees = cashRows.filter((row) => {
    const type = getCashType(row);
    if (excludedTypes.some((excluded) => type.includes(excluded))) return false;
    if (type === "swap" && closedPositionIds.has(String(row["Position ID"] || ""))) return false;
    return true;
  });
  const cashValue = (row) => {
    const source = row.source;
    const date = dateValue(row.Time);
    const type = getCashType(row);
    const useInstrumentCurrency = type.includes("dividend")
      || type.includes("withholding tax")
      || (type.includes("fee") && row.Ticker);
    const currency = useInstrumentCurrency
      ? row.originalCurrency || currencyForTicker(row.Ticker, source.accountCurrency)
      : source.accountCurrency;
    const amount = number(row.Amount);
    const accountCnb = fxFor(source.accountCurrency, date, "CNB", year, rates);
    const currencyCnb = fxFor(currency, date, "CNB", year, rates);
    const originalAmount = Number.isFinite(row.originalAmountOverride)
      ? row.originalAmountOverride
      : amount * accountCnb.value / currencyCnb.value;
    const fx = fxFor(currency, date, rateType, year, rates);
    const conversionRate = originalAmount === 0
      ? accountCnb.value / currencyCnb.value
      : amount / originalAmount;
    return {
      amount,
      currency,
      originalAmount,
      conversionRate,
      fx,
      note: row.originalAmountNote
        || (useInstrumentCurrency && currency !== source.accountCurrency
          ? "Částka odhadnuta z XTB v měně účtu"
          : ""),
      czk: roundMoney(originalAmount * fx.value),
    };
  };

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(dataDir, "Fiscal Report.xlsx"));
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error("Šablona Fiscal Report.xlsx neobsahuje list.");
  sheet.getCell("C2").value = `Fiskální report (Fiscal Report): ${year}/01/01 - ${year}/12/31`;
  sheet.getCell("C4").value = clientName;
  sheet.getCell("B34").value = null;
  sheet.getCell("B8").value = "Původní částky a kurzy jsou dle XTB; modře jsou kurzy ČNB/MF.";
  sheet.getCell("E8").value = "Zeleně jsou částky přepočtené kurzem ČNB/MF do CZK.";
  setFill(sheet.getCell("B8"), RATE_FILL);
  setFill(sheet.getCell("E8"), RESULT_FILL, true);

  const totals = { cfd: 0, stock: 0, fractional: 0, fees: 0, dividends: 0, taxes: 0 };
  const accountSummary = (rows, prepared, category) => {
    const map = new Map();
    rows.forEach((row, index) => {
      const account = row.source;
      const key = `${account.account}|${account.accountCurrency}`;
      const entry = map.get(key) || {
        account, net: 0, loss: 0, profit: 0, accountNet: 0, sale: 0,
        purchase: 0, commission: 0, gross: 0, instrumentCommission: 0, instrumentSale: 0,
        instrumentPurchase: 0, instrumentGross: 0,
      };
      const accountAmount = number(row["Profit/Loss"]);
      const czk = category === "cfd"
        ? roundMoney(accountAmount
          * fxFor(account.accountCurrency, row["Close Time (UTC)"], rateType, year, rates).value)
        : prepared[index].netCzk;
      entry.net += czk;
      if (czk < 0) entry.loss += Math.abs(czk);
      else entry.profit += czk;
      entry.accountNet += accountAmount;
      if (category !== "cfd") {
        const item = prepared[index];
        entry.sale += item.sale;
        entry.purchase += item.purchase;
        entry.commission += item.commission;
        entry.gross += item.sale - item.purchase;
        const instrumentRate = number(row["Close Conversion Rate"]) || 1;
        entry.instrumentCommission += item.commission / instrumentRate;
        entry.instrumentSale += item.instrumentSale;
        entry.instrumentPurchase += item.instrumentPurchase;
        entry.instrumentGross += item.instrumentSale - item.instrumentPurchase;
      }
      map.set(key, entry);
    });
    return [...map.values()];
  };

  const cfdSummary = accountSummary(cfdRows, [], "cfd");
  const stockSummary = accountSummary(stockRows, preparedStocks, "stock");
  const fractionalSummary = accountSummary(fractionalRows, preparedFractionals, "fractional");
  totals.cfd = cfdSummary.reduce((sum, item) => sum + item.net, 0);
  totals.stock = stockSummary.reduce((sum, item) => sum + item.net, 0);
  totals.fractional = fractionalSummary.reduce((sum, item) => sum + item.net, 0);
  totals.fees = fees.reduce((sum, row) => sum + cashValue(row).czk, 0);
  totals.dividends = dividends.reduce((sum, row) => sum + cashValue(row).czk, 0);
  totals.taxes = taxes.reduce((sum, row) => sum + cashValue(row).czk, 0);

  const sumByAccountCurrency = (rows, amountFor) => rows.reduce((totalsByCurrency, row, index) => {
    const column = row.source.accountCurrency === "CZK" ? 0 : 1;
    totalsByCurrency[column] += amountFor(row, index);
    return totalsByCurrency;
  }, [0, 0]);
  const summaryRows = [
    sumByAccountCurrency(cfdRows, (row) => roundMoney(number(row["Profit/Loss"])
      * fxFor(row.source.accountCurrency, row["Close Time (UTC)"], rateType, year, rates).value)),
    sumByAccountCurrency(stockRows, (row, index) => preparedStocks[index].netCzk),
    sumByAccountCurrency(fractionalRows, (row, index) => preparedFractionals[index].netCzk),
    sumByAccountCurrency(fees, (row) => cashValue(row).czk),
    sumByAccountCurrency(dividends, (row) => cashValue(row).czk),
    sumByAccountCurrency(taxes, (row) => cashValue(row).czk),
  ].map(([czkAccounts, foreignCurrencyAccounts]) => [
    czkAccounts,
    foreignCurrencyAccounts,
    czkAccounts + foreignCurrencyAccounts,
  ]);
  const totalCzkAccounts = summaryRows.reduce((sum, row) => sum + row[0], 0);
  const totalForeignCurrencyAccounts = summaryRows.reduce((sum, row) => sum + row[1], 0);
  summaryRows.push([
    totalCzkAccounts,
    totalForeignCurrencyAccounts,
    totalCzkAccounts + totalForeignCurrencyAccounts,
  ]);
  sheet.getCell("C12").value = "CZK účty (výsledek v CZK)";
  sheet.getCell("D12").value = "Cizí měna (přepočteno do CZK)";
  sheet.getCell("E12").value = "Celkem (CZK)";
  summaryRows.forEach((row, index) => {
    [3, 4, 5].forEach((column, bucket) => {
      sheet.getCell(13 + index, column).value = row[bucket];
      setNumberFormat(sheet.getCell(13 + index, column), '#,##0.00;[Red]-#,##0.00');
      setFill(sheet.getCell(13 + index, column), RESULT_FILL, true);
    });
  });
  sheet.getCell("D34").value = "CFD";
  sheet.getCell("D35").value = "Celkem v: CZK (Total in: CZK)";
  sheet.getCell("E35").value = totals.cfd;

  const writeAccountSummary = (records, productType) => records.map((record) => [
      record.account.account,
      record.account.accountCurrency,
      productType,
      record.accountNet,
      record.accountCurrency,
      record.sale,
      record.purchase,
      record.commission,
      record.gross,
      record.accountNet,
      record.accountCurrency,
      record.instrumentSale,
      record.instrumentPurchase,
      record.instrumentCommission,
      record.instrumentGross,
      record.instrumentGross + record.instrumentCommission,
      record.net,
    ]);
  const cfdSummaryRows = cfdSummary.map((record) => [
    record.account.account, record.account.accountCurrency, "CFD",
    record.loss, record.profit, record.net, record.net,
  ]);
  cfdSummaryRows.push([null, null, "Celkem v: CZK (Total in: CZK)", totals.cfd, null, null, totals.cfd]);
  const stockSummaryRows = writeAccountSummary(stockSummary, "STC");
  const fractionalSummaryRows = writeAccountSummary(fractionalSummary, "Fractional Rights");
  const appendConvertedTotal = (rows, total) => {
    if (!rows.length) return;
    const totalRow = Array(rows[0].length).fill(null);
    totalRow[0] = "Celkem v CZK";
    totalRow[totalRow.length - 1] = total;
    rows.push(totalRow);
  };
  appendConvertedTotal(stockSummaryRows, totals.stock);
  appendConvertedTotal(fractionalSummaryRows, totals.fractional);

  const cfdDetailRows = cfdRows.map((row) => {
    const fx = fxFor(row.source.accountCurrency, row["Close Time (UTC)"], rateType, year, rates);
    const profit = number(row["Profit/Loss"]);
    const net = number(row["Profit/Loss"]);
    return [
      row.source.account, row.source.accountCurrency, String(row["Position ID"] || ""),
      row.Product, dateValue(row["Open Time (UTC)"]), row.Type, number(row.Volume),
      row.Ticker, dateValue(row["Close Time (UTC)"]), number(row.Commission),
      number(row.Swap), number(row["Gross Profit"]) || profit, fx.value,
      roundMoney(net * fx.value),
    ];
  });
  const stockDetailRows = preparedStocks.map((item) => {
    const {
      row, purchaseRate, saleRate, accountSaleRate, purchaseCzk, saleCzk, commissionCzk, netCzk,
    } = item;
    const holdDays = Math.max(0, Math.floor((dateValue(row["Close Time (UTC)"]) - dateValue(row["Open Time (UTC)"])) / 86_400_000));
    return [
      row.source.account, row.source.accountCurrency, String(row["Position ID"] || ""),
      row.Product, dateValue(row["Open Time (UTC)"]), row.Type, number(row.Volume), row.Ticker,
      item.instrumentCurrency, item.purchase, item.sale, dateValue(row["Close Time (UTC)"]),
      holdDays, item.commission, number(row["Gross Profit"]), number(row["Profit/Loss"]),
      number(row["Close Conversion Rate"]) || null, number(row["Open Conversion Rate"]) || null,
      purchaseRate.value, purchaseCzk, saleRate.value, saleCzk, accountSaleRate.value,
      commissionCzk, netCzk,
    ];
  });
  const fractionalDetailRows = preparedFractionals.map((item) => {
    const {
      row, purchaseRate, saleRate, accountSaleRate, purchaseCzk, saleCzk, commissionCzk, netCzk,
    } = item;
    const holdDays = Math.max(0, Math.floor((dateValue(row["Close Time (UTC)"]) - dateValue(row["Open Time (UTC)"])) / 86_400_000));
    return [
      row.source.account, row.source.accountCurrency, String(row["Position ID"] || ""),
      row.Product, dateValue(row["Open Time (UTC)"]), row.Type, number(row.Volume), row.Ticker,
      item.instrumentCurrency, item.purchase, item.sale, dateValue(row["Close Time (UTC)"]),
      holdDays, item.commission, number(row["Gross Profit"]), number(row["Profit/Loss"]),
      number(row["Close Conversion Rate"]) || null, number(row["Open Conversion Rate"]) || null,
      purchaseRate.value, purchaseCzk, saleRate.value, saleCzk, accountSaleRate.value,
      commissionCzk, netCzk,
    ];
  });

  const feesSummaryByAccountCurrency = [...fees.reduce((map, row) => {
    const value = cashValue(row);
    const key = `${row.source.account}|${value.currency}`;
    const item = map.get(key) || {
      account: row.source.account,
      currency: value.currency,
      amount: 0,
      czk: 0,
    };
    item.amount += value.originalAmount;
    item.czk += value.czk;
    map.set(key, item);
    return map;
  }, new Map()).values()];
  const feesSummaryRows = feesSummaryByAccountCurrency.map((item) => [
    item.account, item.currency, item.amount, "Přepočet po řádcích", item.czk,
  ]);
  appendConvertedTotal(feesSummaryRows, totals.fees);
  const adjustmentSummaryRows = transfers.map((row) => {
    const value = cashValue(row);
    return [
      row.source.account,
      row.source.accountCurrency,
      row.Type,
      value.amount,
      value.fx.value,
      value.czk,
    ];
  });
  adjustmentSummaryRows.push([
    null,
    null,
    "Celkem v: CZK (Total in: CZK)",
    null,
    null,
    transfers.reduce((sum, row) => sum + cashValue(row).czk, 0),
  ]);

  const sectionRows = [
    {
      title: "Seznam CFD transakcí (List of CFD Transactions)",
      offset: 2,
      normalizeRowStyles: true,
      dateColumns: [6, 10],
      numberFormats: {
        8: "0.00",
        11: '#,##0.00;[Red]-#,##0.00',
        12: '#,##0.00;[Red]-#,##0.00',
        13: '#,##0.00;[Red]-#,##0.00',
        14: "0.0000",
        15: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [14],
      resultColumns: [15],
      values: cfdDetailRows,
    },
    {
      title: "Seznam otevřených obchodů (List of Open Trades)",
      offset: 2,
      normalizeRowStyles: true,
      dateColumns: [6],
      numberFormats: {
        8: "0.00",
        10: '#,##0.00;[Red]-#,##0.00',
        11: '#,##0.00;[Red]-#,##0.00',
        12: '#,##0.00;[Red]-#,##0.00',
        13: '#,##0.00;[Red]-#,##0.00',
        14: "0.0000",
        15: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [14],
      resultColumns: [15],
      values: openRows.map((row) => {
        const fx = fxFor(row.source.accountCurrency, row["Open time (UTC)"], rateType, year, rates);
        const profit = number(row["Net Profit"]);
        return [
          row.source.account, row.source.accountCurrency, String(row["Position ID"] || ""),
          row.Product, dateValue(row["Open time (UTC)"]), row.Type, number(row.Volume), row.Ticker,
          number(row["Open Commission"]), number(row.Swap), number(row["Gross Profit"]),
          profit, fx.value, roundMoney(profit * fx.value),
        ];
      }),
    },
    {
      title: "Seznam poplatků, bonusů, úrokových sazeb peněžních toků (List of Fees, Bonuses, Interest Rates cashflows)",
      offset: 2,
      normalizeRowStyles: true,
      dateColumns: [5],
      numberFormats: {
        6: '#,##0.00;[Red]-#,##0.00',
        9: '#,##0.00;[Red]-#,##0.00',
        10: "0.0000",
        11: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [10],
      resultColumns: [11],
      values: fees.map((row) => {
        const value = cashValue(row);
        return [row.source.account, row.source.accountCurrency, row.Comment, dateValue(row.Time),
          value.amount, row.Type, value.currency, value.originalAmount,
          value.fx.value, value.czk];
      }),
    },
    {
      title: "Seznam dividend a dalších peněžních operací (List of Dividends and other cash operations)",
      offset: 2,
      normalizeRowStyles: true,
      dateColumns: [5],
      numberFormats: {
        6: '#,##0.00;[Red]-#,##0.00',
        7: '#,##0.00;[Red]-#,##0.00',
        9: "0.000",
        11: "0.0000",
        12: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [11],
      resultColumns: [12],
      values: [...dividends, ...taxes].map((row) => {
        const value = cashValue(row);
        return [row.source.account, row.source.accountCurrency, row.Comment, dateValue(row.Time),
          value.amount, value.originalAmount, isinForTicker(row.Ticker, isinCache), value.conversionRate,
          value.note, value.fx.value, value.czk];
      }),
    },
    {
      title: "Úpravy transakcí (Transaction adjustments)",
      offset: 2,
      normalizeRowStyles: true,
      dateColumns: [7],
      numberFormats: {
        8: '#,##0.00;[Red]-#,##0.00',
        10: '#,##0.00;[Red]-#,##0.00',
        11: "0.0000",
        12: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [11],
      resultColumns: [12],
      values: transfers.map((row) => {
        const value = cashValue(row);
        return [row.source.account, row.source.accountCurrency, row.Comment, row.Type, row.Ticker,
          dateValue(row.Time), value.amount, value.currency, value.originalAmount,
          value.fx.value, value.czk];
      }),
    },
    {
      title: "Seznam transakcí s Frakčními právy (List of Fractional Shares Transactions)",
      offset: 3,
      normalizeRowStyles: true,
      dateColumns: [6, 13],
      numberFormats: {
        8: "0.00",
        11: '#,##0.00;[Red]-#,##0.00',
        12: '#,##0.00;[Red]-#,##0.00',
        15: '#,##0.00;[Red]-#,##0.00',
        16: '#,##0.00;[Red]-#,##0.00',
        17: '#,##0.00;[Red]-#,##0.00',
        18: "0.0000",
        19: "0.0000",
        20: "0.0000",
        21: '#,##0.00;[Red]-#,##0.00',
        22: "0.0000",
        23: '#,##0.00;[Red]-#,##0.00',
        24: "0.0000",
        25: '#,##0.00;[Red]-#,##0.00',
        26: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [20, 22, 24],
      resultColumns: [21, 23, 25, 26],
      values: fractionalDetailRows,
    },
    {
      title: "Seznam transakcí na STC/ETF (List of Stock Transactions)",
      offset: 3,
      normalizeRowStyles: true,
      dateColumns: [6, 13],
      numberFormats: {
        8: "0.0000",
        11: '#,##0.00;[Red]-#,##0.00',
        12: '#,##0.00;[Red]-#,##0.00',
        15: '#,##0.00;[Red]-#,##0.00',
        16: '#,##0.00;[Red]-#,##0.00',
        17: '#,##0.00;[Red]-#,##0.00',
        18: "0.0000",
        19: "0.0000",
        20: "0.0000",
        21: '#,##0.00;[Red]-#,##0.00',
        22: "0.0000",
        23: '#,##0.00;[Red]-#,##0.00',
        24: "0.0000",
        25: '#,##0.00;[Red]-#,##0.00',
        26: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [20, 22, 24],
      resultColumns: [21, 23, 25, 26],
      values: stockDetailRows,
    },
    {
      title: "Přehled zisků a ztrát na Frakčních právech (Profit Loss Summary Fractional Rights)",
      offset: 5,
      values: fractionalSummaryRows,
      numberFormats: {
        5: '#,##0.00;[Red]-#,##0.00',
        7: '#,##0.00;[Red]-#,##0.00',
        8: '#,##0.00;[Red]-#,##0.00',
        9: '#,##0.00;[Red]-#,##0.00',
        10: '#,##0.00;[Red]-#,##0.00',
        11: '#,##0.00;[Red]-#,##0.00',
        13: '#,##0.00;[Red]-#,##0.00',
        14: '#,##0.00;[Red]-#,##0.00',
        15: '#,##0.00;[Red]-#,##0.00',
        16: '#,##0.00;[Red]-#,##0.00',
        17: '#,##0.00;[Red]-#,##0.00',
        18: '#,##0.00;[Red]-#,##0.00',
      },
      resultColumns: [18],
    },
    {
      title: "Přehled zisků a ztrát na STC (Profit Loss Summary STC)",
      offset: 4,
      values: stockSummaryRows,
      numberFormats: {
        5: '#,##0.00;[Red]-#,##0.00',
        7: '#,##0.00;[Red]-#,##0.00',
        8: '#,##0.00;[Red]-#,##0.00',
        9: '#,##0.00;[Red]-#,##0.00',
        10: '#,##0.00;[Red]-#,##0.00',
        11: '#,##0.00;[Red]-#,##0.00',
        13: '#,##0.00;[Red]-#,##0.00',
        14: '#,##0.00;[Red]-#,##0.00',
        15: '#,##0.00;[Red]-#,##0.00',
        16: '#,##0.00;[Red]-#,##0.00',
        17: '#,##0.00;[Red]-#,##0.00',
        18: '#,##0.00;[Red]-#,##0.00',
      },
      resultColumns: [18],
    },
    {
      title: "Přehled zisků a ztrát na CFD (Profit Loss Summary CFD)",
      offset: 2,
      values: cfdSummaryRows,
      numberFormats: {
        5: '#,##0.00;[Red]-#,##0.00',
        6: '#,##0.00;[Red]-#,##0.00',
        7: '#,##0.00;[Red]-#,##0.00',
        8: '#,##0.00;[Red]-#,##0.00',
      },
      resultColumns: [8],
    },
    {
      title: "Souhrn zisku a ztráty s úpravami (Profit Loss Summary with adjustments)",
      offset: 3,
      values: adjustmentSummaryRows,
      numberFormats: {
        5: '#,##0.00;[Red]-#,##0.00',
        6: "0.0000",
        7: '#,##0.00;[Red]-#,##0.00',
      },
      rateColumns: [6],
      resultColumns: [7],
    },
    {
      title: "Souhrn peněžních toků v rámci poplatků, prémií, úroků (Fees, Bonuses, Interest Rates cashflows Summary)",
      offset: 2,
      values: feesSummaryRows,
      numberFormats: {
        4: '#,##0.00;[Red]-#,##0.00',
        6: '#,##0.00;[Red]-#,##0.00',
      },
      resultColumns: [6],
    },
  ];

  const summarySections = [
    {
      title: "Přehled zisků a ztrát na CFD (Profit Loss Summary CFD)",
      offset: 2,
      rows: cfdSummaryRows,
    },
    {
      title: "Přehled zisků a ztrát na STC (Profit Loss Summary STC)",
      offset: 4,
      rows: stockSummaryRows,
    },
    {
      title: "Přehled zisků a ztrát na Frakčních právech (Profit Loss Summary Fractional Rights)",
      offset: 5,
      rows: fractionalSummaryRows,
    },
    {
      title: "Souhrn zisku a ztráty s úpravami (Profit Loss Summary with adjustments)",
      offset: 3,
      rows: adjustmentSummaryRows,
    },
    {
      title: "Souhrn peněžních toků v rámci poplatků, prémií, úroků (Fees, Bonuses, Interest Rates cashflows Summary)",
      offset: 2,
      rows: feesSummaryRows,
    },
  ];
  summarySections.forEach(({ rows }) => moveSummaryTotalsLast(rows));

  const orderedSections = sectionRows
    .map((section) => ({ ...section, originalRow: findRow(sheet, section.title) }))
    .sort((left, right) => right.originalRow - left.originalRow);
  for (const section of orderedSections) {
    writeSection(
      sheet,
      section.title,
      section.offset,
      section.values,
      section.dateColumns,
      section.numberFormats,
      section.rateColumns,
      section.resultColumns,
      section.normalizeRowStyles,
    );
  }
  summarySections.forEach(({ title, offset, rows }) => {
    if (!rows.length) return;
    const firstDataRow = findRow(sheet, title) + offset + 1;
    rows.forEach((values, index) => {
      if (isSummaryTotalRow(values)) return;
      const row = sheet.getRow(firstDataRow + index);
      for (let column = 1; column <= sheet.columnCount; column += 1) {
        clearFill(row.getCell(column), "FFFF4500");
      }
    });

    const lastValues = rows[rows.length - 1];
    if (!isSummaryTotalRow(lastValues)) return;
    const labelIndex = lastValues.findIndex((value) => typeof value === "string"
      && /^(Celkem|Součet|Total)\b/i.test(value.trim()));
    const label = sheet.getRow(firstDataRow + rows.length - 1).getCell(labelIndex + 2);
    setFill(label, "FFFF4500");
    const style = JSON.parse(JSON.stringify(label.style || {}));
    style.font = { ...(style.font || {}), bold: true };
    label.style = style;
  });

  const dividendTitleRow = findRow(
    sheet,
    "Seznam dividend a dalších peněžních operací (List of Dividends and other cash operations)",
  );
  const stockTitleRow = findRow(sheet, "Seznam transakcí na STC/ETF (List of Stock Transactions)");
  const fractionalTitleRow = findRow(
    sheet,
    "Seznam transakcí s Frakčními právy (List of Fractional Shares Transactions)",
  );
  const rateName = rateType === "CNB" ? "ČNB" : "MF";
  for (const titleRow of [stockTitleRow, fractionalTitleRow]) {
    if (titleRow < 0) continue;
    const header = sheet.getRow(titleRow + 3);
    cloneStyle(header.getCell(23), header.getCell(21));
    cloneStyle(header.getCell(22), header.getCell(24));
    cloneStyle(header.getCell(23), header.getCell(25));
    cloneStyle(header.getCell(23), header.getCell(26));
    header.getCell(18).value = "Kurz XTB při prodeji";
    header.getCell(19).value = "Kurz XTB při nákupu";
    header.getCell(20).value = `Kurz ${rateName} při nákupu (CZK/j.)`;
    header.getCell(21).value = "Nákup po přepočtu (CZK)";
    header.getCell(22).value = `Kurz ${rateName} při prodeji (CZK/j.)`;
    header.getCell(23).value = "Prodej po přepočtu (CZK)";
    header.getCell(24).value = `Kurz ${rateName} pro poplatek (CZK/j.)`;
    header.getCell(25).value = "Poplatek po přepočtu (CZK)";
    header.getCell(26).value = `Čistý zisk/ztráta dle ${rateName} (CZK)`;
    for (const [column, width] of [[20, 14], [21, 18], [22, 14], [23, 18], [24, 14], [25, 18], [26, 22]]) {
      sheet.getColumn(column).width = width;
    }
  }
  const cfdTitleRow = findRow(sheet, "Seznam CFD transakcí (List of CFD Transactions)");
  if (cfdTitleRow > 0) {
    const header = sheet.getRow(cfdTitleRow + 2);
    cloneStyle(header.getCell(16), header.getCell(15));
    header.getCell(14).value = `Kurz ${rateName} (CZK/1 jednotku)`;
    header.getCell(15).value = "Zisk/ztráta (CZK)";
    header.getCell(16).value = null;
    setFill(header.getCell(16), "FFFFFFFF");
  }
  const openTradesTitleRow = findRow(sheet, "Seznam otevřených obchodů (List of Open Trades)");
  if (openTradesTitleRow > 0) {
    const header = sheet.getRow(openTradesTitleRow + 2);
    cloneStyle(header.getCell(16), header.getCell(15));
    header.getCell(14).value = `Kurz ${rateName} (CZK/1 jednotku)`;
    header.getCell(15).value = "Zisk/ztráta (CZK)";
    header.getCell(16).value = null;
    setFill(header.getCell(16), "FFFFFFFF");
  }
  if (dividendTitleRow > 0) {
    const header = sheet.getRow(dividendTitleRow + 2);
    cloneStyle(header.getCell(13), header.getCell(12));
    header.getCell(6).value = "Částka v měně účtu (XTB)";
    header.getCell(7).value = "Původní částka (USD)";
    header.getCell(8).value = "ISIN";
    header.getCell(9).value = "Odvozený kurz do měny účtu";
    header.getCell(10).value = "Poznámka k částce";
    header.getCell(11).value = `Kurz ${rateName} (CZK/1 jednotku)`;
    header.getCell(12).value = "Po přepočtu (CZK)";
    header.getCell(13).value = null;
    setFill(header.getCell(10), "FFFFFFFF");
    setFontColor(header.getCell(10), "FF000000");
    setFill(header.getCell(13), "FFFFFFFF");
  }
  const feesTitleRow = findRow(
    sheet,
    "Seznam poplatků, bonusů, úrokových sazeb peněžních toků (List of Fees, Bonuses, Interest Rates cashflows)",
  );
  if (feesTitleRow > 0) {
    const header = sheet.getRow(feesTitleRow + 2);
    cloneStyle(header.getCell(12), header.getCell(11));
    header.getCell(8).value = "Měna přepočtu";
    header.getCell(9).value = "Částka v měně";
    header.getCell(10).value = `Kurz ${rateName} (CZK/1 jednotku)`;
    header.getCell(11).value = "Po přepočtu (CZK)";
    header.getCell(12).value = null;
    setFill(header.getCell(12), "FFFFFFFF");
  }
  const adjustmentTitleRow = findRow(sheet, "Úpravy transakcí (Transaction adjustments)");
  if (adjustmentTitleRow > 0) {
    const header = sheet.getRow(adjustmentTitleRow + 2);
    cloneStyle(header.getCell(13), header.getCell(12));
    header.getCell(9).value = "Měna přepočtu";
    header.getCell(10).value = "Částka v měně";
    header.getCell(11).value = `Kurz ${rateName} (CZK/1 jednotku)`;
    header.getCell(12).value = "Po přepočtu (CZK)";
    header.getCell(13).value = null;
    setFill(header.getCell(13), "FFFFFFFF");
  }
  const adjustmentSummaryTitleRow = findRow(
    sheet,
    "Souhrn zisku a ztráty s úpravami (Profit Loss Summary with adjustments)",
  );
  if (adjustmentSummaryTitleRow > 0) {
    sheet.getRow(adjustmentSummaryTitleRow + 3).getCell(6).value = `Kurz ${rateName} (CZK/1 jednotku)`;
  }
  const feesSummaryTitleRow = findRow(
    sheet,
    "Souhrn peněžních toků v rámci poplatků, prémií, úroků (Fees, Bonuses, Interest Rates cashflows Summary)",
  );
  if (feesSummaryTitleRow > 0) {
    const header = sheet.getRow(feesSummaryTitleRow + 2);
    cloneStyle(header.getCell(7), header.getCell(6));
    header.getCell(5).value = "Metoda přepočtu";
    header.getCell(6).value = "Po přepočtu (CZK)";
    header.getCell(7).value = null;
    setFill(header.getCell(5), "FFFFFFFF");
    setFill(header.getCell(7), "FFFFFFFF");
  }
  clearSectionTailColumn(
    sheet,
    "Seznam CFD transakcí (List of CFD Transactions)",
    2,
    16,
  );
  clearSectionTailColumn(
    sheet,
    "Seznam otevřených obchodů (List of Open Trades)",
    2,
    16,
  );
  clearSectionTailColumn(
    sheet,
    "Seznam poplatků, bonusů, úrokových sazeb peněžních toků (List of Fees, Bonuses, Interest Rates cashflows)",
    2,
    12,
  );
  clearSectionTailColumn(
    sheet,
    "Seznam dividend a dalších peněžních operací (List of Dividends and other cash operations)",
    2,
    13,
  );
  clearSectionTailColumn(sheet, "Úpravy transakcí (Transaction adjustments)", 2, 13);
  clearSectionTailColumn(
    sheet,
    "Souhrn peněžních toků v rámci poplatků, prémií, úroků (Fees, Bonuses, Interest Rates cashflows Summary)",
    2,
    7,
  );

  const dividendSummaryStart = findRow(sheet, "Obdržené dividendy podle zemí z hlediska ISIN (Recieved dividends by countries by ISIN)");
  if (dividendSummaryStart > 0) {
    const countryTotals = new Map();
    for (const row of [...dividends, ...taxes]) {
      const key = countryFromIsin(isinForTicker(row.Ticker, isinCache));
      const entry = countryTotals.get(key) || { gross: 0, tax: 0 };
      const amount = cashValue(row).czk;
      if (getCashType(row).includes("withholding")) entry.tax += amount;
      else entry.gross += amount;
      countryTotals.set(key, entry);
    }
    const firstRow = dividendSummaryStart + 4;
    const nextSectionRow = SECTION_TITLES
      .map((title) => findRow(sheet, title))
      .filter((rowNumber) => rowNumber > dividendSummaryStart)
      .reduce((nearest, rowNumber) => Math.min(nearest, rowNumber), sheet.rowCount + 1);
    const totalStyleRowNumber = Array.from(
      { length: Math.max(0, nextSectionRow - firstRow) },
      (_, index) => firstRow + index,
    ).find((rowNumber) => sheet.getRow(rowNumber).values
      .some((value) => value === "Součet (Total)"));
    if (!totalStyleRowNumber) {
      throw new Error("V šabloně chybí zvýrazněný součtový řádek přehledu dividend podle zemí.");
    }

    const columns = [2, 3, 4, 5, 6];
    const saveStyles = (rowNumber) => columns.map((column) => {
      const cell = sheet.getRow(rowNumber).getCell(column);
      return {
        style: JSON.parse(JSON.stringify(cell.style || {})),
        numFmt: cell.numFmt,
      };
    });
    const regularStyles = saveStyles(firstRow);
    const totalStyles = saveStyles(totalStyleRowNumber);
    const applyStyles = (row, styles) => styles.forEach(({ style, numFmt }, index) => {
      const cell = row.getCell(columns[index]);
      cell.style = JSON.parse(JSON.stringify(style));
      if (numFmt) cell.numFmt = numFmt;
    });

    [...countryTotals.entries()].forEach(([country, value], index) => {
      const row = sheet.getRow(firstRow + index);
      row.getCell(2).value = country;
      row.getCell(3).value = "CZK";
      row.getCell(4).value = value.gross;
      row.getCell(5).value = value.tax;
      row.getCell(6).value = value.gross + value.tax;
      applyStyles(row, regularStyles);
    });
    const totalRow = sheet.getRow(firstRow + countryTotals.size);
    totalRow.getCell(2).value = "Součet (Total)";
    totalRow.getCell(3).value = "CZK";
    totalRow.getCell(4).value = [...countryTotals.values()].reduce((sum, item) => sum + item.gross, 0);
    totalRow.getCell(5).value = [...countryTotals.values()].reduce((sum, item) => sum + item.tax, 0);
    totalRow.getCell(6).value = totalRow.getCell(4).value + totalRow.getCell(5).value;
    applyStyles(totalRow, totalStyles);
  }

  ensureSectionSpacing(sheet);
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

module.exports = {
  buildReport,
  dateKey,
  fxFor,
};
