const test = require("node:test");
const assert = require("node:assert/strict");
const { classifyTrades } = require("./report");

function row({
  account = "account-1",
  currency = "USD",
  ticker = "ABC.US",
  type = "BUY",
  volume,
  category = "STC",
  product = "ABC",
}) {
  return {
    source: { account, accountCurrency: currency },
    Ticker: ticker,
    Type: type,
    Volume: volume,
    Category: category,
    Product: product,
  };
}

test("moves all fractional purchases to stocks once the account and ticker total reaches one", () => {
  const first = row({ volume: 0.4 });
  const second = row({ volume: 0.6 });
  const unrelated = row({ ticker: "XYZ.US", volume: 0.4 });
  const result = classifyTrades([first, second, unrelated], [first, second, unrelated]);

  assert.deepEqual(result.stockRows, [first, second]);
  assert.deepEqual(result.fractionalRows, [unrelated]);
});

test("counts open purchases and keeps sub-unit totals as fractional rights", () => {
  const closed = row({ volume: 0.3 });
  const open = row({ volume: 0.5 });
  const otherAccount = row({ account: "account-2", volume: 0.4 });
  const result = classifyTrades([closed, otherAccount], [closed, open, otherAccount]);

  assert.deepEqual(result.stockRows, []);
  assert.deepEqual(result.fractionalRows, [closed, otherAccount]);
});

test("does not count sales or CFDs toward the fractional purchase total", () => {
  const purchase = row({ volume: 0.4 });
  const sale = row({ volume: 0.6, type: "SELL" });
  const cfd = row({ volume: 1, category: "CFD", product: "CFD" });
  const result = classifyTrades([purchase, sale, cfd], [purchase, sale, cfd]);

  assert.deepEqual(result.stockRows, []);
  assert.deepEqual(result.fractionalRows, [purchase, sale]);
  assert.deepEqual(result.cfdRows, [cfd]);
});
