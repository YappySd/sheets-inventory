const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../src/core.js");

const NOW = new Date(2030, 0, 20);
const PRODUCTS = [
  ["SK-28", "Сковорода 28 см", "шт", 5, 520],
  ["pot-3", "Набор кастрюль", "компл", 3, "1 450,00"],
];

const info = () => C.readProducts(PRODUCTS, 2);
const moves = (rows) => C.readMovements(rows, 2, info().products, NOW);

test("numbers as people type them", () => {
  assert.equal(C.parseNumber("1 234,5"), 1234.5);
  assert.equal(C.parseNumber(7), 7);
  for (const bad of ["", "abc", "1,2,3", "1e5", "12 шт"]) assert.ok(Number.isNaN(C.parseNumber(bad)), bad);
});

test("dates: cells, dd.mm.yyyy, ISO, impossible dates", () => {
  assert.deepEqual(C.parseDate("31.01.2030"), new Date(2030, 0, 31));
  assert.deepEqual(C.parseDate("2030-01-31"), new Date(2030, 0, 31));
  assert.deepEqual(C.parseDate(new Date(2030, 0, 5, 15, 30)), new Date(2030, 0, 5));
  for (const bad of ["31.02.2030", "2030-13-01", "вчера", "", null]) assert.equal(C.parseDate(bad), null, String(bad));
});

test("product list: duplicates (case-insensitive), empty rows, bad numbers", () => {
  const r = C.readProducts([...PRODUCTS, ["", "", "", "", ""], ["Sk-28", "Дубль", "шт", 1, 1],
    ["X-1", "", "шт", 0, 0], ["X-2", "Чайник", "шт", -1, 0], ["X-3", "Нож", "шт", 1, "дорого"]], 2);
  assert.deepEqual(r.order, ["SK-28", "POT-3"]);
  assert.equal(r.products["POT-3"].cost, 1450);
  assert.deepEqual(r.errors.map((e) => e.row), [5, 6, 7, 8]);
});

test("movements are validated and sorted by date", () => {
  const r = moves([
    ["10.01.2030", "sk-28", "Приход", 10, "поставка"],
    ["05.01.2030", "POT-3", "прихід", "2", ""],
    ["21.01.2030", "SK-28", "расход", 1, ""],     // future
    ["11.01.2030", "NOPE", "расход", 1, ""],      // unknown sku
    ["11.01.2030", "SK-28", "подарок", 1, ""],    // unknown type
    ["11.01.2030", "SK-28", "расход", 0, ""],     // zero
    ["11.01.2030", "SK-28", "расход", 2000000, ""], // absurd
    ["", "", "", "", ""],
  ]);
  assert.deepEqual(r.moves.map((m) => m.row), [3, 2]);
  assert.deepEqual(r.errors.map((e) => e.row), [4, 5, 6, 7, 8]);
});

test("stock, statuses and value", () => {
  const m = moves([
    ["01.01.2030", "SK-28", "приход", 20, ""], ["10.01.2030", "SK-28", "расход", 16, ""],
    ["01.01.2030", "POT-3", "приход", 3, ""], ["12.01.2030", "POT-3", "списание", 3, "брак"],
  ]).moves;
  const s = C.computeStock(info(), m);
  assert.deepEqual(s.rows.map((r) => [r.sku, r.qty, r.status, r.value]),
    [["SK-28", 4, "заканчивается", 2080], ["POT-3", 0, "нет в наличии", 0]]);
  assert.deepEqual(s.errors, []);
});

test("going below zero is reported with the row", () => {
  const m = moves([["01.01.2030", "SK-28", "приход", 2, ""], ["02.01.2030", "SK-28", "расход", 5, ""]]).moves;
  const s = C.computeStock(info(), m);
  assert.equal(s.errors[0].row, 3);
  assert.match(s.errors[0].message, /в минус \(-3\)/);
  assert.equal(s.rows[0].value, 0); // negative stock is not counted as money
});

test("fractional quantities do not drift", () => {
  const p = C.readProducts([["W", "Крупа", "кг", 0, 10]], 2);
  const rows = Array.from({ length: 10 }, () => ["01.01.2030", "W", "приход", "0,1", ""]);
  const m = C.readMovements(rows, 2, p.products, NOW).moves;
  assert.equal(C.computeStock(p, m).rows[0].qty, 1); // not 0.9999999999999999
});

test("month report separates sales and write-offs", () => {
  const m = moves([
    ["28.12.2029", "SK-28", "приход", 10, ""], ["03.01.2030", "SK-28", "приход", 5, ""],
    ["04.01.2030", "SK-28", "продажа", 6, ""], ["05.01.2030", "SK-28", "списание", 1, "брак"],
  ]).moves;
  const [row] = C.monthReport(info(), m, 2030, 1);
  assert.deepEqual([row.inQty, row.outQty, row.writeOff, row.endQty], [5, 6, 1, 8]);
  assert.equal(C.monthReport(info(), m, 2029, 12)[0].endQty, 10);
});

test("alert text is plain, bounded and lists low items", () => {
  const low = Array.from({ length: 40 }, (_, i) => ({ sku: "S" + i, name: "<b>x</b>", qty: 1, unit: "шт", min: 2 }));
  const text = C.formatAlert("Склад", low);
  assert.match(text, /заканчивается 40 поз/);
  assert.match(text, /…и ещё 10/);
  assert.ok(text.length <= 4000 && text.includes("<b>x</b>")); // sent without parse_mode, shown literally
  assert.equal(C.formatAlert("Склад", []), "");
});
