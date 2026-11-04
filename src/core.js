/**
 * Pure inventory logic: no SpreadsheetApp, no network. Runs in Apps Script (V8) and in Node for tests.
 * Everything that comes from a cell is treated as untrusted text typed by a person.
 */

var MAX_QTY = 1000000;
var TYPES = {
  "приход": 1, "прихід": 1,
  "расход": -1, "витрата": -1, "продажа": -1, "продаж": -1,
  "списание": -1, "списання": -1
};

/** "1 234,5" / 1234.5 / " 12 " -> number, or NaN. */
function parseNumber(value) {
  if (typeof value === "number") return value;
  var text = String(value == null ? "" : value).replace(/[\s\u00a0]/g, "").replace(",", ".");
  if (!/^-?\d+(\.\d+)?$/.test(text)) return NaN;
  return Number(text);
}

/** Date cell, "31.01.2030" or "2030-01-31" -> Date (local midnight), or null. */
function parseDate(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var text = String(value == null ? "" : value).trim();
  var m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(text);
  var y, mo, d;
  if (m) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else {
    m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
    if (!m) return null;
    y = +m[1]; mo = +m[2]; d = +m[3];
  }
  var date = new Date(y, mo - 1, d);
  // Rejects 31.02.2030 and similar: JS would silently roll it over to March.
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
  return date;
}

function cleanText(value, limit) {
  return String(value == null ? "" : value).replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim().slice(0, limit);
}

function skuKey(value) {
  return cleanText(value, 64).toUpperCase();
}

/**
 * Sheets runs a cell that starts with = + - @ as a formula (and so does Excel after export).
 * Text from people is written back with a leading apostrophe; numbers and dates stay as they are.
 */
function safeCell(value) {
  if (typeof value !== "string") return value;
  var first = value.charAt(0);
  if (first === "=" || first === "+" || first === "-" || first === "@" ||
      first === String.fromCharCode(9) || first === String.fromCharCode(13)) {
    return "'" + value;
  }
  return value;
}

/** products: rows of [sku, name, unit, min, cost] without the header. rowOffset = sheet row of rows[0]. */
function readProducts(rows, rowOffset) {
  var products = {}, order = [], errors = [];
  rows.forEach(function (r, i) {
    var row = rowOffset + i;
    if (r.every(function (c) { return c === "" || c == null; })) return;
    var sku = skuKey(r[0]), name = cleanText(r[1], 120);
    var min = r[3] === "" || r[3] == null ? 0 : parseNumber(r[3]);
    var cost = r[4] === "" || r[4] == null ? 0 : parseNumber(r[4]);
    if (!sku) return errors.push({ sheet: "Товары", row: row, message: "нет артикула" });
    if (!name) return errors.push({ sheet: "Товары", row: row, message: "нет названия у " + sku });
    if (products[sku]) return errors.push({ sheet: "Товары", row: row, message: "артикул " + sku + " повторяется" });
    if (!(min >= 0) || min > MAX_QTY) return errors.push({ sheet: "Товары", row: row, message: "мин. остаток должен быть числом от 0" });
    if (!(cost >= 0) || cost > 1e9) return errors.push({ sheet: "Товары", row: row, message: "цена закупки должна быть числом от 0" });
    products[sku] = { sku: sku, name: name, unit: cleanText(r[2], 10) || "шт", min: min, cost: cost };
    order.push(sku);
  });
  return { products: products, order: order, errors: errors };
}

/** movements: rows of [date, sku, type, qty, comment]. now: Date used to reject future dates. */
function readMovements(rows, rowOffset, products, now) {
  var moves = [], errors = [];
  var tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  rows.forEach(function (r, i) {
    var row = rowOffset + i;
    if (r.every(function (c) { return c === "" || c == null; })) return;
    var date = parseDate(r[0]), sku = skuKey(r[1]);
    var sign = TYPES[cleanText(r[2], 20).toLowerCase()], qty = parseNumber(r[3]);
    var err = function (m) { errors.push({ sheet: "Движение", row: row, message: m }); };
    if (!date) return err("неверная дата: нужно 31.01.2030");
    if (date >= tomorrow) return err("дата в будущем");
    if (!products[sku]) return err("нет товара с артикулом " + (sku || "(пусто)"));
    if (!sign) return err("тип должен быть: приход, расход или списание");
    if (!(qty > 0) || qty > MAX_QTY) return err("количество должно быть числом больше 0");
    moves.push({ row: row, date: date, sku: sku, delta: sign * qty, type: cleanText(r[2], 20).toLowerCase(),
                 comment: cleanText(r[4], 200) });
  });
  moves.sort(function (a, b) { return a.date - b.date || a.row - b.row; });
  return { moves: moves, errors: errors };
}

/** Stock per SKU after all valid moves; a move that takes stock below zero is reported once. */
function computeStock(productsInfo, moves) {
  var qty = {}, errors = [];
  productsInfo.order.forEach(function (sku) { qty[sku] = 0; });
  moves.forEach(function (m) {
    qty[m.sku] = round(qty[m.sku] + m.delta);
    if (qty[m.sku] < 0 && m.delta < 0) {
      errors.push({ sheet: "Движение", row: m.row, message: m.sku + " уходит в минус (" + qty[m.sku] + "): проверьте приход" });
    }
  });
  var rows = productsInfo.order.map(function (sku) {
    var p = productsInfo.products[sku], q = qty[sku];
    var status = q <= 0 ? "нет в наличии" : q <= p.min ? "заканчивается" : "в наличии";
    return { sku: sku, name: p.name, unit: p.unit, qty: q, min: p.min, value: round(Math.max(q, 0) * p.cost), status: status };
  });
  return { rows: rows, errors: errors };
}

/** Month report: per SKU in, out (sales + write-offs), stock at the end of the month. month is 1-12. */
function monthReport(productsInfo, moves, year, month) {
  var start = new Date(year, month - 1, 1), end = new Date(year, month, 1);
  var acc = {};
  productsInfo.order.forEach(function (sku) { acc[sku] = { inQty: 0, outQty: 0, writeOff: 0, endQty: 0 }; });
  moves.forEach(function (m) {
    var a = acc[m.sku];
    if (m.date < end) a.endQty = round(a.endQty + m.delta);
    if (m.date < start || m.date >= end) return;
    if (m.delta > 0) a.inQty = round(a.inQty + m.delta);
    else if (m.type === "списание" || m.type === "списання") a.writeOff = round(a.writeOff - m.delta);
    else a.outQty = round(a.outQty - m.delta);
  });
  return productsInfo.order.map(function (sku) {
    var p = productsInfo.products[sku], a = acc[sku];
    return { sku: sku, name: p.name, inQty: a.inQty, outQty: a.outQty, writeOff: a.writeOff, endQty: a.endQty };
  });
}

function lowStock(stockRows) {
  return stockRows.filter(function (r) { return r.status !== "в наличии"; });
}

/** Plain text for Telegram: no parse_mode, so names from cells cannot inject markup. */
function formatAlert(shop, low, limit) {
  limit = limit || 30;
  if (!low.length) return "";
  var lines = ["Склад «" + cleanText(shop, 60) + "»: заканчивается " + low.length + " поз."];
  low.slice(0, limit).forEach(function (r) {
    lines.push("• " + r.sku + " " + r.name + ": " + r.qty + " " + r.unit + " (мин. " + r.min + ")");
  });
  if (low.length > limit) lines.push("…и ещё " + (low.length - limit));
  return lines.join("\n").slice(0, 4000);
}

function round(x) {
  return Math.round(x * 1000) / 1000;
}

if (typeof module !== "undefined") {
  module.exports = { parseNumber: parseNumber, parseDate: parseDate, cleanText: cleanText, safeCell: safeCell,
    readProducts: readProducts, readMovements: readMovements, computeStock: computeStock,
    monthReport: monthReport, lowStock: lowStock, formatAlert: formatAlert, MAX_QTY: MAX_QTY };
}
