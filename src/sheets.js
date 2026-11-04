/**
 * Apps Script adapter: sheets, menu, daily trigger, Telegram. Business rules live in core.js.
 * Entry points (onOpen, recalcStock, ...) are thin; run* functions take the spreadsheet as a
 * parameter so tests can pass a fake one.
 */

var C = typeof require === "function" ? require("./core.js") : globalThis;

var SHEET = { products: "Товары", moves: "Движение", stock: "Остатки", errors: "Ошибки" };
var HEAD = {
  products: ["Артикул", "Название", "Ед.", "Мин. остаток", "Цена закупки"],
  moves: ["Дата", "Артикул", "Тип", "Количество", "Комментарий"],
  stock: ["Артикул", "Название", "Остаток", "Ед.", "Мин.", "Статус", "Стоимость остатка"],
  errors: ["Лист", "Строка", "Что не так"],
  report: ["Артикул", "Название", "Приход", "Расход", "Списание", "Остаток на конец"]
};

// ------------------------------------------------------------------ core run

function readAll(ss, now) {
  var products = C.readProducts(rowsOf(ss.getSheetByName(SHEET.products)), 2);
  var moves = C.readMovements(rowsOf(ss.getSheetByName(SHEET.moves)), 2, products.products, now);
  return { products: products, moves: moves, errors: products.errors.concat(moves.errors) };
}

function runRecalc(ss, now) {
  var data = readAll(ss, now);
  var stock = C.computeStock(data.products, data.moves.moves);
  var errors = data.errors.concat(stock.errors);
  writeTable(sheetOf(ss, SHEET.stock), HEAD.stock, stock.rows.map(function (r) {
    return [r.sku, r.name, r.qty, r.unit, r.min, r.status, r.value];
  }));
  writeTable(sheetOf(ss, SHEET.errors), HEAD.errors, errors.map(function (e) {
    return [e.sheet, e.row, e.message];
  }));
  return { stock: stock.rows, errors: errors };
}

function runMonthReport(ss, now, year, month) {
  var data = readAll(ss, now);
  var rows = C.monthReport(data.products, data.moves.moves, year, month);
  var name = "Отчёт " + year + "-" + (month < 10 ? "0" : "") + month;
  writeTable(sheetOf(ss, name), HEAD.report, rows.map(function (r) {
    return [r.sku, r.name, r.inQty, r.outQty, r.writeOff, r.endQty];
  }));
  return { name: name, rows: rows, errors: data.errors.length };
}

/** fetch(url, options) is UrlFetchApp.fetch in Apps Script, a stub in tests. */
function runDailyCheck(ss, now, props, fetch) {
  var result = runRecalc(ss, now);
  var low = C.lowStock(result.stock);
  var token = props.getProperty("TG_TOKEN"), chat = props.getProperty("TG_CHAT_ID");
  if (!low.length || !token || !chat) return { sent: false, low: low.length };
  var text = C.formatAlert(ss.getName(), low);
  var resp = fetch("https://api.telegram.org/bot" + token + "/sendMessage", {
    method: "post", contentType: "application/json", muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat, text: text })  // no parse_mode: plain text only
  });
  var code = resp.getResponseCode();
  // The URL contains the token: never log it, only the status code.
  if (code !== 200) throw new Error("Telegram ответил " + code + ": проверьте токен и chat id");
  return { sent: true, low: low.length };
}

function validateAlertSettings(token, chat) {
  token = String(token || "").trim(); chat = String(chat || "").trim();
  if (!/^\d{6,12}:[A-Za-z0-9_-]{30,}$/.test(token)) return "Токен не похож на токен бота от @BotFather";
  if (!/^-?\d{3,20}$/.test(chat)) return "Chat id — это число (узнать: @userinfobot)";
  return "";
}

// ----------------------------------------------------------------- helpers

function rowsOf(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  return values.slice(1);  // without the header
}

function sheetOf(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function writeTable(sheet, header, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length)
      .setValues(rows.map(function (r) { return r.map(C.safeCell); }));
  }
  // Generated sheets: a person editing them by hand gets a warning, the script can still write.
  var type = typeof SpreadsheetApp !== "undefined" ? SpreadsheetApp.ProtectionType.SHEET : "SHEET";
  if (!sheet.getProtections(type).length) sheet.protect().setWarningOnly(true);
}

function withLock(fn) {
  // Two people pressing "Пересчитать" at once must not interleave writes.
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) throw new Error("Склад сейчас пересчитывается, попробуйте через минуту");
  try { return fn(); } finally { lock.releaseLock(); }
}

// --------------------------------------------------- Apps Script entry points

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Склад")
    .addItem("Пересчитать остатки", "recalcStock")
    .addItem("Отчёт за этот месяц", "monthReportCurrent")
    .addItem("Отчёт за прошлый месяц", "monthReportPrevious")
    .addSeparator()
    .addItem("Проверить настройки Telegram", "checkAlerts")
    .addItem("Включить ежедневную проверку (9:00)", "installDailyTrigger")
    .addSeparator()
    .addItem("Создать листы с примером", "setupDemo")
    .addToUi();
}

function recalcStock() {
  var ss = SpreadsheetApp.getActive();
  var res = withLock(function () { return runRecalc(ss, new Date()); });
  ss.toast("Позиций: " + res.stock.length + ", ошибок в данных: " + res.errors.length, "Склад", 5);
}

// No ui.prompt/ui.alert dialogs: in testing their answer never reached the script
// (execution stayed "paused"), so every entry point works without typing and reports via toast.

/** offset 0 = this month, -1 = previous month. Returns {year, month} with month 1-12. */
function monthOf(now, offset) {
  var d = new Date(now.getFullYear(), now.getMonth() + offset, 1);
  return { year: d.getFullYear(), month: d.getMonth() + 1 };
}

function monthReportFor(offset) {
  var ss = SpreadsheetApp.getActive(), now = new Date(), m = monthOf(now, offset);
  var res = withLock(function () { return runMonthReport(ss, now, m.year, m.month); });
  ss.toast("Готово: лист «" + res.name + "»", "Склад", 5);
}

function monthReportCurrent() { monthReportFor(0); }
function monthReportPrevious() { monthReportFor(-1); }

/** Telegram settings live in Script Properties (Apps Script → Project Settings), never in the sheet. */
function alertSettingsStatus(props) {
  var token = props.getProperty("TG_TOKEN"), chat = props.getProperty("TG_CHAT_ID");
  if (!token && !chat) return "Telegram не настроен: добавьте TG_TOKEN и TG_CHAT_ID в свойства скрипта";
  return validateAlertSettings(token, chat) || "Telegram настроен";
}

function checkAlerts() {
  var ss = SpreadsheetApp.getActive();
  ss.toast(alertSettingsStatus(PropertiesService.getScriptProperties()), "Склад", 8);
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === "dailyCheck") ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger("dailyCheck").timeBased().everyDays(1).atHour(9).create();
  SpreadsheetApp.getActive().toast("Каждый день около 9:00 скрипт пересчитает остатки и напишет, что заканчивается.", "Склад", 8);
}

function dailyCheck() {
  var ss = SpreadsheetApp.getActive();
  withLock(function () {
    return runDailyCheck(ss, new Date(), PropertiesService.getScriptProperties(), UrlFetchApp.fetch);
  });
}

function setupDemo() {
  var ss = SpreadsheetApp.getActive();
  runSetupDemo(ss, new Date());
  ss.toast("Листы созданы. Меню «Склад» → «Пересчитать остатки».", "Склад", 5);
}

function runSetupDemo(ss, now) {
  var products = ss.getSheetByName(SHEET.products);
  if (products && products.getLastRow() > 1) return false;  // never overwrite real data
  products = products || ss.insertSheet(SHEET.products);
  var moves = ss.getSheetByName(SHEET.moves) || ss.insertSheet(SHEET.moves);
  var d = function (daysAgo) { return new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysAgo); };
  var p = [HEAD.products,
    ["SK-28", "Сковорода 28 см", "шт", 5, 520],
    ["POT-3", "Набор кастрюль 3 шт", "компл", 3, 1450],
    ["KN-20", "Нож шеф 20 см", "шт", 4, 610],
    ["KT-17", "Электрочайник 1,7 л", "шт", 2, 780]];
  var m = [HEAD.moves,
    [d(20), "SK-28", "приход", 20, "поставка"], [d(18), "POT-3", "приход", 6, "поставка"],
    [d(18), "KN-20", "приход", 10, ""], [d(15), "KT-17", "приход", 4, ""],
    [d(9), "SK-28", "расход", 12, "продажи"], [d(7), "POT-3", "расход", 4, ""],
    [d(5), "KN-20", "списание", 1, "брак"], [d(3), "KT-17", "расход", 4, ""],
    [d(2), "SK-28", "расход", 4, ""]];
  products.getRange(1, 1, p.length, p[0].length).setValues(p);
  moves.getRange(1, 1, m.length, m[0].length).setValues(m);
  return true;
}

if (typeof module !== "undefined") {
  module.exports = { runRecalc: runRecalc, runMonthReport: runMonthReport, runDailyCheck: runDailyCheck,
    runSetupDemo: runSetupDemo, validateAlertSettings: validateAlertSettings,
    monthOf: monthOf, alertSettingsStatus: alertSettingsStatus, SHEET: SHEET, HEAD: HEAD };
}
