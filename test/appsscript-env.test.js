// Apps Script has no require/module: all files share one global scope. Load the sources that way.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { FakeSpreadsheet } = require("./fake-sheets.js");

function appsScriptContext(extra = {}) {
  const ctx = vm.createContext({ Date, JSON, Math, Object, String, Number, Array, Error, ...extra });
  for (const f of ["core.js", "sheets.js"]) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, "..", "src", f), "utf8"), ctx, { filename: f });
  }
  return ctx;
}

test("sources load as Apps Script files and work together", () => {
  const ctx = appsScriptContext();
  for (const fn of ["onOpen", "recalcStock", "monthReportCurrent", "monthReportPrevious", "checkAlerts", "installDailyTrigger", "dailyCheck", "setupDemo"]) {
    assert.equal(typeof ctx[fn], "function", fn);
  }
  const ss = new FakeSpreadsheet();
  ctx.runSetupDemo(ss, new Date(2030, 0, 20));
  const res = ctx.runRecalc(ss, new Date(2030, 0, 20));
  assert.equal(res.errors.length, 0);
  assert.equal(res.stock.length, 4);
});

test("menu entry points call the real Apps Script services", () => {
  const toasts = [], menu = [];
  const ss = new FakeSpreadsheet();
  const ctx = appsScriptContext({
    SpreadsheetApp: {
      getActive: () => ss,
      getUi: () => ({ createMenu: () => { const m = { addItem: (t, f) => (menu.push(f), m), addSeparator: () => m, addToUi: () => m }; return m; } }),
      ProtectionType: { SHEET: "SHEET" },
    },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  });
  ss.toast = (msg) => toasts.push(msg);
  ctx.onOpen();
  assert.deepEqual(menu, ["recalcStock", "monthReportCurrent", "monthReportPrevious", "checkAlerts", "installDailyTrigger", "setupDemo"]);
  ctx.setupDemo();
  ctx.recalcStock();
  assert.match(toasts.at(-1), /Позиций: 4, ошибок в данных: 0/);
  ctx.monthReportCurrent();
  assert.match(toasts.at(-1), /Готово: лист «Отчёт \d{4}-\d{2}»/);
  ctx.PropertiesService = { getScriptProperties: () => ({ getProperty: () => null }) };
  ctx.checkAlerts();
  assert.match(toasts.at(-1), /Telegram не настроен/);
  for (const fn of menu) assert.equal(typeof ctx[fn], "function", fn);
});

test("a second recalc while one is running is refused, not interleaved", () => {
  const ss = new FakeSpreadsheet();
  const ctx = appsScriptContext({
    SpreadsheetApp: { getActive: () => ss, ProtectionType: { SHEET: "SHEET" } },
    LockService: { getDocumentLock: () => ({ tryLock: () => false, releaseLock: () => {} }) },
  });
  assert.throws(() => ctx.recalcStock(), /пересчитывается/);
});

test("no dialogs that wait for typed input: they hung in real Google Sheets", () => {
  const src = fs.readFileSync(path.join(__dirname, "..", "src", "sheets.js"), "utf8");
  assert.ok(!/\.prompt\(|\.alert\(/.test(src));
});
