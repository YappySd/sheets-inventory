const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const S = require("../src/sheets.js");
const { FakeSpreadsheet, FakeProps, fakeFetch } = require("./fake-sheets.js");

const NOW = new Date(2030, 0, 20);
// Built at run time so the secret scan does not flag this fake token in the source.
const TOKEN = "123456789:" + "TEST-not-a-real-token".padEnd(38, "x");

function shop(moves = []) {
  return new FakeSpreadsheet("ДомМаркет", {
    "Товары": [S.HEAD.products, ["SK-28", "Сковорода 28 см", "шт", 5, 520], ["KT-17", "Чайник", "шт", 2, 780]],
    "Движение": [S.HEAD.moves, ...moves],
  });
}

test("recalc writes stock and errors sheets", () => {
  const ss = shop([[new Date(2030, 0, 2), "SK-28", "приход", 8, ""], ["03.01.2030", "SK-28", "расход", 4, ""],
                   ["нет даты", "SK-28", "расход", 1, ""]]);
  const res = S.runRecalc(ss, NOW);
  assert.deepEqual(ss.getSheetByName("Остатки").table(), [S.HEAD.stock,
    ["SK-28", "Сковорода 28 см", 4, "шт", 5, "заканчивается", 2080],
    ["KT-17", "Чайник", 0, "шт", 2, "нет в наличии", 0]]);
  assert.deepEqual(ss.getSheetByName("Ошибки").table()[1], ["Движение", 4, "неверная дата: нужно 31.01.2030"]);
  assert.equal(res.errors.length, 1);
});

test("generated sheets are protected with a warning, once", () => {
  const ss = shop();
  S.runRecalc(ss, NOW);
  S.runRecalc(ss, NOW);
  const prot = ss.getSheetByName("Остатки").protections;
  assert.equal(prot.length, 1);
  assert.equal(prot[0].warningOnly, true);
});

test("month report sheet", () => {
  const ss = shop([["02.01.2030", "KT-17", "приход", 3, ""], ["04.01.2030", "KT-17", "расход", 1, ""]]);
  const r = S.runMonthReport(ss, NOW, 2030, 1);
  assert.equal(r.name, "Отчёт 2030-01");
  assert.deepEqual(ss.getSheetByName("Отчёт 2030-01").table()[2], ["KT-17", "Чайник", 3, 1, 0, 2]);
});

test("daily check sends a plain-text Telegram message only when something runs low", () => {
  const fetch = fakeFetch();
  const props = new FakeProps({ TG_TOKEN: TOKEN, TG_CHAT_ID: "42" });
  const ss = shop([["02.01.2030", "SK-28", "приход", 50, ""], ["02.01.2030", "KT-17", "приход", 10, ""]]);
  assert.deepEqual(S.runDailyCheck(ss, NOW, props, fetch), { sent: false, low: 0 });
  ss.getSheetByName("Движение").cells.push(["05.01.2030", "KT-17", "расход", 9, ""]);
  assert.deepEqual(S.runDailyCheck(ss, NOW, props, fetch), { sent: true, low: 1 });
  const call = fetch.calls[0];
  assert.equal(call.url, `https://api.telegram.org/bot${TOKEN}/sendMessage`);
  assert.equal(call.body.chat_id, "42");
  assert.equal(call.body.parse_mode, undefined);
  assert.match(call.body.text, /KT-17 Чайник: 1 шт/);
});

test("daily check without settings does nothing; Telegram error does not leak the token", () => {
  const ss = shop([]);
  assert.equal(S.runDailyCheck(ss, NOW, new FakeProps(), fakeFetch()).sent, false);
  assert.throws(
    () => S.runDailyCheck(ss, NOW, new FakeProps({ TG_TOKEN: TOKEN, TG_CHAT_ID: "42" }), fakeFetch(401)),
    (e) => /401/.test(e.message) && !e.message.includes(TOKEN));
});

test("alert settings are validated before saving", () => {
  assert.equal(S.validateAlertSettings(TOKEN, "123456789"), "");
  assert.equal(S.validateAlertSettings(TOKEN, "-1001234567890"), ""); // group chat
  assert.notEqual(S.validateAlertSettings("abc", "42"), "");
  assert.notEqual(S.validateAlertSettings(TOKEN, "@me"), "");
});

test("demo setup never overwrites real data", () => {
  const empty = new FakeSpreadsheet();
  assert.equal(S.runSetupDemo(empty, NOW), true);
  assert.equal(S.runRecalc(empty, NOW).errors.length, 0);
  assert.equal(S.runSetupDemo(shop(), NOW), false);
});

test("manifest: minimal scopes, outgoing requests only to Telegram", () => {
  const m = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "src", "appsscript.json"), "utf8"));
  assert.ok(m.oauthScopes.includes("https://www.googleapis.com/auth/spreadsheets.currentonly"));
  assert.ok(!m.oauthScopes.some((s) => s.endsWith("/auth/spreadsheets") || s.includes("drive")));
  assert.deepEqual(m.urlFetchWhitelist, ["https://api.telegram.org/"]);
});

test("report months: this month and previous, across the new year", () => {
  assert.deepEqual(S.monthOf(new Date(2030, 0, 31), 0), { year: 2030, month: 1 });
  assert.deepEqual(S.monthOf(new Date(2030, 0, 31), -1), { year: 2029, month: 12 });
  assert.deepEqual(S.monthOf(new Date(2030, 2, 31), -1), { year: 2030, month: 2 }); // not "March 3"
});

test("telegram settings check never echoes the token", () => {
  assert.match(S.alertSettingsStatus(new FakeProps()), /не настроен/);
  assert.equal(S.alertSettingsStatus(new FakeProps({ TG_TOKEN: TOKEN, TG_CHAT_ID: "123456789" })), "Telegram настроен");
  const bad = S.alertSettingsStatus(new FakeProps({ TG_TOKEN: TOKEN, TG_CHAT_ID: "@me" }));
  assert.ok(/[Cc]hat id/.test(bad) && !bad.includes(TOKEN));
});
