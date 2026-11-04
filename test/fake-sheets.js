// A small stand-in for SpreadsheetApp that behaves like Google Sheets where it matters:
// a written string starting with "=" becomes a formula, a leading apostrophe forces plain text.

class FakeRange {
  constructor(sheet, row, col, rows, cols) { Object.assign(this, { sheet, row, col, rows, cols }); }
  getValues() {
    return this.sheet.cells.slice(this.row - 1, this.row - 1 + this.rows)
      .map((r) => (r || []).slice(this.col - 1, this.col - 1 + this.cols).map((c) => (c && c.formula ? c.shown : c)));
  }
  setValues(values) {
    if (values.length !== this.rows || values.some((r) => r.length !== this.cols)) throw new Error("size mismatch");
    values.forEach((r, i) => {
      const row = (this.sheet.cells[this.row - 1 + i] = this.sheet.cells[this.row - 1 + i] || []);
      r.forEach((v, j) => {
        let cell = v;
        if (typeof v === "string" && v.startsWith("'")) cell = v.slice(1);
        else if (typeof v === "string" && /^[=+\-@]/.test(v)) {
          cell = { formula: v, shown: "#ERROR!" };
          this.sheet.formulas.push(v);
        }
        row[this.col - 1 + j] = cell;
      });
    });
    return this;
  }
}

class FakeSheet {
  constructor(name, rows = []) {
    this.name = name;
    this.cells = rows.map((r) => [...r]);
    this.formulas = [];
    this.protections = [];
  }
  getName() { return this.name; }
  getLastRow() { return this.cells.length; }
  getDataRange() {
    const cols = Math.max(0, ...this.cells.map((r) => r.length));
    return new FakeRange(this, 1, 1, this.cells.length, cols);
  }
  getRange(row, col, rows, cols) { return new FakeRange(this, row, col, rows, cols); }
  clearContents() { this.cells = []; return this; }
  getProtections() { return this.protections; }
  protect() {
    const p = { warningOnly: false, setWarningOnly(v) { this.warningOnly = v; return this; } };
    this.protections.push(p);
    return p;
  }
  table() { return this.getDataRange().getValues(); }
}

class FakeSpreadsheet {
  constructor(name = "ДомМаркет склад", sheets = {}) {
    this.name = name;
    this.sheets = {};
    for (const [n, rows] of Object.entries(sheets)) this.sheets[n] = new FakeSheet(n, rows);
  }
  getName() { return this.name; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { return (this.sheets[n] = new FakeSheet(n)); }
  allFormulas() { return Object.values(this.sheets).flatMap((s) => s.formulas); }
}

class FakeProps {
  constructor(values = {}) { this.values = values; }
  getProperty(k) { return this.values[k] || null; }
}

function fakeFetch(code = 200) {
  const calls = [];
  const fn = (url, options) => {
    calls.push({ url, options, body: JSON.parse(options.payload) });
    return { getResponseCode: () => code };
  };
  fn.calls = calls;
  return fn;
}

module.exports = { FakeSpreadsheet, FakeSheet, FakeProps, fakeFetch };
