// A lightweight, functionally faithful mock of the SpreadsheetApp API used
// by src/SheetWriter.gs, so BudgetBook can actually be exercised end to end
// in Node instead of only being reviewed by eye. It does not evaluate
// formulas (there is no live Sheets engine here) - it exists to catch API
// misuse, off-by-one row/column math, and crashes, and to verify that
// values written via setValues() round-trip correctly through
// read_transactions()/read_rules()/scan_accounts().

class MockRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  _cell(r, c) {
    const key = `${r}:${c}`;
    if (!this.sheet.cells.has(key)) this.sheet.cells.set(key, {});
    return this.sheet.cells.get(key);
  }

  getValue() { return this._cell(this.row, this.col).value ?? ""; }

  setValue(v) {
    this._cell(this.row, this.col).value = v;
    this.sheet._trackExtent(this.row, this.col);
    return this;
  }

  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowOut = [];
      for (let c = 0; c < this.numCols; c++) {
        rowOut.push(this._cell(this.row + r, this.col + c).value ?? "");
      }
      out.push(rowOut);
    }
    return out;
  }

  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        this._cell(this.row + r, this.col + c).value = values[r][c];
      }
    }
    this.sheet._trackExtent(this.row + values.length - 1, this.col + (values[0] ? values[0].length - 1 : 0));
    return this;
  }

  setNumberFormat(fmt) { this._forEach((cell) => { cell.numberFormat = fmt; }); return this; }
  setFontFamily(v) { this._forEach((cell) => { cell.fontFamily = v; }); return this; }
  setFontSize(v) { this._forEach((cell) => { cell.fontSize = v; }); return this; }
  setFontWeight(v) { this._forEach((cell) => { cell.fontWeight = v; }); return this; }
  setFontStyle(v) { this._forEach((cell) => { cell.fontStyle = v; }); return this; }
  setFontColor(v) { this._forEach((cell) => { cell.fontColor = v; }); return this; }
  setHorizontalAlignment(v) { this._forEach((cell) => { cell.hAlign = v; }); return this; }
  setVerticalAlignment(v) { this._forEach((cell) => { cell.vAlign = v; }); return this; }
  setWrap(v) { this._forEach((cell) => { cell.wrap = v; }); return this; }
  setBackground(v) { this._forEach((cell) => { cell.bg = v; }); return this; }
  setNote(text) { this._cell(this.row, this.col).note = text; return this; }
  setDataValidation(rule) { this._forEach((cell) => { cell.validation = rule; }); return this; }
  setFormula(text) { this._cell(this.row, this.col).value = text; this.sheet._trackExtent(this.row, this.col); return this; }
  merge() { return this; }

  _forEach(fn) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) fn(this._cell(this.row + r, this.col + c));
    }
  }
}

function _a1ToRowCol(a1) {
  const m = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!m) throw new Error("bad A1 ref: " + a1);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [parseInt(m[2], 10), col];
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this.cells = new Map();
    this.frozenRows = 0;
    this.frozenCols = 0;
    this.columnWidths = {};
    this.rowHeights = {};
    this.maxRow = 0;
    this.maxCol = 0;
    this.charts = [];
  }

  getName() { return this.name; }

  _trackExtent(r, c) {
    if (r > this.maxRow) this.maxRow = r;
    if (c > this.maxCol) this.maxCol = c;
  }

  getRange(a, b, c, d) {
    if (typeof a === "string") {
      if (a.indexOf(":") !== -1) {
        const [ref1, ref2] = a.split(":");
        const [r1, c1] = _a1ToRowCol(ref1);
        const [r2, c2] = _a1ToRowCol(ref2);
        return new MockRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
      }
      const [r, c1] = _a1ToRowCol(a);
      return new MockRange(this, r, c1, 1, 1);
    }
    return new MockRange(this, a, b, c || 1, d || 1);
  }

  getLastRow() { return this.maxRow; }
  getLastColumn() { return this.maxCol; }
  setColumnWidth(col, w) { this.columnWidths[col] = w; return this; }
  setRowHeight(row, h) { this.rowHeights[row] = h; return this; }
  setFrozenRows(n) { this.frozenRows = n; return this; }
  setFrozenColumns(n) { this.frozenCols = n; return this; }

  newChart() {
    const spec = { ranges: [], options: {} };
    const builder = {
      setChartType: (t) => { spec.type = t; return builder; },
      addRange: (r) => { spec.ranges.push(r); return builder; },
      setOption: (k, v) => { spec.options[k] = v; return builder; },
      setPosition: (...pos) => { spec.position = pos; return builder; },
      build: () => spec,
    };
    return builder;
  }

  insertChart(chart) { this.charts.push(chart); return this; }
}

class MockSpreadsheet {
  constructor() {
    this.sheets = [new MockSheet("Sheet1")];
    this.activeSheet = this.sheets[0];
  }

  insertSheet(name) {
    const sheet = new MockSheet(name);
    this.sheets.push(sheet);
    return sheet;
  }

  deleteSheet(sheet) {
    const idx = this.sheets.indexOf(sheet);
    if (idx === -1) throw new Error("deleteSheet: not found: " + sheet.getName());
    if (this.sheets.length <= 1) throw new Error("cannot delete the only sheet");
    this.sheets.splice(idx, 1);
  }

  getSheetByName(name) {
    return this.sheets.find((s) => s.getName() === name) || null;
  }

  getSheets() { return this.sheets.slice(); }

  setActiveSheet(sheet) { this.activeSheet = sheet; }

  moveActiveSheet(pos) {
    const idx = this.sheets.indexOf(this.activeSheet);
    this.sheets.splice(idx, 1);
    this.sheets.splice(pos - 1, 0, this.activeSheet);
  }
}

function installGlobals() {
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => global.__mockSS,
    newDataValidation: () => {
      const spec = {};
      const builder = {
        requireValueInRange: (range, showDropdown) => { spec.range = range; spec.showDropdown = showDropdown; return builder; },
        setAllowInvalid: (v) => { spec.allowInvalid = v; return builder; },
        build: () => spec,
      };
      return builder;
    },
    getUi: () => ({
      alert: () => {},
      createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }),
      ButtonSet: { OK: "OK", YES_NO: "YES_NO" },
      Button: { YES: "YES", NO: "NO" },
    }),
  };
  global.Charts = { ChartType: { LINE: "LINE" } };
  global.__mockSS = new MockSpreadsheet();
}

module.exports = { MockSpreadsheet, MockSheet, MockRange, installGlobals };
