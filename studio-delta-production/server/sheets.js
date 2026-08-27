const { google } = require("googleapis");
const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;

function colA1(n) {
  let s = "";
  let x = n;
  while (x > 0) {
    const m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

function a1(r1, c1, r2, c2) {
  return colA1(c1) + r1 + ":" + colA1(c2) + r2;
}

function serialToDate(serial) {
  const utcMs = Math.round((Number(serial) - 25569) * 86400000);
  return new Date(utcMs - SAST_OFFSET_MS);
}

function isDateSerial(n) {
  return typeof n === "number" && isFinite(n) && n >= 20000 && n <= 120000;
}

function dateToSheetString(d) {
  const sast = new Date(d.getTime() + SAST_OFFSET_MS);
  const y = sast.getUTCFullYear();
  const m = String(sast.getUTCMonth() + 1).padStart(2, "0");
  const day = String(sast.getUTCDate()).padStart(2, "0");
  const hh = String(sast.getUTCHours()).padStart(2, "0");
  const mm = String(sast.getUTCMinutes()).padStart(2, "0");
  const ss = String(sast.getUTCSeconds()).padStart(2, "0");
  return y + "-" + m + "-" + day + " " + hh + ":" + mm + ":" + ss;
}

function coerceRead(v) {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "boolean") return v;
  if (typeof v === "number") {
    if (isDateSerial(v)) return serialToDate(v);
    return v;
  }
  if (v instanceof Date) return v;
  if (typeof v === "string") {
    if (/^\d{4}-\d{2}-\d{2}/.test(v) || /^\d{1,2}\/\d{1,2}\/\d{4}/.test(v)) {
      const d = new Date(v);
      if (!isNaN(d.getTime())) return d;
    }
    return v;
  }
  return v;
}

function coerceWrite(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return dateToSheetString(v);
  if (typeof v === "boolean" || typeof v === "number") return v;
  return String(v);
}

class Range {
  constructor(sheet, r, c, n, m) {
    this.sheet = sheet;
    this.r = r;
    this.c = c;
    this.n = n || 1;
    this.m = m || 1;
  }
  getValues() {
    this.sheet.ensureGrid();
    const out = [];
    for (let i = 0; i < this.n; i++) {
      const row = [];
      const rr = this.r + i - 1;
      const src = this.sheet.grid[rr] || [];
      for (let j = 0; j < this.m; j++) {
        const v = src[this.c + j - 1];
        row.push(v === undefined || v === null ? "" : v);
      }
      out.push(row);
    }
    return out;
  }
  getValue() {
    return this.getValues()[0][0];
  }
  setValues(values) {
    this.sheet.ensureGrid();
    for (let i = 0; i < values.length; i++) {
      const rr = this.r + i - 1;
      while (this.sheet.grid.length <= rr) this.sheet.grid.push([]);
      for (let j = 0; j < values[i].length; j++) {
        const cc = this.c + j - 1;
        while (this.sheet.grid[rr].length <= cc) this.sheet.grid[rr].push("");
        this.sheet.grid[rr][cc] = values[i][j];
      }
    }
    const endRow = this.r + values.length - 1;
    const endCol = this.c + (values[0] ? values[0].length : 1) - 1;
    if (endRow > this.sheet.lastRow) this.sheet.lastRow = endRow;
    if (endCol > this.sheet.lastCol) this.sheet.lastCol = endCol;
    this.sheet.dirty = true;
    return this;
  }
  setValue(v) {
    return this.setValues([[v]]);
  }
  clearContent() {
    const blank = [];
    for (let i = 0; i < this.n; i++) {
      const row = [];
      for (let j = 0; j < this.m; j++) row.push("");
      blank.push(row);
    }
    return this.setValues(blank);
  }
}

class Sheet {
  constructor(book, props) {
    this.book = book;
    this.title = props.title;
    this.sheetId = props.sheetId;
    this.grid = null;
    this.lastRow = 0;
    this.lastCol = 0;
    this.dirty = false;
    this.hidden = !!(props.hidden || (props.gridProperties && props.gridProperties.hidden));
    this.isNew = !!props.isNew;
    this.loadedLastRow = 0;
  }
  ensureGrid() {
    if (!this.grid) this.grid = [];
  }
  getLastRow() {
    this.ensureGrid();
    return this.lastRow;
  }
  getLastColumn() {
    this.ensureGrid();
    return this.lastCol || 1;
  }
  getRange(r, c, n, m) {
    if (typeof r === "string") throw new Error("A1 getRange is not used");
    return new Range(this, r, c, n || 1, m || 1);
  }
  getDataRange() {
    this.ensureGrid();
    const rows = Math.max(1, this.lastRow);
    const cols = Math.max(1, this.lastCol);
    return this.getRange(1, 1, rows, cols);
  }
  appendRow(row) {
    this.ensureGrid();
    const next = this.lastRow + 1;
    this.getRange(next, 1, 1, row.length).setValues([row]);
    return this;
  }
  deleteRow(rowNum) {
    this.ensureGrid();
    const idx = rowNum - 1;
    if (idx >= 0 && idx < this.grid.length) this.grid.splice(idx, 1);
    this.lastRow = Math.max(0, this.lastRow - 1);
    this.dirty = true;
    return this;
  }
  hideSheet() {
    this.hidden = true;
    this.book.pendingHides.push(this.title);
    return this;
  }
}

class Spreadsheet {
  constructor(client, spreadsheetId) {
    this.client = client;
    this.spreadsheetId = spreadsheetId;
    this.sheetsByName = {};
    this.pendingAdds = [];
    this.pendingHides = [];
  }
  async load() {
    const sheetsApi = google.sheets({ version: "v4", auth: this.client });
    this.sheetsApi = sheetsApi;
    const meta = await sheetsApi.spreadsheets.get({
      spreadsheetId: this.spreadsheetId,
      fields: "sheets.properties"
    });
    const props = (meta.data.sheets || []).map((s) => s.properties);
    await Promise.all(props.map(async (p) => {
      const sheet = new Sheet(this, p);
      const res = await sheetsApi.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: "'" + p.title.replace(/'/g, "''") + "'",
        valueRenderOption: "UNFORMATTED_VALUE",
        dateTimeRenderOption: "SERIAL_NUMBER"
      });
      const values = res.data.values || [];
      sheet.grid = values.map((row) => row.map(coerceRead));
      sheet.lastRow = values.length;
      sheet.lastCol = values.reduce((m, row) => Math.max(m, row.length), 0);
      sheet.loadedLastRow = values.length;
      this.sheetsByName[p.title] = sheet;
    }));
    return this;
  }
  getSheetByName(name) {
    return this.sheetsByName[name] || null;
  }
  insertSheet(name) {
    if (this.sheetsByName[name]) return this.sheetsByName[name];
    const sheet = new Sheet(this, { title: name, sheetId: null, isNew: true });
    sheet.grid = [];
    sheet.lastRow = 0;
    sheet.lastCol = 0;
    this.sheetsByName[name] = sheet;
    this.pendingAdds.push(name);
    return sheet;
  }
  flushSyncPlaceholder() {}
  async flush() {
    const sheetsApi = this.sheetsApi;
    const requests = [];
    for (const name of this.pendingAdds) {
      requests.push({ addSheet: { properties: { title: name } } });
    }
    if (requests.length) {
      const added = await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: { requests }
      });
      const replies = added.data.replies || [];
      for (let i = 0; i < replies.length; i++) {
        const props = replies[i].addSheet && replies[i].addSheet.properties;
        if (props) {
          const sheet = this.sheetsByName[props.title];
          if (sheet) {
            sheet.sheetId = props.sheetId;
            sheet.isNew = false;
          }
        }
      }
      this.pendingAdds = [];
    }

    const dimReqs = [];
    for (const title of Object.keys(this.sheetsByName)) {
      const sheet = this.sheetsByName[title];
      if (sheet.sheetId == null) continue;
      if (sheet.dirty && sheet.loadedLastRow > sheet.lastRow) {
        dimReqs.push({
          deleteDimension: {
            range: {
              sheetId: sheet.sheetId,
              dimension: "ROWS",
              startIndex: Math.max(0, sheet.lastRow),
              endIndex: sheet.loadedLastRow
            }
          }
        });
      }
    }
    if (dimReqs.length) {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: { requests: dimReqs }
      });
    }

    const hideReqs = [];
    for (const title of this.pendingHides) {
      const sheet = this.sheetsByName[title];
      if (sheet && sheet.sheetId != null) {
        hideReqs.push({
          updateSheetProperties: {
            properties: { sheetId: sheet.sheetId, hidden: true },
            fields: "hidden"
          }
        });
      }
    }
    this.pendingHides = [];
    if (hideReqs.length) {
      await sheetsApi.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: { requests: hideReqs }
      });
    }

    const data = [];
    for (const title of Object.keys(this.sheetsByName)) {
      const sheet = this.sheetsByName[title];
      if (!sheet.dirty) continue;
      sheet.ensureGrid();
      const rows = sheet.lastRow;
      const cols = Math.max(1, sheet.lastCol);
      if (rows < 1) {
        sheet.dirty = false;
        sheet.loadedLastRow = 0;
        continue;
      }
      const values = [];
      for (let i = 0; i < rows; i++) {
        const src = sheet.grid[i] || [];
        const row = [];
        for (let j = 0; j < cols; j++) row.push(coerceWrite(src[j] === undefined ? "" : src[j]));
        values.push(row);
      }
      data.push({
        range: "'" + title.replace(/'/g, "''") + "'!" + a1(1, 1, rows, cols),
        values
      });
      sheet.dirty = false;
      sheet.loadedLastRow = rows;
    }
    if (data.length) {
      await sheetsApi.spreadsheets.values.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: { valueInputOption: "USER_ENTERED", data }
      });
    }
  }
}

function createSpreadsheetApp(workbook) {
  return {
    openById() { return workbook; },
    getActiveSpreadsheet() { return workbook; },
    flush() { workbook.flushSyncPlaceholder(); }
  };
}

module.exports = {
  Spreadsheet,
  Sheet,
  Range,
  createSpreadsheetApp,
  coerceRead,
  coerceWrite
};
