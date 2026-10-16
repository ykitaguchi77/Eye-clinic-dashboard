/**
 * Apps Script のごく一部を真似る、テスト用のモック（Node でもブラウザでも動く）
 * SpreadsheetApp / Session / LockService / Utilities / HtmlService / DriveApp / UrlFetchApp / ScriptApp / CalendarApp
 */
(function (root) {
  function pad(n, w) { return String(n).padStart(w || 2, '0'); }

  function colIndex(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  class MockSheet {
    constructor(ss, name) {
      this.ss = ss; this.name = name; this.data = []; this.maxRows = 1000; this.maxCols = 26;
      this.formats = {}; this.weights = {}; this.backgrounds = {}; this.colors = {}; this.borders = []; this.widths = {}; this.frozen = 0;
    }
    getName() { return this.name; }
    setName(n) { this.name = n; return this; }
    getMaxRows() { return this.maxRows; }
    getMaxColumns() { return this.maxCols; }
    insertRowsAfter(_after, n) { this.maxRows += n; return this; }
    insertColumnsAfter(_after, n) { this.maxCols += n; return this; }
    getLastRow() {
      for (let r = this.data.length; r >= 1; r--) {
        const row = this.data[r - 1];
        if (row && row.some(v => v !== '' && v !== null && v !== undefined)) return r;
      }
      return 0;
    }
    getLastColumn() {
      let m = 0;
      this.data.forEach(row => { if (row) row.forEach((v, i) => { if (v !== '' && v !== null && v !== undefined) m = Math.max(m, i + 1); }); });
      return m;
    }
    cell(r, c) { const row = this.data[r - 1]; return row && row[c - 1] !== undefined ? row[c - 1] : ''; }
    put(r, c, v) {
      if (r > this.maxRows || c > this.maxCols) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
      while (this.data.length < r) this.data.push([]);
      const row = this.data[r - 1];
      while (row.length < c) row.push('');
      row[c - 1] = v;
    }
    getRange(a, b, c, d) {
      if (typeof a === 'string') {
        const m = /^([A-Z]+)(\d*)(?::([A-Z]+)(\d*))?$/.exec(a);
        if (!m) throw new Error('bad A1 ' + a);
        const c1 = colIndex(m[1]), r1 = m[2] ? Number(m[2]) : 1;
        const c2 = m[3] ? colIndex(m[3]) : c1, r2 = m[4] ? Number(m[4]) : (m[3] || !m[2] ? this.maxRows : r1);
        return new MockRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
      }
      return new MockRange(this, a, b, c || 1, d || 1);
    }
    getDataRange() { return new MockRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    appendRow(vals) {
      const r = this.getLastRow() + 1;
      if (r > this.maxRows) this.maxRows = r;
      vals.forEach((v, i) => this.put(r, i + 1, v));
      return this;
    }
    setFrozenRows(n) { this.frozen = n; return this; }
    setColumnWidth(c, w) { this.widths[c] = w; return this; }
  }

  class MockRange {
    constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
    each(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); }
    setValues(vals) {
      if (vals.length !== this.nr || vals.some(row => row.length !== this.nc)) throw new Error('The number of rows or columns in the data does not match the range.');
      this.each((r, c, i, j) => this.sh.put(r, c, vals[i][j]));
      return this;
    }
    setValue(v) { this.each((r, c) => this.sh.put(r, c, v)); return this; }
    getValues() {
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const row = [];
        for (let j = 0; j < this.nc; j++) row.push(this.sh.cell(this.r + i, this.c + j));
        out.push(row);
      }
      return out;
    }
    getValue() { return this.sh.cell(this.r, this.c); }
    insertCheckboxes() {
      this.each((r, c) => { if (this.sh.cell(r, c) === '') this.sh.put(r, c, false); });
      return this;
    }
    store(map, val2d, single) {
      this.each((r, c, i, j) => { map[r + ',' + c] = single !== undefined ? single : val2d[i][j]; });
      return this;
    }
    setFontWeight(w) { return this.store(this.sh.weights, null, w); }
    setFontWeights(w) { return this.store(this.sh.weights, w); }
    setBackground(b) { return this.store(this.sh.backgrounds, null, b); }
    setBackgrounds(b) { return this.store(this.sh.backgrounds, b); }
    setFontColors(c) { return this.store(this.sh.colors, c); }
    setNumberFormat(f) { return this.store(this.sh.formats, null, f); }
    setDataValidation() { return this; }
    setBorder() { this.sh.borders.push([this.r, this.c, this.nr, this.nc]); return this; }
    merge() { return this; }
  }

  class MockSpreadsheet {
    constructor(name) { this.name = name; this.sheets = []; this.id = 'ss' + Math.random().toString(36).slice(2); }
    getId() { return this.id; }
    getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
    getSheets() { return this.sheets.slice(); }
    insertSheet(n) {
      if (this.getSheetByName(n)) throw new Error('A sheet with the name "' + n + '" already exists.');
      const s = new MockSheet(this, n); this.sheets.push(s); return s;
    }
    deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
    toast() { throw new Error('toast is not available in web app context'); }
  }

  const state = { active: null, created: [], user: 'admin@example.com', effective: 'admin@example.com', calendar: [] };

  const SpreadsheetApp = {
    getActive() { return state.active; },
    create(name) { const ss = new MockSpreadsheet(name); ss.insertSheet('Sheet1'); state.created.push(ss); return ss; },
    flush() {},
    newDataValidation() { const b = { requireValueInList() { return b; }, build() { return {}; } }; return b; },
    BorderStyle: { SOLID: 'SOLID' }
  };
  const Session = {
    getActiveUser() { return { getEmail: () => state.user }; },
    getEffectiveUser() { return { getEmail: () => state.effective }; },
    getScriptTimeZone() { return 'Asia/Tokyo'; }
  };
  const LockService = { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } };
  const Utilities = {
    // テストはプロセスのタイムゾーンを Asia/Tokyo にして動かす前提（ローカル時刻で書式化する）
    formatDate(d, _tz, fmt) {
      return fmt.replace('yyyy', d.getFullYear()).replace('yy', pad(d.getFullYear() % 100)).replace('MM', pad(d.getMonth() + 1))
        .replace('dd', pad(d.getDate())).replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes())).replace('ss', pad(d.getSeconds()));
    },
    base64Encode(bytes) {
      if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
      let s = ''; bytes.forEach(b => { s += String.fromCharCode(b & 255); }); return btoa(s);
    }
  };
  const out = { setTitle() { return out; }, addMetaTag() { return out; } };
  const HtmlService = { createHtmlOutput() { return out; }, createHtmlOutputFromFile() { return out; } };
  const DriveApp = { getFileById(id) { return { setTrashed() { state.created = state.created.filter(s => s.id !== id || (s.trashed = true, true)); } }; } };
  const ScriptApp = { getOAuthToken() { return 'token'; } };
  // 出力のテスト用：作ったスプレッドシートの中身を JSON にして「xlsx の中身」として返す
  const UrlFetchApp = {
    fetch(url) {
      const id = /\/d\/([^/]+)\//.exec(url)[1];
      const ss = state.created.find(s => s.id === id);
      const dump = ss.sheets.map(s => ({
        name: s.name, data: s.data.map(row => row.map(v => v instanceof Date ? { date: Utilities.formatDate(v, '', 'yyyy-MM-dd') } : v)),
        formats: s.formats, weights: s.weights, backgrounds: s.backgrounds, colors: s.colors, borders: s.borders, widths: s.widths, frozen: s.frozen
      }));
      const json = JSON.stringify(dump);
      const bytes = typeof Buffer !== 'undefined' ? Array.from(Buffer.from(json, 'utf8')) : Array.from(new TextEncoder().encode(json));
      return { getResponseCode: () => 200, getContent: () => bytes };
    }
  };
  const CalendarApp = {
    getCalendarById() {
      return {
        createAllDayEvent(title, s, e) { const ev = { id: 'ev' + state.calendar.length, title, s, e }; state.calendar.push(ev); return { getId: () => ev.id }; },
        getEventById(id) {
          const ev = state.calendar.find(x => x.id === id);
          return ev ? { setTitle(t) { ev.title = t; }, setAllDayDates(s, e) { ev.s = s; ev.e = e; }, deleteEvent() { ev.deleted = true; } } : null;
        }
      };
    }
  };

  root.GasMock = {
    state, MockSpreadsheet,
    reset() { state.active = new MockSpreadsheet('test'); state.active.insertSheet('シート1'); state.created = []; state.calendar = []; return state.active; },
    globals: { SpreadsheetApp, Session, LockService, Utilities, HtmlService, DriveApp, ScriptApp, UrlFetchApp, CalendarApp }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
