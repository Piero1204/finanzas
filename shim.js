/* Simula, dentro del navegador, los servicios de Google Apps Script que usa el sistema
   (hojas de cálculo, propiedades, caché, fechas, red). Los datos viven en memoria y el host los guarda en el teléfono. */
(function (G) {
  'use strict';
  const H = G.__host;

  // ---------- SHA-256 / HMAC (síncronos) ----------
  const K = new Uint32Array([0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2]);
  function sha256(bytes) {
    const l = bytes.length, bits = l * 8;
    const n = ((l + 9 + 63) >> 6) << 6, m = new Uint8Array(n);
    m.set(bytes); m[l] = 0x80;
    const dv = new DataView(m.buffer);
    dv.setUint32(n - 8, Math.floor(bits / 4294967296)); dv.setUint32(n - 4, bits >>> 0);
    const h = new Uint32Array([0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19]);
    const w = new Uint32Array(64);
    for (let o = 0; o < n; o += 64) {
      for (let i = 0; i < 16; i++) w[i] = dv.getUint32(o + i * 4);
      for (let i = 16; i < 64; i++) {
        const a = w[i - 15], b = w[i - 2];
        const s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
        const s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
        w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
      }
      let [a, b, c, d, e, f, g, hh] = h;
      for (let i = 0; i < 64; i++) {
        const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
        const ch = (e & f) ^ (~e & g);
        const t1 = (hh + S1 + ch + K[i] + w[i]) >>> 0;
        const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
        const mj = (a & b) ^ (a & c) ^ (b & c);
        const t2 = (S0 + mj) >>> 0;
        hh = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
      }
      h[0] = (h[0] + a) >>> 0; h[1] = (h[1] + b) >>> 0; h[2] = (h[2] + c) >>> 0; h[3] = (h[3] + d) >>> 0;
      h[4] = (h[4] + e) >>> 0; h[5] = (h[5] + f) >>> 0; h[6] = (h[6] + g) >>> 0; h[7] = (h[7] + hh) >>> 0;
    }
    const out = new Uint8Array(32);
    const o2 = new DataView(out.buffer);
    for (let i = 0; i < 8; i++) o2.setUint32(i * 4, h[i]);
    return out;
  }
  function utf8(s) { return new TextEncoder().encode(String(s)); }
  function hmac(keyB, msgB) {
    if (keyB.length > 64) keyB = sha256(keyB);
    const k = new Uint8Array(64); k.set(keyB);
    const ip = new Uint8Array(64 + msgB.length), op = new Uint8Array(96);
    for (let i = 0; i < 64; i++) { ip[i] = k[i] ^ 0x36; op[i] = k[i] ^ 0x5c; }
    ip.set(msgB, 64);
    op.set(sha256(ip), 64);
    return sha256(op);
  }
  function aBytes(x) { return typeof x === 'string' ? utf8(x) : Uint8Array.from(x, function (v) { return v & 255; }); }
  const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  function b64(bytes, web) {
    let s = '';
    for (let i = 0; i < bytes.length; i += 3) {
      const a = bytes[i], b = bytes[i + 1], c = bytes[i + 2];
      s += B64[a >> 2] + B64[((a & 3) << 4) | ((b || 0) >> 4)] + (b === undefined ? '=' : B64[((b & 15) << 2) | ((c || 0) >> 6)]) + (c === undefined ? '=' : B64[c & 63]);
    }
    return web ? s.replace(/\+/g, '-').replace(/\//g, '_') : s;
  }

  // ---------- Hojas ----------
  const store = { sheets: {}, nextId: 1, props: {} };
  G.__store = store;
  const dirty = function () { store.dirty = true; if (H.dirty) H.dirty(); };

  function Range(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  Range.prototype.getValues = function () {
    const o = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = this.sh.get(this.r + i, this.c + j); row.push(v === undefined ? '' : v); } o.push(row); }
    return o;
  };
  Range.prototype.getValue = function () { const v = this.sh.get(this.r, this.c); return v === undefined ? '' : v; };
  Range.prototype.setValues = function (vs) { const s = this; vs.forEach(function (row, i) { row.forEach(function (v, j) { s.sh.set(s.r + i, s.c + j, v); }); }); dirty(); return this; };
  Range.prototype.setValue = function (v) { this.sh.set(this.r, this.c, v); dirty(); return this; };
  Range.prototype.clear = Range.prototype.clearContent = function () { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); dirty(); return this; };
  Range.prototype.getRow = function () { return this.r; };
  Range.prototype.getColumn = function () { return this.c; };
  Range.prototype.getNumRows = function () { return this.nr; };
  Range.prototype.getNumColumns = function () { return this.nc; };
  function wrapRange(rg) { return new Proxy(rg, { get: function (t, k) { if (k in t) return t[k]; return function () { return this_; }; } }); }
  let this_;
  function mk(sh, r, c, nr, nc) { const rg = new Range(sh, r, c, nr, nc); const p = new Proxy(rg, { get: function (t, k) { if (k in t) return t[k]; return function () { return p; }; } }); return p; }
  const rangoVacio = new Proxy({ getValues: function () { return [[]]; }, getValue: function () { return ''; } }, { get: function (t, k) { if (k in t) return t[k]; return function () { return rangoVacio; }; } });

  function Sheet(name, id) { this.name = name; this.id = id; this.d = [[]]; }
  Sheet.prototype = {
    get: function (r, c) { return (this.d[r - 1] || [])[c - 1]; },
    set: function (r, c, v) { while (this.d.length < r) this.d.push([]); const row = this.d[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; dirty(); },
    lastRow: function () { let l = 0; this.d.forEach(function (r, i) { if (r.some(function (x) { return x !== '' && x !== undefined && x !== null; })) l = i + 1; }); return l; },
    width: function () { let w = 1; this.d.forEach(function (r) { if (r.length > w) w = r.length; }); return w; },
    getLastColumn: function () { return this.width(); }, getLastRow: function () { return this.lastRow(); }, getMaxRows: function () { return Math.max(1000, this.d.length); },
    insertRowsAfter: function () {}, getSheetId: function () { return this.id; }, getName: function () { return this.name; },
    deleteRow: function (n) { this.d.splice(n - 1, 1); dirty(); },
    setTabColor: function () {}, setFrozenRows: function () {}, autoResizeColumns: function () {}, setColumnWidths: function () {}, setRowHeight: function () {},
    clear: function () { this.d = [[]]; dirty(); },
    getDataRange: function () { return mk(this, 1, 1, this.lastRow() || 1, this.width()); },
    appendRow: function (vals) { const r = this.lastRow() + 1; for (let i = 0; i < vals.length; i++) this.set(r, i + 1, vals[i]); },
    getRange: function (a, b, c, d) { if (typeof a === 'string') return rangoVacio; return mk(this, a, b, c || 1, d || 1); }
  };
  G.Sheet = Sheet; // para restaurar desde el guardado
  const ss = {
    getSheetByName: function (n) { return store.sheets[n] || null; },
    insertSheet: function (n) { store.sheets[n] = new Sheet(n, store.nextId++); dirty(); return store.sheets[n]; },
    getSheets: function () { return Object.keys(store.sheets).map(function (k) { return store.sheets[k]; }); },
    getId: function () { return 'LOCAL'; }, getUrl: function () { return ''; }
  };
  G.SpreadsheetApp = {
    getActiveSpreadsheet: function () { return ss; }, flush: function () {},
    getUi: function () { throw new Error('Sin interfaz de menús'); },
    newDataValidation: function () { return { requireValueInList: function () { return this; }, setAllowInvalid: function () { return this; }, build: function () { return {}; } }; },
    BorderStyle: { SOLID: 1 }
  };
  G.LockService = { getScriptLock: function () { return { waitLock: function () {}, tryLock: function () { return true; }, releaseLock: function () {} }; } };

  // ---------- Caché y propiedades ----------
  const cacheM = {};
  G.CacheService = { getScriptCache: function () { return {
    get: function (k) { const e = cacheM[k]; if (!e) return null; if (e.x && e.x < Date.now()) { delete cacheM[k]; return null; } return e.v; },
    put: function (k, v, s) { cacheM[k] = { v: String(v), x: s ? Date.now() + s * 1000 : 0 }; },
    remove: function (k) { delete cacheM[k]; }
  }; } };
  G.PropertiesService = { getScriptProperties: function () { return {
    getProperty: function (k) { const v = store.props[k]; return v === undefined ? null : v; },
    setProperty: function (k, v) { store.props[k] = String(v); dirty(); },
    deleteProperty: function (k) { delete store.props[k]; dirty(); }
  }; } };

  G.Session = {
    getScriptTimeZone: function () { try { return Intl.DateTimeFormat().resolvedOptions().timeZone || 'America/Lima'; } catch (e) { return 'America/Lima'; } },
    getActiveUser: function () { return { getEmail: function () { return 'local'; } }; },
    getEffectiveUser: function () { return { getEmail: function () { return 'local'; } }; }
  };

  // ---------- Utilities ----------
  const p2 = function (n) { return String(n).padStart(2, '0'); };
  G.Utilities = {
    sleep: function () {},
    formatDate: function (d, tz, f) {
      return f.replace(/yyyy|MM|dd|HH|mm|ss|H|u/g, function (t) {
        switch (t) {
          case 'yyyy': return String(d.getFullYear()); case 'MM': return p2(d.getMonth() + 1); case 'dd': return p2(d.getDate());
          case 'HH': return p2(d.getHours()); case 'mm': return p2(d.getMinutes()); case 'ss': return p2(d.getSeconds());
          case 'H': return String(d.getHours()); case 'u': return String(d.getDay() === 0 ? 7 : d.getDay());
        }
      });
    },
    computeHmacSha256Signature: function (v, k) { return Array.from(hmac(aBytes(k), aBytes(v)), function (x) { return x > 127 ? x - 256 : x; }); },
    base64Encode: function (x) { return b64(aBytes(x), false); },
    base64EncodeWebSafe: function (x) { return b64(aBytes(x), true); },
    newBlob: function (c, t, n) { return { c: c, t: t, n: n, setName: function (x) { this.n = x; return this; }, getBytes: function () { return Array.from(aBytes(c)); }, getContentType: function () { return t; }, getName: function () { return this.n; } }; }
  };

  // ---------- Red (síncrona; el host decide cómo) ----------
  function resp(r) {
    return { getResponseCode: function () { return r.status; }, getContentText: function () { return r.text || ''; }, getHeaders: function () { return r.headers || {}; },
      getBlob: function () { return { setName: function () { return this; }, getContentType: function () { return 'application/octet-stream'; }, getBytes: function () { return []; } }; } };
  }
  G.UrlFetchApp = {
    fetch: function (u, o) {
      o = o || {};
      let r;
      try { r = H.fetchSync(u, o); } catch (e) { throw new Error('No se pudo conectar (' + e.message + ')'); }
      if (!o.muteHttpExceptions && r.status >= 400) throw new Error('Request failed for ' + u + ' returned code ' + r.status);
      return resp(r);
    },
    fetchAll: function (reqs) { return reqs.map(function (q) { const u = q.url; const o = Object.assign({}, q); delete o.url; return G.UrlFetchApp.fetch(u, o); }); }
  };

  // ---------- Lo que no existe en el teléfono ----------
  const triggerStub = { timeBased: function () { return this; }, everyDays: function () { return this; }, atHour: function () { return this; }, onWeekDay: function () { return this; }, onMonthDay: function () { return this; }, everyMinutes: function () { return this; }, create: function () { return this; } };
  G.ScriptApp = { getProjectTriggers: function () { return []; }, newTrigger: function () { return triggerStub; }, deleteTrigger: function () {}, getOAuthToken: function () { return ''; }, WeekDay: { SUNDAY: 1 } };
  const sinDrive = function () { throw new Error('Esta función no está disponible en la versión local.'); };
  G.DriveApp = { getFileById: sinDrive, getFoldersByName: sinDrive, createFolder: sinDrive };
  G.ContentService = { createTextOutput: function (t) { return { t: t }; } };
  G.HtmlService = {};
  G.Logger = { log: function () {} };
})(globalThis);
