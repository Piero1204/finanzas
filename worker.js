/* Motor local: corre el mismo código del sistema dentro del teléfono y guarda los datos en el propio navegador. */
'use strict';
const DB = 'finanzas-local', ST = 'estado', KEY = 'datos';
let guardando = Promise.resolve(), pendiente = false;

function abrirDB() {
  return new Promise(function (ok, no) {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = function () { r.result.createObjectStore(ST); };
    r.onsuccess = function () { ok(r.result); };
    r.onerror = function () { no(r.error); };
  });
}
function idb(modo, fn) {
  return abrirDB().then(function (db) {
    return new Promise(function (ok, no) {
      const t = db.transaction(ST, modo), rq = fn(t.objectStore(ST));
      t.oncomplete = function () { db.close(); ok(rq && rq.result); };
      t.onerror = t.onabort = function () { db.close(); no(t.error); };
    });
  });
}
function serializar(store) {
  return JSON.stringify({ v: 1, nextId: store.nextId, props: store.props, sheets: Object.keys(store.sheets).map(function (k) { const s = store.sheets[k]; return { name: s.name, id: s.id, d: s.d }; }) },
    function (k, v) { const o = this[k]; return o instanceof Date ? { $d: o.getTime() } : v; });
}
function restaurar(txt) {
  const o = JSON.parse(txt, function (k, v) { return v && typeof v === 'object' && v.$d !== undefined && Object.keys(v).length === 1 ? new Date(v.$d) : v; });
  const st = self.__store;
  st.sheets = {}; st.props = o.props || {}; st.nextId = o.nextId || 1;
  (o.sheets || []).forEach(function (s) { const h = new self.Sheet(s.name, s.id); h.d = s.d; st.sheets[s.name] = h; });
}
function guardar() {
  guardando = guardando.then(function () { const txt = serializar(self.__store); return idb('readwrite', function (s) { return s.put(txt, KEY); }); }).catch(function (e) { self.postMessage({ aviso: 'No se pudo guardar: ' + e }); });
  return guardando;
}

// ---- Red síncrona (la usan Gemini y el tipo de cambio) ----
self.__host = {
  dirty: function () { pendiente = true; },
  fetchSync: function (url, o) {
    const x = new XMLHttpRequest();
    x.open((o.method || 'get').toUpperCase(), url, false);
    const h = o.headers || {};
    Object.keys(h).forEach(function (k) { x.setRequestHeader(k, h[k]); });
    if (o.contentType && !h['Content-Type']) x.setRequestHeader('Content-Type', o.contentType);
    x.send(o.payload === undefined ? null : o.payload);
    return { status: x.status, text: x.responseText };
  }
};

importScripts('shim.js', 'gs-bundle.js', 'prelude.js');

let listo = abrirDB().then(function () { return idb('readonly', function (s) { return s.get(KEY); }); }).then(function (txt) { if (txt) restaurar(txt); });

function sinSecretos(props) {
  const o = {};
  Object.keys(props).forEach(function (k) { if (!/PIN|SECRET|KEY|TOKEN/i.test(k) && k !== 'LOCAL_TAREAS_DIA') o[k] = props[k]; });
  return o;
}
function aleatorio() { const a = new Uint8Array(16); crypto.getRandomValues(a); return Array.from(a, function (b) { return b.toString(16).padStart(2, '0'); }).join(''); }
function pinValido(p) { p = String(p || '').trim(); if (!/^\d{4,8}$/.test(p)) throw new Error('El PIN debe tener de 4 a 8 números.'); return p; }

const ops = {
  estado: function () { return { tienePin: !!self.PropertiesService.getScriptProperties().getProperty('DASHBOARD_PIN') }; },
  crearPin: function (a) {
    const P = self.PropertiesService.getScriptProperties();
    if (P.getProperty('DASHBOARD_PIN')) throw new Error('Ya hay un PIN creado.');
    P.setProperty('DASHBOARD_PIN', pinValido(a.pin)); P.setProperty('SESSION_SECRET', aleatorio());
    return true;
  },
  cambiarPin: function (a) {
    const P = self.PropertiesService.getScriptProperties();
    if (String(a.actual || '').trim() !== P.getProperty('DASHBOARD_PIN')) throw new Error('El PIN actual no es correcto.');
    P.setProperty('DASHBOARD_PIN', pinValido(a.nuevo)); P.setProperty('SESSION_SECRET', aleatorio());
    return true;
  },
  exportar: function () {
    const st = self.__store;
    return serializar({ nextId: st.nextId, props: sinSecretos(st.props), sheets: st.sheets });
  },
  importar: function (a) {
    const P = self.PropertiesService.getScriptProperties();
    const pin = P.getProperty('DASHBOARD_PIN'), sec = P.getProperty('SESSION_SECRET'), gem = P.getProperty('GEMINI_API_KEY');
    let o; try { o = JSON.parse(a.json); } catch (e) { throw new Error('El archivo no es una copia válida.'); }
    if (!o || o.v !== 1 || !Array.isArray(o.sheets)) throw new Error('El archivo no es una copia válida.');
    if (!pin) a.pin = pinValido(a.pin);
    restaurar(a.json);
    const Q = self.PropertiesService.getScriptProperties();
    Q.setProperty('DASHBOARD_PIN', pin || a.pin); Q.setProperty('SESSION_SECRET', sec || aleatorio()); if (gem) Q.setProperty('GEMINI_API_KEY', gem);
    return true;
  },
  borrarTodo: function () { const st = self.__store; st.sheets = {}; st.props = {}; st.nextId = 1; return true; },
  setProp: function (a) { if (!/^(GEMINI_API_KEY|GEMINI_MODEL)$/.test(a.k)) throw new Error('No permitido.'); const P = self.PropertiesService.getScriptProperties(); if (String(a.v || '').trim()) P.setProperty(a.k, a.v); else P.deleteProperty(a.k); return true; },
  getProp: function (a) { if (!/^(GEMINI_API_KEY)$/.test(a.k)) throw new Error('No permitido.'); return !!self.PropertiesService.getScriptProperties().getProperty(a.k); },
  tareas: function () { return self.tareasLocales_(); },
  llamar: function (a) { return self.apiLlamar(a.nombre, a.args, a.token); }
};

let cola = Promise.resolve();
self.onmessage = function (ev) {
  const m = ev.data;
  cola = cola.then(function () { return listo; }).then(function () {
    let res, err;
    pendiente = false;
    try { res = ops[m.op](m.a || {}); } catch (e) { err = e.message || String(e); }
    const g = pendiente || m.op === 'borrarTodo' || m.op === 'importar' ? guardar() : Promise.resolve();
    return g.then(function () { self.postMessage(err ? { id: m.id, err: err } : { id: m.id, res: res }); });
  });
};
