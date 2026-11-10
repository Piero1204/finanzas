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
    const cfg = self.APP_CONFIG || {};
    const bolsa = /^https:\/\/(query[12]\.finance\.yahoo\.com|stooq\.com)\//.test(url);
    if (bolsa) {
      if (!cfg.PROXY_URL) throw new Error('La bolsa no está configurada');
      url = cfg.PROXY_URL + (cfg.PROXY_URL.indexOf('?') === -1 ? '?' : '&') + 'u=' + encodeURIComponent(url);
    }
    x.open((o.method || 'get').toUpperCase(), url, false);
    const h = bolsa ? {} : (o.headers || {});
    Object.keys(h).forEach(function (k) { x.setRequestHeader(k, h[k]); });
    if (o.contentType && !h['Content-Type']) x.setRequestHeader('Content-Type', o.contentType);
    x.send(o.payload === undefined ? null : o.payload);
    return { status: x.status, text: x.responseText };
  }
};

const VQ = self.location.search || '';   // ?v=versión: evita que el navegador reutilice piezas viejas guardadas
importScripts('config.js' + VQ, 'shim.js' + VQ, 'gs-bundle.js' + VQ, 'prelude.js' + VQ);

let listo = abrirDB().then(function () { return idb('readonly', function (s) { return s.get(KEY); }); }).then(function (txt) { if (txt) restaurar(txt); });

/* ---- PIN: nunca se guarda tal cual, solo su huella (PBKDF2 + sal) ---- */
const PIN_ITER = 10000;
function hex(b) { return Array.from(b, function (x) { return x.toString(16).padStart(2, '0'); }).join(''); }
function deHex(h) { const o = new Uint8Array(h.length / 2); for (let i = 0; i < o.length; i++) o[i] = parseInt(h.substr(i * 2, 2), 16); return o; }
function igualSeguro(a, b) { a = String(a); b = String(b); let d = a.length ^ b.length; for (let i = 0; i < Math.max(a.length, b.length); i++) d |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0); return d === 0; }
function huellaPin(pin) {
  const sal = new Uint8Array(16); crypto.getRandomValues(sal);
  return 'pbkdf2$' + PIN_ITER + '$' + hex(sal) + '$' + hex(self.__cripto.pbkdf2(pin, sal, PIN_ITER));
}
function pinCorrecto(pin, guardado) {
  pin = String(pin || '').trim(); guardado = String(guardado || '');
  const p = guardado.split('$');
  if (p[0] === 'pbkdf2' && p.length === 4) return igualSeguro(hex(self.__cripto.pbkdf2(pin, deHex(p[2]), Number(p[1]) || PIN_ITER)), p[3]);
  return !!guardado && igualSeguro(pin, guardado); // PIN antiguo en texto: se convierte al entrar
}
/* Intentos fallidos guardados en el teléfono (recargar la app no los reinicia). Desde el 5.º, espera creciente: 30 s, 1 min, 2 min… hasta 1 h */
function esperaRestante() { return Math.max(0, Number(self.PropertiesService.getScriptProperties().getProperty('PIN_BLOQUEO_HASTA') || 0) - Date.now()); }
function textoEspera(ms) { const s = Math.ceil(ms / 1000); return s < 60 ? s + ' segundos' : Math.ceil(s / 60) + ' minuto' + (Math.ceil(s / 60) === 1 ? '' : 's'); }
function revisarPin(pin) {
  const P = self.PropertiesService.getScriptProperties();
  const guardado = P.getProperty('DASHBOARD_PIN');
  if (!guardado) throw new Error('Aún no has creado un PIN.');
  const espera = esperaRestante();
  if (espera > 0) throw new Error('Demasiados intentos. Espera ' + textoEspera(espera) + '.');
  if (!pinCorrecto(pin, guardado)) {
    const f = Number(P.getProperty('PIN_FALLOS') || 0) + 1;
    P.setProperty('PIN_FALLOS', String(f));
    if (f >= 5) {
      const ms = Math.min(30000 * Math.pow(2, f - 5), 3600000);
      P.setProperty('PIN_BLOQUEO_HASTA', String(Date.now() + ms));
      throw new Error('PIN incorrecto. Por seguridad espera ' + textoEspera(ms) + ' antes de intentar de nuevo.');
    }
    throw new Error('PIN incorrecto.' + (f >= 3 ? (5 - f === 1 ? ' Te queda 1 intento' : ' Te quedan ' + (5 - f) + ' intentos') + ' antes de una espera.' : ''));
  }
  if (P.getProperty('PIN_FALLOS')) P.deleteProperty('PIN_FALLOS');
  if (P.getProperty('PIN_BLOQUEO_HASTA')) P.deleteProperty('PIN_BLOQUEO_HASTA');
  if (String(guardado).indexOf('pbkdf2$') !== 0) P.setProperty('DASHBOARD_PIN', huellaPin(String(pin).trim()));
}
/* Entrada local: reemplaza la comparación en texto del código compartido con la hoja */
self.login_ = function (args) { revisarPin(args.pin); return { token: self.crearToken_() }; };
self.verificarToken_ = function (token) {
  const p = String(token || '').split('.');
  if (p.length !== 2 || !self.PropertiesService.getScriptProperties().getProperty('DASHBOARD_PIN')) return false;
  if (!(Number(p[0]) > Date.now())) return false;
  return igualSeguro(self.firmar_(p[0]), p[1]);
};

function sinSecretos(props) {
  const o = {};
  Object.keys(props).forEach(function (k) { if (!/PIN|SECRET|KEY|TOKEN/i.test(k) && k !== 'LOCAL_TAREAS_DIA') o[k] = props[k]; });
  return o;
}
function aleatorio() { const a = new Uint8Array(16); crypto.getRandomValues(a); return Array.from(a, function (b) { return b.toString(16).padStart(2, '0'); }).join(''); }
function pinValido(p) { p = String(p || '').trim(); if (!/^\d{4,8}$/.test(p)) throw new Error('El PIN debe tener de 4 a 8 números.'); return p; }

const ops = {
  version: function () { return self.APP_V || ''; },
  estado: function () { return { tienePin: !!self.PropertiesService.getScriptProperties().getProperty('DASHBOARD_PIN'), espera: esperaRestante() }; },
  crearPin: function (a) {
    const P = self.PropertiesService.getScriptProperties();
    if (P.getProperty('DASHBOARD_PIN')) throw new Error('Ya hay un PIN creado.');
    P.setProperty('DASHBOARD_PIN', huellaPin(pinValido(a.pin))); P.setProperty('SESSION_SECRET', aleatorio());
    return true;
  },
  cambiarPin: function (a) {
    const P = self.PropertiesService.getScriptProperties();
    const nuevo = pinValido(a.nuevo);
    try { revisarPin(a.actual); } catch (e) { throw new Error(/espera/i.test(e.message) ? e.message : 'El PIN actual no es correcto.'); }
    P.setProperty('DASHBOARD_PIN', huellaPin(nuevo)); P.setProperty('SESSION_SECRET', aleatorio());
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
    const valida = o && o.v === 1 && Array.isArray(o.sheets) && o.sheets.every(function (h) { return h && typeof h.name === 'string' && h.name.length < 100 && Array.isArray(h.d) && h.d.every(Array.isArray); });
    if (!valida) throw new Error('El archivo no es una copia válida.');
    if (!pin) a.pin = pinValido(a.pin);
    // Una copia nunca trae PIN, claves ni bloqueos: si alguien los mete a mano en el archivo, se ignoran
    o.props = sinSecretos(o.props && typeof o.props === 'object' ? o.props : {});
    restaurar(JSON.stringify(o));
    const Q = self.PropertiesService.getScriptProperties();
    Q.setProperty('DASHBOARD_PIN', pin || huellaPin(a.pin)); Q.setProperty('SESSION_SECRET', sec || aleatorio()); if (gem) Q.setProperty('GEMINI_API_KEY', gem);
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
    try { if (!Object.prototype.hasOwnProperty.call(ops, m.op)) throw new Error('Operación no válida.'); res = ops[m.op](m.a || {}); } catch (e) { err = e.message || String(e); }
    const g = pendiente || m.op === 'borrarTodo' || m.op === 'importar' ? guardar() : Promise.resolve();
    return g.then(function () { self.postMessage(err ? { id: m.id, err: err } : { id: m.id, res: res }); });
  });
};
