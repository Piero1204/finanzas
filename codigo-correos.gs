// ======================================================================
// FINANZAS · lectura de consumos con tarjeta desde Gmail (OPCIONAL)
// En Apps Script de tu hoja: + (Archivos) > Script > nómbralo "Correos", pega TODO esto, guarda
// y ejecuta una vez "activarCorreosTarjetas". Necesita el código principal (codigo-hoja.gs) en el mismo proyecto.
// ======================================================================

/* ==== 33_CorreosTarjetas.gs ==== */
/**
 * ============================================================
 * CONSUMOS CON TARJETA DESDE TUS CORREOS
 * ============================================================
 * Cada 5 minutos revisa el Gmail de la cuenta dueña de esta hoja. Cuando llega el aviso de un banco de que usaste
 * una de tus tarjetas, lo registra como consumo en esa tarjeta (Tarjetas + Movimientos) y te avisa por Telegram.
 * - Solo registra si el correo es de una tarjeta que tienes en la hoja «Tarjetas» (busca el nombre del banco).
 * - Excluye los bancos de Ajustes › correos.excluir (por defecto: BBVA y Banca di Asti).
 * - Los pagos HACIA la tarjeta, devoluciones y rechazos no se registran como gasto.
 * - Cada correo se procesa una sola vez (queda anotado en la hoja «Correos», con lo que se hizo).
 * - Sirve también si los avisos llegan reenviados desde otra cuenta de Gmail.
 * Para activarlo ejecuta una vez activarCorreosTarjetas (pide permiso para leer Gmail).
 */

var CORREOS_ENC_ = ['MensajeID', 'Fecha', 'Remitente', 'Asunto', 'Resultado', 'Tarjeta', 'Monto', 'Moneda', 'Comercio', 'IDTarjeta'];
var CORREOS_EXCLUIR_DEF_ = 'bbva, banca di asti, bancadiasti';
var BANCOS_CONOCIDOS_ = ['interbank', 'bcp', 'scotiabank', 'bbva', 'diners', 'falabella', 'ripley', 'oh', 'banbif', 'pichincha', 'santander', 'amex', 'american express', 'mibanco', 'gnb', 'banca di asti', 'intesa', 'unicredit', 'revolut', 'n26'];
var PALABRAS_CORREO_ = '(consumo OR compra OR cargo OR transaccion OR transacción OR operacion OR operación OR realizaste OR "has realizado" OR acquisto OR pagamento OR addebito OR purchase OR "tarjeta de credito" OR "tarjeta de crédito")';

function activarCorreosTarjetas() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'revisarCorreosTarjetas') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('revisarCorreosTarjetas').timeBased().everyMinutes(5).create();
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('CORREOS_DESDE')) P.setProperty('CORREOS_DESDE', String(Date.now() - 6 * 3600000)); // no reimporta tu historial: solo lo de las últimas 6 h en adelante
  if (getAjuste_('correos.excluir', '') === '') setAjuste_('correos.excluir', CORREOS_EXCLUIR_DEF_);
  hojaCorreos_();
  GmailApp.getInboxUnreadCount(); // fuerza el permiso de Gmail ahora
  registrarLog_('Correos', 'Lectura de correos de tarjetas activada');
  avisar_('✅ Listo: cada 5 minutos se leen los avisos de consumo de tus tarjetas (' + tarjetasActivas_().map(function (t) { return t.banco; }).join(', ') + ').\nExcluidos: ' + getAjuste_('correos.excluir', CORREOS_EXCLUIR_DEF_));
}

function desactivarCorreosTarjetas() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'revisarCorreosTarjetas') ScriptApp.deleteTrigger(t); });
  avisar_('Lectura de correos desactivada.');
}

function hojaCorreos_() {
  const ss = getSS_();
  return ss.getSheetByName('Correos') || crearHojaBase_(ss, 'Correos', CORREOS_ENC_);
}

/** Disparador: revisa los correos nuevos */
function revisarCorreosTarjetas() {
  ejecutarSeguro_('Correos', function () { procesarCorreosTarjetas_(); });
}

function procesarCorreosTarjetas_() {
  const P = PropertiesService.getScriptProperties();
  const desde = Number(P.getProperty('CORREOS_DESDE') || (Date.now() - 6 * 3600000));
  const hoja = hojaCorreos_();
  const n = hoja.getLastRow();
  const vistos = {};
  if (n > 1) hoja.getRange(Math.max(2, n - 1500), 1, Math.min(n - 1, 1500), 1).getValues().forEach(function (f) { vistos[String(f[0])] = 1; });
  const consulta = 'after:' + Math.floor((desde - 86400000) / 1000) + ' ' + PALABRAS_CORREO_;
  const hilos = GmailApp.search(consulta, 0, 40);
  const tarjetas = tarjetasActivas_();
  const excluir = String(getAjuste_('correos.excluir', CORREOS_EXCLUIR_DEF_)).split(',').map(normalizarTxt_).filter(String);
  const resumen = [];
  hilos.forEach(function (h) {
    h.getMessages().forEach(function (m) {
      const id = m.getId();
      if (vistos[id] || m.getDate().getTime() < desde) return;
      vistos[id] = 1;
      const correo = { id: id, fecha: m.getDate(), de: m.getFrom(), asunto: m.getSubject(), cuerpo: m.getPlainBody() || textoDeHtml_(m.getBody()) };
      const r = analizarCorreoTarjeta_(correo, tarjetas, excluir);
      let idTj = '';
      if (r.registrar) {
        try {
          idTj = registrarConsumoTarjetaDetallado_({ tarjetaId: r.tarjeta.id, fecha: correo.fecha, tipo: 'Gasto Personal', monto: r.monto, moneda: r.moneda, descripcion: r.comercio, notas: 'Desde correo' }).idTarjeta;
          r.resultado = 'Registrado';
          resumen.push('💳 ' + r.moneda + ' ' + r.monto.toFixed(2) + ' · ' + r.comercio + ' (' + r.tarjeta.banco + ')');
        } catch (e) { r.resultado = 'Error: ' + e.message; }
      }
      hoja.appendRow([id, correo.fecha, String(correo.de).substring(0, 120), String(correo.asunto).substring(0, 150), r.resultado,
        r.tarjeta ? r.tarjeta.banco : '', r.monto || '', r.moneda || '', r.comercio || '', idTj]);
    });
  });
  if (resumen.length) {
    try { enviarMensajeTelegram('Registré desde tus correos:\n' + resumen.join('\n')); } catch (e) { /* Telegram opcional */ }
  }
  return resumen.length;
}

function textoDeHtml_(html) {
  return String(html || '').replace(/<style[\s\S]*?<\/style>|<script[\s\S]*?<\/script>/gi, ' ').replace(/<br\s*\/?>|<\/(p|div|tr|td|li|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#36;/g, '$').replace(/&euro;/g, '€').replace(/[ \t]+/g, ' ');
}

/** Decide qué hacer con un correo. Devuelve { registrar, resultado, tarjeta, monto, moneda, comercio } */
function analizarCorreoTarjeta_(c, tarjetas, excluir) {
  const cuerpo = String(c.cuerpo || '').substring(0, 6000);
  const cabeza = normalizarTxt_(c.de + ' ' + c.asunto + ' ' + cuerpo.substring(0, 1500));
  const todo = normalizarTxt_(c.de + ' ' + c.asunto + ' ' + cuerpo);
  const out = { registrar: false, resultado: '', tarjeta: null, monto: 0, moneda: '', comercio: '' };

  const ex = excluir.filter(function (x) { return contienePalabra_(cabeza, x); })[0];
  if (ex) { out.resultado = 'Excluido (' + ex + ')'; return out; }

  out.tarjeta = tarjetaDeCorreo_(todo, tarjetas);
  if (!out.tarjeta) { out.resultado = 'Ignorado: no es de tus tarjetas'; return out; }

  const tema = normalizarTxt_(c.asunto + ' ' + cuerpo.substring(0, 800));
  if (/(rechaz|denegad|no se pudo|no fue aprobad|no procesad|declinad|cancelad|anulad|negato|rifiutat)/.test(tema)) { out.resultado = 'Ignorado: rechazado o anulado'; return out; }
  if (/(pago (recibido|de tu tarjeta|a tu tarjeta|de tarjeta)|pagaste tu tarjeta|abono|devolucion|extorno|reembolso|estado de cuenta|codigo de verificacion|clave dinamica|token)/.test(tema)
      && !/(realizaste (un|una) (consumo|compra)|consumo con tu|compra con tu)/.test(tema)) {
    out.resultado = 'Ignorado: no es un consumo'; return out;
  }

  let d = null;
  if (iaProveedor_()) { try { d = datosCorreoIA_(c.asunto, cuerpo); } catch (e) { d = null; } }
  if (d && d.es_consumo === false) { out.resultado = 'Ignorado: la IA dice que no es un consumo'; return out; }
  const r = datosCorreoReglas_(c.asunto + '\n' + cuerpo);
  // la IA solo vale si el monto que da aparece de verdad en el correo
  if (d && d.monto > 0 && r.montos.indexOf(redondear_(d.monto)) !== -1) { out.monto = redondear_(d.monto); out.moneda = d.moneda || r.moneda; }
  else { out.monto = r.monto; out.moneda = r.moneda; }
  out.comercio = String((d && d.comercio) || r.comercio || ('Consumo ' + out.tarjeta.banco)).replace(/\s+/g, ' ').trim().substring(0, 80);
  if (!(out.monto > 0) || CONFIG.MONEDAS.indexOf(out.moneda) === -1) { out.resultado = 'Revisar: no encontré el monto'; return out; }
  out.registrar = true;
  return out;
}

function contienePalabra_(texto, palabra) {
  const p = palabra.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s*');
  return new RegExp('(^|[^a-z0-9])' + p + '([^a-z0-9]|$)').test(texto);
}

/** La tarjeta cuyo banco aparece en el correo (el nombre del banco pesa más que palabras como «Amex» u «Oro») */
function tarjetaDeCorreo_(texto, tarjetas) {
  let mejor = null, puntos = 0;
  tarjetas.forEach(function (t) {
    let p = 0;
    normalizarTxt_(t.banco).split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 2 && ['tarjeta', 'credito', 'de', 'del', 'la'].indexOf(w) === -1; }).forEach(function (w) {
      if (contienePalabra_(texto, w)) p += BANCOS_CONOCIDOS_.indexOf(w) !== -1 && w !== 'amex' ? 10 : 1;
    });
    if (p > puntos || (p === puntos && p > 0 && t.principal)) { mejor = t; puntos = p; }
  });
  return puntos >= 10 ? mejor : null; // sin el nombre del banco no se adivina
}

function numeroCorreo_(s) {
  s = String(s).replace(/\s/g, '');
  if (/,\d{2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,56
  else s = s.replace(/,/g, '');                                     // 1,234.56
  return redondear_(Number(s) || 0);
}

/** Lectura por reglas: monto con su moneda, y el comercio si el banco lo indica */
function datosCorreoReglas_(texto) {
  const t = String(texto);
  const MON = { 's/': 'PEN', 's/.': 'PEN', 'pen': 'PEN', 'soles': 'PEN', 'us$': 'USD', 'usd': 'USD', '$': 'USD', 'dolares': 'USD', 'dólares': 'USD', '€': 'EUR', 'eur': 'EUR', 'euro': 'EUR', 'euros': 'EUR' };
  const cand = [];
  const num = '(\\d{1,3}(?:[.,]\\d{3})*(?:[.,]\\d{1,2})|\\d+(?:[.,]\\d{1,2})?)';
  const sim = '(S\\/\\.?|US\\$|USD|PEN|EUR|€|\\$)';
  let m, re1 = new RegExp(sim + '\\s?' + num, 'gi'), re2 = new RegExp(num + '\\s?(soles|d[oó]lares|euros?|EUR|USD|PEN|€)', 'gi');
  while ((m = re1.exec(t))) cand.push({ i: m.index, moneda: MON[m[1].toLowerCase()], monto: numeroCorreo_(m[2]) });
  while ((m = re2.exec(t))) cand.push({ i: m.index, moneda: MON[m[2].toLowerCase()], monto: numeroCorreo_(m[1]) });
  const validos = cand.filter(function (x) { return x.moneda && x.monto > 0; });
  // se prefiere el monto que aparece junto a «monto/importe/total/consumo»
  const etiquetado = validos.filter(function (x) { return /(monto|importe|total|consumo|compra|cargo|valor|amount|importo)[^\n\d]{0,25}$/i.test(t.substring(Math.max(0, x.i - 40), x.i)); })[0];
  const elegido = etiquetado || validos.sort(function (a, b) { return a.i - b.i; })[0] || null;
  let comercio = '';
  const mc = t.match(/(?:empresa|comercio|establecimiento|negocio|lugar|merchant|esercente|en el comercio)\s*[:\-]?\s*([^\n\r]{2,60})/i);
  if (mc) comercio = mc[1];
  else { const me = t.match(/\ben\s+([A-Z0-9][A-Z0-9 *&.\-]{2,40})(?:\s|$|,|\.)/); if (me) comercio = me[1]; }
  return { monto: elegido ? elegido.monto : 0, moneda: elegido ? elegido.moneda : '', comercio: comercio.trim(), montos: validos.map(function (x) { return x.monto; }) };
}

/** Lectura con la IA (si configuraste una clave): más robusta con formatos de banco que cambian */
function datosCorreoIA_(asunto, cuerpo) {
  const sys = 'Lees avisos de bancos. Responde SOLO un JSON: {"es_consumo":bool,"monto":number,"moneda":"PEN"|"USD"|"EUR","comercio":string}. ' +
    'es_consumo=true solo si el correo avisa de una compra o consumo hecho CON la tarjeta. Pagos a la tarjeta, devoluciones, rechazos, estados de cuenta o publicidad = false. ' +
    'monto es el importe del consumo tal como aparece; comercio es el nombre del establecimiento (corto).';
  const j = JSON.parse(extraerJSON_(iaTexto_(sys, 'Asunto: ' + asunto + '\n\n' + String(cuerpo).substring(0, 3500), 200)));
  return { es_consumo: j.es_consumo !== false, monto: numero_(j.monto), moneda: String(j.moneda || '').toUpperCase(), comercio: String(j.comercio || '') };
}

/** Prueba manual desde el editor: muestra qué haría con los últimos correos sin registrar nada */
function probarCorreosTarjetas() {
  const tarjetas = tarjetasActivas_();
  const excluir = String(getAjuste_('correos.excluir', CORREOS_EXCLUIR_DEF_)).split(',').map(normalizarTxt_).filter(String);
  const l = [];
  GmailApp.search('newer_than:30d ' + PALABRAS_CORREO_, 0, 15).forEach(function (h) {
    const m = h.getMessages().pop();
    const r = analizarCorreoTarjeta_({ de: m.getFrom(), asunto: m.getSubject(), cuerpo: m.getPlainBody() || textoDeHtml_(m.getBody()) }, tarjetas, excluir);
    l.push(m.getSubject().substring(0, 50) + ' → ' + (r.registrar ? 'REGISTRARÍA ' + r.moneda + ' ' + r.monto + ' · ' + r.comercio + ' en ' + r.tarjeta.banco : r.resultado));
  });
  Logger.log(l.join('\n'));
  avisar_(l.join('\n') || 'No encontré correos de consumo en los últimos 30 días.');
}
