// ======================================================================
// FINANZAS · código para tu hoja de Google (Extensiones > Apps Script)
// Pega TODO este texto en Code.gs, guarda y ejecuta la función "instalarTodo".
// ======================================================================
/* ==== 00_Config.gs ==== */
/**
 * ============================================================
 * CONFIGURACIÓN GLOBAL — Sistema de Finanzas Personales
 * ============================================================
 * Edita SOLO este archivo para ajustar parámetros del sistema.
 * El resto de módulos lee de aquí, nunca hardcodean valores.
 */

var CONFIG = {

  // --- Secretos: NO van aquí. Se guardan en Propiedades del script (ver prop_ más abajo) ---

  // --- Monedas soportadas ---
  MONEDAS: ['PEN', 'USD', 'EUR'],

  // --- Cuentas por moneda (banco/efectivo), reflejan tu realidad real ---
  CUENTAS_POR_MONEDA: {
    PEN: ['BCP', 'Interbank', 'Efectivo'],
    USD: ['Interbank', 'Revolut', 'Efectivo'],
    EUR: ['Revolut', 'Intesa Sanpaolo', 'Efectivo']
  },

  // --- Tarjeta de crédito ---
  TARJETA_CREDITO: {
    BANCO: 'Interbank',
    DIA_CIERRE: 20,  // día del mes en que cierra el ciclo de facturación
    DIA_COBRO: 15    // día del mes SIGUIENTE al cierre en que se cobra
  },

  // --- Centros de costo (negocios / ámbitos) ---
  // --- Tipos de cuenta ('Efectivo' es el único que se considera dinero en efectivo; el resto es dinero en banco/app) ---
  TIPOS_CUENTA: ['Banco', 'Efectivo', 'Ahorros', 'Inversion', 'Billetera Digital'],

  CENTROS_COSTO: ['Personal', 'Alquileres', 'Importaciones Italia', 'Galletas'],

  // --- Categorías de tributación SUNAT ---
  CATEGORIAS_SUNAT: ['Primera', 'Segunda', 'Tercera', 'Cuarta', 'No aplica'],

  // --- Tasas SUNAT vigentes ---
  // IMPORTANTE: SUNAT actualiza la UIT cada año (normalmente en enero).
  // Revisa y actualiza estos valores una vez al año en ConfiguracionTributaria.
  TRIBUTARIO: {
    ANIO_VIGENTE: 2026,
    UIT: 5350, // referencial — confirmar valor oficial vigente del ejercicio
    TASA_PRIMERA_CATEGORIA: 0.05,    // efectiva sobre renta bruta mensual (alquileres)
    TASA_SEGUNDA_CATEGORIA: 0.05,    // efectiva sobre renta bruta (intereses por préstamos otorgados)
    RETENCION_CUARTA_CATEGORIA: 0.08 // sobre recibos por honorarios que superen el umbral
  },

  // --- API de tipo de cambio referencial (para consolidar patrimonio, NO para tus cambios reales) ---
  API_TIPO_CAMBIO_URL: 'https://api.exchangerate-api.com/v4/latest/PEN' // reemplazable por otra fuente
};


// ===============================================================
// UTILIDADES COMPARTIDAS (usadas por todos los módulos)
// ===============================================================

var MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
var SIMBOLOS_MONEDA = { PEN: 'S/', USD: '$', EUR: '€' };
var TIPOS_CUENTA_ = ['Banco', 'Efectivo', 'Ahorros', 'Inversion', 'Billetera Digital'];
var TIPOS_MOVIMIENTO = ['Ingreso Activo', 'Ingreso Pasivo', 'Gasto Personal', 'Gasto Fijo', 'Gasto Inversion', 'Transferencia', 'Ajuste'];

// ===============================================================
// PROPIEDADES DEL SCRIPT (Configuración del proyecto > Propiedades de la secuencia de comandos)
// ===============================================================
//   TELEGRAM_BOT_TOKEN       token de @BotFather
//   TELEGRAM_CHAT_ID         tu chat id personal
//   TELEGRAM_WEBHOOK_SECRET  texto aleatorio (solo letras/números) que tú inventas
//   ANTHROPIC_API_KEY        (opcional) IA para recomendaciones y para entender mensajes de texto
//   GEMINI_API_KEY           (opcional) necesaria solo para entender AUDIOS (y como IA alterna si no hay Anthropic)
//   DASHBOARD_PIN            (opcional) PIN de respaldo para entrar al dashboard si Google no te reconoce (ej. iPad)
//   SESSION_SECRET           (opcional) texto aleatorio para firmar sesiones del PIN
//   AI_MODEL / GEMINI_MODEL  (opcional) para cambiar de modelo sin tocar código
function prop_(clave, porDefecto) {
  const v = PropertiesService.getScriptProperties().getProperty(clave);
  return (v === null || v === undefined || String(v).trim() === '') ? (porDefecto === undefined ? '' : porDefecto) : String(v).trim();
}
function tgToken_() { return prop_('TELEGRAM_BOT_TOKEN'); }
function tgChat_() { return prop_('TELEGRAM_CHAT_ID'); }
function tgSecret_() { return prop_('TELEGRAM_WEBHOOK_SECRET'); }

/** Devuelve el spreadsheet activo (útil para testear módulos por separado) */
function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Devuelve una hoja o lanza un error claro si falta (suele pasar si no corriste crearEstructuraCompleta) */
function getHoja_(nombre) {
  const hoja = getSS_().getSheetByName(nombre);
  if (!hoja) throw new Error('Falta la hoja "' + nombre + '". Corre ⚙️ Finanzas > Crear estructura completa.');
  return hoja;
}

/** Valores de una hoja (incluye la fila de encabezados en la posición 0). Hoja vacía => solo encabezados */
var MEMO_LECTURA_ = null; // solo durante pantallas de SOLO lectura (dashboard, etc.): cada hoja se lee una vez
function leerHoja_(nombre) {
  if (MEMO_LECTURA_ && nombre !== 'Cotizaciones' && nombre !== 'Log') {
    if (!MEMO_LECTURA_[nombre]) MEMO_LECTURA_[nombre] = getHoja_(nombre).getDataRange().getValues();
    return MEMO_LECTURA_[nombre];
  }
  return getHoja_(nombre).getDataRange().getValues();
}

/**
 * Convierte CUALQUIER cosa escrita a mano en una celda a número.
 * Entiende: 6516.99 | "6,516.99" | "S/.6,516.99" | "S/ 6.516,99" | "$ 1 200,50" | "(250)" | "-40".
 * Nunca devuelve NaN: si no hay número, devuelve 0.
 */
function numero_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined || v === '' || v instanceof Date) return 0;
  let s = String(v).trim();
  if (!s) return 0;
  const negativo = /^\(.*\)$/.test(s) || /^[^0-9]*-/.test(s);
  s = s.replace(/[^0-9.,]/g, '').replace(/^[.,]+/, '').replace(/[.,]+$/, '');
  if (!s) return 0;

  const ultPunto = s.lastIndexOf('.');
  const ultComa = s.lastIndexOf(',');
  let normal;
  if (ultPunto !== -1 && ultComa !== -1) {
    // El último separador es el decimal; el otro es de miles
    const decimal = ultPunto > ultComa ? '.' : ',';
    const miles = decimal === '.' ? ',' : '.';
    normal = s.split(miles).join('').replace(decimal, '.');
  } else if (ultComa !== -1) {
    const partes = s.split(',');
    const unaSola = partes.length === 2;
    const pareceMiles = !unaSola || (partes[1].length === 3 && partes[0].length >= 1 && partes[0].length <= 3 && partes[0] !== '0');
    normal = pareceMiles ? partes.join('') : partes[0] + '.' + partes[1];
  } else if (ultPunto !== -1) {
    const partes = s.split('.');
    normal = partes.length === 2 ? s : partes.join(''); // "1.234.567" = miles; "12.5" = decimal
  } else {
    normal = s;
  }
  const n = parseFloat(normal);
  if (!isFinite(n)) return 0;
  return negativo ? -n : n;
}

/** Redondea a 2 decimales (evita 0.1+0.2 = 0.30000000000000004 en los saldos) */
function redondear_(n) {
  return Math.round((Number(n) + (n >= 0 ? 1e-9 : -1e-9)) * 100) / 100;
}

/** Valida que un monto sea un número positivo; devuelve el número ya limpio */
function validarMonto_(valor, nombre) {
  const n = numero_(valor);
  if (!(n > 0)) throw new Error((nombre || 'El monto') + ' debe ser un número mayor a 0 (recibí: "' + valor + '").');
  return redondear_(n);
}

/** Convierte a Date: acepta Date, "2026-10-15", "15/10/2026", "15-10-26". Devuelve null si no se puede */
function fecha_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2}))?/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 12, m[5] ? +m[5] : 0);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    const anio = +m[3] < 100 ? 2000 + +m[3] : +m[3];
    return new Date(anio, +m[2] - 1, +m[1], 12, 0);
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Suma n meses a una fecha sin desbordar (31-ene + 1 mes = 28/29-feb, no 3-mar) */
function sumarMeses_(fecha, n) {
  const dia = fecha.getDate();
  const r = new Date(fecha.getFullYear(), fecha.getMonth() + n, 1, fecha.getHours(), fecha.getMinutes());
  const ultimo = new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate();
  r.setDate(Math.min(dia, ultimo));
  return r;
}

function ultimoDiaMes_(fecha) {
  return new Date(fecha.getFullYear(), fecha.getMonth() + 1, 0).getDate();
}

function zona_() { return Session.getScriptTimeZone() || 'America/Lima'; }

function fechaISO_(d) {
  const f = fecha_(d);
  return f ? Utilities.formatDate(f, zona_(), 'yyyy-MM-dd') : '';
}

function nombreMes_(fecha) { return MESES_ES[fecha.getMonth()]; }

function formatoMonto_(monto, moneda) {
  const n = redondear_(numero_(monto));
  const partes = Math.abs(n).toFixed(2).split('.');
  partes[0] = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '-' : '') + (SIMBOLOS_MONEDA[moneda] || moneda || '') + ' ' + partes.join('.');
}

function esIngreso_(tipo) { return String(tipo).indexOf('Ingreso') === 0; }
function esGasto_(tipo) { return String(tipo).indexOf('Gasto') === 0; }

/** Escapa caracteres que rompen el Markdown clásico de Telegram */
function escaparMd_(t) {
  return String(t === undefined || t === null ? '' : t).replace(/([_*`\[])/g, '\\$1');
}

/** Evita que reglas/valores libres rompan las listas desplegables de Movimientos */
function limpiarCentroCosto_(c) {
  return c && CONFIG.CENTROS_COSTO.indexOf(c) !== -1 ? c : (c ? 'Personal' : '');
}
function limpiarCategoriaSunat_(c) {
  return c && CONFIG.CATEGORIAS_SUNAT.indexOf(c) !== -1 ? c : 'No aplica';
}

/**
 * Ejecuta fn con el candado del script (evita que dos registros simultáneos —por ejemplo
 * Telegram y el dashboard al mismo tiempo— se pisen y dupliquen IDs o corrompan saldos).
 * Es re-entrante: una función con candado puede llamar a otra con candado.
 */
var LOCK_PROFUNDIDAD_ = 0;
function conLock_(fn) {
  if (LOCK_PROFUNDIDAD_ > 0) return fn();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  LOCK_PROFUNDIDAD_++;
  try {
    return fn();
  } finally {
    LOCK_PROFUNDIDAD_--;
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/** Para triggers: nunca falla en silencio. Registra el error en Log y te avisa por Telegram */
function ejecutarSeguro_(nombre, fn) {
  try {
    return fn();
  } catch (e) {
    try { registrarLog_('ERROR ' + nombre, e.message); } catch (e2) { /* ignorar */ }
    try {
      // Un mismo error se avisa una sola vez cada 6 horas (si no, un fallo en un trigger de 15 min = mensajes a cada rato)
      const cache = CacheService.getScriptCache();
      const clave = 'err_' + hash_(nombre + e.message);
      if (!cache.get(clave)) { cache.put(clave, '1', 21600); enviarMensajeTelegram('❌ Error en ' + nombre + ': ' + e.message); }
    } catch (e3) { /* ignorar */ }
  }
}

/** Muestra un aviso en el Sheet si hay interfaz; si no (trigger/editor), solo lo deja en el Log */
function avisar_(mensaje) {
  try {
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    Logger.log(mensaje);
  }
}

/** Busca una cuenta por ID. Devuelve { id, moneda, banco, tipo, saldo, activa, numeroFila } o null */
function obtenerCuenta_(cuentaId) {
  const datos = leerHoja_('Cuentas');
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === String(cuentaId) && datos[i][0] !== '') {
      return {
        id: datos[i][0], moneda: String(datos[i][1]).trim(), banco: String(datos[i][2]).trim(), tipo: String(datos[i][3]).trim(),
        saldo: numero_(datos[i][5]), saldoInicial: numero_(datos[i][4]), activa: String(datos[i][6]).trim().toUpperCase() !== 'NO',
        color: String(datos[i][7] || ''), numeroFila: i + 1
      };
    }
  }
  return null;
}

/** Valida que la cuenta exista, esté activa y sea de la misma moneda que la operación */
function validarCuentaMoneda_(cuentaId, moneda) {
  const c = obtenerCuenta_(cuentaId);
  if (!c) throw new Error('La cuenta ' + cuentaId + ' no existe en la hoja Cuentas.');
  if (!c.activa) throw new Error('La cuenta ' + c.banco + ' (' + c.moneda + ') está marcada como inactiva.');
  if (moneda && c.moneda !== moneda) {
    throw new Error('La cuenta ' + c.banco + ' es en ' + c.moneda + ' pero la operación es en ' + moneda + '.');
  }
  return c;
}

function validarMoneda_(moneda) {
  const m = String(moneda || '').trim().toUpperCase();
  if (CONFIG.MONEDAS.indexOf(m) === -1) throw new Error('Moneda no soportada: "' + moneda + '".');
  return m;
}

/**
 * Fecha para un registro del usuario: si viene solo la fecha y es HOY, usa la hora actual (así aparece
 * en orden y no queda "en el futuro" a las 12:00). Otras fechas quedan a las 12:00.
 */
function fechaConHora_(v) {
  const f = fecha_(v);
  if (!f) return new Date();
  const s = String(v instanceof Date ? '' : v).trim();
  const soloFecha = /^\d{4}-\d{2}-\d{2}$/.test(s) || /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$/.test(s);
  if (soloFecha || v instanceof Date && v.getHours() === 12 && v.getMinutes() === 0) {
    const ahora = new Date();
    if (f.getFullYear() === ahora.getFullYear() && f.getMonth() === ahora.getMonth() && f.getDate() === ahora.getDate()) return ahora;
  }
  return f;
}

/** ¿La fecha es de un día posterior a hoy? (entonces se guarda como movimiento programado) */
function esFuturo_(f) {
  const h = new Date();
  return f > new Date(h.getFullYear(), h.getMonth(), h.getDate(), 23, 59, 59, 999);
}

/** Ajustes editables desde la app (hoja Ajustes: Clave | Valor) */
var AJUSTES_MEM_ = null;
function getAjuste_(clave, porDefecto) {
  if (AJUSTES_MEM_ === null) {
    const hoja = getSS_().getSheetByName('Ajustes');
    AJUSTES_MEM_ = {};
    if (hoja) {
      const datos = hoja.getDataRange().getValues();
      for (let i = 1; i < datos.length; i++) if (datos[i][0] !== '' && datos[i][1] !== '') AJUSTES_MEM_[String(datos[i][0])] = datos[i][1];
    }
  }
  const v = AJUSTES_MEM_[clave];
  return v === undefined ? (porDefecto === undefined ? '' : porDefecto) : v;
}
function setAjuste_(clave, valor) {
  AJUSTES_MEM_ = null;
  const hoja = getHoja_('Ajustes');
  const datos = hoja.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === clave) { hoja.getRange(i + 1, 2).setValue(valor); return; }
  }
  hoja.appendRow([clave, valor]);
}


/* ==== 01_Setup.gs ==== */
/**
 * ============================================================
 * SETUP — Crea toda la estructura de hojas del sistema
 * ============================================================
 * Ejecuta crearEstructuraCompleta() UNA SOLA VEZ (o desde el menú
 * "⚙️ Finanzas" que aparece al abrir el Sheet) para dejar listo
 * todo el modelo de datos. Es seguro volver a ejecutarlo: no borra
 * datos ya cargados en Cuentas ni en ConfiguracionTributaria.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Finanzas')
    .addItem('Crear estructura completa (setup inicial)', 'crearEstructuraCompleta')
    .addItem('Cargar reglas de categorización iniciales', 'poblarReglasCategorizacionIniciales')
    .addItem('Instalar triggers automáticos', 'instalarTriggers')
    .addItem('Actualizar tipo de cambio referencial', 'actualizarTipoCambioReferencial')
    .addSeparator()
    .addItem('Ajustar saldo de una cuenta (saldo real de hoy)', 'ajustarSaldoDesdeMenu')
    .addItem('Auditar saldos contra el historial', 'auditarSaldos')
    .addItem('Revisar el sistema (diagnóstico)', 'diagnosticarSistema')
    .addSeparator()
    .addItem('Telegram: conectar webhook', 'configurarWebhookTelegram')
    .addItem('Telegram: reparar (limpiar cola)', 'repararTelegramMenu')
    .addItem('Telegram: enviar mensaje de prueba', 'probarTelegram')
    .addItem('Instalar alertas de programados (cada 15 min)', 'instalarTriggers')
    .addItem('Procesar recurrentes ahora', 'procesarRecurrentesDelDia')
    .addItem('Enviar reporte del mes en curso (prueba)', 'enviarReporteMesActualDePrueba')
    .addSeparator()
    .addItem('Migración: crear hoja Importacion', 'crearHojaImportacion')
    .addItem('Migración: importar historial', 'migrarDesdeImportacion')
    .addToUi();
}

var ESTRUCTURA_VERSION_ = '4';

/** Si pegaste código nuevo, crea las hojas nuevas y actualiza validaciones solo (una vez por versión) */
function asegurarEstructura_() {
  if (PropertiesService.getScriptProperties().getProperty('ESTRUCTURA_VERSION') === ESTRUCTURA_VERSION_) return;
  conLock_(function () { crearEstructuraCompleta(); });
}

function crearEstructuraCompleta() {
  const ss = getSS_();

  crearHojaCuentas(ss);
  crearHojaMovimientos(ss);
  crearHojaCambiosDivisa(ss);
  crearHojaTarjetaCredito(ss);
  crearHojaPrestamosOtorgados(ss);
  crearHojaDeudas(ss);
  crearHojaCronogramaCuotas(ss);
  crearHojaMetas(ss);
  crearHojaPresupuestos(ss);
  crearHojaRecurrentes(ss);
  crearHojaReglasCategorizacion(ss);
  crearHojaTipoCambioReferencial(ss);
  crearHojaAjustesSaldo(ss);
  crearHojaIntegraciones(ss);
  crearHojaLog(ss);
  crearHojaProgramados(ss);
  crearHojaAjustes(ss);
  crearHojaInversiones(ss);
  crearHojaCotizaciones(ss);
  crearHojaCategorias(ss);

  poblarCuentasIniciales(ss);
  poblarReglasCategorizacionIniciales(ss);

  PropertiesService.getScriptProperties().setProperty('ESTRUCTURA_VERSION', ESTRUCTURA_VERSION_);
  registrarLog_('Setup', 'Estructura completa creada/verificada');
  avisar_('✅ Estructura lista. Revisa las pestañas nuevas del Sheet.');
}

/** Utilidad base: crea (o re-encabeza) una hoja con encabezados fijos, sin borrar filas de datos */
function crearHojaBase_(ss, nombre, encabezados) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
  }
  hoja.getRange(1, 1, 1, encabezados.length)
    .setValues([encabezados])
    .setFontWeight('bold')
    .setBackground('#1c1c1e')
    .setFontColor('#ffffff');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, encabezados.length);
  return hoja;
}

/** Utilidad: aplica una lista desplegable de validación a una columna completa */
function aplicarValidacion_(hoja, columna, opciones, filas) {
  filas = filas || Math.max(hoja.getMaxRows() - 1, 1000);
  if (hoja.getMaxRows() < filas + 1) hoja.insertRowsAfter(hoja.getMaxRows(), filas + 1 - hoja.getMaxRows());
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  hoja.getRange(2, columna, filas, 1).setDataValidation(regla);
}

// ---------------------------------------------------------------
// DEFINICIÓN DE CADA HOJA
// ---------------------------------------------------------------

function crearHojaCuentas(ss) {
  const hoja = crearHojaBase_(ss, 'Cuentas', [
    'ID', 'Moneda', 'Banco', 'Tipo', 'SaldoInicial', 'SaldoActual', 'Activa', 'Color'
  ]);
  aplicarValidacion_(hoja, 2, CONFIG.MONEDAS);
  aplicarValidacion_(hoja, 4, TIPOS_CUENTA_);
  aplicarValidacion_(hoja, 7, ['SI', 'NO']);
}

function crearHojaMovimientos(ss) {
  const hoja = crearHojaBase_(ss, 'Movimientos', [
    'ID', 'Fecha', 'FechaEfectiva', 'Tipo', 'Monto', 'Moneda', 'CuentaID',
    'CentroCosto', 'Categoria', 'Subcategoria', 'Descripcion', 'MedioPago',
    'CategoriaSUNAT', 'Origen', 'Notas', 'Proyecto'
  ]);
  aplicarValidacion_(hoja, 4, [
    'Ingreso Activo', 'Ingreso Pasivo', 'Gasto Personal', 'Gasto Fijo',
    'Gasto Inversion', 'Transferencia', 'Ajuste'
  ]);
  aplicarValidacion_(hoja, 6, CONFIG.MONEDAS);
  aplicarValidacion_(hoja, 8, CONFIG.CENTROS_COSTO);
  aplicarValidacion_(hoja, 12, ['Efectivo', 'Debito', 'Tarjeta de Credito']);
  aplicarValidacion_(hoja, 13, CONFIG.CATEGORIAS_SUNAT);
  aplicarValidacion_(hoja, 14, ['Manual', 'Bot Airbnb', 'Bot Ventas', 'Bot Finanzas', 'Tarjeta Credito', 'Recurrente']);
}

function crearHojaCambiosDivisa(ss) {
  crearHojaBase_(ss, 'CambiosDivisa', [
    'ID', 'Fecha', 'MonedaOrigen', 'MontoOrigen', 'CuentaOrigenID',
    'MonedaDestino', 'MontoDestino', 'CuentaDestinoID', 'TasaAplicada', 'Comision'
  ]);
}

function crearHojaTarjetaCredito(ss) {
  const hoja = crearHojaBase_(ss, 'TarjetaCredito', [
    'ID', 'FechaConsumo', 'Descripcion', 'Monto', 'Moneda', 'CicloCierre',
    'FechaCobro', 'Estado', 'EsDesgravamen', 'CategoriaSUNAT', 'CentroCosto',
    'Categoria', 'Subcategoria', 'TarjetaID'
  ]);
  aplicarValidacion_(hoja, 8, ['Pendiente', 'Facturado', 'Pagado']);
  aplicarValidacion_(hoja, 9, ['SI', 'NO']);
}

function crearHojaPrestamosOtorgados(ss) {
  const hoja = crearHojaBase_(ss, 'PrestamosOtorgados', [
    'ID', 'Deudor', 'Moneda', 'MontoPrestado', 'TasaInteres', 'TipoTasa',
    'Frecuencia', 'FechaInicio', 'PlazoMeses', 'Estado', 'SaldoPendiente', 'CuentaOrigenID', 'Concepto'
  ]);
  aplicarValidacion_(hoja, 6, ['Mensual', 'Anual']);
  aplicarValidacion_(hoja, 7, ['Mensual', 'Al vencimiento']);
  aplicarValidacion_(hoja, 10, ['Vigente', 'Pagado', 'Atrasado']);
}

function crearHojaDeudas(ss) {
  const hoja = crearHojaBase_(ss, 'Deudas', [
    'ID', 'Acreedor', 'Moneda', 'MontoRecibido', 'TasaInteres', 'TipoTasa',
    'FechaInicio', 'PlazoMeses', 'Estado', 'SaldoPendiente', 'CuentaDestinoID'
  ]);
  aplicarValidacion_(hoja, 6, ['TEA', 'TCEA', 'Mensual']);
  aplicarValidacion_(hoja, 9, ['Vigente', 'Pagado', 'Atrasado']);
}

function crearHojaCronogramaCuotas(ss) {
  const hoja = crearHojaBase_(ss, 'CronogramaCuotas', [
    'ID', 'ReferenciaID', 'Tipo', 'NumeroCuota', 'FechaVencimiento',
    'MontoCapital', 'MontoInteres', 'Estado'
  ]);
  aplicarValidacion_(hoja, 3, ['Otorgado', 'Recibido']);
  aplicarValidacion_(hoja, 8, ['Pendiente', 'Pagado', 'Atrasado']);
}

function crearHojaMetas(ss) {
  const hoja = crearHojaBase_(ss, 'Metas', [
    'ID', 'Periodo', 'Moneda', 'MontoMeta', 'FechaInicio', 'FechaFin', 'CrecimientoPct'
  ]);
  aplicarValidacion_(hoja, 2, ['Diaria', 'Semanal', 'Mensual', 'Anual']);
  aplicarValidacion_(hoja, 3, CONFIG.MONEDAS);
}

function crearHojaPresupuestos(ss) {
  const hoja = crearHojaBase_(ss, 'Presupuestos', [
    'ID', 'Categoria', 'Subcategoria', 'Moneda', 'MontoLimite', 'MesAplicable'
  ]);
  aplicarValidacion_(hoja, 4, CONFIG.MONEDAS);
}

function crearHojaRecurrentes(ss) {
  const hoja = crearHojaBase_(ss, 'Recurrentes', [
    'ID', 'Descripcion', 'Tipo', 'Monto', 'Moneda', 'DiaDelMes',
    'Categoria', 'Subcategoria', 'CentroCosto', 'CuentaID', 'Activo', 'UltimaEjecucion', 'Confirmar'
  ]);
  aplicarValidacion_(hoja, 3, ['Ingreso', 'Gasto']);
  aplicarValidacion_(hoja, 5, CONFIG.MONEDAS);
  aplicarValidacion_(hoja, 11, ['SI', 'NO']);
}

function crearHojaReglasCategorizacion(ss) {
  crearHojaBase_(ss, 'ReglasCategorizacion', [
    'ID', 'PalabraClave', 'Categoria', 'Subcategoria', 'CentroCosto', 'Prioridad'
  ]);
}

function crearHojaTipoCambioReferencial(ss) {
  crearHojaBase_(ss, 'TipoCambioReferencial', [
    'Fecha', 'USD_PEN', 'EUR_PEN', 'USD_EUR'
  ]);
}

function crearHojaAjustesSaldo(ss) {
  crearHojaBase_(ss, 'AjustesSaldo', [
    'ID', 'Fecha', 'CuentaID', 'MontoAjuste', 'Motivo'
  ]);
}

function crearHojaIntegraciones(ss) {
  const hoja = crearHojaBase_(ss, 'Integraciones', [
    'ID', 'SistemaOrigen', 'SpreadsheetIDOrigen', 'HojaOrigen', 'Activo', 'UltimaSync'
  ]);
  aplicarValidacion_(hoja, 2, ['Bot Airbnb', 'Bot Ventas Galletas']);
  aplicarValidacion_(hoja, 5, ['SI', 'NO']);
}

function crearHojaProgramados(ss) {
  const hoja = crearHojaBase_(ss, 'Programados', [
    'ID', 'Tipo', 'Monto', 'Moneda', 'Medio', 'CuentaID', 'Descripcion', 'Categoria', 'Subcategoria', 'FechaProgramada',
    'Alertas', 'AlertasEnviadas', 'Estado', 'FechaRealizada', 'MovimientoID', 'UltimoAviso', 'Notas'
  ]);
  aplicarValidacion_(hoja, 13, ['Pendiente', 'Realizado', 'Cancelado']);
}

function crearHojaAjustes(ss) {
  crearHojaBase_(ss, 'Ajustes', ['Clave', 'Valor']);
}

function crearHojaInversiones(ss) {
  const hoja = crearHojaBase_(ss, 'Inversiones', ['ID', 'Fecha', 'Operacion', 'Ticker', 'Cantidad', 'Precio', 'Moneda', 'Comision', 'CuentaID', 'MovimientoID', 'Notas']);
  aplicarValidacion_(hoja, 3, ['Compra', 'Venta', 'Dividendo']);
  aplicarValidacion_(hoja, 7, CONFIG.MONEDAS);
}

function crearHojaCotizaciones(ss) {
  crearHojaBase_(ss, 'Cotizaciones', ['Ticker', 'Precio', 'Moneda', 'Nombre', 'CambioPct', 'Fecha', 'Previo']);
}

function crearHojaLog(ss) {
  crearHojaBase_(ss, 'Log', ['FechaHora', 'Accion', 'Detalle']);
}

// ---------------------------------------------------------------
// DATOS INICIALES
// ---------------------------------------------------------------

/** Crea las cuentas iniciales según tus bancos reales (solo si la hoja está vacía) */
function poblarCuentasIniciales(ss) {
  ss = ss || getSS_();
  const hoja = ss.getSheetByName('Cuentas');
  if (hoja.getLastRow() > 1) return; // ya tiene datos: no sobrescribir

  const filas = [];
  let id = 1;
  Object.keys(CONFIG.CUENTAS_POR_MONEDA).forEach(function (moneda) {
    CONFIG.CUENTAS_POR_MONEDA[moneda].forEach(function (banco) {
      const tipo = banco === 'Efectivo' ? 'Efectivo' : 'Banco';
      filas.push([id++, moneda, banco, tipo, 0, 0, 'SI']);
    });
  });
  hoja.getRange(2, 1, filas.length, 7).setValues(filas);
}

/**
 * Precarga reglas de categorización con las palabras clave que ya aparecen
 * de forma recurrente en tu propio historial (habitaciones, suscripciones,
 * transporte, negocio de galletas, importaciones, etc.). Puedes editar,
 * borrar o agregar filas libremente después desde la hoja ReglasCategorizacion.
 * Solo se ejecuta si la hoja está vacía.
 */
function poblarReglasCategorizacionIniciales(ss) {
  ss = ss || getSS_();
  const hoja = ss.getSheetByName('ReglasCategorizacion');
  if (hoja.getLastRow() > 1) return;

  const reglas = [
    // PalabraClave, Categoria, Subcategoria, CentroCosto, Prioridad
    ['habitacion', 'Alquiler', 'Ingreso Habitacion', 'Alquileres', 1],
    ['habitación', 'Alquiler', 'Ingreso Habitacion', 'Alquileres', 1],
    ['mini dep', 'Alquiler', 'Ingreso Habitacion', 'Alquileres', 1],
    ['garantia', 'Alquiler', 'Garantia', 'Alquileres', 2],
    ['icloud', 'Suscripciones', 'iCloud', 'Personal', 1],
    ['apple music', 'Suscripciones', 'Apple Music', 'Personal', 1],
    ['pasta gallo', 'Negocio Galletas', 'Insumos', 'Galletas', 1],
    ['galletas', 'Negocio Galletas', 'Insumos', 'Galletas', 2],
    ['cuota italia', 'Importaciones Italia', 'Cuota', 'Importaciones Italia', 1],
    ['productos italia', 'Importaciones Italia', 'Venta Productos', 'Importaciones Italia', 1],
    ['taxi', 'Transporte', 'Taxi', 'Personal', 2],
    ['tren', 'Transporte', 'Tren', 'Personal', 2],
    ['bus', 'Transporte', 'Bus', 'Personal', 2],
    ['uber', 'Transporte', 'Uber', 'Personal', 2],
    ['pedidos ya', 'Alimentacion', 'Delivery', 'Personal', 1],
    ['pedidosya', 'Alimentacion', 'Delivery', 'Personal', 1],
    ['gym', 'Salud', 'Gimnasio', 'Personal', 1],
    ['seguro desgravamen', 'Tarjeta de Credito', 'Seguro Desgravamen', 'Personal', 1],
    ['itf', 'Impuestos', 'ITF', 'Personal', 1],
    ['corte de pelo', 'Cuidado Personal', 'Corte de pelo', 'Personal', 2],
    ['corte pelo', 'Cuidado Personal', 'Corte de pelo', 'Personal', 2],
    ['perfume', 'Cuidado Personal', 'Perfume', 'Personal', 2],
    ['skincare', 'Cuidado Personal', 'Skincare', 'Personal', 2],
    ['facebook ads', 'Marketing', 'Publicidad', 'Galletas', 1]
  ];

  const filas = reglas.map(function (r, i) { return [i + 1].concat(r); });
  hoja.getRange(2, 1, filas.length, 6).setValues(filas);
}

/** Utilidad de auditoría: registra una línea en la hoja Log */
function registrarLog_(accion, detalle) {
  const hoja = getSS_().getSheetByName('Log');
  if (!hoja) return;
  hoja.appendRow([new Date(), accion, detalle]);
}

/** Prueba rápida del bot: si no llega el mensaje, revisa token, chat id y Log */
function probarTelegram() {
  if (!telegramConfigurado_()) { avisar_('Agrega TELEGRAM_BOT_TOKEN y TELEGRAM_CHAT_ID en las propiedades del script'); return; }
  avisar_(enviarMensajeTelegram('✅ Prueba de conexión desde tu sistema de finanzas.') ? 'Mensaje enviado. Revisa Telegram.' : 'Telegram rechazó el envío. Mira la hoja Log.');
}

/** Revisa problemas típicos: hojas faltantes, cuentas duplicadas, montos escritos como texto, monedas inválidas, etc. */
function diagnosticarSistema() {
  const problemas = [];
  const ss = getSS_();
  ['Programados', 'Ajustes', 'Categorias', 'Inversiones', 'Cotizaciones', 'Cuentas', 'Movimientos', 'CambiosDivisa', 'TarjetaCredito', 'PrestamosOtorgados', 'Deudas', 'CronogramaCuotas', 'Metas',
    'Presupuestos', 'Recurrentes', 'ReglasCategorizacion', 'TipoCambioReferencial',
    'AjustesSaldo', 'Integraciones', 'Log'].forEach(function (n) { if (!ss.getSheetByName(n)) problemas.push('Falta la hoja ' + n); });
  if (problemas.length) { avisar_('Problemas:\n- ' + problemas.join('\n- ')); return; }

  const cuentas = leerHoja_('Cuentas');
  const vistos = {};
  for (let i = 1; i < cuentas.length; i++) {
    if (cuentas[i][0] === '') continue;
    const clave = cuentas[i][1] + '|' + cuentas[i][2];
    if (vistos[clave]) problemas.push('Cuenta duplicada: ' + cuentas[i][2] + ' ' + cuentas[i][1]);
    vistos[clave] = true;
    if (CONFIG.MONEDAS.indexOf(String(cuentas[i][1]).trim()) === -1) problemas.push('Fila ' + (i + 1) + ' de Cuentas: moneda inválida');
    [4, 5].forEach(function (c) {
      if (typeof cuentas[i][c] === 'string' && cuentas[i][c] !== '') problemas.push('Cuentas fila ' + (i + 1) + ': el saldo está como texto ("' + cuentas[i][c] + '"); el sistema lo interpreta como ' + numero_(cuentas[i][c]));
    });
  }
  const movs = leerHoja_('Movimientos');
  const ids = {};
  for (let i = 1; i < movs.length; i++) {
    if (movs[i][0] === '') continue;
    if (ids[movs[i][0]]) problemas.push('ID de movimiento repetido: ' + movs[i][0]);
    ids[movs[i][0]] = true;
    if (typeof movs[i][4] === 'string') problemas.push('Movimientos fila ' + (i + 1) + ': monto como texto ("' + movs[i][4] + '")');
    if (!fecha_(movs[i][1])) problemas.push('Movimientos fila ' + (i + 1) + ': fecha ilegible');
    if (TIPOS_MOVIMIENTO.indexOf(String(movs[i][3]).trim()) === -1) problemas.push('Movimientos fila ' + (i + 1) + ': tipo desconocido "' + movs[i][3] + '"');
    if (problemas.length > 25) break;
  }
  if (!tgToken_()) problemas.push('Falta la propiedad TELEGRAM_BOT_TOKEN');
  if (!tgChat_()) problemas.push('Falta la propiedad TELEGRAM_CHAT_ID');
  if (!tgSecret_()) problemas.push('Falta la propiedad TELEGRAM_WEBHOOK_SECRET');
  if (!iaProveedor_()) problemas.push('Sin IA (opcional): agrega ANTHROPIC_API_KEY o GEMINI_API_KEY para recomendaciones y mensajes libres');
  if (!prop_('GEMINI_API_KEY')) problemas.push('Sin GEMINI_API_KEY (opcional): los audios de Telegram no se podrán entender');
  if (!prop_('DASHBOARD_PIN')) problemas.push('Sin DASHBOARD_PIN (opcional): no habrá entrada de respaldo si Google no te reconoce en algún dispositivo');
  if (obtenerUltimaTasaReferencial_().respaldo) problemas.push('Aún no hay tipo de cambio referencial (corre "Actualizar tipo de cambio referencial")');
  if (zona_() !== 'America/Lima') problemas.push('La zona horaria del proyecto es ' + zona_() + ' (debería ser America/Lima): los triggers saldrán a otra hora');
  if (ScriptApp.getProjectTriggers().length === 0) problemas.push('No hay triggers instalados');
  const auditoria = Object.keys(calcularSaldosDesdeHistorial_()).length; // fuerza lectura; diferencias: usa "Auditar saldos"

  avisar_(problemas.length ? 'Revisión:\n- ' + problemas.join('\n- ') : '✅ Todo en orden (' + auditoria + ' cuentas revisadas).');
}


/* ==== 02_Movimientos.gs ==== */
/**
 * ============================================================
 * MOVIMIENTOS — Motor central de registro de ingresos y gastos
 * ============================================================
 * Convenciones (todo el sistema depende de ellas):
 *  - Ingresos y Gastos: monto SIEMPRE positivo. Afectan ganancia y metas.
 *  - Transferencia    : monto CON SIGNO (+ entra a la cuenta, − sale). No es ingreso ni gasto.
 *                       Se usa para cambios de divisa, capital de préstamos/deudas y pago de tarjeta.
 *  - Ajuste           : monto CON SIGNO. Corrige el saldo de una cuenta (ver ajustarSaldoManual).
 *  - Gasto con tarjeta de crédito: CuentaID vacío y MedioPago 'Tarjeta de Credito'. Se reconoce como
 *    gasto en su FechaEfectiva (día de cobro). El efectivo sale recién al pagar el estado de cuenta,
 *    que es una Transferencia (así el gasto NO se cuenta dos veces).
 *  - Reportes, metas y presupuestos usan FechaEfectiva. El efectivo de las cuentas usa Fecha.
 */

// Posiciones de columna (base 0) de la hoja Movimientos
var M = { ID: 0, FECHA: 1, EFECTIVA: 2, TIPO: 3, MONTO: 4, MONEDA: 5, CUENTA: 6, CENTRO: 7, CAT: 8, SUB: 9, DESC: 10, MEDIO: 11, SUNAT: 12, ORIGEN: 13, NOTAS: 14, PROYECTO: 15 };

/** Lee Movimientos como objetos limpios (montos ya convertidos a número, fechas a Date) */
function leerMovimientos_() {
  const datos = leerHoja_('Movimientos');
  const lista = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (f[M.ID] === '' && f[M.TIPO] === '') continue;
    const fecha = fecha_(f[M.FECHA]);
    const efectiva = fecha_(f[M.EFECTIVA]) || fecha;
    if (!fecha && !efectiva) continue;
    lista.push({
      fila: i + 1, id: f[M.ID], fecha: fecha || efectiva, efectiva: efectiva || fecha, tipo: String(f[M.TIPO]).trim(),
      monto: numero_(f[M.MONTO]), moneda: String(f[M.MONEDA]).trim().toUpperCase(), cuentaId: f[M.CUENTA],
      centro: f[M.CENTRO], categoria: f[M.CAT], subcategoria: f[M.SUB], descripcion: String(f[M.DESC] || ''),
      medio: String(f[M.MEDIO]).trim(), sunat: f[M.SUNAT], origen: f[M.ORIGEN], notas: String(f[M.NOTAS] || ''), proyecto: f[M.PROYECTO] === undefined || f[M.PROYECTO] === null ? '' : String(f[M.PROYECTO])
    });
  }
  return lista;
}

/** Efecto en el efectivo de la cuenta de un movimiento (0 si no toca ninguna cuenta) */
function efectoCaja_(m) {
  if (m.cuentaId === '' || m.cuentaId === null || m.cuentaId === undefined) return 0;
  if (m.medio === 'Tarjeta de Credito' && esGasto_(m.tipo)) return 0; // aún no sale plata de la cuenta
  if (esIngreso_(m.tipo)) return m.monto;
  if (esGasto_(m.tipo)) return -m.monto;
  if (m.tipo === 'Transferencia' || m.tipo === 'Ajuste') return m.monto; // con signo
  return 0;
}

// ---------------------------------------------------------------
// UTILIDADES GENERALES
// ---------------------------------------------------------------

/** Genera el siguiente ID numérico correlativo para una hoja (columna A) */
function generarID_(nombreHoja) {
  const hoja = getHoja_(nombreHoja);
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return 1;
  const ids = hoja.getRange(2, 1, ultimaFila - 1, 1).getValues().map(function (r) { return numero_(r[0]); });
  return Math.max.apply(null, ids.concat([0])) + 1;
}

/** Suma (o resta, si delta es negativo) un monto al SaldoActual de una cuenta */
function actualizarSaldoCuenta_(cuentaId, delta) {
  const cuenta = obtenerCuenta_(cuentaId);
  if (!cuenta) throw new Error('Cuenta no encontrada: ' + cuentaId);
  getHoja_('Cuentas').getRange(cuenta.numeroFila, 6).setValue(redondear_(cuenta.saldo + delta));
}

/**
 * Autocategoriza una descripción según ReglasCategorizacion (coincidencia "contiene",
 * sin importar mayúsculas/tildes). Gana la prioridad más baja; a igual prioridad, la
 * palabra clave más larga (más específica). Sin coincidencia devuelve campos vacíos.
 */
function sinTildes_(t) {
  return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function autoCategorizar_(descripcion) {
  const datos = leerHoja_('ReglasCategorizacion');
  const texto = sinTildes_(descripcion);
  let mejor = null;
  for (let i = 1; i < datos.length; i++) {
    const clave = sinTildes_(datos[i][1]).trim();
    if (!clave || texto.indexOf(clave) === -1) continue;
    const prioridad = numero_(datos[i][5]) || 999;
    if (!mejor || prioridad < mejor.prioridad || (prioridad === mejor.prioridad && clave.length > mejor.largo)) {
      mejor = { categoria: datos[i][2], subcategoria: datos[i][3], centroCosto: limpiarCentroCosto_(datos[i][4]), prioridad: prioridad, largo: clave.length };
    }
  }
  return mejor || { categoria: '', subcategoria: '', centroCosto: '' };
}

/**
 * Fecha de cierre y de cobro de un consumo de tarjeta: el ciclo cierra el día DIA_CIERRE
 * (el consumo de ese mismo día entra) y se cobra el día DIA_COBRO del mes siguiente al cierre.
 */
function calcularFechaCobroTarjeta_(fechaConsumo, moneda, tarjeta) {
  const cfg = tarjeta || tarjetaPrincipal_();
  const dia = fechaConsumo.getDate();
  const mesCierre = dia > cfg.cierre ? fechaConsumo.getMonth() + 1 : fechaConsumo.getMonth();
  const anio = fechaConsumo.getFullYear();
  const enMes = function (m, d) { return Math.min(d, new Date(anio, m + 1, 0).getDate()); };
  return {
    fechaCierre: new Date(anio, mesCierre, enMes(mesCierre, cfg.cierre), 12, 0),
    fechaCobro: new Date(anio, mesCierre + 1, enMes(mesCierre + 1, cfg.cobro), 12, 0)
  };
}

/** Compara solo año/mes/día de dos fechas (ignora hora) */
function mismaFecha_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function completarCategoria_(datos) {
  let categoria = datos.categoria, subcategoria = datos.subcategoria, centro = datos.centroCosto;
  if (!categoria) {
    const auto = autoCategorizar_(datos.descripcion);
    categoria = auto.categoria;
    subcategoria = subcategoria || auto.subcategoria;
    centro = centro || auto.centroCosto;
  }
  registrarCategoriaSiNueva_(categoria, subcategoria, datos.tipo);
  return { categoria: categoria || '', subcategoria: subcategoria || '', centro: limpiarCentroCosto_(centro) };
}

// ---------------------------------------------------------------
// REGISTRO DE INGRESOS Y GASTOS (efectivo / débito)
// ---------------------------------------------------------------

/**
 * Registra un ingreso, o un gasto pagado en efectivo/débito (para tarjeta usa registrarConsumoTarjeta).
 * datos = { fecha, tipo, monto, moneda, cuentaId, descripcion, medioPago ('Efectivo'|'Debito'),
 *           centroCosto, categoria, subcategoria, categoriaSUNAT, origen, notas }
 * Devuelve el ID del movimiento creado.
 */
function registrarMovimiento(datos) {
  return conLock_(function () {
    const tipoIn = String(datos.tipo || '').trim();
    if (!esIngreso_(tipoIn) && !esGasto_(tipoIn) || TIPOS_MOVIMIENTO.indexOf(tipoIn) === -1) {
      throw new Error('Tipo de movimiento inválido: "' + datos.tipo + '".');
    }
    const monto = validarMonto_(datos.monto);
    const moneda = validarMoneda_(datos.moneda);
    validarCuentaMoneda_(datos.cuentaId, moneda);
    const fecha = fechaConHora_(datos.fecha);
    const cuentaMov = obtenerCuenta_(datos.cuentaId);
    const medio = cuentaMov.tipo === 'Efectivo' ? 'Efectivo' : 'Debito'; // el medio lo define el tipo de cuenta
    const cat = completarCategoria_(datos);
    const tipo = tipoIngresoPorCategoria_(tipoIn, cat.categoria); // activo/pasivo lo define la categoría del ingreso
    const origen = ['Manual', 'Bot Airbnb', 'Bot Ventas', 'Bot Finanzas', 'Tarjeta Credito', 'Recurrente'].indexOf(datos.origen) !== -1 ? datos.origen : 'Manual';
    const id = generarID_('Movimientos');

    getHoja_('Movimientos').appendRow([
      id, fecha, fecha, tipo, monto, moneda, datos.cuentaId,
      cat.centro, cat.categoria, cat.subcategoria, String(datos.descripcion || '').trim(),
      medio, 'No aplica', origen, datos.notas || ''
    ]);
    actualizarSaldoCuenta_(datos.cuentaId, esIngreso_(tipo) ? monto : -monto);

    registrarLog_('Movimiento', 'ID ' + id + ' — ' + tipo + ' ' + monto + ' ' + moneda + ' (' + (datos.descripcion || '') + ')');
    return id;
  });
}

/**
 * Registra una Transferencia (monto CON SIGNO) y mueve el saldo de la cuenta en el mismo paso,
 * para que movimiento y saldo nunca queden desfasados.
 * o = { fecha, monto (con signo), moneda, cuentaId, descripcion, categoria, subcategoria, notas, medioPago }
 */
function registrarTransferencia_(o) {
  return conLock_(function () {
    const monto = redondear_(numero_(o.monto));
    if (monto === 0) throw new Error('La transferencia no puede ser 0.');
    const moneda = validarMoneda_(o.moneda);
    validarCuentaMoneda_(o.cuentaId, moneda);
    const fecha = fechaConHora_(o.fecha);
    const id = generarID_('Movimientos');
    const cta = obtenerCuenta_(o.cuentaId);
    getHoja_('Movimientos').appendRow([
      id, fecha, fecha, 'Transferencia', monto, moneda, o.cuentaId,
      'Personal', o.categoria || 'Transferencias', o.subcategoria || '', o.descripcion || '',
      cta.tipo === 'Efectivo' ? 'Efectivo' : 'Debito', 'No aplica', 'Manual', o.notas || ''
    ]);
    actualizarSaldoCuenta_(o.cuentaId, monto);
    return id;
  });
}

// ---------------------------------------------------------------
// TARJETA DE CRÉDITO
// ---------------------------------------------------------------

/**
 * Registra un consumo con tarjeta de crédito (opción predeterminada de gasto).
 * No toca ningún saldo: el gasto se reconoce en la FechaEfectiva (día de cobro) y el efectivo
 * sale de tu cuenta cuando pagas el estado de cuenta (registrarPagoTarjeta).
 * datos = { fecha, tipo (Gasto*), monto, moneda, descripcion, esDesgravamen, centroCosto, categoria,
 *           subcategoria, categoriaSUNAT }
 */
function registrarConsumoTarjeta(datos) {
  return registrarConsumoTarjetaDetallado_(datos).idTarjeta;
}

/** Igual que registrarConsumoTarjeta pero devuelve { idTarjeta, idMov } (el ID del movimiento espejo) */
function registrarConsumoTarjetaDetallado_(datos) {
  return conLock_(function () {
    const tipo = datos.tipo || 'Gasto Personal';
    if (!esGasto_(tipo) || TIPOS_MOVIMIENTO.indexOf(tipo) === -1) throw new Error('Un consumo de tarjeta debe ser un gasto (recibí: "' + tipo + '").');
    const monto = validarMonto_(datos.monto);
    const moneda = validarMoneda_(datos.moneda);
    const fecha = fechaConHora_(datos.fecha);
    const tarjeta = tarjetaPorId_(datos.tarjetaId);
    const ciclos = calcularFechaCobroTarjeta_(fecha, moneda, tarjeta);
    const cat = completarCategoria_(datos);
    const sunat = limpiarCategoriaSunat_(datos.categoriaSUNAT);
    const descripcion = String(datos.descripcion || '').trim();

    const idTarjeta = generarID_('TarjetaCredito');
    getHoja_('TarjetaCredito').appendRow([
      idTarjeta, fecha, descripcion, monto, moneda, ciclos.fechaCierre, ciclos.fechaCobro, 'Pendiente',
      datos.esDesgravamen ? 'SI' : 'NO', sunat, cat.centro, cat.categoria, cat.subcategoria, tarjeta.id
    ]);

    const idMov = generarID_('Movimientos');
    getHoja_('Movimientos').appendRow([
      idMov, fecha, ciclos.fechaCobro, tipo, monto, moneda, '',
      cat.centro, cat.categoria, cat.subcategoria, descripcion,
      'Tarjeta de Credito', sunat, 'Tarjeta Credito', 'Ref. TarjetaCredito #' + idTarjeta + (datos.notas ? ' · ' + datos.notas : '')
    ]);

    FILAS_TARJETA_MEM_ = null;
    registrarLog_('ConsumoTarjeta', 'ID ' + idTarjeta + ' — ' + monto + ' ' + moneda + ' cobra el ' + fechaISO_(ciclos.fechaCobro));
    return { idTarjeta: idTarjeta, idMov: idMov };
  });
}

/** Estados de cuenta pendientes: [{ fechaCobro (Date), moneda, total, consumos }] ordenados por fecha */
function obtenerEstadosTarjetaPendientes_() {
  const datos = leerHoja_('TarjetaCredito');
  const filaTarjeta = mapaFilaTarjeta_();
  const mapa = {};
  for (let i = 1; i < datos.length; i++) {
    const estado = String(datos[i][7]).trim();
    if (estado === 'Pagado' || datos[i][0] === '') continue;
    const cobro = fecha_(datos[i][6]);
    if (!cobro) continue;
    const moneda = String(datos[i][4]).trim().toUpperCase();
    const tid = filaTarjeta[String(datos[i][0])];
    const clave = tid + '|' + fechaISO_(cobro) + '|' + moneda;
    if (!mapa[clave]) mapa[clave] = { tarjetaId: tid, tarjetaNombre: tarjetaPorId_(tid).banco, fechaCobro: cobro, fechaCierre: fecha_(datos[i][5]), moneda: moneda, total: 0, consumos: 0 };
    mapa[clave].total = redondear_(mapa[clave].total + numero_(datos[i][3]));
    mapa[clave].consumos++;
  }
  return Object.keys(mapa).map(function (k) { return mapa[k]; })
    .sort(function (a, b) { return a.fechaCobro - b.fechaCobro; });
}

/**
 * Paga el estado de cuenta (todos los consumos pendientes de esa fecha de cobro, en la moneda de la
 * cuenta elegida). Genera UNA Transferencia negativa desde la cuenta —no un gasto, porque el gasto ya
 * se reconoció al consumir— y marca los consumos como Pagado.
 */
function registrarPagoTarjeta(fechaCobro, cuentaId, fechaPago, tarjetaId) {
  return conLock_(function () {
    const cobro = fecha_(fechaCobro);
    if (!cobro) throw new Error('Fecha de cobro inválida: ' + fechaCobro);
    const cuenta = validarCuentaMoneda_(cuentaId, null);
    const hoja = getHoja_('TarjetaCredito');
    const datos = hoja.getDataRange().getValues();
    let total = 0;
    const filas = [];
    const mapaTj = mapaFilaTarjeta_();

    for (let i = 1; i < datos.length; i++) {
      const estado = String(datos[i][7]).trim();
      const f = fecha_(datos[i][6]);
      if (estado !== 'Pagado' && f && mismaFecha_(f, cobro) && String(datos[i][4]).trim().toUpperCase() === cuenta.moneda && (!tarjetaId || mapaTj[String(datos[i][0])] === String(tarjetaId))) {
        total += numero_(datos[i][3]);
        filas.push(i + 1);
      }
    }
    if (filas.length === 0) {
      throw new Error('No hay consumos pendientes en ' + cuenta.moneda + ' para el cobro del ' + fechaISO_(cobro) + '.');
    }
    total = redondear_(total);

    registrarTransferencia_({
      fecha: fecha_(fechaPago) || new Date(), monto: -total, moneda: cuenta.moneda, cuentaId: cuentaId,
      descripcion: 'Pago tarjeta de crédito ' + tarjetaPorId_(tarjetaId).banco + ' (cobro ' + fechaISO_(cobro) + ')',
      categoria: 'Tarjeta de Credito', subcategoria: 'Pago de estado de cuenta',
      notas: 'Pago de ' + filas.length + ' consumos. Los gastos ya se contaron al consumir.'
    });
    filas.forEach(function (fila) { hoja.getRange(fila, 8).setValue('Pagado'); });

    registrarLog_('PagoTarjeta', total + ' ' + cuenta.moneda + ' pagados desde cuenta ' + cuentaId + ' (' + filas.length + ' consumos)');
    return total;
  });
}

// ---------------------------------------------------------------
// AJUSTE MANUAL DE SALDO
// ---------------------------------------------------------------

/**
 * Fuerza el saldo de una cuenta al valor real de hoy. Deja un movimiento tipo 'Ajuste' (con signo),
 * así el saldo sigue siendo reproducible desde el historial y recalcular nunca lo deshace.
 */
function ajustarSaldoManual(cuentaId, saldoReal) {
  return conLock_(function () {
    const cuenta = obtenerCuenta_(cuentaId);
    if (!cuenta) throw new Error('Cuenta no encontrada: ' + cuentaId);
    const objetivo = redondear_(numero_(saldoReal));
    const diferencia = redondear_(objetivo - cuenta.saldo);
    if (diferencia === 0) return 0;

    const id = generarID_('Movimientos');
    getHoja_('Movimientos').appendRow([
      id, new Date(), new Date(), 'Ajuste', diferencia, cuenta.moneda, cuenta.id,
      'Personal', 'Ajuste de saldo', '', 'Ajuste manual a saldo real', 'Efectivo', 'No aplica', 'Manual', 'Saldo previo ' + cuenta.saldo
    ]);
    getHoja_('Cuentas').getRange(cuenta.numeroFila, 6).setValue(objetivo);
    getHoja_('AjustesSaldo').appendRow([generarID_('AjustesSaldo'), new Date(), cuenta.id, diferencia, 'Ajuste manual a saldo real']);
    registrarLog_('AjusteSaldo', 'Cuenta ' + cuenta.id + ' ajustada de ' + cuenta.saldo + ' a ' + objetivo);
    return diferencia;
  });
}


/* ==== 03_CambioDivisa.gs ==== */
/**
 * ============================================================
 * CAMBIO DE DIVISA
 * ============================================================
 */

/**
 * Registra un cambio de divisa (ej. comprar dólares con soles). Mueve el saldo de las dos cuentas
 * y guarda la tasa real aplicada, para calcular tu tasa promedio ponderada.
 * En Movimientos deja DOS transferencias con signo (sale de origen, entra a destino): no cuentan
 * como ingreso ni gasto.
 *
 * datos = { fecha, monedaOrigen, montoOrigen, cuentaOrigenId, monedaDestino, montoDestino,
 *           cuentaDestinoId, comision (opcional, en moneda destino; se registra como gasto aparte) }
 */
function registrarCambioDivisa(datos) {
  return conLock_(function () {
    const monedaOrigen = validarMoneda_(datos.monedaOrigen);
    const monedaDestino = validarMoneda_(datos.monedaDestino);
    if (monedaOrigen === monedaDestino) throw new Error('La moneda de origen y destino deben ser distintas.');
    const montoOrigen = validarMonto_(datos.montoOrigen, 'El monto origen');
    const montoDestino = validarMonto_(datos.montoDestino, 'El monto destino');
    const comision = numero_(datos.comision) > 0 ? redondear_(numero_(datos.comision)) : 0;
    validarCuentaMoneda_(datos.cuentaOrigenId, monedaOrigen);
    validarCuentaMoneda_(datos.cuentaDestinoId, monedaDestino);

    const fecha = fecha_(datos.fecha) || new Date();
    const tasa = montoDestino / montoOrigen;
    const id = generarID_('CambiosDivisa');

    getHoja_('CambiosDivisa').appendRow([
      id, fecha, monedaOrigen, montoOrigen, datos.cuentaOrigenId,
      monedaDestino, montoDestino, datos.cuentaDestinoId, tasa, comision
    ]);

    const detalle = 'Cambio ' + montoOrigen + ' ' + monedaOrigen + ' -> ' + montoDestino + ' ' + monedaDestino;
    registrarTransferencia_({
      fecha: fecha, monto: -montoOrigen, moneda: monedaOrigen, cuentaId: datos.cuentaOrigenId,
      descripcion: detalle, categoria: 'Cambio de Divisa', subcategoria: monedaOrigen + ' a ' + monedaDestino,
      notas: 'Ref. CambiosDivisa #' + id + ' (sale)', medioPago: 'Efectivo'
    });
    registrarTransferencia_({
      fecha: fecha, monto: montoDestino, moneda: monedaDestino, cuentaId: datos.cuentaDestinoId,
      descripcion: detalle, categoria: 'Cambio de Divisa', subcategoria: monedaOrigen + ' a ' + monedaDestino,
      notas: 'Ref. CambiosDivisa #' + id + ' (entra)', medioPago: 'Efectivo'
    });

    if (comision > 0) {
      registrarMovimiento({
        fecha: fecha, tipo: 'Gasto Personal', monto: comision, moneda: monedaDestino, cuentaId: datos.cuentaDestinoId,
        descripcion: 'Comisión de cambio ' + monedaOrigen + ' a ' + monedaDestino, medioPago: 'Debito',
        categoria: 'Comisiones', subcategoria: 'Cambio de divisa', centroCosto: 'Personal',
        notas: 'Ref. CambiosDivisa #' + id
      });
    }

    registrarLog_('CambioDivisa', 'ID ' + id + ' — ' + montoOrigen + ' ' + monedaOrigen + ' -> ' + montoDestino + ' ' + monedaDestino + ' (tasa ' + tasa.toFixed(4) + ')');
    return id;
  });
}

/**
 * Tu tasa promedio ponderada histórica entre dos monedas (destino por cada unidad de origen).
 * Ej.: tasaPromedioPonderada('PEN','USD') => dólares que obtuviste por cada sol cambiado.
 */
function tasaPromedioPonderada(monedaOrigen, monedaDestino) {
  const datos = leerHoja_('CambiosDivisa');
  let totalOrigen = 0, totalDestino = 0;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][2] === monedaOrigen && datos[i][5] === monedaDestino) {
      totalOrigen += numero_(datos[i][3]);
      totalDestino += numero_(datos[i][6]);
    }
  }
  return totalOrigen === 0 ? null : totalDestino / totalOrigen;
}


/* ==== 04_PrestamosDeudas.gs ==== */
/**
 * ============================================================
 * PRÉSTAMOS Y DEUDAS
 * ============================================================
 * "PrestamosOtorgados": dinero que TÚ prestas (ganas interés)
 * "Deudas": dinero que TÚ recibes prestado (pagas interés)
 *
 * Regla de oro: el desembolso o la devolución de CAPITAL es una Transferencia (con signo),
 * nunca ingreso ni gasto. Solo el INTERÉS cobrado o pagado afecta tu ganancia real.
 */

/** Acepta 3 (=3%) o 0.03 (=3%): todo lo mayor a 1 se interpreta como porcentaje */
function normalizarTasa_(t) {
  const n = numero_(t);
  if (!(n >= 0)) throw new Error('La tasa de interés no puede ser negativa.');
  return n > 1 ? n / 100 : n;
}

// ---------------- PRÉSTAMOS QUE TÚ OTORGAS ----------------

/**
 * datos = { deudor, moneda, montoPrestado, tasaInteres (0.03 = 3%), tipoTasa ('Mensual'|'Anual'),
 *           frecuencia ('Mensual'|'Al vencimiento'), fechaInicio, plazoMeses, cuentaOrigenId }
 */
function registrarPrestamoOtorgado(datos) {
  return conLock_(function () {
    const moneda = validarMoneda_(datos.moneda);
    const monto = validarMonto_(datos.montoPrestado, 'El monto prestado');
    const sinCuenta = datos.cuentaOrigenId === undefined || datos.cuentaOrigenId === '' || datos.cuentaOrigenId === null;
    if (!sinCuenta) validarCuentaMoneda_(datos.cuentaOrigenId, moneda);
    asegurarColumnaConcepto_();
    const tipoTasa = datos.tipoTasa === 'Anual' ? 'Anual' : 'Mensual';
    const frecuencia = datos.frecuencia === 'Al vencimiento' ? 'Al vencimiento' : 'Mensual';
    const fecha = fecha_(datos.fechaInicio) || new Date();
    const id = generarID_('PrestamosOtorgados');

    getHoja_('PrestamosOtorgados').appendRow([
      id, datos.deudor, moneda, monto, normalizarTasa_(datos.tasaInteres),
      tipoTasa, frecuencia, fecha, Math.max(1, Math.round(numero_(datos.plazoMeses)) || 1), 'Vigente',
      monto, datos.cuentaOrigenId || '', String(datos.concepto || '')
    ]);

    // El desembolso sale de tu cuenta pero no es gasto: el dinero pasa de "efectivo" a "por cobrar".
    // Si la compra se pagó con tarjeta (sin cuenta) el gasto ya quedó registrado en la tarjeta: solo queda la deuda.
    if (datos.cuentaOrigenId !== undefined && datos.cuentaOrigenId !== '' && datos.cuentaOrigenId !== null)
    registrarTransferencia_({
      fecha: fecha, monto: -monto, moneda: moneda, cuentaId: datos.cuentaOrigenId,
      descripcion: datos.concepto ? 'Compra para ' + datos.deudor + ': ' + datos.concepto : 'Préstamo otorgado a ' + datos.deudor, categoria: 'Prestamos y Deudas',
      notas: 'Ref. PrestamoOtorgado #' + id
    });

    registrarLog_('PrestamoOtorgado', 'ID ' + id + ' — ' + monto + ' ' + moneda + ' a ' + datos.deudor);
    return id;
  });
}

/** Interés mensual esperado de un préstamo otorgado (sobre el saldo pendiente) */
function interesMensualPrestamo_(fila) {
  const tasa = numero_(fila[4]);
  const mensual = fila[5] === 'Anual' ? Math.pow(1 + tasa, 1 / 12) - 1 : tasa;
  return redondear_(numero_(fila[10]) * mensual);
}

/** Cobro de interés de un préstamo que otorgaste. SÍ es ingreso pasivo (renta de 2da categoría) */
function registrarCobroInteresPrestamo(prestamoId, monto, cuentaDestinoId, fecha) {
  const prestamo = buscarFilaPorId_('PrestamosOtorgados', prestamoId);
  if (!prestamo) throw new Error('Préstamo no encontrado: ' + prestamoId);

  return registrarMovimiento({
    fecha: fecha, tipo: 'Ingreso Pasivo', monto: monto, moneda: prestamo.fila[2], cuentaId: cuentaDestinoId,
    descripcion: 'Interés cobrado — préstamo a ' + prestamo.fila[1], medioPago: 'Debito',
    categoria: 'Interes Prestamo Otorgado', subcategoria: String(prestamo.fila[1]), centroCosto: 'Personal',
    categoriaSUNAT: 'Segunda', notas: 'Ref. PrestamoOtorgado #' + prestamoId
  });
}

/** Devolución de CAPITAL de un préstamo otorgado (no es ingreso: reduce el saldo por cobrar) */
function registrarDevolucionCapitalPrestamo(prestamoId, monto, cuentaDestinoId, fecha) {
  return conLock_(function () {
    const p = buscarFilaPorId_('PrestamosOtorgados', prestamoId);
    if (!p) throw new Error('Préstamo no encontrado: ' + prestamoId);
    const m = validarMonto_(monto);
    const moneda = p.fila[2];
    const pendiente = numero_(p.fila[10]);
    if (m > pendiente + 0.005) throw new Error('Ese monto (' + m + ') supera el capital pendiente (' + pendiente + ').');
    validarCuentaMoneda_(cuentaDestinoId, moneda);

    const hoja = getHoja_('PrestamosOtorgados');
    const nuevoSaldo = redondear_(pendiente - m);
    hoja.getRange(p.numeroFila, 11).setValue(nuevoSaldo);
    if (nuevoSaldo <= 0.005) hoja.getRange(p.numeroFila, 10).setValue('Pagado');

    registrarTransferencia_({
      fecha: fecha, monto: m, moneda: moneda, cuentaId: cuentaDestinoId,
      descripcion: 'Devolución de capital — préstamo a ' + p.fila[1], categoria: 'Prestamos y Deudas',
      notas: 'Ref. PrestamoOtorgado #' + prestamoId
    });
    return nuevoSaldo;
  });
}

/**
 * Recordatorio diario: avisa cuando toca (o se atrasó 3 días) el cobro de interés mensual de un
 * préstamo vigente y aún no registraste ese interés este mes.
 */
function recordatorioInteresesPrestamos_() {
  const hoy = new Date();
  const datos = leerHoja_('PrestamosOtorgados');
  let movs = null;

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (String(f[9]).trim() === 'Pagado' || f[0] === '' || f[6] !== 'Mensual') continue;
    const inicio = fecha_(f[7]);
    if (!inicio) continue;
    const diaCobro = Math.min(inicio.getDate(), ultimoDiaMes_(hoy));
    if (hoy.getDate() !== diaCobro && hoy.getDate() !== Math.min(diaCobro + 3, ultimoDiaMes_(hoy))) continue;
    if (inicio > hoy) continue;

    movs = movs || leerMovimientos_();
    const yaCobrado = movs.some(function (m) {
      return m.categoria === 'Interes Prestamo Otorgado' && String(m.subcategoria) === String(f[1]) &&
        m.fecha.getMonth() === hoy.getMonth() && m.fecha.getFullYear() === hoy.getFullYear();
    });
    if (!yaCobrado) {
      enviarMensajeTelegram('💸 Hoy toca cobrar el interés de *' + escaparMd_(f[1]) + '*: ' +
        formatoMonto_(interesMensualPrestamo_(f), f[2]) + ' (préstamo #' + f[0] + '). Aún no lo registras este mes.');
    }
  }
}

// ---------------- DEUDAS QUE TÚ RECIBES ----------------

/**
 * datos = { acreedor, moneda, montoRecibido, tasaInteres, tipoTasa ('TEA'|'TCEA'|'Mensual'),
 *           fechaInicio, plazoMeses, cuentaDestinoId }
 * Genera el cronograma de cuotas (sistema francés, cuota fija).
 */
function registrarDeudaRecibida(datos) {
  return conLock_(function () {
    const moneda = validarMoneda_(datos.moneda);
    const monto = validarMonto_(datos.montoRecibido, 'El monto recibido');
    validarCuentaMoneda_(datos.cuentaDestinoId, moneda);
    const tipoTasa = ['TEA', 'TCEA', 'Mensual'].indexOf(datos.tipoTasa) !== -1 ? datos.tipoTasa : 'TEA';
    const plazo = Math.round(numero_(datos.plazoMeses));
    if (!(plazo >= 1)) throw new Error('El plazo debe ser de al menos 1 mes.');
    const tasa = normalizarTasa_(datos.tasaInteres);
    const fecha = fecha_(datos.fechaInicio) || new Date();
    const id = generarID_('Deudas');

    getHoja_('Deudas').appendRow([id, datos.acreedor, moneda, monto, tasa, tipoTasa, fecha, plazo, 'Vigente', monto, datos.cuentaDestinoId]);

    registrarTransferencia_({
      fecha: fecha, monto: monto, moneda: moneda, cuentaId: datos.cuentaDestinoId,
      descripcion: 'Deuda recibida de ' + datos.acreedor, categoria: 'Prestamos y Deudas', notas: 'Ref. Deuda #' + id
    });
    generarCronogramaCuotas_(id, 'Recibido', monto, tasa, tipoTasa, plazo, fecha);

    registrarLog_('DeudaRecibida', 'ID ' + id + ' — ' + monto + ' ' + moneda + ' de ' + datos.acreedor);
    return id;
  });
}

/** Cronograma con sistema francés. La última cuota ajusta centavos para que el capital sume exacto */
function generarCronogramaCuotas_(referenciaId, tipo, monto, tasaInteres, tipoTasa, plazoMeses, fechaInicio) {
  const i = convertirATasaMensual_(tasaInteres, tipoTasa);
  const cuotaFija = i === 0 ? monto / plazoMeses : monto * i / (1 - Math.pow(1 + i, -plazoMeses));

  const hoja = getHoja_('CronogramaCuotas');
  const idBase = generarID_('CronogramaCuotas');
  let saldo = monto;
  const filas = [];

  for (let n = 1; n <= plazoMeses; n++) {
    const interes = redondear_(saldo * i);
    let capital = n === plazoMeses ? redondear_(saldo) : redondear_(cuotaFija - interes);
    saldo = redondear_(saldo - capital);
    filas.push([idBase + n - 1, referenciaId, tipo, n, sumarMeses_(fechaInicio, n), capital, interes, 'Pendiente']);
  }
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
}

/** Convierte TEA/TCEA/mensual a tasa efectiva mensual (la TCEA incluye comisiones: úsala como referencia) */
function convertirATasaMensual_(tasa, tipoTasa) {
  if (tipoTasa === 'Mensual') return tasa;
  return Math.pow(1 + tasa, 1 / 12) - 1;
}

/** Paga una cuota: el interés es gasto financiero real; el capital es transferencia (reduce el pasivo) */
function registrarPagoCuotaDeuda(cronogramaFilaId, cuentaOrigenId, fecha) {
  return conLock_(function () {
    const cuota = buscarFilaPorId_('CronogramaCuotas', cronogramaFilaId);
    if (!cuota) throw new Error('Cuota no encontrada: ' + cronogramaFilaId);
    if (String(cuota.fila[7]).trim() === 'Pagado') throw new Error('Esa cuota ya está pagada.');

    const referenciaId = cuota.fila[1];
    const capital = numero_(cuota.fila[5]);
    const interes = numero_(cuota.fila[6]);
    const deuda = buscarFilaPorId_('Deudas', referenciaId);
    if (!deuda) throw new Error('La deuda #' + referenciaId + ' de esa cuota no existe.');
    const moneda = deuda.fila[2];
    validarCuentaMoneda_(cuentaOrigenId, moneda);
    const f = fecha_(fecha) || new Date();

    if (interes > 0) {
      registrarMovimiento({
        fecha: f, tipo: 'Gasto Fijo', monto: interes, moneda: moneda, cuentaId: cuentaOrigenId,
        descripcion: 'Interés pagado — deuda con ' + deuda.fila[1], medioPago: 'Debito',
        categoria: 'Intereses Pagados', subcategoria: String(deuda.fila[1]), centroCosto: 'Personal',
        notas: 'Ref. CronogramaCuotas #' + cronogramaFilaId
      });
    }
    if (capital > 0) {
      registrarTransferencia_({
        fecha: f, monto: -capital, moneda: moneda, cuentaId: cuentaOrigenId,
        descripcion: 'Amortización de capital — deuda con ' + deuda.fila[1], categoria: 'Prestamos y Deudas',
        notas: 'Ref. CronogramaCuotas #' + cronogramaFilaId
      });
    }

    const hojaDeudas = getHoja_('Deudas');
    const nuevoSaldo = redondear_(numero_(deuda.fila[9]) - capital);
    hojaDeudas.getRange(deuda.numeroFila, 10).setValue(Math.max(nuevoSaldo, 0));
    if (nuevoSaldo <= 0.005) hojaDeudas.getRange(deuda.numeroFila, 9).setValue('Pagado');

    getHoja_('CronogramaCuotas').getRange(cuota.numeroFila, 8).setValue('Pagado');
    registrarLog_('PagoCuotaDeuda', 'Cuota #' + cronogramaFilaId + ' pagada — capital ' + capital + ', interés ' + interes);
  });
}

// ---------------- UTILIDADES COMPARTIDAS ----------------

/** Busca una fila por ID en cualquier hoja. Devuelve { fila: [...], numeroFila } o null */
function buscarFilaPorId_(nombreHoja, id) {
  const datos = leerHoja_(nombreHoja);
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== '' && String(datos[i][0]) === String(id)) return { fila: datos[i], numeroFila: i + 1 };
  }
  return null;
}


/* ==== 05_Recurrentes.gs ==== */
/**
 * ============================================================
 * RECURRENTES
 * ============================================================
 * Corre diariamente (ver 12_Triggers.gs).
 *  - Gasto recurrente (suscripciones, cuotas) → se registra solo desde el día programado. Si el
 *    sistema no corrió ese día (por ejemplo falló el trigger), lo registra al día siguiente: no se pierde.
 *    En meses cortos (día 31 en febrero) se ejecuta el último día del mes.
 *  - Ingreso recurrente (alquiler, etc.) → NO se inventa. Si 3 días después del día esperado aún no
 *    llegó, te avisa por Telegram una sola vez ese mes.
 * Columna UltimaEjecucion: en gastos guarda cuándo se registró; en ingresos, cuándo se te avisó.
 */

var DIAS_GRACIA_INGRESO_RECURRENTE = 3;

function procesarRecurrentesDelDia() {
  const hoy = new Date();
  const datos = leerHoja_('Recurrentes');
  let movs = null;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[0] === '' || String(fila[10]).trim().toUpperCase() !== 'SI') continue;

    const dia = Math.min(Math.max(Math.round(numero_(fila[5])), 1), 31);
    const diaEfectivo = Math.min(dia, ultimoDiaMes_(hoy));
    const tipo = String(fila[2]).trim();

    try {
      if (String(fila[12] || '').trim().toUpperCase() === 'SI') {
        // Con confirmación: el día de pago se crea un programado pendiente; tú confirmas cuándo se pagó de verdad
        if (hoy.getDate() >= diaEfectivo) crearProgramadoDeRecurrente_(fila, i + 1, hoy, diaEfectivo);
      } else if (tipo === 'Gasto' && hoy.getDate() >= diaEfectivo) {
        registrarMovimientoRecurrente_(fila, i + 1, hoy, diaEfectivo);
      } else if (tipo === 'Ingreso' && hoy.getDate() >= Math.min(diaEfectivo + DIAS_GRACIA_INGRESO_RECURRENTE, ultimoDiaMes_(hoy))) {
        movs = movs || leerMovimientos_();
        verificarIngresoRecurrenteRecibido_(fila, i + 1, hoy, movs);
      }
    } catch (e) {
      registrarLog_('ERROR Recurrente', 'Fila ' + (i + 1) + ' (' + fila[1] + '): ' + e.message);
      enviarMensajeTelegram('⚠️ No pude procesar el recurrente *' + escaparMd_(fila[1]) + '*: ' + escaparMd_(e.message));
      getHoja_('Recurrentes').getRange(i + 1, 12).setValue(hoy); // evita reintentar y avisar todos los días
    }
  }
}

/** Recurrente con confirmación: crea el pendiente del mes (una sola vez) con aviso en el momento y recordatorios diarios si se atrasa */
function crearProgramadoDeRecurrente_(fila, numeroFila, hoy, diaEfectivo) {
  if (yaSeEjecutoEsteMes_(fila[11], hoy)) return;
  const tipo = String(fila[2]).trim();
  const sinCuenta = fila[9] === '' || fila[9] === null;
  const id = crearProgramado({
    tipo: tipo === 'Ingreso' ? tipoIngresoPorCategoria_('Ingreso Activo', fila[6]) : 'Gasto Fijo', monto: fila[3], moneda: fila[4], medio: sinCuenta ? 'Tarjeta' : 'Debito', cuentaId: sinCuenta ? '' : fila[9],
    descripcion: fila[1], categoria: fila[6], subcategoria: fila[7], fecha: new Date(hoy.getFullYear(), hoy.getMonth(), diaEfectivo, 9, 0), alertas: [0],
    notas: 'Recurrente #' + fila[0]
  });
  getHoja_('Recurrentes').getRange(numeroFila, 12).setValue(hoy);
  return id;
}

function registrarMovimientoRecurrente_(fila, numeroFila, hoy, diaEfectivo) {
  if (yaSeEjecutoEsteMes_(fila[11], hoy)) return;

  const fechaMov = new Date(hoy.getFullYear(), hoy.getMonth(), diaEfectivo, 12, 0);
  if (fila[9] === '' || fila[9] === null) {
    // Sin cuenta = se paga con la tarjeta de crédito
    registrarConsumoTarjeta({ fecha: fechaMov, tipo: 'Gasto Fijo', monto: fila[3], moneda: fila[4], descripcion: fila[1] + ' (recurrente)', categoria: fila[6], subcategoria: fila[7], centroCosto: fila[8] });
  } else {
    registrarMovimiento({
      fecha: fechaMov, tipo: 'Gasto Fijo', monto: fila[3], moneda: fila[4], cuentaId: fila[9],
      descripcion: fila[1] + ' (recurrente)', categoria: fila[6], subcategoria: fila[7], centroCosto: fila[8], origen: 'Recurrente'
    });
  }
  getHoja_('Recurrentes').getRange(numeroFila, 12).setValue(hoy);
}

/** ¿Ya llegó este mes un INGRESO cuya descripción contiene el nombre del recurrente? Si no, avisa 1 vez */
function verificarIngresoRecurrenteRecibido_(fila, numeroFila, hoy, movs) {
  if (yaSeEjecutoEsteMes_(fila[11], hoy)) return; // ya se avisó este mes

  const clave = sinTildes_(fila[1]);
  const yaLlego = movs.some(function (m) {
    return esIngreso_(m.tipo) && sinTildes_(m.descripcion).indexOf(clave) !== -1 &&
      m.fecha.getMonth() === hoy.getMonth() && m.fecha.getFullYear() === hoy.getFullYear();
  });
  if (yaLlego) return;

  enviarMensajeTelegram('⚠️ Aún no se registra el ingreso esperado de *' + escaparMd_(fila[1]) + '* (' +
    formatoMonto_(fila[3], fila[4]) + ') de este mes. Puede que aún no te paguen — revisa.\n_Tip: registra el ingreso con ese nombre en la descripción._');
  getHoja_('Recurrentes').getRange(numeroFila, 12).setValue(hoy);
}

function yaSeEjecutoEsteMes_(ultimaEjecucion, fecha) {
  const u = fecha_(ultimaEjecucion);
  if (!u) return false;
  return u.getMonth() === fecha.getMonth() && u.getFullYear() === fecha.getFullYear();
}


/* ==== 06_MetasPresupuestos.gs ==== */
/**
 * ============================================================
 * METAS Y PRESUPUESTOS
 * ============================================================
 * Las metas miden INGRESOS; los presupuestos miden GASTOS. Ambos usan FechaEfectiva.
 */

/** Fecha de inicio y fin de un periodo (Diaria/Semanal/Mensual/Anual) relativo a una fecha. Semana: lunes–domingo */
function obtenerRangoFecha_(periodo, fechaRef) {
  const f = fechaRef || new Date();
  let inicio, fin;

  if (periodo === 'Diaria') {
    inicio = new Date(f.getFullYear(), f.getMonth(), f.getDate());
    fin = new Date(f.getFullYear(), f.getMonth(), f.getDate(), 23, 59, 59, 999);
  } else if (periodo === 'Semanal') {
    const diaSemana = f.getDay() === 0 ? 7 : f.getDay();
    inicio = new Date(f.getFullYear(), f.getMonth(), f.getDate() - diaSemana + 1);
    fin = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate() + 6, 23, 59, 59, 999);
  } else if (periodo === 'Mensual') {
    inicio = new Date(f.getFullYear(), f.getMonth(), 1);
    fin = new Date(f.getFullYear(), f.getMonth() + 1, 0, 23, 59, 59, 999);
  } else { // Anual
    inicio = new Date(f.getFullYear(), 0, 1);
    fin = new Date(f.getFullYear(), 11, 31, 23, 59, 59, 999);
  }
  return { inicio: inicio, fin: fin };
}

/** Ingresos de un periodo/moneda vs. la meta configurada. `movs` opcional para no releer la hoja */
function calcularProgresoMeta(periodo, moneda, fechaRef, movs) {
  const rango = obtenerRangoFecha_(periodo, fechaRef);
  movs = movs || leerMovimientos_();

  let actual = 0;
  movs.forEach(function (m) {
    if (m.moneda === moneda && esIngreso_(m.tipo) && m.efectiva >= rango.inicio && m.efectiva <= rango.fin) actual += m.monto;
  });
  actual = redondear_(actual);

  const metaFila = buscarMetaVigente_(periodo, moneda, rango.inicio);
  let montoMeta = metaFila ? numero_(metaFila[3]) : 0;
  const crecimiento = metaFila ? numero_(metaFila[6]) : 0;
  let base = 0;
  if (crecimiento > 0) {
    // Meta relativa: ganar X% más que el periodo anterior (si el anterior no tuvo ingresos, rige el monto fijo)
    const rp = obtenerRangoFecha_(periodo, new Date(rango.inicio.getTime() - 86400000));
    movs.forEach(function (m) {
      if (m.moneda === moneda && esIngreso_(m.tipo) && m.efectiva >= rp.inicio && m.efectiva <= rp.fin) base += m.monto;
    });
    base = redondear_(base);
    if (base > 0) montoMeta = redondear_(base * (1 + crecimiento / 100));
  }
  return {
    actual: actual, meta: montoMeta, porcentaje: montoMeta > 0 ? (actual / montoMeta) * 100 : 0,
    falta: Math.max(redondear_(montoMeta - actual), 0), moneda: moneda, periodo: periodo, id: metaFila ? metaFila[0] : null,
    crecimiento: crecimiento, base: base,
    proxima: crecimiento > 0 && montoMeta > 0 ? redondear_(montoMeta * (1 + crecimiento / 100)) : 0
  };
}

/** Meta vigente para esa fecha. FechaInicio/FechaFin vacías = sin límite. Si hay varias, gana la de inicio más reciente */
function buscarMetaVigente_(periodo, moneda, fecha) {
  const datos = leerHoja_('Metas');
  let elegida = null, inicioElegida = null;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).trim() !== periodo || String(datos[i][2]).trim() !== moneda) continue;
    const inicio = fecha_(datos[i][4]);
    let fin = fecha_(datos[i][5]);
    if (fin) fin = new Date(fin.getFullYear(), fin.getMonth(), fin.getDate(), 23, 59, 59, 999);
    if (inicio && fecha < new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate())) continue;
    if (fin && fecha > fin) continue;
    if (!elegida || (inicio && (!inicioElegida || inicio > inicioElegida))) { elegida = datos[i]; inicioElegida = inicio; }
  }
  return elegida;
}

/** ¿Aplica este presupuesto al mes de `fecha`? MesAplicable vacío = todos los meses; acepta fecha, 1–12 o "2026-10" */
function presupuestoAplicaAlMes_(mesAplicable, fecha) {
  if (mesAplicable === '' || mesAplicable === null || mesAplicable === undefined) return true;
  if (mesAplicable instanceof Date) return mesAplicable.getMonth() === fecha.getMonth() && mesAplicable.getFullYear() === fecha.getFullYear();
  const s = String(mesAplicable).trim();
  const ym = s.match(/^(\d{4})-(\d{1,2})$/);
  if (ym) return +ym[1] === fecha.getFullYear() && +ym[2] === fecha.getMonth() + 1;
  const n = parseInt(s, 10);
  if (n >= 1 && n <= 12) return n === fecha.getMonth() + 1;
  return true;
}

/** Estado de cada presupuesto del mes de `fechaRef`: [{ categoria, subcategoria, moneda, limite, gastado }] */
function obtenerEstadoPresupuestos_(fechaRef, movs) {
  const ref = fechaRef || new Date();
  const rango = obtenerRangoFecha_('Mensual', ref);
  const presupuestos = leerHoja_('Presupuestos');
  movs = movs || leerMovimientos_();
  const resultado = [];

  for (let i = 1; i < presupuestos.length; i++) {
    const categoria = String(presupuestos[i][1]).trim();
    if (!categoria || !presupuestoAplicaAlMes_(presupuestos[i][5], ref)) continue;
    const subcategoria = String(presupuestos[i][2]).trim();
    const moneda = String(presupuestos[i][3]).trim();
    const limite = numero_(presupuestos[i][4]);

    let gastado = 0;
    movs.forEach(function (m) {
      if (esGasto_(m.tipo) && m.moneda === moneda && m.categoria === categoria &&
        (!subcategoria || m.subcategoria === subcategoria) && m.efectiva >= rango.inicio && m.efectiva <= rango.fin) gastado += m.monto;
    });
    resultado.push({ id: presupuestos[i][0], categoria: categoria, subcategoria: subcategoria, moneda: moneda, limite: limite, gastado: redondear_(gastado) });
  }
  return resultado;
}

/** Avisa por Telegram los presupuestos del mes que ya se pasaron del límite */
function verificarPresupuestos() {
  obtenerEstadoPresupuestos_().forEach(function (p) {
    if (p.limite > 0 && p.gastado > p.limite * 0.8 && p.gastado <= p.limite) {
      enviarMensajeTelegram('⚠️ Ya usaste el ' + Math.round(p.gastado / p.limite * 100) + '% del presupuesto de *' + escaparMd_(p.categoria) +
        (p.subcategoria ? ' / ' + escaparMd_(p.subcategoria) : '') + '*: ' + formatoMonto_(p.gastado, p.moneda) + ' de ' + formatoMonto_(p.limite, p.moneda) + '.');
    }
    if (p.limite > 0 && p.gastado > p.limite) {
      enviarMensajeTelegram('🚨 Presupuesto superado en *' + escaparMd_(p.categoria) +
        (p.subcategoria ? ' / ' + escaparMd_(p.subcategoria) : '') + '*: llevas ' + formatoMonto_(p.gastado, p.moneda) +
        ' de un límite de ' + formatoMonto_(p.limite, p.moneda) + ' este mes.');
    }
  });
}


/* ==== 08_TipoCambio.gs ==== */
/**
 * ============================================================
 * TIPO DE CAMBIO REFERENCIAL Y PATRIMONIO NETO CONSOLIDADO
 * ============================================================
 * Esta tasa es SOLO para expresar tu patrimonio (y totales mixtos) en un único número en soles.
 * Tus cambios reales usan la tasa que tú aplicas en registrarCambioDivisa().
 */

function actualizarTipoCambioReferencial() {
  try {
    const respuesta = UrlFetchApp.fetch(CONFIG.API_TIPO_CAMBIO_URL, { muteHttpExceptions: true });
    if (respuesta.getResponseCode() !== 200) throw new Error('La API respondió ' + respuesta.getResponseCode());
    const data = JSON.parse(respuesta.getContentText());
    if (!data.rates || !(data.rates.USD > 0) || !(data.rates.EUR > 0)) throw new Error('Respuesta sin tasas USD/EUR');

    const usdPen = 1 / data.rates.USD;
    const eurPen = 1 / data.rates.EUR;
    getHoja_('TipoCambioReferencial').appendRow([new Date(), usdPen, eurPen, usdPen / eurPen]);
    registrarLog_('TipoCambio', 'Actualizado: USD/PEN ' + usdPen.toFixed(4) + ', EUR/PEN ' + eurPen.toFixed(4));
  } catch (e) {
    registrarLog_('TipoCambio', 'ERROR al actualizar: ' + e.message);
  }
}

/** Última tasa referencial. Si aún no hay datos, usa un respaldo (marcado con respaldo: true) */
function obtenerUltimaTasaReferencial_() {
  const datos = leerHoja_('TipoCambioReferencial');
  for (let i = datos.length - 1; i >= 1; i--) {
    const usd = numero_(datos[i][1]), eur = numero_(datos[i][2]);
    if (usd > 0 && eur > 0) return { usdPen: usd, eurPen: eur, fecha: datos[i][0], respaldo: false };
  }
  return { usdPen: 3.7, eurPen: 4.0, fecha: null, respaldo: true };
}

function convertirAPEN_(monto, moneda, tasa) {
  if (moneda === 'USD') return monto * tasa.usdPen;
  if (moneda === 'EUR') return monto * tasa.eurPen;
  return monto;
}

/**
 * Patrimonio neto en soles, con su desglose:
 *   cuentas + préstamos por cobrar + valor de tus acciones − deudas pendientes − consumos de tarjeta aún no pagados
 * (los consumos de tarjeta ya son deuda aunque la cuenta todavía no se haya movido).
 */
function calcularPatrimonioDetalle_() {
  const tasa = obtenerUltimaTasaReferencial_();
  const d = { cuentas: 0, porCobrar: 0, deudas: 0, tarjeta: 0, inversiones: 0 };

  const cuentas = leerHoja_('Cuentas');
  for (let i = 1; i < cuentas.length; i++) {
    if (cuentas[i][0] === '') continue;
    d.cuentas += convertirAPEN_(numero_(cuentas[i][5]), String(cuentas[i][1]).trim(), tasa);
  }
  const prestamos = leerHoja_('PrestamosOtorgados');
  for (let i = 1; i < prestamos.length; i++) {
    if (prestamos[i][0] !== '' && String(prestamos[i][9]).trim() !== 'Pagado') {
      d.porCobrar += convertirAPEN_(numero_(prestamos[i][10]), String(prestamos[i][2]).trim(), tasa);
    }
  }
  const deudas = leerHoja_('Deudas');
  for (let i = 1; i < deudas.length; i++) {
    if (deudas[i][0] !== '' && String(deudas[i][8]).trim() !== 'Pagado') {
      d.deudas += convertirAPEN_(numero_(deudas[i][9]), String(deudas[i][2]).trim(), tasa);
    }
  }
  obtenerEstadosTarjetaPendientes_().forEach(function (e) {
    d.tarjeta += convertirAPEN_(e.total, e.moneda, tasa);
  });

  d.inversiones = valorInversionesPEN_(tasa);

  d.total = redondear_(d.cuentas + d.porCobrar + d.inversiones - d.deudas - d.tarjeta);
  d.cuentas = redondear_(d.cuentas); d.porCobrar = redondear_(d.porCobrar);
  d.deudas = redondear_(d.deudas); d.inversiones = redondear_(d.inversiones); d.tarjeta = redondear_(d.tarjeta);
  d.tasaRespaldo = tasa.respaldo;
  d.tasa = { usdPen: tasa.usdPen, eurPen: tasa.eurPen };
  return d;
}

function calcularPatrimonioNetoConsolidado() {
  return calcularPatrimonioDetalle_().total;
}


/* ==== 09_Telegram.gs ==== */
/**
 * ============================================================
 * TELEGRAM — mensajes, botones, comandos, texto libre y audios
 * ============================================================
 * Token, chat id y secreto viven en las Propiedades del script (ver 00_Config.gs).
 */

function telegramConfigurado_() {
  return !!tgToken_() && !!tgChat_();
}

/**
 * Envía un mensaje (Markdown). `botones` opcional: filas de botones [[{ text, data }]] (callback).
 * Parte textos largos y, si Telegram rechaza el Markdown, reenvía como texto plano.
 */
var MODO_RESPUESTA_ = false; // true mientras respondo a un mensaje tuyo: esas respuestas nunca se filtran

function enviarMensajeTelegram(texto, botones) {
  if (!MODO_RESPUESTA_) {
    // Antirrepetición: el mismo texto no se manda dos veces en 6 horas, y como máximo 8 avisos automáticos por hora
    try {
      const cache = CacheService.getScriptCache();
      const k = 'tgm_' + hash_(texto);
      if (cache.get(k)) { registrarLog_('Telegram', 'Aviso repetido omitido'); return true; }
      const hk = 'tgc_' + Utilities.formatDate(new Date(), zona_(), 'yyyyMMddHH');
      const n = Number(cache.get(hk) || 0);
      if (n >= 8) { registrarLog_('Telegram', 'Tope de avisos por hora alcanzado; se omitió uno'); return false; }
      cache.put(k, '1', 21600);
      cache.put(hk, String(n + 1), 3600);
    } catch (e) { /* si el caché falla, se envía igual */ }
  }
  if (!telegramConfigurado_()) { registrarLog_('Telegram', 'No configurado: agrega TELEGRAM_BOT_TOKEN y TELEGRAM_CHAT_ID en las propiedades del script'); return false; }
  const trozos = [];
  let resto = String(texto);
  while (resto.length > 3800) {
    let corte = resto.lastIndexOf('\n', 3800);
    if (corte < 1000) corte = 3800;
    trozos.push(resto.substring(0, corte));
    resto = resto.substring(corte);
  }
  trozos.push(resto);

  let ok = true;
  trozos.forEach(function (t, i) {
    const b = i === trozos.length - 1 ? botones : null;
    let r = enviarTextoTelegram_(t, 'Markdown', b);
    if (r.getResponseCode() !== 200) r = enviarTextoTelegram_(t.replace(/\\([_*`\[])/g, '$1'), null, b);
    if (r.getResponseCode() !== 200) { ok = false; registrarLog_('Telegram', 'Error ' + r.getResponseCode() + ': ' + r.getContentText().substring(0, 200)); }
  });
  return ok;
}

function enviarTextoTelegram_(texto, parseMode, botones) {
  const payload = { chat_id: tgChat_(), text: texto };
  if (parseMode) payload.parse_mode = parseMode;
  if (botones && botones.length) {
    payload.reply_markup = { inline_keyboard: botones.map(function (fila) { return fila.map(function (b) { return { text: b.text, callback_data: b.data }; }); }) };
  }
  return UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/sendMessage', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  });
}

function enviarDocumentoTelegram(blob, caption) {
  if (!telegramConfigurado_()) { registrarLog_('Telegram', 'No configurado: no se envió el documento'); return false; }
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/sendDocument', {
    method: 'post', payload: { chat_id: tgChat_(), document: blob, caption: caption || '' }, muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) {
    registrarLog_('Telegram', 'Error enviando documento ' + r.getResponseCode() + ': ' + r.getContentText().substring(0, 200));
    return false;
  }
  return true;
}

function responderCallbackTelegram_(callbackId, texto) {
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/answerCallbackQuery', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ callback_query_id: callbackId, text: texto || '' })
    });
  } catch (e) { /* no crítico */ }
}

/**
 * Registra el webhook. Pide la URL de la implementación de la aplicación web (termina en /exec).
 * El secreto viaja en la URL (?tg=...) porque Apps Script no puede leer cabeceras.
 * Menú ⚙️ Finanzas > Telegram: conectar webhook.
 */
function configurarWebhookTelegram() {
  const ui = SpreadsheetApp.getUi();
  if (!telegramConfigurado_() || !tgSecret_()) { ui.alert('Primero agrega TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID y TELEGRAM_WEBHOOK_SECRET en las propiedades del script.'); return; }
  const r = ui.prompt('Webhook de Telegram', 'Pega la URL de tu aplicación web (termina en /exec):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const base = r.getResponseText().trim();
  if (base.indexOf('https://script.google.com/') !== 0) { ui.alert('La URL debe empezar con https://script.google.com/'); return; }
  const url = base + (base.indexOf('?') === -1 ? '?' : '&') + 'tg=' + encodeURIComponent(tgSecret_());
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/setWebhook?drop_pending_updates=true&max_connections=1&allowed_updates=' +
    encodeURIComponent('["message","callback_query"]') + '&url=' + encodeURIComponent(url), { muteHttpExceptions: true });
  ui.alert(resp.getResponseCode() === 200 ? '✅ Webhook conectado. Escríbele /ayuda a tu bot.' : '❌ Telegram respondió: ' + resp.getContentText());
}

/**
 * Entrada del webhook (también la usa el dashboard para otras cosas, ver 13_WebApp.gs).
 * Solo responde si la URL trae el secreto correcto Y el mensaje viene de tu chat.
 * Se guarda el update_id: Telegram reintenta si tardas y así no se registra dos veces lo mismo.
 */
function doPost(e) {
  // Puerta de la app local (PWA): ver 29_ApiLocal.gs
  if (e && e.parameter && e.parameter.api === '1') return apiHttp_(e);
  MODO_RESPUESTA_ = true;
  try { return doPostInterno_(e); } finally { MODO_RESPUESTA_ = false; }
}

/**
 * Reclama un update de Telegram: devuelve true solo la primera vez que lo ve.
 * Guarda el último update_id procesado en las propiedades del script (los ids solo crecen) bajo un candado,
 * y además los últimos message_id como segunda red. Si cambias de bot (otro token) se reinicia solo.
 */
function reclamarUpdateTelegram_(update) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) { registrarLog_('Telegram', 'No pude obtener el candado para el update ' + update.update_id); return false; }
  try {
    const props = PropertiesService.getScriptProperties();
    const marca = hash_(tgToken_());
    const id = Number(update.update_id);
    if (isFinite(id)) {
      const g = String(props.getProperty('TG_LAST') || '').split(':');
      if (g[0] === marca && id <= Number(g[1])) return false;
      props.setProperty('TG_LAST', marca + ':' + id);
    }
    const msg = update.message;
    if (msg && msg.message_id) {
      let vistos = [];
      try { vistos = JSON.parse(props.getProperty('TG_MSGS') || '[]'); } catch (e) { vistos = []; }
      const clave = marca + '/' + msg.message_id;
      if (vistos.indexOf(clave) !== -1) return false;
      vistos.push(clave);
      props.setProperty('TG_MSGS', JSON.stringify(vistos.slice(-60)));
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

function doPostInterno_(e) {
  const ok = ContentService.createTextOutput('ok');
  try {
    if (!e || !e.parameter || !tgSecret_() || e.parameter.tg !== tgSecret_()) return ok;
    const update = JSON.parse(e.postData.contents);

    // Telegram reintenta el mismo update cuando tarda o recibe una redirección (Apps Script responde 302).
    // Por eso se marca ANTES de procesar y de forma permanente (no en caché): un reintento nunca vuelve a registrar.
    if (!reclamarUpdateTelegram_(update)) return ok;

    if (update.callback_query) { procesarCallbackTelegram_(update.callback_query); return ok; }

    const mensaje = update.message;
    if (!mensaje || String(mensaje.chat.id) !== String(tgChat_())) return ok;

    try {
      if (mensaje.voice || mensaje.audio) procesarAudioTelegram_((mensaje.voice || mensaje.audio).file_id);
      else if (mensaje.photo && mensaje.photo.length) procesarFotoTelegram_(mensaje.photo[mensaje.photo.length - 1].file_id, mensaje.caption || '');
      else if (mensaje.document && /^image\//.test(String(mensaje.document.mime_type || ''))) procesarFotoTelegram_(mensaje.document.file_id, mensaje.caption || '');
      else if (mensaje.text) procesarComandoTelegram_(mensaje.text.trim());
    } catch (err) {
      enviarMensajeTelegram('❌ ' + escaparMd_(err.message));
    }
  } catch (err) {
    registrarLog_('ERROR doPost', err.message);
  }
  return ok;
}

/** Botones bajo los recordatorios: "Ya se realizó" / "Cancelar" */
function procesarCallbackTelegram_(cb) {
  if (!cb.message || String(cb.message.chat.id) !== String(tgChat_())) return;
  const partes = String(cb.data || '').split(':');
  try {
    if (partes[0] === 'prog_ok') {
      const atras = Math.max(0, Math.min(30, Math.round(numero_(partes[2]))));
      const cuando = new Date(); cuando.setDate(cuando.getDate() - atras);
      if (atras > 0) cuando.setHours(12, 0, 0, 0);
      realizarProgramado(partes[1], atras > 0 ? { fecha: cuando } : {});
      responderCallbackTelegram_(cb.id, 'Registrado ✅');
      enviarMensajeTelegram('✅ Registrado' + (atras > 0 ? ' con fecha de hace ' + atras + (atras === 1 ? ' día' : ' días') : ' con la fecha de hoy') + '. ' + escaparMd_(textoPuntualidadProgramado_(partes[1])));
    } else if (partes[0] === 'pend') {
      resolverPendienteTelegram_(cb, partes[1] === 'si', partes[2]);
    } else if (partes[0] === 'undo') {
      responderCallbackTelegram_(cb.id, deshacerLote_(partes[1]) ? 'Deshecho' : 'Ya estaba deshecho');
    } else if (partes[0] === 'cat') {
      const gastoCat = partes[2] === 'g';
      const nombreCat = categoriasOrdenadas_(gastoCat)[Number(partes[3])];
      if (!nombreCat) { responderCallbackTelegram_(cb.id, 'Esa lista cambió, escríbeme de nuevo'); return; }
      editarMovimiento(partes[1], { categoria: nombreCat, subcategoria: '' });
      responderCallbackTelegram_(cb.id, 'Listo');
      enviarMensajeTelegram('🏷️ Categoría: ' + escaparMd_(nombreCat) + ' ✅');
    } else if (partes[0] === 'prog_no') {
      cancelarProgramado(partes[1]);
      responderCallbackTelegram_(cb.id, 'Cancelado');
      enviarMensajeTelegram('🗑 Programado cancelado.');
    }
  } catch (err) {
    responderCallbackTelegram_(cb.id, 'No se pudo');
    enviarMensajeTelegram('❌ ' + escaparMd_(err.message));
  }
}

var AYUDA_TELEGRAM_ =
  '*Puedes escribirme normal, mandarme un audio o la foto de un recibo*, por ejemplo: "gasté 5 soles en huevos", "me pagaron 600 del alquiler", "el viernes pago 80 dólares de internet".\n\n' +
  '*Soy inteligente*: detecto solo si es gasto o ingreso, la categoría, la cuenta o tarjeta, y si dividiste la cuenta ("cena 80 soles con la BCP, dividí con Ana y Luis"). Si me falta el monto u otro dato, te lo pregunto. Escribe "cancela" para descartar.\n\n' +
  '*Comandos*\n' +
  '/saldo — saldos y patrimonio\n' +
  '/meta — progreso de metas\n' +
  '/mes — resumen del mes\n' +
  '/tarjeta — estados de cuenta pendientes\n' +
  '/proximos — movimientos programados\n' +
  '/gasto 25.50 PEN almuerzo — gasto (con el medio predeterminado de esa moneda; tarjeta si no lo cambiaste)\n' +
  '/efectivo 25.50 PEN almuerzo — gasto en efectivo\n' +
  '/ingreso 300 PEN venta galletas — ingreso activo\n' +
  '/pasivo 600 PEN alquiler habitación — ingreso pasivo\n' +
  '/deshacer — borra lo último que registré (puedes repetirlo para ir hacia atrás)\n' +
  '/estado — revisa si el bot y la IA están funcionando';

function procesarComandoTelegram_(texto) {
  if (texto.charAt(0) !== '/') return procesarMensajeLibreTelegram_(texto);

  const partes = texto.split(/\s+/);
  const comando = partes[0].toLowerCase().replace(/@.*$/, '');

  if (comando === '/start' || comando === '/ayuda' || comando === '/help') return enviarMensajeTelegram(AYUDA_TELEGRAM_);
  if (comando === '/saldo') return enviarMensajeTelegram(construirMensajeSaldos_());
  if (comando === '/meta') return enviarMensajeTelegram(construirMensajeMetas_());
  if (comando === '/mes') return enviarMensajeTelegram(construirMensajeMes_());
  if (comando === '/tarjeta') return enviarMensajeTelegram(construirMensajeTarjeta_());
  if (comando === '/proximos') return enviarMensajeTelegram(construirMensajeProximos_());
  if (comando === '/deshacer') return deshacerUltimoBot_();
  if (comando === '/estado') return enviarMensajeTelegram(construirMensajeEstado_());

  if (['/gasto', '/efectivo', '/ingreso', '/pasivo'].indexOf(comando) !== -1) {
    const monto = numero_(partes[1]);
    const moneda = String(partes[2] || '').toUpperCase();
    const descripcion = partes.slice(3).join(' ').trim();

    if (!(monto > 0) || CONFIG.MONEDAS.indexOf(moneda) === -1 || !descripcion) {
      return enviarMensajeTelegram('Formato: ' + comando + ' MONTO MONEDA DESCRIPCION\nEj: ' + comando + ' 25.50 PEN almuerzo con amigos');
    }
    const auto = autoCategorizar_(descripcion);
    const gasto = comando === '/gasto' || comando === '/efectivo';
    const r = registrarOProgramar_({
      tipo: gasto ? 'Gasto Personal' : (comando === '/pasivo' ? 'Ingreso Pasivo' : 'Ingreso Activo'),
      monto: monto, moneda: moneda, descripcion: descripcion, medio: comando === '/efectivo' ? 'Efectivo' : ''
    });
    const lote = guardarDeshacer_([{ k: 'mov', id: r.idMov }]);
    return enviarMensajeTelegram('✅ Registrado: ' + (gasto ? 'gasto' : 'ingreso') + ' de ' + formatoMonto_(monto, moneda) + ' — "' + escaparMd_(descripcion) + '"' +
      (auto.categoria ? ' (' + escaparMd_(auto.categoria) + (auto.subcategoria ? ' / ' + escaparMd_(auto.subcategoria) : '') + ')'
        : '\n⚠️ Sin categoría reconocida: edítala en la app o agrega una regla en Categorías.'),
      [[{ text: '↩️ Deshacer', data: 'undo:' + lote }]]);
  }

  enviarMensajeTelegram('No reconozco ese comando. Escribe /ayuda.');
}

function construirMensajeSaldos_() {
  const cuentas = leerHoja_('Cuentas');
  let texto = '💰 *Tus saldos actuales*\n';
  for (let i = 1; i < cuentas.length; i++) {
    if (cuentas[i][0] === '' || String(cuentas[i][6]).trim().toUpperCase() === 'NO') continue;
    texto += '\n' + escaparMd_(cuentas[i][2]) + ' (' + cuentas[i][1] + '): ' + formatoMonto_(cuentas[i][5], String(cuentas[i][1]).trim());
  }
  const d = calcularPatrimonioDetalle_();
  texto += '\n\n📊 *Patrimonio neto estimado:* ' + formatoMonto_(d.total, 'PEN');
  if (d.inversiones > 0) texto += '\n📈 Inversiones: ' + formatoMonto_(d.inversiones, 'PEN');
  if (d.tarjeta > 0) texto += '\n💳 Tarjeta por pagar: ' + formatoMonto_(d.tarjeta, 'PEN');
  if (d.tasaRespaldo) texto += '\n_(tipo de cambio de respaldo: aún no hay uno actualizado)_';
  return texto;
}

function construirMensajeMetas_() {
  const movs = leerMovimientos_();
  let texto = '🎯 *Progreso de metas*\n';
  let alguna = false;
  ['Diaria', 'Semanal', 'Mensual'].forEach(function (periodo) {
    CONFIG.MONEDAS.forEach(function (moneda) {
      const p = calcularProgresoMeta(periodo, moneda, null, movs);
      if (p.meta === 0) return;
      alguna = true;
      texto += '\n' + periodo + ' (' + moneda + '): ' + formatoMonto_(p.actual, moneda) + ' / ' + formatoMonto_(p.meta, moneda) + ' (' + p.porcentaje.toFixed(0) + '%)';
    });
  });
  return alguna ? texto : texto + '\nAún no configuraste metas: créalas en la app (Más > Metas).';
}

function construirMensajeTarjeta_() {
  const estados = obtenerEstadosTarjetaPendientes_();
  if (estados.length === 0) return '💳 No tienes consumos de tarjeta pendientes de pago.';
  let texto = '💳 *Tarjetas — por pagar*\n';
  estados.forEach(function (e) {
    texto += '\n' + escaparMd_(e.tarjetaNombre) + ' · ' + fechaISO_(e.fechaCobro) + ': ' + formatoMonto_(e.total, e.moneda) + ' (' + e.consumos + ' consumos)';
  });
  tarjetasResumen_().forEach(function (t) {
    if (t.linea > 0) texto += '\n\n' + escaparMd_(t.banco) + ': usada ' + formatoMonto_(t.usada, t.moneda) + ' de ' + formatoMonto_(t.linea, t.moneda) + ' (' + t.utilizacion.toFixed(0) + '%)';
  });
  return texto;
}

function construirMensajeProximos_() {
  const r = resumenProgramados_(60);
  if (!r.lista.length) return '📅 No tienes movimientos programados pendientes.';
  let t = '📅 *Programados*\n';
  r.lista.slice(0, 15).forEach(function (p) {
    t += '\n' + (p.atrasado ? '⚠️ ' : '• ') + p.fecha.replace('T', ' ') + ' — ' + (esIngreso_(p.tipo) ? '+' : '−') + formatoMonto_(p.monto, p.moneda) + ' ' + escaparMd_(p.descripcion);
  });
  return t;
}

/** /estado: diagnóstico rápido desde el propio chat */
function construirMensajeEstado_() {
  let t = '🩺 *Estado del sistema*\n';
  t += '\nIA: ' + (iaProveedor_() ? 'conectada (' + iaProveedor_() + (iaProveedor_() === 'gemini' ? ', modelo ' + escaparMd_(geminiModelo_()) : '') + ')' : 'sin configurar');
  try {
    const w = infoWebhookTelegram_();
    t += '\nWebhook: ' + (w.url ? 'conectado' : 'NO conectado');
    t += '\nMensajes en cola: ' + (w.pendientes || 0);
    if (w.ultimoError) t += '\n⚠️ Último error de entrega: ' + escaparMd_(String(w.ultimoError).substring(0, 160));
  } catch (e) { t += '\nWebhook: no pude consultarlo'; }
  t += '\nDeshacer disponible: ' + leerPilaDeshacer_().length + ' registro(s)';
  return t;
}

/** Consulta cómo ve Telegram tu webhook (sin exponer el secreto) */
function infoWebhookTelegram_() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/getWebhookInfo', { muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText());
  if (!j.ok) throw new Error(j.description || 'Telegram no respondió');
  const i = j.result || {};
  return { url: i.url || '', pendientes: i.pending_update_count || 0, ultimoError: i.last_error_message || '', ultimoErrorFecha: i.last_error_date ? new Date(i.last_error_date * 1000).toISOString() : '' };
}

/**
 * Repara el webhook sin pedirte nada: vuelve a registrar la MISMA URL y descarta los mensajes viejos atascados
 * (lo que provoca registros repetidos). Menú ⚙️ Finanzas > Telegram: reparar, o botón en Más > Estado del sistema.
 */
function repararTelegram() {
  const r0 = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/getWebhookInfo', { muteHttpExceptions: true });
  const info = (JSON.parse(r0.getContentText()).result || {});
  if (!info.url) throw new Error('No hay webhook conectado. Usa Telegram: conectar webhook.');
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/setWebhook?drop_pending_updates=true&max_connections=1&allowed_updates=' +
    encodeURIComponent('["message","callback_query"]') + '&url=' + encodeURIComponent(info.url), { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Telegram respondió: ' + resp.getContentText().substring(0, 200));
  registrarLog_('Telegram', 'Webhook reparado (cola descartada: ' + (info.pending_update_count || 0) + ')');
  return { descartados: info.pending_update_count || 0 };
}

function repararTelegramMenu() {
  const ui = SpreadsheetApp.getUi();
  try { const r = repararTelegram(); ui.alert('✅ Webhook reparado. Mensajes viejos descartados: ' + r.descartados); }
  catch (e) { ui.alert('❌ ' + e.message); }
}


/* ==== 10_Notificaciones.gs ==== */
/**
 * ============================================================
 * NOTIFICACIONES PROGRAMADAS (ver 12_Triggers.gs)
 * ============================================================
 * Todas están envueltas en ejecutarSeguro_: si algo falla te llega el error por Telegram.
 */

/** Líneas de meta por moneda para un periodo; '' si no hay meta configurada en ninguna moneda */
function lineasMeta_(periodo, fechaRef, textoCumplida, textoFalta) {
  const movs = leerMovimientos_();
  let lineas = '';
  CONFIG.MONEDAS.forEach(function (moneda) {
    const p = calcularProgresoMeta(periodo, moneda, fechaRef, movs);
    if (p.meta === 0) return;
    lineas += '\n' + (p.actual >= p.meta
      ? '✅ ' + moneda + ': ' + formatoMonto_(p.actual, moneda) + ' de ' + formatoMonto_(p.meta, moneda) + ' ' + textoCumplida
      : '• ' + moneda + ': ' + formatoMonto_(p.actual, moneda) + ' de ' + formatoMonto_(p.meta, moneda) + ' — ' + textoFalta + ' ' + formatoMonto_(p.falta, moneda));
  });
  return lineas;
}

/** Avisa si hoy o mañana se cobra un estado de cuenta de la tarjeta */
function recordatorioTarjeta_() {
  if (String(getAjuste_('notif.tarjeta', 'SI')) === 'NO') return;
  const hoy = new Date();
  const manana = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() + 1);
  obtenerEstadosTarjetaPendientes_().forEach(function (e) {
    if (mismaFecha_(e.fechaCobro, hoy) || mismaFecha_(e.fechaCobro, manana)) {
      enviarMensajeTelegram('💳 ' + (mismaFecha_(e.fechaCobro, hoy) ? 'Hoy' : 'Mañana') + ' se cobra tu tarjeta ' + e.tarjetaNombre +
        ': ' + formatoMonto_(e.total, e.moneda) + '. Deja saldo y luego regístralo con el pago de tarjeta.');
    }
  });
}

function notificacionDiaria() {
  ejecutarSeguro_('notificacionDiaria', function () {
    procesarRecurrentesDelDia();
    if (String(getAjuste_('notif.intereses', 'SI')) !== 'NO') recordatorioInteresesPrestamos_();
    recordatorioTarjeta_();

    const lineas = String(getAjuste_('notif.meta_diaria', 'SI')) === 'NO' ? '' : lineasMeta_('Diaria', null, '¡meta diaria cumplida!', 'te faltan');
    if (lineas) enviarMensajeTelegram('📅 *Meta diaria*' + lineas);
  });
}

function notificacionSemanal() {
  ejecutarSeguro_('notificacionSemanal', function () {
    verificarPresupuestos();
    const lineas = String(getAjuste_('notif.meta_semanal', 'SI')) === 'NO' ? '' : lineasMeta_('Semanal', null, '¡meta semanal cumplida!', 'te faltaron');
    if (lineas) enviarMensajeTelegram('📆 *Semana cerrada*' + lineas);
    const resumen = resumenSemanalTexto_();
    if (resumen) enviarMensajeTelegram(resumen);
  });
}

/** Corre el día 1: cierra el MES ANTERIOR (meta y PDF) */
function notificacionMensual() {
  ejecutarSeguro_('notificacionMensual', function () {
    const hoy = new Date();
    const mesCerrado = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 15);
    const lineas = lineasMeta_('Mensual', mesCerrado, '¡meta mensual cumplida!', 'te faltaron');
    if (lineas) enviarMensajeTelegram('🗓️ *Mes cerrado: ' + nombreMes_(mesCerrado) + ' ' + mesCerrado.getFullYear() + '*' + lineas);

    generarYEnviarReporteMensual();
  });
}

function notificacionAnual() {
  ejecutarSeguro_('notificacionAnual', function () {
    if (new Date().getMonth() !== 0) return; // el trigger corre el día 2 de cada mes; solo actúa en enero
    generarYEnviarReporteAnual();
  });
}


/* ==== 11_Reportes.gs ==== */
/**
 * ============================================================
 * REPORTES — PDF mensual / anual por Telegram (estética de tu Excel: azul, tipografía serif)
 * ============================================================
 * Rellena las hojas "Plantilla Reporte Mensual" / "Plantilla Reporte Anual", las exporta como PDF y
 * las envía. Todo se calcula por FechaEfectiva (los gastos de tarjeta caen en el mes en que se cobran).
 * Capital inicial/final = saldos reales de las cuentas reconstruidos hacia atrás con los movimientos.
 */

var COLOR_REPORTE = { banda: '#0b5394', claro: '#cfe2f3', medio: '#6fa8dc', texto: '#0b3d6b', borde: '#9fc5e8' };
var FUENTE_REPORTE = 'Cambria';

function generarYEnviarReporteMensual(fechaRef) {
  const hoy = fechaRef || new Date();
  const mesCerrado = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 15);
  const hoja = prepararReporteMensual_(mesCerrado);
  enviarHojaComoPDF_(hoja, 'Reporte ' + nombreMes_(mesCerrado) + ' ' + mesCerrado.getFullYear());
}

function generarYEnviarReporteAnual(fechaRef) {
  const hoy = fechaRef || new Date();
  const anio = new Date(hoy.getFullYear() - 1, 5, 15);
  const hoja = prepararReporteAnual_(anio);
  enviarHojaComoPDF_(hoja, 'Reporte Anual ' + anio.getFullYear());
}

/** Utilidades para probar desde el menú (mes en curso / año en curso, sin esperar al cierre) */
function enviarReporteMesActualDePrueba() {
  enviarHojaComoPDF_(prepararReporteMensual_(new Date()), 'Reporte ' + nombreMes_(new Date()) + ' ' + new Date().getFullYear() + ' (parcial)');
}

function asegurarHojaPlantilla_(nombre) {
  let hoja = getSS_().getSheetByName(nombre);
  if (!hoja) {
    hoja = getSS_().insertSheet(nombre);
    hoja.setTabColor(COLOR_REPORTE.banda);
  }
  hoja.getDataRange().breakApart();
  hoja.clear();
  return hoja;
}

/** Capital (suma de saldos de cuentas) por moneda justo ANTES de `fecha`, reconstruido desde los movimientos */
function capitalAntesDe_(fecha, movs) {
  const cap = { PEN: 0, USD: 0, EUR: 0 };
  const cuentas = leerHoja_('Cuentas');
  for (let i = 1; i < cuentas.length; i++) {
    const mon = String(cuentas[i][1]).trim();
    if (cuentas[i][0] !== '' && cap.hasOwnProperty(mon)) cap[mon] += numero_(cuentas[i][5]);
  }
  movs.forEach(function (m) {
    if (m.fecha >= fecha && cap.hasOwnProperty(m.moneda)) cap[m.moneda] -= efectoCaja_(m);
  });
  CONFIG.MONEDAS.forEach(function (k) { cap[k] = redondear_(cap[k]); });
  return cap;
}

function vacioPorMoneda_() { return { PEN: 0, USD: 0, EUR: 0 }; }

/** Agrupa el periodo: totales por sección y moneda, listas de movimientos y gastos por categoría */
function analizarPeriodo_(rango, movs) {
  const secciones = { 'Ingresos activos': [], 'Ingresos pasivos': [], 'Gastos personales': [], 'Inversión': [], 'Gastos fijos': [] };
  const claveTipo = { 'Ingreso Activo': 'Ingresos activos', 'Ingreso Pasivo': 'Ingresos pasivos', 'Gasto Personal': 'Gastos personales', 'Gasto Inversion': 'Inversión', 'Gasto Fijo': 'Gastos fijos' };
  const totales = {};
  Object.keys(secciones).forEach(function (s) { totales[s] = vacioPorMoneda_(); });
  const porCategoria = {};
  const tasa = obtenerUltimaTasaReferencial_();

  movs.forEach(function (m) {
    if (m.efectiva < rango.inicio || m.efectiva > rango.fin) return;
    const sec = claveTipo[m.tipo];
    if (!sec || !totales[sec].hasOwnProperty(m.moneda)) return;
    secciones[sec].push(m);
    totales[sec][m.moneda] += m.monto;
    if (esGasto_(m.tipo)) {
      const cat = m.categoria || '(sin categoría)';
      porCategoria[cat] = (porCategoria[cat] || 0) + convertirAPEN_(m.monto, m.moneda, tasa);
    }
  });
  Object.keys(secciones).forEach(function (s) {
    secciones[s].sort(function (a, b) { return a.efectiva - b.efectiva; });
    CONFIG.MONEDAS.forEach(function (k) { totales[s][k] = redondear_(totales[s][k]); });
  });

  const ing = vacioPorMoneda_(), egr = vacioPorMoneda_();
  CONFIG.MONEDAS.forEach(function (k) {
    ing[k] = redondear_(totales['Ingresos activos'][k] + totales['Ingresos pasivos'][k]);
    egr[k] = redondear_(totales['Gastos personales'][k] + totales['Inversión'][k] + totales['Gastos fijos'][k]);
  });
  return { secciones: secciones, totales: totales, ingresos: ing, egresos: egr, porCategoria: porCategoria };
}

/** Aplica estilo de banda (encabezado) a un rango */
function estiloBanda_(rango, tamano) {
  rango.setBackground(COLOR_REPORTE.banda).setFontColor('#ffffff').setFontWeight('bold').setFontFamily(FUENTE_REPORTE)
    .setFontSize(tamano || 11).setVerticalAlignment('middle');
}

function estiloCeldas_(rango, fondo, negrita) {
  rango.setBackground(fondo).setFontColor(COLOR_REPORTE.texto).setFontFamily(FUENTE_REPORTE).setFontSize(10)
    .setFontWeight(negrita ? 'bold' : 'normal')
    .setBorder(true, true, true, true, true, true, COLOR_REPORTE.borde, SpreadsheetApp.BorderStyle.SOLID);
}

function formatoDinero_(hoja, fila, columna, filas, columnas) {
  hoja.getRange(fila, columna, filas, columnas).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
}

// ---------------------------------------------------------------
// REPORTE MENSUAL
// ---------------------------------------------------------------

function prepararReporteMensual_(fechaRef) {
  const hoja = asegurarHojaPlantilla_('Plantilla Reporte Mensual');
  const rango = obtenerRangoFecha_('Mensual', fechaRef);
  const movs = leerMovimientos_();
  const a = analizarPeriodo_(rango, movs);
  const capIni = capitalAntesDe_(rango.inicio, movs);
  const capFin = capitalAntesDe_(new Date(rango.fin.getTime() + 1), movs);

  hoja.setColumnWidths(1, 1, 70); hoja.setColumnWidths(2, 1, 260); hoja.setColumnWidths(3, 3, 105); hoja.setColumnWidths(6, 1, 70);
  hoja.getRange('A1').setValue(nombreMes_(fechaRef) + ' ' + fechaRef.getFullYear());
  hoja.getRange('A1:F1').merge(); estiloBanda_(hoja.getRange('A1:F1'), 16); hoja.setRowHeight(1, 36);

  let f = 3;
  hoja.getRange(f, 1, 1, 5).setValues([['Resumen', '', 'S/', '$', '€']]); estiloBanda_(hoja.getRange(f, 1, 1, 5));
  const metas = {}, faltan = {};
  const movsMeta = movs;
  CONFIG.MONEDAS.forEach(function (k) {
    const p = calcularProgresoMeta('Mensual', k, fechaRef, movsMeta);
    metas[k] = p.meta; faltan[k] = p.meta > 0 ? p.falta : 0;
  });
  const kpis = [
    ['Capital inicial', capIni], ['Ingresos', a.ingresos], ['Egresos', a.egresos],
    ['Ganancia', { PEN: redondear_(a.ingresos.PEN - a.egresos.PEN), USD: redondear_(a.ingresos.USD - a.egresos.USD), EUR: redondear_(a.ingresos.EUR - a.egresos.EUR) }],
    ['Meta de ingresos', metas], ['Faltante para la meta', faltan], ['Capital final', capFin]
  ];
  kpis.forEach(function (k) {
    f++;
    hoja.getRange(f, 1, 1, 5).setValues([[k[0], '', k[1].PEN, k[1].USD, k[1].EUR]]);
    hoja.getRange(f, 1, 1, 2).merge();
    estiloCeldas_(hoja.getRange(f, 1, 1, 5), COLOR_REPORTE.claro, k[0] === 'Ganancia' || k[0] === 'Capital final');
    formatoDinero_(hoja, f, 3, 1, 3);
  });

  // Gastos por categoría (equivalente en soles), de mayor a menor
  f += 2;
  hoja.getRange(f, 1, 1, 6).setValues([['Gastos por categoría', '', 'Equiv. S/', '', '', '%']]); estiloBanda_(hoja.getRange(f, 1, 1, 6));
  const cats = Object.keys(a.porCategoria).sort(function (x, y) { return a.porCategoria[y] - a.porCategoria[x]; });
  const totalCat = cats.reduce(function (s, c) { return s + a.porCategoria[c]; }, 0);
  if (cats.length === 0) { f++; hoja.getRange(f, 1).setValue('Sin gastos en el periodo'); estiloCeldas_(hoja.getRange(f, 1, 1, 6), COLOR_REPORTE.claro, false); }
  cats.forEach(function (c, i) {
    f++;
    hoja.getRange(f, 1, 1, 6).setValues([[c, '', redondear_(a.porCategoria[c]), '', '', totalCat ? a.porCategoria[c] / totalCat : 0]]);
    hoja.getRange(f, 1, 1, 2).merge();
    estiloCeldas_(hoja.getRange(f, 1, 1, 6), i === 0 ? '#ffe5b4' : COLOR_REPORTE.claro, i === 0);
    formatoDinero_(hoja, f, 3, 1, 1);
    hoja.getRange(f, 6).setNumberFormat('0.0%').setHorizontalAlignment('right');
  });
  if (cats.length > 0) { f++; hoja.getRange(f, 1).setValue('Mayor gasto: ' + cats[0] + ' (' + formatoMonto_(a.porCategoria[cats[0]], 'PEN') + ' equiv.)'); hoja.getRange(f, 1, 1, 6).merge().setFontFamily(FUENTE_REPORTE).setFontStyle('italic').setFontColor(COLOR_REPORTE.texto); }

  // Detalle por sección
  Object.keys(a.secciones).forEach(function (sec) {
    f += 2;
    hoja.getRange(f, 1, 1, 6).setValues([[sec, '', 'S/', '$', '€', '']]); estiloBanda_(hoja.getRange(f, 1, 1, 6));
    f++;
    hoja.getRange(f, 1, 1, 5).setValues([['Fecha', 'Descripción', '', '', '']]); estiloCeldas_(hoja.getRange(f, 1, 1, 5), COLOR_REPORTE.medio, true);
    const lista = a.secciones[sec];
    if (lista.length > 0) {
      const filas = lista.map(function (m) {
        return [m.efectiva.getDate() + '/' + (m.efectiva.getMonth() + 1), m.descripcion || '(sin descripción)',
          m.moneda === 'PEN' ? m.monto : '', m.moneda === 'USD' ? m.monto : '', m.moneda === 'EUR' ? m.monto : ''];
      });
      hoja.getRange(f + 1, 1, filas.length, 1).setNumberFormat('@');
      hoja.getRange(f + 1, 1, filas.length, 5).setValues(filas);
      estiloCeldas_(hoja.getRange(f + 1, 1, filas.length, 5), '#ffffff', false);
      formatoDinero_(hoja, f + 1, 3, filas.length, 3);
      f += filas.length;
    }
    f++;
    const t = a.totales[sec];
    hoja.getRange(f, 1, 1, 5).setValues([['Total', '', t.PEN, t.USD, t.EUR]]);
    hoja.getRange(f, 1, 1, 2).merge();
    estiloCeldas_(hoja.getRange(f, 1, 1, 5), COLOR_REPORTE.claro, true);
    formatoDinero_(hoja, f, 3, 1, 3);
  });

  SpreadsheetApp.flush();
  return hoja;
}

// ---------------------------------------------------------------
// REPORTE ANUAL
// ---------------------------------------------------------------

function prepararReporteAnual_(fechaRef) {
  const hoja = asegurarHojaPlantilla_('Plantilla Reporte Anual');
  const anio = fechaRef.getFullYear();
  const movs = leerMovimientos_();
  const rangoAnual = obtenerRangoFecha_('Anual', fechaRef);
  const anualA = analizarPeriodo_(rangoAnual, movs);
  const capIni = capitalAntesDe_(rangoAnual.inicio, movs);
  const capFin = capitalAntesDe_(new Date(rangoAnual.fin.getTime() + 1), movs);

  hoja.setColumnWidths(1, 1, 110); hoja.setColumnWidths(2, 9, 88);
  hoja.getRange('A1').setValue('Año ' + anio); hoja.getRange('A1:J1').merge(); estiloBanda_(hoja.getRange('A1:J1'), 16); hoja.setRowHeight(1, 36);

  // Tabla mensual: Mes | Ingresos S/ $ € | Egresos S/ $ € | Ganancia S/ $ €
  hoja.getRange(3, 1, 2, 10).setValues([
    ['', 'Ingresos', '', '', 'Egresos', '', '', 'Ganancia', '', ''],
    ['Mes', 'S/', '$', '€', 'S/', '$', '€', 'S/', '$', '€']]);
  estiloBanda_(hoja.getRange(3, 1, 2, 10));
  hoja.getRange(3, 2, 1, 3).merge(); hoja.getRange(3, 5, 1, 3).merge(); hoja.getRange(3, 8, 1, 3).merge();
  hoja.getRange(3, 1, 2, 10).setHorizontalAlignment('center');

  const tot = { ing: vacioPorMoneda_(), egr: vacioPorMoneda_() };
  const filas = [];
  for (let mes = 0; mes < 12; mes++) {
    const r = obtenerRangoFecha_('Mensual', new Date(anio, mes, 15));
    const a = analizarPeriodo_(r, movs);
    const fila = [MESES_ES[mes].charAt(0).toUpperCase() + MESES_ES[mes].slice(1)];
    ['PEN', 'USD', 'EUR'].forEach(function (k) { fila.push(a.ingresos[k]); tot.ing[k] += a.ingresos[k]; });
    ['PEN', 'USD', 'EUR'].forEach(function (k) { fila.push(a.egresos[k]); tot.egr[k] += a.egresos[k]; });
    ['PEN', 'USD', 'EUR'].forEach(function (k) { fila.push(redondear_(a.ingresos[k] - a.egresos[k])); });
    filas.push(fila);
  }
  const filaTotal = ['Total'];
  ['PEN', 'USD', 'EUR'].forEach(function (k) { filaTotal.push(redondear_(tot.ing[k])); });
  ['PEN', 'USD', 'EUR'].forEach(function (k) { filaTotal.push(redondear_(tot.egr[k])); });
  ['PEN', 'USD', 'EUR'].forEach(function (k) { filaTotal.push(redondear_(tot.ing[k] - tot.egr[k])); });
  filas.push(filaTotal);

  hoja.getRange(5, 1, filas.length, 10).setValues(filas);
  estiloCeldas_(hoja.getRange(5, 1, 12, 10), COLOR_REPORTE.claro, false);
  estiloCeldas_(hoja.getRange(17, 1, 1, 10), COLOR_REPORTE.medio, true);
  formatoDinero_(hoja, 5, 2, 13, 9);

  // Capital y resumen
  let f = 19;
  hoja.getRange(f, 1, 1, 4).setValues([['Resumen anual', 'S/', '$', '€']]); estiloBanda_(hoja.getRange(f, 1, 1, 4));
  [['Capital inicial', capIni], ['Capital final', capFin]].forEach(function (k) {
    f++;
    hoja.getRange(f, 1, 1, 4).setValues([[k[0], k[1].PEN, k[1].USD, k[1].EUR]]);
    estiloCeldas_(hoja.getRange(f, 1, 1, 4), COLOR_REPORTE.claro, true);
    formatoDinero_(hoja, f, 2, 1, 3);
  });

  // Distribución del gasto por sección (equivalente en soles)
  const tasa = obtenerUltimaTasaReferencial_();
  const secs = ['Gastos personales', 'Inversión', 'Gastos fijos'];
  const eq = secs.map(function (s) {
    return CONFIG.MONEDAS.reduce(function (acc, k) { return acc + convertirAPEN_(anualA.totales[s][k], k, tasa); }, 0);
  });
  const totalEq = eq.reduce(function (x, y) { return x + y; }, 0);
  f += 2;
  hoja.getRange(f, 1, 1, 3).setValues([['Distribución del gasto', 'Equiv. S/', '%']]); estiloBanda_(hoja.getRange(f, 1, 1, 3));
  secs.forEach(function (s, i) {
    f++;
    hoja.getRange(f, 1, 1, 3).setValues([[s, redondear_(eq[i]), totalEq ? eq[i] / totalEq : 0]]);
    estiloCeldas_(hoja.getRange(f, 1, 1, 3), COLOR_REPORTE.claro, false);
    hoja.getRange(f, 2).setNumberFormat('#,##0.00'); hoja.getRange(f, 3).setNumberFormat('0.0%');
  });

  // Top categorías del año
  f += 2;
  hoja.getRange(f, 1, 1, 3).setValues([['Gastos por categoría', 'Equiv. S/', '%']]); estiloBanda_(hoja.getRange(f, 1, 1, 3));
  const cats = Object.keys(anualA.porCategoria).sort(function (x, y) { return anualA.porCategoria[y] - anualA.porCategoria[x]; });
  const totalCat = cats.reduce(function (s, c) { return s + anualA.porCategoria[c]; }, 0);
  cats.forEach(function (c, i) {
    f++;
    hoja.getRange(f, 1, 1, 3).setValues([[c, redondear_(anualA.porCategoria[c]), totalCat ? anualA.porCategoria[c] / totalCat : 0]]);
    estiloCeldas_(hoja.getRange(f, 1, 1, 3), i === 0 ? '#ffe5b4' : COLOR_REPORTE.claro, i === 0);
    hoja.getRange(f, 2).setNumberFormat('#,##0.00'); hoja.getRange(f, 3).setNumberFormat('0.0%');
  });

  SpreadsheetApp.flush();
  return hoja;
}

// ---------------------------------------------------------------
// EXPORTAR Y ENVIAR
// ---------------------------------------------------------------

/** Exporta una hoja como PDF (A4) y la envía por Telegram. Falla con mensaje claro si Google no entrega un PDF */
function enviarHojaComoPDF_(hoja, tituloArchivo) {
  const ss = getSS_();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&gid=' + hoja.getSheetId() + '&size=A4&portrait=' + (hoja.getName().indexOf('Anual') !== -1 ? 'false' : 'true') +
    '&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';

  const respuesta = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  const tipo = String(respuesta.getHeaders()['Content-Type'] || '');
  if (respuesta.getResponseCode() !== 200 || tipo.indexOf('pdf') === -1) {
    throw new Error('Google no entregó el PDF (código ' + respuesta.getResponseCode() + '). Revisa que el script tenga permisos de Drive.');
  }
  const blob = respuesta.getBlob().setName(tituloArchivo + '.pdf');
  const enviado = enviarDocumentoTelegram(blob, '📄 ' + tituloArchivo);
  registrarLog_('Reporte', tituloArchivo + (enviado ? ' generado y enviado por Telegram' : ' generado pero NO se pudo enviar'));
}


/* ==== 12_Triggers.gs ==== */
/**
 * ============================================================
 * TRIGGERS — Instala los disparadores automáticos del sistema
 * ============================================================
 * Ejecuta instalarTriggers() UNA VEZ (menú ⚙️ Finanzas). Es seguro repetirlo: borra solo los
 * triggers de este sistema y los recrea, sin duplicar. Las horas usan la zona horaria del
 * proyecto (appsscript.json ya la fija en America/Lima).
 */

var FUNCIONES_TRIGGER_ = ['revisarAlertasProgramados', 'revisarAccionesAlertas', 'notificacionDiaria', 'notificacionSemanal', 'notificacionMensual', 'notificacionAnual', 'actualizarTipoCambioReferencial', 'respaldoSemanal'];

function instalarTriggers() {
  eliminarTriggersExistentes_();

  ScriptApp.newTrigger('notificacionDiaria').timeBased().everyDays(1).atHour(21).create();
  ScriptApp.newTrigger('notificacionSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(21).create();
  ScriptApp.newTrigger('notificacionMensual').timeBased().onMonthDay(1).atHour(9).create();
  ScriptApp.newTrigger('notificacionAnual').timeBased().onMonthDay(2).atHour(10).create(); // solo actúa en enero
  ScriptApp.newTrigger('respaldoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
  ScriptApp.newTrigger('actualizarTipoCambioReferencial').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('revisarAlertasProgramados').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('revisarAccionesAlertas').timeBased().everyMinutes(30).create();

  avisar_(
    '✅ Triggers instalados (hora ' + zona_() + '):\n' +
    '• Diario 9pm — recurrentes, recordatorios de tarjeta/intereses y meta diaria\n' +
    '• Domingo 9pm — presupuestos y meta semanal\n' +
    '• Domingo 3am — copia de seguridad en tu Drive (carpeta Respaldos Finanzas)\n' +
    '• Día 1, 9am — cierre del mes anterior: meta y PDF\n' +
    '• Día 2 de enero, 10am — PDF anual\n' +
    '• Todos los días 7am — tipo de cambio referencial\n' +
    '• Cada 15 minutos — alertas de movimientos programados\n' +
    '• Cada 30 minutos (horario de mercado) — alertas de acciones que suben o bajan mucho'
  );
}

function eliminarTriggersExistentes_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (FUNCIONES_TRIGGER_.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
}


/* ==== 13_WebApp.gs ==== */
/**
 * ============================================================
 * WEB APP — la app (Dashboard.html) y su API
 * ============================================================
 * Implementar > Nueva implementación > Aplicación web
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario  (el mismo despliegue sirve para la app y para el webhook de Telegram)
 * Seguridad: la página se puede abrir, pero NINGÚN dato se entrega sin el PIN (propiedad DASHBOARD_PIN del script).
 * Se pide el PIN SIEMPRE que abres la app: la sesión no se guarda en el dispositivo.
 * Cada vez que cambies el código: Implementar > Administrar implementaciones > editar > Nueva versión.
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.tg) return ContentService.createTextOutput('ok'); // ping del webhook
  let plantilla;
  try { plantilla = HtmlService.createTemplateFromFile('Dashboard'); }
  catch (x) { return ContentService.createTextOutput('✅ Tu hoja de finanzas está lista. Copia esta misma dirección (termina en /exec) y pégala en la app.'); }
  return plantilla
    .evaluate()
    .setTitle('Finanzas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------
// ACCESO
// ---------------------------------------------------------------

/** Solo se entra con el PIN (DASHBOARD_PIN). La sesión vive únicamente en la memoria de la pestaña: al abrir la app de nuevo vuelve a pedirlo. */
function usuarioActual_(token) {
  if (token && verificarToken_(token)) return { email: '(PIN)', nombre: 'Yo', rol: 'Admin', via: 'pin' };
  return null;
}

function secretoSesion_() {
  return prop_('SESSION_SECRET') || (prop_('DASHBOARD_PIN') + '|' + tgSecret_() + '|sesion');
}

function firmar_(texto) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, secretoSesion_()));
}

function crearToken_() {
  const exp = String(Date.now() + 12 * 3600000);
  return exp + '.' + firmar_(exp);
}

function verificarToken_(token) {
  const p = String(token).split('.');
  if (p.length !== 2 || !prop_('DASHBOARD_PIN')) return false;
  if (Number(p[0]) < Date.now()) return false;
  return firmar_(p[0]) === p[1];
}

function login_(args) {
  const pinCorrecto = prop_('DASHBOARD_PIN');
  if (!pinCorrecto) throw new Error('No hay un PIN configurado. Agrega DASHBOARD_PIN en las propiedades del script.');
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('pin_fallos') || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos. Espera 15 minutos.');
  if (String(args.pin || '').trim() !== pinCorrecto) {
    cache.put('pin_fallos', String(fallos + 1), 900);
    throw new Error('PIN incorrecto.');
  }
  return { token: crearToken_() };
}

// Endpoints que modifican datos / solo para Admin
var API_SOLO_LECTURA_RAPIDA_ = ['dashboard', 'categoria', 'tarjetas', 'tarjeta', 'objetivos'];
var API_ESCRITURA_ = ['registrar', 'editarMovimiento', 'eliminarMovimiento', 'traspaso', 'cambioDivisa', 'pagarTarjeta', 'crearCuenta', 'editarCuenta',
  'eliminarCuenta', 'ajustarSaldo', 'editarProgramado', 'cancelarProgramado', 'eliminarProgramado', 'realizarProgramado', 'guardarMeta', 'eliminarMeta',
  'guardarPresupuesto', 'eliminarPresupuesto', 'guardarRecurrente', 'eliminarRecurrente', 'guardarRegla', 'eliminarRegla', 'guardarAjustes',
  'operacionInversion', 'eliminarOperacion', 'crearPrestamo', 'guardarProyecto', 'eliminarProyecto', 'asignarProyecto', 'guardarFondo', 'respaldoAhora', 'enviarExportacion', 'compraParaOtro', 'cobrarPersona', 'editarPrestamo', 'cerrarPrestamo', 'renombrarPersona', 'cobrarInteres', 'devolverCapital', 'crearDeuda', 'pagarCuota',
  'configurarTarjeta','guardarTarjeta','eliminarTarjeta','repararTelegram', 'guardarObjetivos', 'guardarCategoria', 'guardarNaturaleza', 'guardarMetaAhorro', 'eliminarMetaAhorro', 'moverMetaAhorro', 'eliminarCategoria', 'guardarSubcategoria', 'eliminarSubcategoria', 'guardarWatchlist', 'registrarDividido', 'registrarDivision', 'eliminarDivision', 'importarDesdeLocal', 'convertirParaOtro'];
var API_ADMIN_ = ['enviarReporte', 'estadoSistema', 'importarDesdeLocal', 'guardarClaveIA', 'guardarTelegram'];

/**
 * Único punto de entrada desde la página: apiLlamar(nombre, args, token). Devuelve siempre un objeto simple:
 *   { ok: true, data } | { __auth: true, ... } | { __error: 'mensaje' }
 */
function apiLlamar(nombre, args, token) {
  try {
    args = args || {};
    if (nombre === 'login') return { ok: true, data: login_(args) };

    const u = usuarioActual_(token);
    if (nombre === 'sesion') {
      return { ok: true, data: { autorizado: !!u, usuario: u, pinDisponible: !!prop_('DASHBOARD_PIN') } };
    }
    if (!u) return { __auth: true };
    asegurarEstructura_();

    const fn = globalThis['api_' + nombre];
    if (typeof fn !== 'function') throw new Error('Operación desconocida: ' + nombre);
    if (API_ESCRITURA_.indexOf(nombre) !== -1 && u.rol === 'Lector') throw new Error('Tu acceso es de solo lectura.');
    if (API_ADMIN_.indexOf(nombre) !== -1 && u.rol !== 'Admin') throw new Error('Solo un administrador puede hacer esto.');
    if (API_ESCRITURA_.indexOf(nombre) !== -1 && esDuplicado_(args.token)) return { ok: true, data: null, duplicado: true };

    let r;
    if (API_SOLO_LECTURA_RAPIDA_.indexOf(nombre) !== -1) {
      MEMO_LECTURA_ = {};
      try { r = fn(args, u); } finally { MEMO_LECTURA_ = null; }
    } else r = fn(args, u);
    return { ok: true, data: JSON.parse(JSON.stringify(r === undefined ? null : r)) };
  } catch (e) {
    try { if (args && args.token) CacheService.getScriptCache().remove('op_' + args.token); } catch (x) { /* ignorar */ }
    registrarLogSeguro_('ERROR api ' + nombre, e.message);
    return { __error: e.message };
  }
}

function registrarLogSeguro_(a, d) { try { registrarLog_(a, d); } catch (e) { /* ignorar */ } }

/** Evita registrar dos veces lo mismo si tocas "Guardar" dos veces o el celular reintenta la llamada */
function esDuplicado_(token) {
  if (!token) return false;
  const cache = CacheService.getScriptCache();
  const clave = 'op_' + token;
  if (cache.get(clave)) return true;
  cache.put(clave, '1', 600);
  return false;
}

// ---------------------------------------------------------------
// LECTURA: dashboard
// ---------------------------------------------------------------

function obtenerCuentasActivas_() {
  const datos = leerHoja_('Cuentas');
  const cuentas = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== '' && String(datos[i][6]).trim().toUpperCase() !== 'NO') {
      cuentas.push({
        id: datos[i][0], moneda: String(datos[i][1]).trim(), banco: String(datos[i][2]).trim(),
        tipo: String(datos[i][3]).trim() || 'Banco', saldo: redondear_(numero_(datos[i][5])), color: String(datos[i][7] || '')
      });
    }
  }
  return cuentas;
}

function cuentasArchivadas_() {
  const datos = leerHoja_('Cuentas');
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] !== '' && String(datos[i][6]).trim().toUpperCase() === 'NO') {
      out.push({ id: datos[i][0], moneda: String(datos[i][1]).trim(), banco: String(datos[i][2]).trim(), tipo: String(datos[i][3]).trim() || 'Banco', saldo: redondear_(numero_(datos[i][5])), color: String(datos[i][7] || '') });
    }
  }
  return out;
}

function movimientoParaApp_(m, nombres) {
  return {
    id: m.id, fecha: fechaHoraISO_(m.fecha), efectiva: fechaISO_(m.efectiva), tipo: m.tipo, monto: m.monto, moneda: m.moneda,
    categoria: String(m.categoria || ''), subcategoria: String(m.subcategoria || ''), descripcion: m.descripcion, medioPago: m.medio,
    cuentaId: m.cuentaId, proyecto: String(m.proyecto || ''), cuenta: m.cuentaId !== '' ? (nombres[String(m.cuentaId)] || '') : (m.medio === 'Tarjeta de Credito' ? nombreTarjetaDeMov_(m) : '')
  };
}

/** Solo movimientos que ya ocurrieron (fecha <= ahora), del más reciente al más antiguo */
function movimientosReales_(movs) {
  const ahora = new Date();
  return movs.filter(function (m) { return m.fecha <= ahora; })
    .sort(function (a, b) { return (b.fecha - a.fecha) || (numero_(b.id) - numero_(a.id)); });
}

/** Resumen del periodo: totales reales por moneda, categorías (equivalente en soles y por moneda) */
function resumenPeriodo_(periodo, fechaRef, movs, tasa) {
  const rango = obtenerRangoFecha_(periodo, fechaRef);
  const porMoneda = {};
  CONFIG.MONEDAS.forEach(function (k) { porMoneda[k] = { ingresos: 0, egresos: 0 }; });
  const categorias = {}, crudas = {};
  let ingresosEq = 0, egresosEq = 0, ingActivo = 0, ingPasivo = 0;

  movs.forEach(function (m) {
    if (m.efectiva < rango.inicio || m.efectiva > rango.fin || !porMoneda[m.moneda]) return;
    const eq = convertirAPEN_(m.monto, m.moneda, tasa);
    if (esIngreso_(m.tipo)) {
      porMoneda[m.moneda].ingresos += m.monto; ingresosEq += eq;
      if (m.tipo === 'Ingreso Pasivo') ingPasivo += eq; else ingActivo += eq;
    }
    else if (esGasto_(m.tipo)) {
      porMoneda[m.moneda].egresos += m.monto; egresosEq += eq;
      const cat = m.categoria || 'Sin categoría';
      categorias[cat] = (categorias[cat] || 0) + eq;
      crudas[cat + '|' + m.moneda] = (crudas[cat + '|' + m.moneda] || 0) + m.monto;
    }
  });

  const lista = Object.keys(categorias).map(function (c) {
    return { categoria: c, monto: redondear_(categorias[c]), porcentaje: egresosEq > 0 ? categorias[c] / egresosEq * 100 : 0 };
  }).sort(function (a, b) { return b.monto - a.monto; });
  const porCategoriaMoneda = Object.keys(crudas).map(function (k) {
    const p = k.split('|');
    return { categoria: p[0], moneda: p[1], monto: redondear_(crudas[k]) };
  });

  CONFIG.MONEDAS.forEach(function (k) {
    porMoneda[k].ingresos = redondear_(porMoneda[k].ingresos);
    porMoneda[k].egresos = redondear_(porMoneda[k].egresos);
  });
  return {
    ingresos: redondear_(ingresosEq), egresos: redondear_(egresosEq), ganancia: redondear_(ingresosEq - egresosEq),
    ingActivo: redondear_(ingActivo), ingPasivo: redondear_(ingPasivo),
    categorias: lista, mayorGasto: lista.length ? lista[0] : null, porMoneda: porMoneda, porCategoriaMoneda: porCategoriaMoneda
  };
}

/** Ingresos/egresos por mes (últimos n) y por moneda, y capital en cuentas a fin de cada mes (en soles) */
function seriesMensuales_(movs, tasa, n) {
  const hoy = new Date();
  const serie = [], capital = [];
  for (let k = n - 1; k >= 0; k--) {
    const ref = new Date(hoy.getFullYear(), hoy.getMonth() - k, 15);
    const rango = obtenerRangoFecha_('Mensual', ref);
    const pm = {};
    CONFIG.MONEDAS.forEach(function (c) { pm[c] = { ingresos: 0, egresos: 0 }; });
    movs.forEach(function (m) {
      if (m.efectiva < rango.inicio || m.efectiva > rango.fin || !pm[m.moneda]) return;
      if (esIngreso_(m.tipo)) pm[m.moneda].ingresos += m.monto; else if (esGasto_(m.tipo)) pm[m.moneda].egresos += m.monto;
    });
    const etiqueta = Utilities.formatDate(ref, zona_(), 'yyyy-MM');
    serie.push({ mes: etiqueta, porMoneda: pm });
    const finMes = k === 0 ? new Date(Date.now() + 1) : new Date(rango.fin.getTime() + 1);
    const cap = capitalAntesDe_(finMes, movs.filter(function (m) { return m.fecha <= new Date(); }));
    capital.push({ mes: etiqueta, PEN: cap.PEN, USD: cap.USD, EUR: cap.EUR });
  }
  return { serie: serie, capital: capital };
}

/** Ejecuta fn y devuelve null si falla (para que un dato secundario no tumbe el dashboard) */
function seguro_(fn) { try { return fn(); } catch (e) { registrarLog_('ERROR dashboard', e.message); return null; } }

function api_dashboard(args, usuario) {
  const movs = leerMovimientos_();
  const tasa = obtenerUltimaTasaReferencial_();
  const hoy = new Date();
  const nombres = mapaNombresCuentas_();

  // Cotizaciones recientes (como mucho cada 2 min) para que el patrimonio incluya el valor de mercado de tus acciones
  try {
    const pos = calcularPosiciones_(leerOperacionesInversion_(), false);
    const tk = Object.keys(pos).filter(function (t) { return pos[t].cantidad > 0; });
    if (tk.length) actualizarCotizaciones_(tk, 2);
  } catch (e) { /* si Yahoo falla se usan los últimos precios guardados */ }

  const metas = [];
  ['Diaria', 'Semanal', 'Mensual'].forEach(function (periodo) {
    CONFIG.MONEDAS.forEach(function (moneda) {
      const p = calcularProgresoMeta(periodo, moneda, hoy, movs);
      if (p.meta > 0 || p.crecimiento > 0) metas.push(p);
    });
  });

  const prog = resumenProgramados_(7);
  const series = seriesMensuales_(movs, tasa, 12);
  const mesPrev = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 15);
  const anioPrev = new Date(hoy.getFullYear() - 1, 5, 15);
  const mensual = resumenPeriodo_('Mensual', hoy, movs, tasa);
  mensual.prev = resumenPeriodo_('Mensual', mesPrev, movs, tasa);
  const anual = resumenPeriodo_('Anual', hoy, movs, tasa);
  anual.prev = resumenPeriodo_('Anual', anioPrev, movs, tasa);

  return {
    hoy: fechaISO_(hoy), ahora: fechaHoraISO_(hoy), usuario: usuario,
    cuentas: obtenerCuentasActivas_(), archivadas: cuentasArchivadas_(),
    patrimonio: calcularPatrimonioDetalle_(),
    metas: metas, presupuestos: obtenerEstadoPresupuestos_(hoy, movs),
    tarjeta: obtenerEstadosTarjetaPendientes_().map(function (e) { return { fechaCobro: fechaISO_(e.fechaCobro), moneda: e.moneda, total: e.total, consumos: e.consumos, tarjetaId: e.tarjetaId, banco: e.tarjetaNombre, cerrado: !!e.fechaCierre && e.fechaCierre <= hoy }; }),
    tarjetas: tarjetasResumen_(),
    distribucion: { Mensual: distribucionIngresos_('Mensual', hoy, movs, tasa), Anual: distribucionIngresos_('Anual', hoy, movs, tasa) },
    proximos: prog.lista, netoProyectado: prog.netoProyectado,
    resumen: { Mensual: mensual, Anual: anual },
    plan: seguro_(function () { return api_plan(); }), metasAhorro: seguro_(function () { return api_metasAhorro(); }) || [],
    series: series.serie, capital: series.capital,
    recientes: movimientosReales_(movs).slice(0, 12).map(function (m) { return movimientoParaApp_(m, nombres); }),
    ajustes: ajustesParaApp_(), catalogo: catalogoCategorias_(), proyectos: leerProyectos_().filter(function (p) { return p.estado !== 'Cerrado'; }).map(function (p) { return { id: p.id, nombre: p.nombre }; }), fondo: fondoEmergencia_(movs, tasa),  tiposCuenta: TIPOS_CUENTA_,
    iaDisponible: !!iaProveedor_(), tasaRespaldo: tasa.respaldo
  };
}

/** Lista de movimientos con búsqueda y filtros. args = { q, tipo, moneda, cuentaId, desde, hasta, pagina } */
function api_movimientos(args) {
  const nombres = mapaNombresCuentas_();
  const q = sinTildes_(args.q || '');
  const desde = fecha_(args.desde), hasta = fecha_(args.hasta);
  const hastaFin = hasta ? new Date(hasta.getFullYear(), hasta.getMonth(), hasta.getDate(), 23, 59, 59, 999) : null;
  const filtrados = movimientosReales_(leerMovimientos_()).filter(function (m) {
    if (args.tipo && args.tipo !== 'Todos') {
      if (args.tipo === 'Ingresos' && !esIngreso_(m.tipo)) return false;
      if (args.tipo === 'Gastos' && !esGasto_(m.tipo)) return false;
      if (args.tipo === 'Transferencias' && m.tipo !== 'Transferencia' && m.tipo !== 'Ajuste') return false;
    }
    if (args.tipoExacto && m.tipo !== args.tipoExacto) return false;
    if (args.proyecto === 'none' && String(m.proyecto || '')) return false;
    if (args.proyecto && args.proyecto !== 'none' && String(m.proyecto || '') !== String(args.proyecto)) return false;
    if (args.moneda && m.moneda !== args.moneda) return false;
    if (args.cuentaId !== undefined && args.cuentaId !== '' && String(m.cuentaId) !== String(args.cuentaId)) return false;
    if (desde && m.fecha < desde) return false;
    if (hastaFin && m.fecha > hastaFin) return false;
    if (q && sinTildes_(m.descripcion + ' ' + m.categoria + ' ' + m.subcategoria).indexOf(q) === -1) return false;
    if (args.categoria !== undefined && args.categoria !== '') {
      const cm = String(m.categoria || '').trim() || 'Sin categoría';
      if (cm !== args.categoria) return false;
      if (args.subcategoria && String(m.subcategoria || '').trim() !== args.subcategoria) return false;
    }
    return true;
  });
  const tasa = obtenerUltimaTasaReferencial_();
  const orden = args.orden === 'monto' || args.orden === 'categoria' ? args.orden : 'fecha';
  const porFecha = function (a, b) { return (b.fecha - a.fecha) || (numero_(b.id) - numero_(a.id)); };
  let grupos = null;
  if (orden === 'monto') filtrados.sort(function (a, b) { return (convertirAPEN_(b.monto, b.moneda, tasa) - convertirAPEN_(a.monto, a.moneda, tasa)) || porFecha(a, b); });
  else if (orden === 'categoria') {
    const cat = function (m) { return String(m.categoria || '').trim() || 'Sin categoría'; };
    filtrados.sort(function (a, b) { const x = cat(a), y = cat(b); return x < y ? -1 : x > y ? 1 : porFecha(a, b); });
    const mapa = {};
    filtrados.forEach(function (m) {
      if (m.tipo === 'Transferencia' || m.tipo === 'Ajuste') return;
      const g = mapa[cat(m)] || (mapa[cat(m)] = { categoria: cat(m), n: 0, totales: { PEN: 0, USD: 0, EUR: 0 } });
      g.n++; g.totales[m.moneda] = redondear_((g.totales[m.moneda] || 0) + m.monto);
    });
    grupos = Object.keys(mapa).map(function (k) { return mapa[k]; });
  } else filtrados.sort(porFecha);
  const porPagina = Math.min(200, Math.max(10, Number(args.porPagina) || 40)), pagina = Math.max(1, Number(args.pagina) || 1);
  return {
    total: filtrados.length, pagina: pagina, hayMas: filtrados.length > pagina * porPagina, orden: orden, grupos: grupos,
    items: filtrados.slice((pagina - 1) * porPagina, pagina * porPagina).map(function (m) { return movimientoParaApp_(m, nombres); })
  };
}

/**
 * Detalle de una categoría: total del periodo (en soles equivalentes y por moneda), % del total de su tipo,
 * evolución de los últimos 6 meses y desglose por subcategoría. args = { categoria, tipo: 'Gasto'|'Ingreso', desde, hasta }
 */
function api_categoria(args) {
  const tipo = tipoCat_(args.tipo), cat = String(args.categoria || '').trim();
  const sinCat = cat === 'Sin categoría';
  const tasa = obtenerUltimaTasaReferencial_();
  const hoy = new Date();
  const desde = fecha_(args.desde) || new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const hastaD = fecha_(args.hasta) || hoy;
  const hasta = new Date(hastaD.getFullYear(), hastaD.getMonth(), hastaD.getDate(), 23, 59, 59, 999);
  const de = function (m) { return tipo === 'Ingreso' ? esIngreso_(m.tipo) : esGasto_(m.tipo); };
  const catDe = function (m) { return String(m.categoria || '').trim() || 'Sin categoría'; };
  const reales = movimientosReales_(leerMovimientos_()).filter(de);
  let totalTipo = 0, total = 0, n = 0;
  const porMoneda = { PEN: 0, USD: 0, EUR: 0 }, subs = {};
  reales.forEach(function (m) {
    if (m.efectiva < desde || m.efectiva > hasta) return;
    const pen = convertirAPEN_(m.monto, m.moneda, tasa);
    totalTipo += pen;
    if (catDe(m) !== (sinCat ? 'Sin categoría' : cat)) return;
    total += pen; n++;
    porMoneda[m.moneda] = redondear_((porMoneda[m.moneda] || 0) + m.monto);
    const sb = String(m.subcategoria || '').trim() || 'Sin subcategoría';
    const s = subs[sb] || (subs[sb] = { sub: sb, total: 0, n: 0 });
    s.total += pen; s.n++;
  });
  const serie = [];
  for (let k = 5; k >= 0; k--) {
    const ini = new Date(hoy.getFullYear(), hoy.getMonth() - k, 1), fin = new Date(hoy.getFullYear(), hoy.getMonth() - k + 1, 0, 23, 59, 59, 999);
    let t = 0;
    reales.forEach(function (m) { if (m.efectiva >= ini && m.efectiva <= fin && catDe(m) === (sinCat ? 'Sin categoría' : cat)) t += convertirAPEN_(m.monto, m.moneda, tasa); });
    serie.push({ mes: Utilities.formatDate(ini, zona_(), 'yyyy-MM'), total: redondear_(t) });
  }
  return {
    categoria: cat, tipo: tipo, desde: fechaISO_(desde), hasta: fechaISO_(hastaD), total: redondear_(total), n: n, porMoneda: porMoneda,
    porcentaje: totalTipo > 0 ? total / totalTipo * 100 : 0, serie: serie,
    subcategorias: Object.keys(subs).map(function (k) { return { sub: subs[k].sub, total: redondear_(subs[k].total), n: subs[k].n }; }).sort(function (a, b) { return b.total - a.total; })
  };
}

function api_detalleMovimiento(args) { return obtenerDetalleMovimiento(args.id); }
function api_editarMovimiento(args) {
  const r = editarMovimiento(args.id, args.cambios || {});
  if (args.cambios && args.cambios.proyecto !== undefined) asignarProyecto_([args.id], args.cambios.proyecto);
  return r;
}
function api_eliminarMovimiento(args) { eliminarMovimiento(args.id); return true; }

// ---------------------------------------------------------------
// ESCRITURA: registrar, traspasos, cambios, tarjeta
// ---------------------------------------------------------------

/**
 * Gasto o ingreso (o programado). args = { clase:'gasto'|'ingreso', tipo, monto, moneda, descripcion, medio ('Tarjeta'|'Efectivo'|'Debito'),
 *   cuentaId, fecha ('yyyy-mm-dd' o 'yyyy-mm-ddThh:mm'), categoria, subcategoria, programar (bool), alertas ([minutos]) }
 */
function api_registrar(args) {
  const gasto = args.clase === 'gasto';
  const tipo = args.tipo || (gasto ? 'Gasto Personal' : 'Ingreso Activo');
  if (gasto && !esGasto_(tipo)) throw new Error('Tipo de gasto inválido.');
  if (!gasto && !esIngreso_(tipo)) throw new Error('Tipo de ingreso inválido.');
  const resReg = registrarOProgramar_({
    tipo: tipo, monto: args.monto, moneda: args.moneda, descripcion: args.descripcion, medio: args.medio, cuentaId: args.cuentaId,
    fecha: args.fecha, categoria: args.categoria, subcategoria: args.subcategoria, programar: !!args.programar, alertas: args.alertas, notas: args.notas,
    cuotas: args.cuotas, interes: args.interes, tipoInteres: args.tipoInteres, tarjetaId: args.tarjetaId
  });
  if (args.proyecto !== undefined && args.proyecto !== '' && resReg && resReg.idMov !== '' && resReg.idMov !== undefined) asignarProyecto_([resReg.idMov], args.proyecto);
  return resReg;
}

function api_traspaso(args) { return transferirEntreCuentas(args); }
function api_cambioDivisa(args) { return registrarCambioDivisa(args); }
function api_pagarTarjeta(args) { return registrarPagoTarjeta(args.fechaCobro, args.cuentaId, args.fechaPago, args.tarjetaId); }

// ---------------------------------------------------------------
// CUENTAS
// ---------------------------------------------------------------

function api_crearCuenta(args) { return crearCuenta(args); }
function api_editarCuenta(args) { editarCuenta(args.id, args.cambios || {}); return true; }
function api_eliminarCuenta(args) { eliminarCuenta(args.id); return true; }
function api_ajustarSaldo(args) { return ajustarSaldoManual(args.id, args.saldo); }

// ---------------------------------------------------------------
// PROGRAMADOS
// ---------------------------------------------------------------

function api_programados(args) {
  const l = listarProgramados_(args.estado || '');
  if (!args.estado || args.estado === 'Pendiente') return l.concat(proximosRecurrentes_()).sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
  return l;
}
function api_editarProgramado(args) { editarProgramado(args.id, args.cambios || {}); return true; }
function api_cancelarProgramado(args) { cancelarProgramado(args.id); return true; }
function api_eliminarProgramado(args) { eliminarProgramado(args.id); return true; }
function api_realizarProgramado(args) { return realizarProgramado(args.id, { fecha: args.fecha, monto: args.monto, cuentaId: args.cuentaId }); }

// ---------------------------------------------------------------
// INVERSIONES, IA
// ---------------------------------------------------------------

function api_inversiones(args) { return resumenInversiones_(args.refrescar !== false); }
function api_historialAccion(args) { return historialAccion_(args.ticker, args.rango || '1M'); }
function api_operacionInversion(args) { return registrarOperacionInversion(args); }
function api_eliminarOperacion(args) { eliminarOperacionInversion(args.id); return true; }
function api_dividendosSugeridos() { return dividendosSugeridos_(); }
function api_recomendaciones(args) { return obtenerRecomendaciones_(!!args.forzar); }

function api_estadoSistema() {
  return {
    telegram: telegramConfigurado_(), webhookSecret: !!tgSecret_(), ia: iaProveedor_(), audio: !!prop_('GEMINI_API_KEY'), pin: !!prop_('DASHBOARD_PIN'),
    triggers: ScriptApp.getProjectTriggers().length, zona: zona_(), webhook: webhookResumen_(), modeloIA: iaProveedor_() === 'gemini' ? geminiModelo_() : ''
  };
}

function webhookResumen_() {
  try { const w = infoWebhookTelegram_(); return { conectado: !!w.url, pendientes: w.pendientes, ultimoError: w.ultimoError, ultimoErrorFecha: w.ultimoErrorFecha }; }
  catch (e) { return { conectado: false, error: e.message }; }
}

function api_repararTelegram() { return repararTelegram(); }

/** Envía el PDF a tu Telegram. tipo: 'mes_actual' | 'mes_anterior' | 'anual' */
function api_enviarReporte(args) {
  if (args.tipo === 'anual') generarYEnviarReporteAnual(new Date(new Date().getFullYear() + 1, 0, 2));
  else if (args.tipo === 'mes_anterior') generarYEnviarReporteMensual();
  else enviarReporteMesActualDePrueba();
  return true;
}


/* ==== 14_Migracion.gs ==== */
/**
 * ============================================================
 * MIGRACIÓN Y MANTENIMIENTO DE SALDOS
 * ============================================================
 * Dos usos:
 *  1) Importar historial antiguo (hoja "Importacion") para que reportes y categorías lo vean.
 *     Estas filas se guardan SIN cuenta (CuentaID vacío): no mueven ningún saldo. Tus saldos reales los
 *     fijas tú en la hoja Cuentas (o con ⚙️ Finanzas > Ajustar saldo de una cuenta).
 *  2) Auditar saldos: comparar el saldo guardado de cada cuenta contra lo que dice el historial.
 *
 * Columnas de "Importacion":
 *   Fecha | Tipo | Monto | Moneda | Descripcion | CuentaBanco (se ignora) | Categoria | Subcategoria | CentroCosto
 * Tipo: Ingreso Activo | Ingreso Pasivo | Gasto Personal | Gasto Fijo | Gasto Inversion
 */

function crearHojaImportacion() {
  const ss = getSS_();
  let hoja = ss.getSheetByName('Importacion');
  if (hoja && hoja.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    if (ui.alert('La hoja Importacion ya tiene datos. ¿Borrarlos y empezar de cero?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  if (!hoja) hoja = ss.insertSheet('Importacion');
  hoja.clear();
  const encabezados = ['Fecha', 'Tipo', 'Monto', 'Moneda', 'Descripcion', 'CuentaBanco', 'Categoria', 'Subcategoria', 'CentroCosto'];
  hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
  hoja.setFrozenRows(1);
  avisar_('✅ Hoja "Importacion" lista. Pega tus datos desde la fila 2 (columna A en adelante).');
}

function migrarDesdeImportacion() {
  return conLock_(function () {
    const hojaImport = getSS_().getSheetByName('Importacion');
    if (!hojaImport) throw new Error('No existe la hoja "Importacion". Corre crearHojaImportacion() primero.');

    const datos = hojaImport.getDataRange().getValues();
    const hojaMov = getHoja_('Movimientos');
    let siguienteId = generarID_('Movimientos');
    const filasNuevas = [];
    const rechazadas = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (fila.every(function (c) { return c === ''; })) continue;
      const fecha = fecha_(fila[0]);
      const tipo = String(fila[1]).trim();
      const monto = Math.abs(numero_(fila[2]));
      const moneda = String(fila[3]).trim().toUpperCase();
      const tiposValidos = ['Ingreso Activo', 'Ingreso Pasivo', 'Gasto Personal', 'Gasto Fijo', 'Gasto Inversion'];

      if (!fecha || tiposValidos.indexOf(tipo) === -1 || !(monto > 0) || CONFIG.MONEDAS.indexOf(moneda) === -1) {
        rechazadas.push('fila ' + (i + 1));
        continue;
      }
      const descripcion = String(fila[4] || '').trim();
      let categoria = fila[6], subcategoria = fila[7], centro = fila[8];
      if (!categoria) {
        const auto = autoCategorizar_(descripcion);
        categoria = auto.categoria;
        subcategoria = subcategoria || auto.subcategoria;
        centro = centro || auto.centroCosto;
      }
      filasNuevas.push([
        siguienteId++, fecha, fecha, tipo, monto, moneda, '',
        limpiarCentroCosto_(centro), categoria || '', subcategoria || '', descripcion,
        'Efectivo', 'No aplica', 'Manual', 'Migración histórica (sin cuenta)'
      ]);
    }

    if (filasNuevas.length > 0) {
      hojaMov.getRange(hojaMov.getLastRow() + 1, 1, filasNuevas.length, 15).setValues(filasNuevas);
    }
    registrarLog_('Migracion', filasNuevas.length + ' movimientos importados, ' + rechazadas.length + ' rechazados');

    avisar_('✅ Migración: ' + filasNuevas.length + ' movimientos importados sin tocar tus saldos.' +
      (rechazadas.length ? '\n\n⚠️ ' + rechazadas.length + ' filas rechazadas por datos incompletos o inválidos: ' + rechazadas.slice(0, 15).join(', ') + (rechazadas.length > 15 ? '…' : '') : '') +
      '\n\nNo la corras dos veces: duplicaría el historial.');
  });
}

/** Saldo que debería tener cada cuenta según SaldoInicial + historial. Devuelve { cuentaId: saldo } */
function calcularSaldosDesdeHistorial_() {
  const cuentas = leerHoja_('Cuentas');
  const saldos = {};
  for (let i = 1; i < cuentas.length; i++) {
    if (cuentas[i][0] !== '') saldos[String(cuentas[i][0])] = numero_(cuentas[i][4]);
  }
  leerMovimientos_().forEach(function (m) {
    const k = String(m.cuentaId);
    if (saldos[k] !== undefined) saldos[k] += efectoCaja_(m);
  });
  return saldos;
}

/**
 * Auditoría de saldos (menú). Compara SaldoActual contra SaldoInicial + movimientos y muestra las diferencias.
 * NO modifica nada salvo que confirmes; si confirmas, deja SaldoActual igual al historial.
 * Útil para detectar si editaste un movimiento a mano. Si escribiste tus saldos reales directo en SaldoActual
 * y no en SaldoInicial, verás diferencias: en ese caso responde NO y usa "Ajustar saldo de una cuenta".
 */
function auditarSaldos() {
  const esperados = calcularSaldosDesdeHistorial_();
  const cuentas = leerHoja_('Cuentas');
  const dif = [];
  for (let i = 1; i < cuentas.length; i++) {
    if (cuentas[i][0] === '') continue;
    const esperado = redondear_(esperados[String(cuentas[i][0])]);
    const actual = redondear_(numero_(cuentas[i][5]));
    if (Math.abs(esperado - actual) > 0.005) {
      dif.push({ fila: i + 1, texto: cuentas[i][2] + ' (' + cuentas[i][1] + '): guardado ' + actual + ' · historial ' + esperado, esperado: esperado });
    }
  }
  const ui = SpreadsheetApp.getUi();
  if (dif.length === 0) { ui.alert('✅ Todos los saldos coinciden con el historial.'); return; }
  const resp = ui.alert('Diferencias encontradas:\n\n' + dif.map(function (d) { return d.texto; }).join('\n') +
    '\n\n¿Reemplazar SaldoActual por el valor del historial? (NO recomendado si tus saldos reales los escribiste a mano)', ui.ButtonSet.YES_NO);
  if (resp === ui.Button.YES) {
    conLock_(function () {
      dif.forEach(function (d) { getHoja_('Cuentas').getRange(d.fila, 6).setValue(d.esperado); });
      registrarLog_('Auditoria', 'Saldos reemplazados por el historial en ' + dif.length + ' cuentas');
    });
  }
}

/** Menú: pide cuenta y saldo real, y deja un movimiento Ajuste */
function ajustarSaldoDesdeMenu() {
  const ui = SpreadsheetApp.getUi();
  const cuentas = leerHoja_('Cuentas');
  const lista = [];
  for (let i = 1; i < cuentas.length; i++) if (cuentas[i][0] !== '') lista.push(cuentas[i][0] + ' = ' + cuentas[i][2] + ' ' + cuentas[i][1] + ' (saldo ' + numero_(cuentas[i][5]) + ')');
  const r1 = ui.prompt('Ajustar saldo', 'ID de la cuenta:\n' + lista.join('\n'), ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Ajustar saldo', 'Saldo REAL de hoy (como lo ves en tu banco/app):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const dif = ajustarSaldoManual(r1.getResponseText().trim(), r2.getResponseText());
  ui.alert(dif === 0 ? 'El saldo ya coincidía.' : '✅ Ajustado (diferencia ' + dif + '). Quedó un movimiento tipo Ajuste.');
}


/* ==== 15_Cuentas.gs ==== */
/**
 * ============================================================
 * CUENTAS — crear, editar, archivar, mover dinero entre cuentas y valores predeterminados
 * ============================================================
 * Tipos: Banco, Efectivo, Ahorros, Inversion, Billetera Digital. Solo 'Efectivo' cuenta como efectivo:
 * un gasto "en efectivo" solo puede salir de una cuenta Efectivo, y uno "con débito" de las demás.
 */

function crearCuenta(d) {
  return conLock_(function () {
    const moneda = validarMoneda_(d.moneda);
    const nombre = String(d.banco || d.nombre || '').trim();
    if (!nombre) throw new Error('Ponle un nombre a la cuenta.');
    const tipo = TIPOS_CUENTA_.indexOf(d.tipo) !== -1 ? d.tipo : 'Banco';
    const saldo = redondear_(numero_(d.saldoInicial));

    const cuentas = leerHoja_('Cuentas');
    for (let i = 1; i < cuentas.length; i++) {
      if (String(cuentas[i][1]).trim() === moneda && sinTildes_(cuentas[i][2]).trim() === sinTildes_(nombre).trim()) {
        throw new Error('Ya existe una cuenta "' + nombre + '" en ' + moneda + '.');
      }
    }
    const id = generarID_('Cuentas');
    getHoja_('Cuentas').appendRow([id, moneda, nombre, tipo, saldo, saldo, 'SI', String(d.color || '')]);
    registrarLog_('Cuenta', 'Creada #' + id + ' ' + nombre + ' (' + moneda + ', ' + tipo + ') saldo inicial ' + saldo);
    return id;
  });
}

/** Edita nombre, tipo, color o estado. La moneda y el tipo solo cambian si la cuenta aún no tiene movimientos */
function editarCuenta(id, cambios) {
  return conLock_(function () {
    const c = obtenerCuenta_(id);
    if (!c) throw new Error('Cuenta no encontrada.');
    const hoja = getHoja_('Cuentas');
    const tieneMovs = leerMovimientos_().some(function (m) { return String(m.cuentaId) === String(id); });

    if (cambios.banco !== undefined && String(cambios.banco).trim()) hoja.getRange(c.numeroFila, 3).setValue(String(cambios.banco).trim());
    if (cambios.tipo !== undefined && cambios.tipo !== c.tipo) {
      if (TIPOS_CUENTA_.indexOf(cambios.tipo) === -1) throw new Error('Tipo de cuenta inválido.');
      const cambiaEfectivo = (cambios.tipo === 'Efectivo') !== (c.tipo === 'Efectivo');
      if (tieneMovs && cambiaEfectivo) throw new Error('No puedes pasar una cuenta con movimientos de/para Efectivo. Crea una cuenta nueva y muévele el dinero.');
      hoja.getRange(c.numeroFila, 4).setValue(cambios.tipo);
    }
    if (cambios.moneda !== undefined && cambios.moneda !== c.moneda) {
      if (tieneMovs) throw new Error('No puedes cambiar la moneda de una cuenta con movimientos.');
      hoja.getRange(c.numeroFila, 2).setValue(validarMoneda_(cambios.moneda));
    }
    if (cambios.color !== undefined) hoja.getRange(c.numeroFila, 8).setValue(String(cambios.color));
    if (cambios.activa !== undefined) {
      if (!cambios.activa && Math.abs(c.saldo) > 0.005) throw new Error('Para archivarla primero mueve su saldo (' + c.saldo + ') a otra cuenta.');
      hoja.getRange(c.numeroFila, 7).setValue(cambios.activa ? 'SI' : 'NO');
    }
    registrarLog_('Cuenta', 'Editada #' + id);
  });
}

/** Elimina una cuenta SOLO si nunca tuvo movimientos; si ya tiene historial, se archiva */
function eliminarCuenta(id) {
  return conLock_(function () {
    const c = obtenerCuenta_(id);
    if (!c) throw new Error('Cuenta no encontrada.');
    if (leerMovimientos_().some(function (m) { return String(m.cuentaId) === String(id); })) {
      throw new Error('Esta cuenta tiene movimientos. Archívala en vez de eliminarla (así no se pierde tu historial).');
    }
    getHoja_('Cuentas').deleteRow(c.numeroFila);
    registrarLog_('Cuenta', 'Eliminada #' + id + ' ' + c.banco);
  });
}

/**
 * Mueve dinero entre dos de tus cuentas (ej. efectivo → BCP porque fuiste a depositar). No es ingreso
 * ni gasto. Si las monedas son distintas se registra como cambio de divisa (necesita montoDestino).
 * d = { cuentaOrigenId, cuentaDestinoId, monto, montoDestino (solo si cambian de moneda), fecha, nota }
 */
function transferirEntreCuentas(d) {
  return conLock_(function () {
    if (String(d.cuentaOrigenId) === String(d.cuentaDestinoId)) throw new Error('Elige dos cuentas distintas.');
    const o = validarCuentaMoneda_(d.cuentaOrigenId, null);
    const t = validarCuentaMoneda_(d.cuentaDestinoId, null);
    const monto = validarMonto_(d.monto);

    if (o.moneda !== t.moneda) {
      return registrarCambioDivisa({
        fecha: d.fecha, monedaOrigen: o.moneda, montoOrigen: monto, cuentaOrigenId: o.id,
        monedaDestino: t.moneda, montoDestino: d.montoDestino, cuentaDestinoId: t.id
      });
    }
    if (o.saldo < monto - 0.005 && o.tipo !== 'Banco') {
      // En bancos se permite sobregiro "lógico" (el saldo puede estar desfasado); en efectivo avisamos
      throw new Error('No tienes ' + monto + ' en ' + o.banco + ' (saldo ' + o.saldo + ').');
    }
    const token = generarID_('Movimientos');
    const desc = String(d.nota || '').trim() || ('Traspaso: ' + o.banco + ' → ' + t.banco);
    const base = { fecha: d.fecha, moneda: o.moneda, descripcion: desc, categoria: 'Transferencia entre cuentas', notas: 'Traspaso #' + token };
    registrarTransferencia_(Object.assign({}, base, { monto: -monto, cuentaId: o.id, subcategoria: 'Sale de ' + o.banco }));
    registrarTransferencia_(Object.assign({}, base, { monto: monto, cuentaId: t.id, subcategoria: 'Entra a ' + t.banco }));
    registrarLog_('Traspaso', monto + ' ' + o.moneda + ' de ' + o.banco + ' a ' + t.banco);
    return token;
  });
}

// ---------------------------------------------------------------
// VALORES PREDETERMINADOS (por moneda): medio de pago y cuenta
// ---------------------------------------------------------------

/** Medio predeterminado para un GASTO en esa moneda: 'Tarjeta' (por defecto) | 'Efectivo' | 'Debito' */
function medioGastoPorDefecto_(moneda) {
  const m = String(getAjuste_('medio.gasto.' + moneda, 'Tarjeta'));
  return ['Tarjeta', 'Efectivo', 'Debito'].indexOf(m) !== -1 ? m : 'Tarjeta';
}

/** Medio predeterminado para un INGRESO: 'Efectivo' | 'Debito' */
function medioIngresoPorDefecto_(moneda) {
  return getAjuste_('medio.ingreso.' + moneda, 'Debito') === 'Efectivo' ? 'Efectivo' : 'Debito';
}

/** Cuenta predeterminada para ese medio ('Efectivo'|'Debito') y moneda; si no hay configurada, la primera que sirva */
function cuentaPorDefecto_(moneda, medio) {
  const cuentas = leerHoja_('Cuentas');
  const configurada = getAjuste_('cuenta.' + moneda + '.' + medio, '');
  let primera = null;
  for (let i = 1; i < cuentas.length; i++) {
    if (cuentas[i][0] === '' || String(cuentas[i][1]).trim() !== moneda) continue;
    if (String(cuentas[i][6]).trim().toUpperCase() === 'NO') continue;
    const esEfectivo = String(cuentas[i][3]).trim() === 'Efectivo';
    if ((medio === 'Efectivo') !== esEfectivo) continue;
    if (configurada !== '' && String(cuentas[i][0]) === String(configurada)) return cuentas[i][0];
    if (primera === null) primera = cuentas[i][0];
  }
  if (primera === null) throw new Error('No tienes una cuenta ' + (medio === 'Efectivo' ? 'de efectivo' : 'bancaria') + ' activa en ' + moneda + '. Crea una en la pestaña Cuentas.');
  return primera;
}


/* ==== 16_Programados.gs ==== */
/**
 * ============================================================
 * MOVIMIENTOS PROGRAMADOS (ingresos o gastos que aún no ocurren)
 * ============================================================
 * Viven en la hoja Programados y NO tocan saldos, metas ni "movimientos recientes" hasta que los
 * marcas como realizados. Al realizarlos se crea el movimiento real con la fecha en que realmente
 * ocurrió (puede ser distinta a la programada: el detalle te dice si fue a tiempo o con cuántos días de atraso).
 * Alertas: lista de "minutos antes" (0 = en el momento, 60 = 1 hora antes, 1440 = 1 día antes...).
 * Un trigger revisa cada 15 minutos y te avisa por Telegram con botones para marcar "Ya se realizó".
 */

var P = { ID: 0, TIPO: 1, MONTO: 2, MONEDA: 3, MEDIO: 4, CUENTA: 5, DESC: 6, CAT: 7, SUB: 8, FECHA: 9, ALERTAS: 10, ENVIADAS: 11, ESTADO: 12, REALIZADA: 13, MOV: 14, AVISO: 15, NOTAS: 16 };

function fechaHoraISO_(d) {
  const f = fecha_(d);
  return f ? Utilities.formatDate(f, zona_(), "yyyy-MM-dd'T'HH:mm") : '';
}

function parsearAlertas_(v) {
  let lista = Array.isArray(v) ? v : String(v === undefined || v === null ? '' : v).split(',');
  const out = [];
  lista.forEach(function (x) {
    if (x === '' || x === null) return;
    const n = Math.round(Number(x));
    if (isFinite(n) && n >= 0 && n <= 60 * 24 * 60 && out.indexOf(n) === -1) out.push(n);
  });
  return out.sort(function (a, b) { return b - a; });
}

function textoAlerta_(min) {
  if (min === 0) return 'En el momento';
  if (min < 60) return min + ' min antes';
  if (min < 1440) return (min / 60) + (min === 60 ? ' hora antes' : ' horas antes');
  const d = min / 1440;
  return d + (d === 1 ? ' día antes' : ' días antes');
}

function serializarProgramado_(f, nombresCuentas) {
  const prog = fecha_(f[P.FECHA]);
  const real = fecha_(f[P.REALIZADA]);
  const estado = String(f[P.ESTADO]).trim() || 'Pendiente';
  let dias = null;
  if (estado === 'Realizado' && prog && real) {
    dias = Math.round((new Date(real.getFullYear(), real.getMonth(), real.getDate()) - new Date(prog.getFullYear(), prog.getMonth(), prog.getDate())) / 86400000);
  }
  const ahora = new Date();
  return {
    id: f[P.ID], tipo: String(f[P.TIPO]).trim(), monto: numero_(f[P.MONTO]), moneda: String(f[P.MONEDA]).trim(),
    medio: String(f[P.MEDIO]).trim(), cuentaId: f[P.CUENTA], cuentaNombre: f[P.CUENTA] !== '' && nombresCuentas ? (nombresCuentas[String(f[P.CUENTA])] || '') : '',
    descripcion: String(f[P.DESC] || ''), categoria: String(f[P.CAT] || ''), subcategoria: String(f[P.SUB] || ''),
    fecha: fechaHoraISO_(prog), alertas: parsearAlertas_(f[P.ALERTAS]), estado: estado,
    fechaRealizada: fechaHoraISO_(real), movimientoId: f[P.MOV], notas: String(f[P.NOTAS] || ''),
    diasDiferencia: dias,
    atrasado: estado === 'Pendiente' && prog ? prog < ahora : false,
    diasAtraso: estado === 'Pendiente' && prog && prog < ahora ? Math.floor((ahora - prog) / 86400000) : 0
  };
}

function mapaNombresCuentas_() {
  const m = {};
  const cu = leerHoja_('Cuentas');
  for (let i = 1; i < cu.length; i++) if (cu[i][0] !== '') m[String(cu[i][0])] = cu[i][2];
  return m;
}

function listarProgramados_(estado) {
  const datos = leerHoja_('Programados');
  const nombres = mapaNombresCuentas_();
  const lista = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][P.ID] === '') continue;
    const s = serializarProgramado_(datos[i], nombres);
    if (!estado || s.estado === estado) lista.push(s);
  }
  return lista.sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
}

/** Valida y normaliza los campos comunes de un programado */
function normalizarProgramado_(d) {
  const tipo = String(d.tipo || '').trim();
  if ((!esIngreso_(tipo) && !esGasto_(tipo)) || TIPOS_MOVIMIENTO.indexOf(tipo) === -1) throw new Error('Tipo inválido: "' + d.tipo + '".');
  const monto = validarMonto_(d.monto);
  const moneda = validarMoneda_(d.moneda);
  const fecha = fecha_(d.fecha);
  if (!fecha) throw new Error('Indica la fecha del movimiento programado.');
  let medio = String(d.medio || d.medioPago || '').trim();
  if (esGasto_(tipo)) { if (['Tarjeta', 'Efectivo', 'Debito'].indexOf(medio) === -1) medio = medioGastoPorDefecto_(moneda); }
  else medio = medio === 'Efectivo' ? 'Efectivo' : 'Debito';

  let cuentaId = '';
  if (medio !== 'Tarjeta') {
    cuentaId = d.cuentaId !== undefined && d.cuentaId !== '' ? d.cuentaId : cuentaPorDefecto_(moneda, medio);
    const c = validarCuentaMoneda_(cuentaId, moneda);
    medio = c.tipo === 'Efectivo' ? 'Efectivo' : 'Debito';
  }
  let alertas;
  if (d.alertas === undefined) alertas = parsearAlertas_(getAjuste_('alertas.default', '1440'));
  else alertas = parsearAlertas_(d.alertas);

  return { tipo: tipo, monto: monto, moneda: moneda, fecha: fecha, medio: medio, cuentaId: cuentaId, alertas: alertas };
}

function crearProgramado(d) {
  return conLock_(function () {
    const n = normalizarProgramado_(d);
    const cat = completarCategoria_({ categoria: d.categoria, subcategoria: d.subcategoria, centroCosto: '', descripcion: d.descripcion });
    const id = generarID_('Programados');
    getHoja_('Programados').appendRow([
      id, tipoIngresoPorCategoria_(n.tipo, cat.categoria), n.monto, n.moneda, n.medio, n.cuentaId, String(d.descripcion || '').trim(), cat.categoria, cat.subcategoria,
      n.fecha, n.alertas.join(','), '', 'Pendiente', '', '', '', String(d.notas || '')
    ]);
    registrarLog_('Programado', 'Creado #' + id + ' ' + n.tipo + ' ' + n.monto + ' ' + n.moneda + ' para ' + fechaHoraISO_(n.fecha));
    return id;
  });
}

function editarProgramado(id, cambios) {
  return conLock_(function () {
    const p = buscarFilaPorId_('Programados', id);
    if (!p) throw new Error('Programado no encontrado.');
    if (String(p.fila[P.ESTADO]).trim() !== 'Pendiente') throw new Error('Solo puedes editar movimientos programados que siguen pendientes.');
    const f = p.fila;
    const base = {
      tipo: f[P.TIPO], monto: f[P.MONTO], moneda: f[P.MONEDA], medio: f[P.MEDIO], cuentaId: f[P.CUENTA],
      fecha: f[P.FECHA], alertas: parsearAlertas_(f[P.ALERTAS]), descripcion: f[P.DESC], categoria: f[P.CAT], subcategoria: f[P.SUB], notas: f[P.NOTAS]
    };
    const d = Object.assign({}, base, cambios);
    if (cambios.moneda && cambios.cuentaId === undefined && cambios.medio === undefined) d.cuentaId = ''; // cambió de moneda: cuenta por defecto
    const n = normalizarProgramado_(d);
    const reiniciarAvisos = cambios.fecha !== undefined || cambios.alertas !== undefined;
    getHoja_('Programados').getRange(p.numeroFila, 1, 1, 17).setValues([[
      id, n.tipo, n.monto, n.moneda, n.medio, n.cuentaId, String(d.descripcion || '').trim(), d.categoria || '', d.subcategoria || '',
      n.fecha, n.alertas.join(','), reiniciarAvisos ? '' : f[P.ENVIADAS], 'Pendiente', '', '', reiniciarAvisos ? '' : f[P.AVISO], String(d.notas || '')
    ]]);
  });
}

function cambiarEstadoProgramado_(id, estado) {
  return conLock_(function () {
    const p = buscarFilaPorId_('Programados', id);
    if (!p) throw new Error('Programado no encontrado.');
    if (String(p.fila[P.ESTADO]).trim() !== 'Pendiente') throw new Error('Ese movimiento ya no está pendiente.');
    getHoja_('Programados').getRange(p.numeroFila, P.ESTADO + 1).setValue(estado);
  });
}
function cancelarProgramado(id) { cambiarEstadoProgramado_(id, 'Cancelado'); }

function eliminarProgramado(id) {
  return conLock_(function () {
    const p = buscarFilaPorId_('Programados', id);
    if (!p) throw new Error('Programado no encontrado.');
    if (String(p.fila[P.ESTADO]).trim() === 'Realizado') throw new Error('Ya se realizó; elimina el movimiento real si quieres borrarlo.');
    getHoja_('Programados').deleteRow(p.numeroFila);
  });
}

/**
 * Marca como realizado: crea el movimiento real. opts = { fecha (cuándo ocurrió; por defecto ahora),
 * monto (si el monto real fue distinto), cuentaId }
 */
function realizarProgramado(id, opts) {
  opts = opts || {};
  return conLock_(function () {
    const p = buscarFilaPorId_('Programados', id);
    if (!p) throw new Error('Programado no encontrado.');
    const f = p.fila;
    if (String(f[P.ESTADO]).trim() !== 'Pendiente') throw new Error('Ese movimiento ya no está pendiente.');

    const fechaReal = opts.fecha ? fechaConHora_(opts.fecha) : new Date();
    if (esFuturo_(fechaReal)) throw new Error('No puedes marcar como realizado con una fecha futura.');
    const monto = opts.monto !== undefined && opts.monto !== '' ? validarMonto_(opts.monto) : numero_(f[P.MONTO]);
    const tipo = String(f[P.TIPO]).trim(), moneda = String(f[P.MONEDA]).trim(), medio = String(f[P.MEDIO]).trim();
    const nota = 'Programado #' + id;
    let movId;

    if (esGasto_(tipo) && medio === 'Tarjeta') {
      movId = registrarConsumoTarjetaDetallado_({
        fecha: fechaReal, tipo: tipo, monto: monto, moneda: moneda, descripcion: f[P.DESC],
        categoria: f[P.CAT], subcategoria: f[P.SUB], notas: nota
      }).idMov;
    } else {
      movId = registrarMovimiento({
        fecha: fechaReal, tipo: tipo, monto: monto, moneda: moneda, cuentaId: opts.cuentaId || f[P.CUENTA],
        descripcion: f[P.DESC], categoria: f[P.CAT], subcategoria: f[P.SUB], notas: nota
      });
    }
    const hoja = getHoja_('Programados');
    hoja.getRange(p.numeroFila, P.ESTADO + 1, 1, 3).setValues([['Realizado', fechaReal, movId]]);
    registrarLog_('Programado', '#' + id + ' realizado como movimiento #' + movId);
    return movId;
  });
}

/** Resumen de lo programado pendiente: próximos N días + atrasados, y neto proyectado por moneda */
function resumenProgramados_(dias) {
  const pendientes = listarProgramados_('Pendiente');
  const limite = new Date(Date.now() + dias * 86400000);
  const neto = { PEN: 0, USD: 0, EUR: 0 };
  const lista = pendientes.concat(proximosRecurrentes_()).filter(function (p) { return p.atrasado || fecha_(p.fecha) <= limite; })
    .sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
  lista.forEach(function (p) {
    if (neto[p.moneda] !== undefined) neto[p.moneda] += esIngreso_(p.tipo) ? p.monto : -p.monto;
  });
  CONFIG.MONEDAS.forEach(function (k) { neto[k] = redondear_(neto[k]); });
  return { lista: lista, netoProyectado: neto, totalPendientes: pendientes.length };
}

// ---------------------------------------------------------------
// ALERTAS (trigger cada 15 minutos)
// ---------------------------------------------------------------

function textoRelativo_(fecha) {
  const min = Math.round((fecha - new Date()) / 60000);
  const abs = Math.abs(min);
  let t;
  if (abs < 60) t = abs + ' min';
  else if (abs < 1440) t = Math.round(abs / 60) + ' h';
  else t = Math.round(abs / 1440) + (Math.round(abs / 1440) === 1 ? ' día' : ' días');
  return min >= 0 ? 'en ' + t : 'hace ' + t;
}

function revisarAlertasProgramados() {
  ejecutarSeguro_('revisarAlertasProgramados', function () {
    if (!getSS_().getSheetByName('Programados')) return; // aún no se creó la estructura: nada que revisar
    const datos = leerHoja_('Programados');
    const ahora = new Date();
    const hoja = getHoja_('Programados');
    const nombres = mapaNombresCuentas_();

    for (let i = 1; i < datos.length; i++) {
      const f = datos[i];
      if (f[P.ID] === '' || String(f[P.ESTADO]).trim() !== 'Pendiente') continue;
      const prog = fecha_(f[P.FECHA]);
      if (!prog) continue;
      const alertas = parsearAlertas_(f[P.ALERTAS]);
      const enviadas = parsearAlertas_(f[P.ENVIADAS]);
      const nuevas = [];

      alertas.forEach(function (min) {
        if (enviadas.indexOf(min) !== -1) return;
        const momento = prog.getTime() - min * 60000;
        if (ahora.getTime() >= momento && ahora.getTime() - momento < 3 * 86400000) nuevas.push(min);
      });
      if (nuevas.length) {
        // Marca primero para no repetir aunque el envío tarde; un solo mensaje aunque se junten varias alertas
        hoja.getRange(i + 1, P.ENVIADAS + 1).setValue(enviadas.concat(nuevas).join(','));
        enviarProgramadoTelegram_(f, '⏰ *Recordatorio* — ' + textoRelativo_(prog), nombres);
        continue;
      }
      // Atrasado: aviso una vez por día, a partir de las 9am y 12 h después de la hora programada
      const aviso = fecha_(f[P.AVISO]);
      const hoyAvisado = aviso && mismaFecha_(aviso, ahora);
      if (ahora.getTime() - prog.getTime() > 12 * 3600000 && ahora.getHours() >= 9 && !hoyAvisado && alertas.length > 0) {
        hoja.getRange(i + 1, P.AVISO + 1).setValue(ahora);
        enviarProgramadoTelegram_(f, '⚠️ *Programado atrasado* — ' + textoRelativo_(prog), nombres);
      }
    }
  });
}

function enviarProgramadoTelegram_(f, encabezado, nombres) {
  const tipo = String(f[P.TIPO]).trim();
  const moneda = String(f[P.MONEDA]).trim();
  const medio = String(f[P.MEDIO]).trim();
  const donde = medio === 'Tarjeta' ? 'tarjeta de crédito' : (nombres[String(f[P.CUENTA])] || 'cuenta');
  const texto = encabezado + '\n' + (esIngreso_(tipo) ? '💰 Ingreso' : '💸 Gasto') + ' programado: *' + escaparMd_(f[P.DESC] || '(sin descripción)') + '*\n' +
    formatoMonto_(f[P.MONTO], moneda) + ' · ' + escaparMd_(donde) + '\n📅 ' + fechaHoraISO_(f[P.FECHA]).replace('T', ' ');
  enviarMensajeTelegram(texto, [[
    { text: '✅ Ya se realizó', data: 'prog_ok:' + f[P.ID] },
    { text: '📅 Fue ayer', data: 'prog_ok:' + f[P.ID] + ':1' },
    { text: '🗑 Cancelar', data: 'prog_no:' + f[P.ID] }
  ]]);
}


// ---------------------------------------------------------------
// PUNTUALIDAD: ¿se pagó a tiempo, tarde o antes?
// ---------------------------------------------------------------

function textoDiasPuntualidad_(dias) {
  if (dias === null || dias === undefined) return '';
  if (dias === 0) return 'A tiempo (el día programado)';
  const a = Math.abs(dias);
  return dias > 0 ? 'Con ' + a + (a === 1 ? ' día' : ' días') + ' de atraso' : 'Adelantado ' + a + (a === 1 ? ' día' : ' días');
}

function textoPuntualidadProgramado_(id) {
  const p = buscarFilaPorId_('Programados', id);
  if (!p) return '';
  return textoDiasPuntualidad_(serializarProgramado_(p.fila, null).diasDiferencia);
}

/** Historial de puntualidad: por concepto (recurrente o descripción) con promedio de días de atraso (+) o adelanto (−) */
function api_puntualidad() {
  const datos = leerHoja_('Programados');
  const grupos = {}, recientes = [];
  let n = 0, suma = 0, aTiempo = 0;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][P.ID] === '') continue;
    const s = serializarProgramado_(datos[i], null);
    if (s.estado !== 'Realizado' || s.diasDiferencia === null) continue;
    const clave = sinTildes_(s.descripcion).trim().toLowerCase() + '|' + s.tipo.split(' ')[0];
    const g = grupos[clave] || (grupos[clave] = { descripcion: s.descripcion, tipo: /^Ingreso/.test(s.tipo) ? 'Ingreso' : 'Gasto', n: 0, suma: 0, tarde: 0, aTiempo: 0, temprano: 0, ultimo: '', ultimoDias: 0 });
    g.n++; g.suma += s.diasDiferencia;
    if (s.diasDiferencia > 0) g.tarde++; else if (s.diasDiferencia < 0) g.temprano++; else g.aTiempo++;
    if (s.fechaRealizada > g.ultimo) { g.ultimo = s.fechaRealizada; g.ultimoDias = s.diasDiferencia; }
    n++; suma += s.diasDiferencia; if (s.diasDiferencia <= 0) aTiempo++;
    recientes.push({ id: s.id, descripcion: s.descripcion, tipo: g.tipo, moneda: s.moneda, monto: s.monto, programada: s.fecha, realizada: s.fechaRealizada, dias: s.diasDiferencia });
  }
  recientes.sort(function (a, b) { return a.realizada < b.realizada ? 1 : -1; });
  const items = Object.keys(grupos).map(function (k) { const g = grupos[k]; g.promedio = Math.round(g.suma / g.n * 10) / 10; return g; })
    .sort(function (a, b) { return b.promedio - a.promedio; });
  return {
    n: n, promedio: n ? Math.round(suma / n * 10) / 10 : 0, pctATiempo: n ? Math.round(aTiempo / n * 100) : 0,
    items: items, recientes: recientes.slice(0, 20),
    pendientesAtrasados: listarProgramados_('Pendiente').filter(function (x) { return x.atrasado; }).length
  };
}


/* ==== 17_EdicionMovimientos.gs ==== */
/**
 * ============================================================
 * DETALLE, EDICIÓN Y ELIMINACIÓN DE MOVIMIENTOS
 * ============================================================
 * Cada edición revierte el efecto anterior en la cuenta y aplica el nuevo, así los saldos siempre cuadran.
 * Qué se puede hacer según el tipo de movimiento:
 *  - Ingreso / Gasto normal ......... todo (monto, moneda, cuenta, fecha, categoría...)
 *  - Gasto con tarjeta .............. todo mientras su estado de cuenta no esté pagado (después, solo textos)
 *  - Transferencia / Ajuste ......... solo descripción, categoría y notas; se pueden eliminar
 *  - Ligados a préstamo o deuda ..... solo textos (borrarlos descuadraría el cronograma)
 */

function extraerRef_(notas, patron) {
  const m = String(notas || '').match(patron);
  return m ? m[1] : null;
}

function clasificarMovimiento_(m) {
  if (m.medio === 'Tarjeta de Credito' && esGasto_(m.tipo)) return 'tarjeta';
  if (m.tipo === 'Transferencia' || m.tipo === 'Ajuste') {
    if (/Ref\. (PrestamoOtorgado|Deuda|CronogramaCuotas) #/.test(m.notas)) return 'ligado';
    return 'transferencia';
  }
  if (/Ref\. CronogramaCuotas #/.test(m.notas)) return 'ligado';
  return 'normal';
}

function obtenerDetalleMovimiento(id) {
  const movs = leerMovimientos_();
  const m = movs.filter(function (x) { return String(x.id) === String(id); })[0];
  if (!m) throw new Error('Movimiento no encontrado.');
  const clase = clasificarMovimiento_(m);
  const cuenta = m.cuentaId !== '' ? obtenerCuenta_(m.cuentaId) : null;

  const det = {
    id: m.id, fecha: fechaHoraISO_(m.fecha), efectiva: fechaISO_(m.efectiva), tipo: m.tipo, monto: m.monto, moneda: m.moneda,
    cuentaId: m.cuentaId, cuentaNombre: cuenta ? cuenta.banco : (m.medio === 'Tarjeta de Credito' ? 'Tarjeta ' + nombreTarjetaDeMov_(m) : ''),
    centro: String(m.centro || ''), categoria: String(m.categoria || ''), subcategoria: String(m.subcategoria || ''),
    proyecto: String(m.proyecto || ''), descripcion: m.descripcion, medio: m.medio, sunat: String(m.sunat || ''), origen: String(m.origen || ''), notas: m.notas,
    clase: clase, programado: null, tarjeta: null, editable: 'completo', eliminable: true, aviso: ''
  };

  // ¿Vino de un movimiento programado? (y si se hizo a tiempo o con atraso)
  const prog = listarProgramados_('Realizado').filter(function (p) { return String(p.movimientoId) === String(m.id); })[0];
  if (prog) {
    let texto;
    if (prog.diasDiferencia === 0) texto = 'Se realizó en la fecha programada.';
    else if (prog.diasDiferencia > 0) texto = 'Se realizó ' + prog.diasDiferencia + (prog.diasDiferencia === 1 ? ' día' : ' días') + ' DESPUÉS de lo programado (atraso).';
    else texto = 'Se adelantó ' + Math.abs(prog.diasDiferencia) + (prog.diasDiferencia === -1 ? ' día' : ' días') + ' respecto a lo programado.';
    det.programado = { id: prog.id, fechaProgramada: prog.fecha, fechaRealizada: prog.fechaRealizada, diasDiferencia: prog.diasDiferencia, texto: texto };
  }

  if (clase === 'tarjeta') {
    const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
    const t = tid ? buscarFilaPorId_('TarjetaCredito', tid) : null;
    if (t) {
      det.tarjeta = { id: tid, estado: String(t.fila[7]).trim(), cierre: fechaISO_(t.fila[5]), cobro: fechaISO_(t.fila[6]), desgravamen: String(t.fila[8]) === 'SI' };
      const gc = String(m.notas).match(/Cuotas #(\w+) \((\d+)\/(\d+)\)/);
      if (gc) { det.cuotas = { grupo: gc[1], k: +gc[2], n: +gc[3] }; det.editable = 'texto'; det.aviso = 'Compra en cuotas (' + gc[2] + ' de ' + gc[3] + '). Solo puedes editar textos; si la borras se borran todas las cuotas que aún no se pagaron.'; }
      if (det.tarjeta.estado === 'Pagado') { det.editable = 'texto'; det.aviso = 'Este consumo ya se pagó con un estado de cuenta: solo puedes editar textos.'; det.eliminable = false; }
    }
  } else if (clase === 'transferencia') {
    det.editable = 'texto';
    det.aviso = 'Las transferencias solo permiten editar textos. Si te equivocaste de monto, elimínala y créala de nuevo.';
  } else if (clase === 'ligado') {
    det.editable = 'texto'; det.eliminable = false;
    det.aviso = 'Está ligado a un préstamo o deuda: solo puedes editar textos.';
  }
  try {
    const dd = detalleDivision_(m);
    if (dd) { det.division = dd; det.aviso = (det.aviso ? det.aviso + ' ' : '') + 'Gasto dividido: para cambiar el monto, la moneda o la cuenta, primero deshaz la división.'; }
  } catch (e) { /* sin datos de división */ }
  return det;
}

function validarCamposTexto_(c, fila) {
  // devuelve las columnas de texto ya resueltas (para no perder las que no se editaron)
  return {
    descripcion: c.descripcion !== undefined ? String(c.descripcion).trim() : fila.descripcion,
    categoria: c.categoria !== undefined ? String(c.categoria).trim() : String(fila.categoria || ''),
    subcategoria: c.subcategoria !== undefined ? String(c.subcategoria).trim() : String(fila.subcategoria || ''),
    centro: c.centro !== undefined ? limpiarCentroCosto_(c.centro) : String(fila.centro || ''),
    sunat: c.sunat !== undefined ? limpiarCategoriaSunat_(c.sunat) : (fila.sunat || 'No aplica'),
    notas: c.notas !== undefined ? String(c.notas) : fila.notas
  };
}

function editarMovimiento(id, c) {
  return conLock_(function () {
    const movs = leerMovimientos_();
    const m = movs.filter(function (x) { return String(x.id) === String(id); })[0];
    if (!m) throw new Error('Movimiento no encontrado.');
    const clase = clasificarMovimiento_(m);
    const hoja = getHoja_('Movimientos');
    const t = validarCamposTexto_(c, m);
    // La marca de "gasto dividido" vive en las notas: si editas las notas no se puede perder
    const marcaDiv = String(m.notas).match(/Dividido #\w+ \(total [\d.]+\)/);
    if (marcaDiv && !/Dividido #\w+ \(total [\d.]+\)/.test(t.notas)) t.notas = (t.notas ? t.notas + ' · ' : '') + marcaDiv[0];

    const soloTexto = function () {
      hoja.getRange(m.fila, M.CENTRO + 1, 1, 8).setValues([[t.centro, t.categoria, t.subcategoria, t.descripcion, m.medio, t.sunat, m.origen, t.notas]]);
    };

    // --- Transferencias, ajustes y ligados a préstamo/deuda: solo textos ---
    if (clase === 'transferencia' || clase === 'ligado') { soloTexto(); return obtenerDetalleMovimiento(id); }

    let nuevoTipo = c.tipo !== undefined ? String(c.tipo).trim() : m.tipo;
    if (esIngreso_(nuevoTipo) && c.categoria !== undefined) nuevoTipo = tipoIngresoPorCategoria_(nuevoTipo, String(c.categoria || '').trim());
    if ((!esIngreso_(nuevoTipo) && !esGasto_(nuevoTipo)) || TIPOS_MOVIMIENTO.indexOf(nuevoTipo) === -1) throw new Error('Tipo inválido.');
    const nuevoMonto = c.monto !== undefined ? validarMonto_(c.monto) : m.monto;
    const nuevaMoneda = c.moneda !== undefined ? validarMoneda_(c.moneda) : m.moneda;
    const nuevaFecha = c.fecha !== undefined && c.fecha !== '' ? fechaConHora_(c.fecha) : m.fecha;
    if (esFuturo_(nuevaFecha)) throw new Error('La fecha es futura. Para algo que aún no ocurre usa "Programar".');
    if (marcaDiv) {
      const cambiaCuenta = c.cuentaId !== undefined && c.cuentaId !== '' && String(c.cuentaId) !== String(m.cuentaId);
      if (nuevoMonto !== m.monto || nuevaMoneda !== m.moneda || cambiaCuenta) throw new Error('Este gasto está dividido: para cambiar el monto, la moneda o la cuenta primero deshaz la división.');
    }

    // --- Gasto con tarjeta ---
    if (clase === 'tarjeta') {
      const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
      const tarj = tid ? buscarFilaPorId_('TarjetaCredito', tid) : null;
      if (!tarj) throw new Error('No encuentro el consumo de tarjeta ligado a este movimiento.');
      if (/Cuotas #\w+ \(\d+\/\d+\)/.test(m.notas)) {
        if (nuevoMonto !== m.monto || nuevaMoneda !== m.moneda || nuevaFecha.getTime() !== m.fecha.getTime() || nuevoTipo !== m.tipo) throw new Error('Una compra en cuotas solo permite cambiar textos. Elimínala y vuelve a crearla.');
        soloTexto(); return obtenerDetalleMovimiento(id);
      }
      if (String(tarj.fila[7]).trim() === 'Pagado') {
        if (nuevoMonto !== m.monto || nuevaMoneda !== m.moneda || nuevaFecha.getTime() !== m.fecha.getTime() || nuevoTipo !== m.tipo) {
          throw new Error('Ese consumo ya fue pagado: solo puedes cambiar textos.');
        }
        soloTexto(); return obtenerDetalleMovimiento(id);
      }
      if (!esGasto_(nuevoTipo)) throw new Error('Un consumo de tarjeta es un gasto.');
      const ciclos = calcularFechaCobroTarjeta_(nuevaFecha, nuevaMoneda, tarjetaPorId_(mapaFilaTarjeta_()[String(tid)]));
      hoja.getRange(m.fila, 2, 1, 14).setValues([[
        nuevaFecha, ciclos.fechaCobro, nuevoTipo, nuevoMonto, nuevaMoneda, '', t.centro, t.categoria, t.subcategoria, t.descripcion,
        'Tarjeta de Credito', t.sunat, m.origen, t.notas
      ]]);
      getHoja_('TarjetaCredito').getRange(tarj.numeroFila, 2, 1, 12).setValues([[
        nuevaFecha, t.descripcion, nuevoMonto, nuevaMoneda, ciclos.fechaCierre, ciclos.fechaCobro, tarj.fila[7], tarj.fila[8], t.sunat, t.centro, t.categoria, t.subcategoria
      ]]);
      registrarLog_('EditarMovimiento', '#' + id + ' (tarjeta)');
      return obtenerDetalleMovimiento(id);
    }

    // --- Ingreso / gasto normal ---
    const nuevaCuentaId = c.cuentaId !== undefined && c.cuentaId !== '' ? c.cuentaId : m.cuentaId;
    const cuentaNueva = validarCuentaMoneda_(nuevaCuentaId, nuevaMoneda);
    const medio = cuentaNueva.tipo === 'Efectivo' ? 'Efectivo' : 'Debito';
    const efectoViejo = efectoCaja_(m);
    const efectoNuevo = esIngreso_(nuevoTipo) ? nuevoMonto : -nuevoMonto;

    actualizarSaldoCuenta_(m.cuentaId, -efectoViejo);
    try {
      actualizarSaldoCuenta_(nuevaCuentaId, efectoNuevo);
    } catch (e) {
      actualizarSaldoCuenta_(m.cuentaId, efectoViejo);
      throw e;
    }
    hoja.getRange(m.fila, 2, 1, 14).setValues([[
      nuevaFecha, nuevaFecha, nuevoTipo, nuevoMonto, nuevaMoneda, nuevaCuentaId, t.centro, t.categoria, t.subcategoria, t.descripcion,
      medio, t.sunat, m.origen, t.notas
    ]]);
    registrarLog_('EditarMovimiento', '#' + id);
    return obtenerDetalleMovimiento(id);
  });
}

/** Borra una fila de Movimientos revirtiendo su efecto en la cuenta */
function borrarFilaMovimiento_(movId) {
  const fila = buscarFilaPorId_('Movimientos', movId);
  if (!fila) return;
  const f = fila.fila;
  const m = {
    cuentaId: f[M.CUENTA], medio: String(f[M.MEDIO]).trim(), tipo: String(f[M.TIPO]).trim(), monto: numero_(f[M.MONTO])
  };
  const efecto = efectoCaja_(m);
  if (efecto !== 0) actualizarSaldoCuenta_(m.cuentaId, -efecto);
  getHoja_('Movimientos').deleteRow(fila.numeroFila);
}

function eliminarMovimiento(id) {
  return conLock_(function () {
    const movs = leerMovimientos_();
    const m = movs.filter(function (x) { return String(x.id) === String(id); })[0];
    if (!m) throw new Error('Movimiento no encontrado.');
    const clase = clasificarMovimiento_(m);

    // Operaciones de inversión: se borran completas (compra + comisión)
    const inv = extraerRef_(m.notas, /Inversion #(\d+)/);
    if (inv) { eliminarOperacionInversion(inv); return; }

    if (clase === 'ligado') throw new Error('Está ligado a un préstamo o deuda y no se puede borrar desde aquí.');

    // Gasto dividido: al borrarlo también se borran las deudas que generó (si nadie te ha devuelto nada)
    if (/Dividido #\w+ \(total [\d.]+\)/.test(m.notas)) {
      if (clase === 'tarjeta') {
        const tidD = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
        const tjD = tidD ? buscarFilaPorId_('TarjetaCredito', tidD) : null;
        if (tjD && String(tjD.fila[7]).trim() === 'Pagado') throw new Error('Ese consumo ya se pagó con un estado de cuenta; no se puede borrar.');
      }
      quitarDeudasDivision_(m.id);
    }

    if (clase === 'tarjeta' && /Cuotas #\w+ \(\d+\/\d+\)/.test(m.notas)) {
      eliminarGrupoCuotas_(extraerRef_(m.notas, /Cuotas #(\w+) \(/));
    } else if (clase === 'tarjeta') {
      const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
      const tarj = tid ? buscarFilaPorId_('TarjetaCredito', tid) : null;
      if (tarj && String(tarj.fila[7]).trim() === 'Pagado') throw new Error('Ese consumo ya se pagó con un estado de cuenta; no se puede borrar.');
      borrarFilaMovimiento_(id);
      if (tarj) getHoja_('TarjetaCredito').deleteRow(buscarFilaPorId_('TarjetaCredito', tid).numeroFila);
    } else if (clase === 'transferencia') {
      const traspaso = extraerRef_(m.notas, /Traspaso #(\d+)/);
      const cambio = extraerRef_(m.notas, /Ref\. CambiosDivisa #(\d+)/);
      if (traspaso) {
        movs.filter(function (x) { return /Traspaso #/.test(x.notas) && extraerRef_(x.notas, /Traspaso #(\d+)/) === traspaso; })
          .forEach(function (x) { borrarFilaMovimiento_(x.id); });
      } else if (cambio) {
        movs.filter(function (x) { return extraerRef_(x.notas, /Ref\. CambiosDivisa #(\d+)/) === cambio; })
          .forEach(function (x) { borrarFilaMovimiento_(x.id); });
        const cd = buscarFilaPorId_('CambiosDivisa', cambio);
        if (cd) getHoja_('CambiosDivisa').deleteRow(cd.numeroFila);
      } else if (m.categoria === 'Tarjeta de Credito' && m.subcategoria === 'Pago de estado de cuenta') {
        // Deshace un pago de tarjeta: los consumos de ese cobro vuelven a "Pendiente"
        const fechaCobro = extraerRef_(m.descripcion, /cobro (\d{4}-\d{2}-\d{2})/);
        borrarFilaMovimiento_(id);
        if (fechaCobro) {
          const hojaT = getHoja_('TarjetaCredito');
          const datos = hojaT.getDataRange().getValues();
          for (let i = 1; i < datos.length; i++) {
            if (String(datos[i][7]).trim() === 'Pagado' && fechaISO_(datos[i][6]) === fechaCobro && String(datos[i][4]).trim() === m.moneda) {
              hojaT.getRange(i + 1, 8).setValue('Pendiente');
            }
          }
        }
      } else {
        borrarFilaMovimiento_(id);
      }
    } else {
      borrarFilaMovimiento_(id);
    }

    // Si venía de un programado, vuelve a quedar pendiente
    const prog = listarProgramados_('Realizado').filter(function (p) { return String(p.movimientoId) === String(id); })[0];
    if (prog) {
      const pf = buscarFilaPorId_('Programados', prog.id);
      if (pf) getHoja_('Programados').getRange(pf.numeroFila, P.ESTADO + 1, 1, 3).setValues([['Pendiente', '', '']]);
    }
    registrarLog_('EliminarMovimiento', '#' + id);
  });
}


/* ==== 18_Inversiones.gs ==== */
/**
 * ============================================================
 * INVERSIONES — acciones/ETF: compras, ventas, dividendos y cotizaciones
 * ============================================================
 * Modelo: tienes una cuenta (tipo Inversion o la que uses en tu broker). Una COMPRA saca el dinero de esa
 * cuenta como Transferencia (no es gasto) y suma la posición; el patrimonio incluye el valor de mercado de
 * tus acciones. Una VENTA devuelve el dinero a la cuenta. Un DIVIDENDO es ingreso pasivo (2da categoría).
 * Las comisiones sí son gasto. Costo de la posición = método de costo promedio.
 * Cotizaciones: Yahoo Finance (sin clave). Se guardan en la hoja Cotizaciones por si Yahoo no responde.
 * Los precios son de mercado con un posible retraso de unos minutos según la bolsa: sirve para seguimiento, no para operar.
 */

var INV = { ID: 0, FECHA: 1, OP: 2, TICKER: 3, CANT: 4, PRECIO: 5, MONEDA: 6, COMISION: 7, CUENTA: 8, MOV: 9, NOTAS: 10 };
var YAHOO_ = 'https://query1.finance.yahoo.com/v8/finance/chart/';
var OPCIONES_YAHOO_ = { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0 (compatible; FinanzasApp/1.0)' } };

function limpiarTicker_(t) {
  const s = String(t || '').trim().toUpperCase().replace(/[^A-Z0-9.\-=^]/g, '');
  if (!s) throw new Error('Escribe el símbolo (ticker).');
  return s;
}

function leerOperacionesInversion_() {
  const datos = leerHoja_('Inversiones');
  const ops = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (f[INV.ID] === '') continue;
    ops.push({
      id: f[INV.ID], fila: i + 1, fecha: fecha_(f[INV.FECHA]) || new Date(), operacion: String(f[INV.OP]).trim(), ticker: String(f[INV.TICKER]).trim().toUpperCase(),
      cantidad: numero_(f[INV.CANT]), precio: numero_(f[INV.PRECIO]), moneda: String(f[INV.MONEDA]).trim(), comision: numero_(f[INV.COMISION]),
      cuentaId: f[INV.CUENTA], movId: f[INV.MOV], notas: String(f[INV.NOTAS] || '')
    });
  }
  return ops.sort(function (a, b) { return a.fecha - b.fecha || numero_(a.id) - numero_(b.id); });
}

/** Posiciones abiertas por ticker (costo promedio). Lanza error si alguna venta deja cantidad negativa */
function calcularPosiciones_(ops, estricto) {
  const pos = {};
  ops.forEach(function (o) {
    const p = pos[o.ticker] || (pos[o.ticker] = { ticker: o.ticker, moneda: o.moneda, cantidad: 0, costo: 0, realizado: 0, dividendos: 0, primeraCompra: null });
    if (o.operacion === 'Compra') {
      p.cantidad += o.cantidad;
      p.costo += o.cantidad * o.precio + o.comision;
      if (!p.primeraCompra) p.primeraCompra = o.fecha;
    } else if (o.operacion === 'Venta') {
      if (o.cantidad > p.cantidad + 1e-9) {
        if (estricto) throw new Error('Esa operación dejaría ' + o.ticker + ' con cantidad negativa.');
        o = Object.assign({}, o, { cantidad: p.cantidad });
      }
      const promedio = p.cantidad > 0 ? p.costo / p.cantidad : 0;
      p.realizado += o.cantidad * o.precio - o.comision - promedio * o.cantidad;
      p.costo -= promedio * o.cantidad;
      p.cantidad -= o.cantidad;
      if (p.cantidad < 1e-9) { p.cantidad = 0; p.costo = 0; }
    } else if (o.operacion === 'Dividendo') {
      p.dividendos += o.cantidad * o.precio - o.comision;
    }
  });
  return pos;
}

function registrarOperacionInversion(d) {
  return conLock_(function () {
    const operacion = ['Compra', 'Venta', 'Dividendo'].indexOf(d.operacion) !== -1 ? d.operacion : null;
    if (!operacion) throw new Error('Operación inválida (Compra, Venta o Dividendo).');
    const ticker = limpiarTicker_(d.ticker);
    const cantidad = numero_(d.cantidad);
    const precio = numero_(d.precio);
    if (!(cantidad > 0)) throw new Error(operacion === 'Dividendo' ? 'Indica cuántas acciones tenías.' : 'La cantidad debe ser mayor a 0.');
    if (!(precio > 0)) throw new Error(operacion === 'Dividendo' ? 'Indica el dividendo por acción.' : 'El precio debe ser mayor a 0.');
    const moneda = validarMoneda_(d.moneda);
    const comision = numero_(d.comision) > 0 ? redondear_(numero_(d.comision)) : 0;
    const cuenta = validarCuentaMoneda_(d.cuentaId, moneda);
    const fecha = fechaConHora_(d.fecha);
    if (esFuturo_(fecha)) throw new Error('La fecha no puede ser futura.');

    const ops = leerOperacionesInversion_();
    if (operacion === 'Venta') {
      const pos = calcularPosiciones_(ops, false)[ticker];
      if (!pos || pos.cantidad < cantidad - 1e-9) throw new Error('Solo tienes ' + (pos ? pos.cantidad : 0) + ' de ' + ticker + '.');
    }

    const id = generarID_('Inversiones');
    const total = redondear_(cantidad * precio);
    const nota = 'Inversion #' + id;
    let movId = '';

    if (operacion === 'Compra') {
      movId = registrarTransferencia_({ fecha: fecha, monto: -total, moneda: moneda, cuentaId: cuenta.id, descripcion: 'Compra ' + cantidad + ' ' + ticker + ' @ ' + precio, categoria: 'Inversiones', subcategoria: ticker, notas: nota });
    } else if (operacion === 'Venta') {
      movId = registrarTransferencia_({ fecha: fecha, monto: total, moneda: moneda, cuentaId: cuenta.id, descripcion: 'Venta ' + cantidad + ' ' + ticker + ' @ ' + precio, categoria: 'Inversiones', subcategoria: ticker, notas: nota });
    } else {
      movId = registrarMovimiento({ fecha: fecha, tipo: 'Ingreso Pasivo', monto: total, moneda: moneda, cuentaId: cuenta.id, descripcion: 'Dividendo ' + ticker, categoria: 'Dividendos', subcategoria: ticker, centroCosto: 'Personal', categoriaSUNAT: 'Segunda', notas: nota + (d.claveDiv ? ' ' + d.claveDiv : '') });
    }
    if (comision > 0) {
      registrarMovimiento({ fecha: fecha, tipo: 'Gasto Personal', monto: comision, moneda: moneda, cuentaId: cuenta.id, descripcion: (operacion === 'Dividendo' ? 'Retención/impuesto dividendo ' : 'Comisión ') + ticker, categoria: operacion === 'Dividendo' ? 'Impuestos' : 'Comisiones', subcategoria: 'Inversiones', centroCosto: 'Personal', notas: nota });
    }
    getHoja_('Inversiones').appendRow([id, fecha, operacion, ticker, cantidad, precio, moneda, comision, cuenta.id, movId, String(d.notas || '') + (d.claveDiv ? ' ' + d.claveDiv : '')]);
    registrarLog_('Inversion', operacion + ' ' + cantidad + ' ' + ticker + ' @ ' + precio + ' ' + moneda);
    return id;
  });
}

function eliminarOperacionInversion(id) {
  return conLock_(function () {
    const ops = leerOperacionesInversion_();
    const op = ops.filter(function (o) { return String(o.id) === String(id); })[0];
    if (!op) throw new Error('Operación no encontrada.');
    calcularPosiciones_(ops.filter(function (o) { return String(o.id) !== String(id); }), true); // valida que no deje cantidades negativas
    leerMovimientos_().filter(function (m) { return new RegExp('Inversion #' + id + '(\\D|$)').test(m.notas); })
      .forEach(function (m) { borrarFilaMovimiento_(m.id); });
    const fila = buscarFilaPorId_('Inversiones', id);
    if (fila) getHoja_('Inversiones').deleteRow(fila.numeroFila);
    registrarLog_('Inversion', 'Eliminada operación #' + id);
  });
}

// ---------------------------------------------------------------
// COTIZACIONES
// ---------------------------------------------------------------

function leerCotizaciones_() {
  const datos = leerHoja_('Cotizaciones');
  const mapa = {};
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    mapa[String(datos[i][0]).toUpperCase()] = {
      fila: i + 1, ticker: String(datos[i][0]).toUpperCase(), precio: numero_(datos[i][1]), moneda: String(datos[i][2]), nombre: String(datos[i][3]),
      cambioPct: numero_(datos[i][4]), fecha: fecha_(datos[i][5]), previo: numero_(datos[i][6])
    };
  }
  return mapa;
}

function parsearCotizacionYahoo_(texto) {
  const j = JSON.parse(texto);
  const r = j.chart && j.chart.result && j.chart.result[0];
  if (!r || !r.meta || !(r.meta.regularMarketPrice > 0)) return null;
  const meta = r.meta;
  const previo = meta.chartPreviousClose || meta.previousClose || 0;
  return {
    precio: meta.regularMarketPrice, moneda: meta.currency || '', nombre: meta.shortName || meta.longName || meta.symbol || '',
    previo: previo, cambioPct: previo > 0 ? (meta.regularMarketPrice - previo) / previo * 100 : 0
  };
}

/** Refresca las cotizaciones más viejas que maxEdadMin. Devuelve el mapa de cotizaciones actualizado */
function actualizarCotizaciones_(tickers, maxEdadMin) {
  const cache = leerCotizaciones_();
  const ahora = new Date();
  const pedir = tickers.filter(function (t) {
    const c = cache[t];
    return !c || !c.fecha || (ahora - c.fecha) / 60000 >= maxEdadMin;
  });
  if (pedir.length === 0) return cache;

  let respuestas = [];
  try {
    respuestas = UrlFetchApp.fetchAll(pedir.map(function (t) {
      return Object.assign({ url: YAHOO_ + encodeURIComponent(t) + '?range=1d&interval=1d' }, OPCIONES_YAHOO_);
    }));
  } catch (e) {
    registrarLog_('Cotizaciones', 'Error de red: ' + e.message);
    return cache;
  }
  const hoja = getHoja_('Cotizaciones');
  pedir.forEach(function (t, i) {
    try {
      if (respuestas[i].getResponseCode() !== 200) return;
      const q = parsearCotizacionYahoo_(respuestas[i].getContentText());
      if (!q) return;
      const fila = [t, q.precio, q.moneda, q.nombre, q.cambioPct, ahora, q.previo];
      if (cache[t]) hoja.getRange(cache[t].fila, 1, 1, 7).setValues([fila]);
      else hoja.appendRow(fila);
      cache[t] = { fila: cache[t] ? cache[t].fila : 0, ticker: t, precio: q.precio, moneda: q.moneda, nombre: q.nombre, cambioPct: q.cambioPct, fecha: ahora, previo: q.previo };
    } catch (e) { registrarLog_('Cotizaciones', t + ': ' + e.message); }
  });
  return cache;
}

/** Valor de mercado de las posiciones, en soles. Usa cotizaciones guardadas (si no hay, el costo) */
function valorInversionesPEN_(tasa, cotizaciones) {
  const pos = calcularPosiciones_(leerOperacionesInversion_(), false);
  const cot = cotizaciones || leerCotizaciones_();
  let total = 0;
  Object.keys(pos).forEach(function (t) {
    const p = pos[t];
    if (p.cantidad <= 0) return;
    const precio = cot[t] && cot[t].precio > 0 ? cot[t].precio : (p.costo / p.cantidad);
    total += convertirAPEN_(precio * p.cantidad, p.moneda, tasa);
  });
  return redondear_(total);
}

// ---------------------------------------------------------------
// DATOS PARA LA APP
// ---------------------------------------------------------------

function resumenInversiones_(refrescar) {
  const ops = leerOperacionesInversion_();
  const posiciones = calcularPosiciones_(ops, false);
  const tickers = Object.keys(posiciones).filter(function (t) { return posiciones[t].cantidad > 0; });
  const cot = refrescar ? actualizarCotizaciones_(tickers, 1) : leerCotizaciones_();
  const tasa = obtenerUltimaTasaReferencial_();

  const lista = [];
  const totales = {};
  let valorPEN = 0, costoPEN = 0, plPEN = 0, divPEN = 0, realizadoPEN = 0;
  const sinCotizacion = [];

  Object.keys(posiciones).forEach(function (t) {
    const p = posiciones[t];
    divPEN += convertirAPEN_(p.dividendos, p.moneda, tasa);
    realizadoPEN += convertirAPEN_(p.realizado, p.moneda, tasa);
    if (p.cantidad <= 0) return;
    const c = cot[t];
    const tienePrecio = c && c.precio > 0;
    if (!tienePrecio) sinCotizacion.push(t);
    const precio = tienePrecio ? c.precio : p.costo / p.cantidad;
    const valor = precio * p.cantidad;
    const pl = valor - p.costo;
    lista.push({
      ticker: t, nombre: c ? c.nombre : '', moneda: p.moneda, cantidad: p.cantidad, costoMedio: p.costo / p.cantidad, costoTotal: redondear_(p.costo),
      precio: precio, valor: redondear_(valor), pl: redondear_(pl), plPct: p.costo > 0 ? pl / p.costo * 100 : 0,
      cambioDiaPct: c ? c.cambioPct : 0, cambioDia: c && c.previo > 0 ? redondear_((precio - c.previo) * p.cantidad) : 0,
      dividendos: redondear_(p.dividendos), realizado: redondear_(p.realizado), actualizado: c && c.fecha ? fechaHoraISO_(c.fecha) : '', sinPrecio: !tienePrecio
    });
    valorPEN += convertirAPEN_(valor, p.moneda, tasa);
    costoPEN += convertirAPEN_(p.costo, p.moneda, tasa);
    plPEN += convertirAPEN_(pl, p.moneda, tasa);
    const tm = totales[p.moneda] || (totales[p.moneda] = { valor: 0, costo: 0, pl: 0 });
    tm.valor += valor; tm.costo += p.costo; tm.pl += pl;
  });
  lista.forEach(function (p) { p.peso = valorPEN > 0 ? convertirAPEN_(p.valor, p.moneda, tasa) / valorPEN * 100 : 0; });
  lista.sort(function (a, b) { return convertirAPEN_(b.valor, b.moneda, tasa) - convertirAPEN_(a.valor, a.moneda, tasa); });
  Object.keys(totales).forEach(function (k) { Object.keys(totales[k]).forEach(function (f) { totales[k][f] = redondear_(totales[k][f]); }); });

  return {
    posiciones: lista, totales: totales, valorPEN: redondear_(valorPEN), costoPEN: redondear_(costoPEN), plPEN: redondear_(plPEN),
    plPctPEN: costoPEN > 0 ? plPEN / costoPEN * 100 : 0, dividendosPEN: redondear_(divPEN), realizadoPEN: redondear_(realizadoPEN),
    sinCotizacion: sinCotizacion, ahora: fechaHoraISO_(new Date()),
    operaciones: ops.slice().reverse().slice(0, 60).map(function (o) {
      return { id: o.id, fecha: fechaISO_(o.fecha), operacion: o.operacion, ticker: o.ticker, cantidad: o.cantidad, precio: o.precio, moneda: o.moneda, comision: o.comision, notas: o.notas };
    })
  };
}

var RANGOS_YAHOO_ = { '1D': ['1d', '5m'], '1S': ['5d', '30m'], '1M': ['1mo', '1d'], '6M': ['6mo', '1d'], '1A': ['1y', '1d'], '5A': ['5y', '1wk'] };

/** Serie de precios para el gráfico de una acción. rango: 1D, 1S, 1M, 6M, 1A, 5A */
/** Descarga con reintento y con la otra dirección de Yahoo (query1 → query2). Devuelve { res, error } */
function yahooChart_(t, query) {
  let ultimo = '';
  const hosts = ['https://query1.finance.yahoo.com/v8/finance/chart/', 'https://query2.finance.yahoo.com/v8/finance/chart/'];
  for (let i = 0; i < hosts.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(hosts[i] + encodeURIComponent(t) + query, OPCIONES_YAHOO_);
      const code = resp.getResponseCode();
      if (code === 200) {
        const j = JSON.parse(resp.getContentText());
        const res = j.chart && j.chart.result && j.chart.result[0];
        if (res && res.timestamp) return { res: res, error: '' };
        ultimo = 'Yahoo no devolvió datos para ' + t;
      } else ultimo = 'Yahoo respondió ' + code;
    } catch (e) { ultimo = e.message; }
    if (i === 0) Utilities.sleep(400);
  }
  return { res: null, error: ultimo };
}

/** Respaldo gratuito (Stooq) para acciones de EE. UU.: cierres diarios [[ms, cierre]] de los últimos `dias` */
function stooqCierres_(ticker, dias) {
  try {
    const s = String(ticker).toLowerCase().replace(/\./g, '-');
    const resp = UrlFetchApp.fetch('https://stooq.com/q/d/l/?s=' + encodeURIComponent(s) + '.us&i=d', { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    const lineas = resp.getContentText().trim().split('\n');
    if (lineas.length < 3 || lineas[0].indexOf('Date') !== 0) return null;
    const out = [];
    for (let i = Math.max(1, lineas.length - dias); i < lineas.length; i++) {
      const c = lineas[i].split(',');
      const ms = new Date(c[0] + 'T16:00:00-05:00').getTime(), cierre = Number(c[4]);
      if (isFinite(ms) && cierre > 0) out.push([ms, Math.round(cierre * 10000) / 10000]);
    }
    return out.length ? out : null;
  } catch (e) { return null; }
}

var DIAS_RANGO_ = { '1M': 24, '6M': 130, '1A': 255, '5A': 1300 };

function historialAccion_(ticker, rango) {
  const t = limpiarTicker_(ticker);
  const r = RANGOS_YAHOO_[rango] || RANGOS_YAHOO_['1M'];
  const clave = 'hist_' + t + '_' + rango;
  const cache = CacheService.getScriptCache();
  const guardado = cache.get(clave);
  if (guardado) return JSON.parse(guardado);

  let puntos = [], moneda = '', fuente = 'Yahoo';
  const y = yahooChart_(t, '?range=' + r[0] + '&interval=' + r[1]);
  if (y.res) {
    const cierres = y.res.indicators.quote[0].close;
    y.res.timestamp.forEach(function (ts, i) { if (cierres[i] !== null && cierres[i] !== undefined) puntos.push([ts * 1000, Math.round(cierres[i] * 10000) / 10000]); });
    moneda = y.res.meta.currency || '';
  }
  if (!puntos.length && DIAS_RANGO_[rango]) {
    const s = stooqCierres_(t, DIAS_RANGO_[rango]);
    if (s) { puntos = s; fuente = 'Stooq'; moneda = 'USD'; }
  }
  if (!puntos.length) throw new Error('No pude obtener el historial de ' + t + (y.error ? ' (' + y.error + ')' : '') + '.');
  if (puntos.length > 220) {
    const paso = puntos.length / 220;
    puntos = puntos.filter(function (p, i) { return Math.floor(i % paso) === 0 || i === puntos.length - 1; });
  }
  const salida = { ticker: t, rango: rango, puntos: puntos, moneda: moneda, fuente: fuente };
  try { cache.put(clave, JSON.stringify(salida), rango === '1D' ? 120 : 600); } catch (e) { /* demasiado grande: sin caché */ }
  return salida;
}

/** Dividendos recientes de tus acciones que aún no registraste (estimados: cantidad que tenías × dividendo por acción) */
function dividendosSugeridos_() {
  const ops = leerOperacionesInversion_();
  const posiciones = calcularPosiciones_(ops, false);
  const tickers = Object.keys(posiciones).filter(function (t) { return posiciones[t].primeraCompra; });
  if (tickers.length === 0) return [];
  const sugeridos = [];

  let respuestas;
  try {
    respuestas = UrlFetchApp.fetchAll(tickers.map(function (t) {
      return Object.assign({ url: YAHOO_ + encodeURIComponent(t) + '?range=1y&interval=1d&events=div' }, OPCIONES_YAHOO_);
    }));
  } catch (e) { return []; }

  tickers.forEach(function (t, i) {
    try {
      if (respuestas[i].getResponseCode() !== 200) return;
      const res = JSON.parse(respuestas[i].getContentText()).chart.result[0];
      const divs = res.events && res.events.dividends ? res.events.dividends : {};
      Object.keys(divs).forEach(function (k) {
        const ev = divs[k];
        const fechaEx = new Date(ev.date * 1000);
        const claveDiv = 'ex:' + t + ':' + fechaISO_(fechaEx);
        if (fechaEx < posiciones[t].primeraCompra) return;
        if (ops.some(function (o) { return o.operacion === 'Dividendo' && o.notas.indexOf(claveDiv) !== -1; })) return;
        let acciones = 0;
        ops.forEach(function (o) {
          if (o.ticker !== t || o.fecha > fechaEx) return;
          if (o.operacion === 'Compra') acciones += o.cantidad; else if (o.operacion === 'Venta') acciones -= o.cantidad;
        });
        if (acciones <= 0) return;
        sugeridos.push({ ticker: t, moneda: posiciones[t].moneda, fechaEx: fechaISO_(fechaEx), porAccion: ev.amount, acciones: acciones, total: redondear_(acciones * ev.amount), claveDiv: claveDiv });
      });
    } catch (e) { /* ticker sin dividendos */ }
  });
  return sugeridos.sort(function (a, b) { return a.fechaEx < b.fechaEx ? 1 : -1; });
}


/* ==== 19_IA.gs ==== */
/**
 * ============================================================
 * IA — recomendaciones financieras y mensajes libres (texto o audio) por Telegram
 * ============================================================
 * Propiedades del script:
 *   ANTHROPIC_API_KEY  → IA para recomendaciones y para entender texto ("gasté 5 soles en huevos")
 *   GEMINI_API_KEY     → necesaria para entender AUDIOS (Anthropic no transcribe audio); también sirve
 *                        como IA alterna para texto si no tienes la de Anthropic.
 * Sin ninguna clave la app sigue funcionando: las recomendaciones salen de reglas fijas (sin IA).
 * La IA solo recibe un resumen numérico de tus finanzas (nunca tus claves ni tokens).
 */

function iaProveedor_() {
  if (prop_('ANTHROPIC_API_KEY')) return 'anthropic';
  if (prop_('GEMINI_API_KEY')) return 'gemini';
  return '';
}

function llamarAnthropic_(system, user, maxTokens) {
  const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': prop_('ANTHROPIC_API_KEY'), 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: prop_('AI_MODEL', 'claude-haiku-5-5'), max_tokens: maxTokens || 1200, system: system, messages: [{ role: 'user', content: user }] })
  });
  if (r.getResponseCode() !== 200) throw new Error('Anthropic respondió ' + r.getResponseCode() + ': ' + r.getContentText().substring(0, 200));
  const j = JSON.parse(r.getContentText());
  return (j.content || []).map(function (b) { return b.text || ''; }).join('');
}

function geminiModelo_() {
  return prop_('GEMINI_MODEL_AUTO') || prop_('GEMINI_MODEL') || 'gemini-flash-latest';
}

/** Si Google retira el modelo (404), busca el flash más reciente disponible para tu clave y lo recuerda */
function geminiDescubrirModelo_() {
  const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { muteHttpExceptions: true, headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') } });
  if (r.getResponseCode() !== 200) return '';
  const lista = (JSON.parse(r.getContentText()).models || []).filter(function (m) {
    return (m.supportedGenerationMethods || []).indexOf('generateContent') !== -1 && /\/gemini-[\d.]+-flash$/.test(m.name);
  }).map(function (m) { return { nombre: m.name.replace('models/', ''), version: parseFloat(m.name.match(/gemini-([\d.]+)-flash/)[1]) }; })
    .sort(function (a, b) { return b.version - a.version; });
  if (!lista.length) return '';
  PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL_AUTO', lista[0].nombre);
  return lista[0].nombre;
}

/** Modelos a probar, en orden: el elegido y luego los demás flash / flash-lite de tu clave (cada modelo tiene su propia cuota gratuita) */
function geminiModelos_() {
  const cache = CacheService.getScriptCache();
  const g = cache.get('gem_modelos');
  let lista = [geminiModelo_()];
  if (g) { try { lista = lista.concat(JSON.parse(g)); } catch (e) { /* se recalcula */ } }
  else {
    try {
      const r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { muteHttpExceptions: true, headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') } });
      if (r.getResponseCode() === 200) {
        const m = (JSON.parse(r.getContentText()).models || []).filter(function (x) {
          return (x.supportedGenerationMethods || []).indexOf('generateContent') !== -1 && /\/gemini-[\d.]+-flash(-lite)?$/.test(x.name);
        }).map(function (x) { const mm = x.name.match(/gemini-([\d.]+)-flash(-lite)?$/); return { n: x.name.replace('models/', ''), v: parseFloat(mm[1]), lite: !!mm[2] }; })
          .sort(function (p, q) { return (p.lite - q.lite) || (q.v - p.v); });
        const extra = m.map(function (x) { return x.n; }).concat(['gemini-flash-latest', 'gemini-flash-lite-latest']);
        cache.put('gem_modelos', JSON.stringify(extra), 21600);
        lista = lista.concat(extra);
      }
    } catch (e) { lista = lista.concat(['gemini-flash-latest', 'gemini-flash-lite-latest']); }
  }
  const vistos = {};
  return lista.filter(function (x) { if (!x || vistos[x]) return false; vistos[x] = 1; return true; }).slice(0, 5);
}

function geminiLlamadaCruda_(modelo, cuerpo) {
  return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': prop_('GEMINI_API_KEY') }, payload: JSON.stringify(cuerpo)
  });
}

function geminiTextoDe_(r) {
  const j = JSON.parse(r.getContentText());
  const parts = j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts;
  if (!parts) throw new Error('Gemini no devolvió texto.');
  return parts.map(function (p) { return p.text || ''; }).join('');
}

/**
 * Llama a Gemini y devuelve el texto. Si un modelo está retirado (404) o agotó su cuota (429) prueba el siguiente;
 * si la búsqueda web falla, reintenta sin ella; y si es solo el límite por minuto, espera unos segundos una vez.
 */
function geminiGenerar_(cuerpo, conBusqueda) {
  const modelos = geminiModelos_();
  let ultimo = '', diario = false, espera = 0, retirado = false;
  for (let ronda = 0; ronda < 2; ronda++) {
    for (let i = 0; i < modelos.length; i++) {
      const modelo = modelos[i];
      const variantes = conBusqueda ? [Object.assign({}, cuerpo, { tools: [{ google_search: {} }] }), cuerpo] : [cuerpo];
      for (let v = 0; v < variantes.length; v++) {
        const r = geminiLlamadaCruda_(modelo, variantes[v]);
        const codigo = r.getResponseCode();
        if (codigo === 200) {
          const t = geminiTextoDe_(r);
          if (retirado && modelo !== geminiModelo_()) PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL_AUTO', modelo);   // el anterior ya no existe: recuérdalo
          return t;
        }
        const texto = r.getContentText();
        ultimo = codigo + ': ' + texto.substring(0, 200);
        if (codigo === 429) {
          if (/PerDay|per day|daily/i.test(texto)) diario = true;
          const m = texto.match(/retry in ([\d.]+)s/i) || texto.match(/"retryDelay":\s*"(\d+)s"/);
          if (m) espera = Math.max(espera, Math.ceil(Number(m[1])));
        }
        if (codigo === 404) { retirado = true; try { CacheService.getScriptCache().remove('gem_modelos'); } catch (e) { /* opcional */ } }
        if (variantes.length > 1 && v === 0) continue;   // reintenta este modelo sin búsqueda web
        break;
      }
    }
    // Todos dieron 429: si es el límite por minuto, espera y prueba otra vez (una sola vez)
    if (ronda === 0 && espera > 0 && espera <= 20 && !diario) { Utilities.sleep(espera * 1000 + 500); continue; }
    break;
  }
  if (/^429/.test(ultimo)) {
    throw new Error(diario
      ? 'Gemini agotó la cuota gratuita DIARIA de tu clave (se reinicia en la madrugada, hora de Lima). Mientras tanto el bot entiende mensajes simples sin IA.'
      : 'Gemini está saturado en este momento (límite por minuto). Prueba de nuevo en 1 minuto.');
  }
  throw new Error('Gemini respondió ' + ultimo);
}

function llamarGemini_(system, partes, maxTokens) {
  return geminiGenerar_({
    systemInstruction: system ? { parts: [{ text: system }] } : undefined,
    contents: [{ role: 'user', parts: partes }],
    generationConfig: { maxOutputTokens: Math.max(2048, (maxTokens || 1200) * 3), temperature: 0.2 }
  }, false);
}

function iaTexto_(system, user, maxTokens) {
  const prov = iaProveedor_();
  if (prov === 'anthropic') return llamarAnthropic_(system, user, maxTokens);
  if (prov === 'gemini') return llamarGemini_(system, [{ text: user }], maxTokens);
  throw new Error('No hay IA configurada. Agrega ANTHROPIC_API_KEY (o GEMINI_API_KEY) en las propiedades del script.');
}

function transcribirAudio_(blob) {
  if (!prop_('GEMINI_API_KEY')) throw new Error('Para entender audios agrega GEMINI_API_KEY en las propiedades del script (gratis en aistudio.google.com).');
  const mime = blob.getContentType() || 'audio/ogg';
  const texto = llamarGemini_(
    'Transcribe el audio literalmente en español. Devuelve SOLO la transcripción, sin comentarios.',
    [{ text: 'Transcribe este audio.' }, { inline_data: { mime_type: mime.indexOf('audio') === 0 ? mime : 'audio/ogg', data: Utilities.base64Encode(blob.getBytes()) } }], 400);
  return texto.trim();
}

function extraerJSON_(t) {
  const s = String(t || '');
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a === -1 || b <= a) throw new Error('La IA no devolvió un formato válido.');
  return JSON.parse(s.substring(a, b + 1));
}

// ---------------------------------------------------------------
// RESUMEN FINANCIERO + RECOMENDACIONES
// ---------------------------------------------------------------

function construirResumenFinanciero_() {
  const movs = leerMovimientos_();
  const tasa = obtenerUltimaTasaReferencial_();
  const hoy = new Date();
  const mesAnt = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 15);
  const actual = resumenPeriodo_('Mensual', hoy, movs, tasa);
  const previo = resumenPeriodo_('Mensual', mesAnt, movs, tasa);
  const patrimonio = calcularPatrimonioDetalle_();

  const meses = [];
  for (let k = 5; k >= 0; k--) {
    const r = resumenPeriodo_('Mensual', new Date(hoy.getFullYear(), hoy.getMonth() - k, 15), movs, tasa);
    meses.push({ mes: nombreMes_(new Date(hoy.getFullYear(), hoy.getMonth() - k, 15)), ingresos: r.ingresos, egresos: r.egresos });
  }
  const metas = [];
  ['Diaria', 'Semanal', 'Mensual'].forEach(function (p) {
    CONFIG.MONEDAS.forEach(function (m) { const x = calcularProgresoMeta(p, m, hoy, movs); if (x.meta > 0) metas.push(x); });
  });
  const liquido = { PEN: 0, USD: 0, EUR: 0 };
  obtenerCuentasActivas_().forEach(function (c) { if (c.tipo !== 'Inversion' && liquido[c.moneda] !== undefined) liquido[c.moneda] += c.saldo; });
  const inv = resumenInversiones_(false);

  return {
    fecha: fechaISO_(hoy), diaDelMes: hoy.getDate(), diasDelMes: ultimoDiaMes_(hoy),
    patrimonioPEN: patrimonio, liquidoPorMoneda: liquido,
    mesActual: { ingresos: actual.ingresos, egresos: actual.egresos, ganancia: actual.ganancia, categorias: actual.categorias.slice(0, 8) },
    mesAnterior: { ingresos: previo.ingresos, egresos: previo.egresos, ganancia: previo.ganancia, categorias: previo.categorias.slice(0, 8) },
    ultimosMeses: meses, presupuestos: obtenerEstadoPresupuestos_(hoy, movs), metas: metas,
    tarjetaPendiente: obtenerEstadosTarjetaPendientes_().map(function (e) { return { cobra: fechaISO_(e.fechaCobro), moneda: e.moneda, total: e.total }; }),
    programados: resumenProgramados_(30).netoProyectado,
    inversiones: { valorPEN: inv.valorPEN, plPct: inv.plPctPEN, pesos: inv.posiciones.slice(0, 6).map(function (p) { return { ticker: p.ticker, peso: Math.round(p.peso) }; }) },
    nota: 'Importes en soles equivalentes (tipo de cambio referencial) salvo liquidoPorMoneda/presupuestos/metas, que van en su moneda.'
  };
}

/** Recomendaciones calculadas con reglas fijas (funcionan sin IA y sin internet) */
function generarInsightsReglas_(r) {
  const out = [];
  const fmt = function (n) { return formatoMonto_(n, 'PEN'); };
  const a = r.mesActual, p = r.mesAnterior;

  if (a.ingresos > 0) {
    const tasaAhorro = a.ganancia / a.ingresos * 100;
    if (a.ganancia < 0) out.push({ nivel: 'alerta', icono: 'trending-down', titulo: 'Gastas más de lo que ingresas', texto: 'Este mes llevas ' + fmt(a.egresos) + ' de egresos contra ' + fmt(a.ingresos) + ' de ingresos (equiv. en soles).' });
    else if (tasaAhorro < 10) out.push({ nivel: 'info', icono: 'piggy-bank', titulo: 'Ahorro bajo este mes', texto: 'Estás conservando ' + tasaAhorro.toFixed(0) + '% de tus ingresos. Una referencia habitual es apuntar a 20%.' });
    else if (tasaAhorro >= 20) out.push({ nivel: 'ok', icono: 'piggy-bank', titulo: '¡Buen ritmo de ahorro!', texto: 'Conservas ' + tasaAhorro.toFixed(0) + '% de tus ingresos este mes.' });
  } else if (a.egresos > 0) {
    out.push({ nivel: 'info', icono: 'wallet', titulo: 'Aún no registras ingresos este mes', texto: 'Ya llevas ' + fmt(a.egresos) + ' en egresos. Registra tus ingresos para ver tu ahorro real.' });
  }
  if (a.categorias.length && a.egresos > 0) {
    const c = a.categorias[0];
    const pct = c.monto / a.egresos * 100;
    if (pct >= 35 && a.categorias.length > 1) out.push({ nivel: 'info', icono: 'pie-chart', titulo: 'Una categoría domina tus gastos', texto: c.categoria + ' concentra ' + pct.toFixed(0) + '% de tus egresos del mes (' + fmt(c.monto) + ').' });
  }
  a.categorias.forEach(function (c) {
    const ant = p.categorias.filter(function (x) { return x.categoria === c.categoria; })[0];
    if (ant && ant.monto >= 50 && c.monto > ant.monto * 1.3 && c.monto - ant.monto >= 30) {
      out.push({ nivel: 'info', icono: 'arrow-up-right', titulo: c.categoria + ' subió vs el mes pasado', texto: 'Llevas ' + fmt(c.monto) + ' contra ' + fmt(ant.monto) + ' el mes anterior (+' + ((c.monto / ant.monto - 1) * 100).toFixed(0) + '%).' });
    }
  });
  r.presupuestos.forEach(function (b) {
    if (b.limite <= 0) return;
    const pct = b.gastado / b.limite * 100;
    if (pct >= 100) out.push({ nivel: 'alerta', icono: 'octagon-alert', titulo: 'Presupuesto superado: ' + b.categoria, texto: 'Llevas ' + formatoMonto_(b.gastado, b.moneda) + ' de ' + formatoMonto_(b.limite, b.moneda) + '.' });
    else if (pct >= 80) out.push({ nivel: 'info', icono: 'gauge', titulo: 'Presupuesto casi al límite: ' + b.categoria, texto: 'Vas en ' + pct.toFixed(0) + '% (' + formatoMonto_(b.gastado, b.moneda) + ' de ' + formatoMonto_(b.limite, b.moneda) + ').' });
  });
  r.metas.filter(function (m) { return m.periodo === 'Mensual'; }).forEach(function (m) {
    const avance = r.diaDelMes / r.diasDelMes;
    if (avance >= 0.3 && m.actual < m.meta) {
      const proyeccion = m.actual / avance;
      out.push({ nivel: proyeccion >= m.meta ? 'ok' : 'info', icono: 'target', titulo: 'Meta mensual en ' + m.moneda,
        texto: 'Llevas ' + formatoMonto_(m.actual, m.moneda) + ' de ' + formatoMonto_(m.meta, m.moneda) + '. A este ritmo cerrarías el mes cerca de ' + formatoMonto_(proyeccion, m.moneda) + '.' });
    }
  });
  r.tarjetaPendiente.forEach(function (t) {
    const liq = r.liquidoPorMoneda[t.moneda] || 0;
    if (liq < t.total) out.push({ nivel: 'alerta', icono: 'credit-card', titulo: 'Tu tarjeta supera tu saldo en ' + t.moneda, texto: 'Se cobra ' + formatoMonto_(t.total, t.moneda) + ' el ' + t.cobra + ' y tienes ' + formatoMonto_(liq, t.moneda) + ' disponibles.' });
  });
  if (r.inversiones.pesos.length && r.inversiones.pesos[0].peso >= 50 && r.inversiones.pesos.length > 1) {
    out.push({ nivel: 'info', icono: 'chart-pie', titulo: 'Cartera concentrada', texto: r.inversiones.pesos[0].ticker + ' pesa ' + r.inversiones.pesos[0].peso + '% de tus inversiones. Diversificar reduce el riesgo.' });
  }
  const prog = listarProgramados_('Pendiente').filter(function (x) { return x.atrasado; });
  if (prog.length) out.push({ nivel: 'info', icono: 'calendar-clock', titulo: prog.length + (prog.length === 1 ? ' movimiento programado atrasado' : ' movimientos programados atrasados'), texto: 'Márcalos como realizados o reprográmalos para que tus proyecciones sean correctas.' });
  if (out.length === 0) out.push({ nivel: 'ok', icono: 'sparkles', titulo: 'Todo en orden', texto: 'No veo alertas ahora. Sigue registrando para que las recomendaciones sean más precisas.' });
  return out;
}

function recomendacionesIA_(r) {
  const system = SYSTEM_ASESOR_ + ' ' +
    'Con base SOLO en los datos JSON, da entre 4 y 6 observaciones accionables, concretas y breves, en español. Empieza por lo MÁS incómodo y urgente, sin suavizarlo. ' +
    'Cita cifras de los datos. No inventes datos. ' +
    'Responde SOLO con JSON: {"recomendaciones":[{"titulo":"","texto":"","prioridad":"alta|media|baja"}]}';
  const j = extraerJSON_(iaTexto_(system, JSON.stringify(r), 1200));
  return (j.recomendaciones || []).slice(0, 6).map(function (x) {
    return { titulo: String(x.titulo || ''), texto: String(x.texto || ''), prioridad: ['alta', 'media', 'baja'].indexOf(x.prioridad) !== -1 ? x.prioridad : 'media' };
  });
}

/** forzar = true ignora el caché de 6 horas (la IA cuesta, por eso se guarda) */
function obtenerRecomendaciones_(forzar) {
  const r = construirResumenFinanciero_();
  const salida = { reglas: generarInsightsReglas_(r), ia: null, iaDisponible: !!iaProveedor_(), generado: '', error: '' };
  if (!salida.iaDisponible) return salida;

  const cache = CacheService.getScriptCache();
  const guardado = cache.get('ia_rec');
  if (guardado && !forzar) { const g = JSON.parse(guardado); salida.ia = g.ia; salida.generado = g.generado; return salida; }
  try {
    salida.ia = recomendacionesIA_(r);
    salida.generado = fechaHoraISO_(new Date());
    cache.put('ia_rec', JSON.stringify({ ia: salida.ia, generado: salida.generado }), 21600);
  } catch (e) {
    salida.error = e.message;
  }
  return salida;
}

// ---------------------------------------------------------------
// MENSAJES LIBRES (Telegram): "gasté 5 soles en huevos"
// ---------------------------------------------------------------

function categoriasConocidas_() {
  const p = categoriasPorTipo_();
  return ['gastos: ' + Object.keys(p.Gasto).slice(0, 40).join(', '), 'ingresos: ' + Object.keys(p.Ingreso).slice(0, 25).join(', ')];
}

/** Texto con tus categorías (y subcategorías), cuentas, tarjetas y personas, para que la IA elija con datos reales */
function contextoIA_() {
  const p = categoriasPorTipo_();
  const fmt = function (m) { return Object.keys(m).slice(0, 45).map(function (c) { return c + (m[c].length ? ' (' + m[c].slice(0, 8).join(', ') + ')' : ''); }).join('; '); };
  let cuentas = '';
  try { cuentas = obtenerCuentasActivas_().map(function (c) { return c.banco + ' [' + c.moneda + ', ' + c.tipo + ']'; }).join('; '); } catch (e) { /* sin cuentas */ }
  let personas = '';
  try { personas = Object.keys(mapaPersonas_()).map(function (k) { return mapaPersonas_()[k]; }).slice(0, 30).join(', '); } catch (e) { /* sin personas */ }
  return { gastos: fmt(p.Gasto), ingresos: fmt(p.Ingreso), cuentas: cuentas, tarjetas: tarjetasActivas_().map(function (t) { return t.banco; }).join(', '), personas: personas };
}

/**
 * Convierte el mensaje en acciones. `pendiente` (opcional) = { json, pregunta, original } cuando el usuario está respondiendo
 * a una pregunta del bot: la IA completa lo que faltaba y devuelve el JSON final.
 */
function interpretarMensajeIA_(texto, pendiente) {
  const ahora = new Date();
  const dias = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  const ctx = contextoIA_();
  const system =
    'Eres el asistente de finanzas personales de un usuario en Perú. Convierte su mensaje en acciones. Hoy es ' + dias[ahora.getDay()] + ' ' + fechaISO_(ahora) +
    ' y son las ' + Utilities.formatDate(ahora, zona_(), 'HH:mm') + '.\n' +
    'Responde SOLO con JSON: {"acciones":[...],"pregunta":null,"no_entendi":false}\n' +
    'Cada acción es una de:\n' +
    '{"accion":"registrar","tipo":"gasto|ingreso|null","monto":número|null,"moneda":"PEN|USD|EUR","descripcion":"texto corto"|null,"medio":"tarjeta|efectivo|debito|null","tarjeta":"nombre o null","cuenta":"nombre de la cuenta o null","fecha":"YYYY-MM-DD"|"YYYY-MM-DDTHH:MM"|null,"clase":"personal|fijo|inversion|activo|pasivo|null","categoria":string|null,"subcategoria":string|null,"dividir":null|{"modo":"igual|porcentaje|monto","personas":[{"nombre":"Ana","pct":null,"monto":null}]}}\n' +
    '{"accion":"consulta","tema":"saldo|metas|tarjeta|mes"}\n' +
    'REGLAS DE LECTURA: "soles/sol/S/" = PEN; "dólares/usd/$" = USD; "euros/€" = EUR; si no dice la moneda usa PEN. ' +
    'Gasto o ingreso se deduce por el sentido ("gasté, pagué, compré, me cobraron" = gasto; "me pagaron, cobré, me depositaron, vendí" = ingreso). ' +
    'tarjeta: si dice con cuál tarjeta, pon ese nombre; si no, null. Tarjetas: [' + ctx.tarjetas + ']. ' +
    'cuenta: si menciona una cuenta o billetera ("de mi BCP", "con yape", "de Interbank") pon su nombre tal como esté en esta lista, si no null. Cuentas: [' + ctx.cuentas + ']. ' +
    'medio: "con tarjeta/crédito" = tarjeta; "efectivo/cash/en mano" = efectivo; "yape/plin/transferencia/débito/depósito" = debito; si no lo dice, null. ' +
    'fecha: null si es ahora; convierte "ayer", "el viernes", "mañana", "el 5 de noviembre" a fecha exacta. Si ocurrirá en el futuro, pon esa fecha futura. ' +
    'clase: gasto → personal (por defecto), fijo (suscripciones, alquiler, servicios) o inversion (negocio, compras para vender); ingreso → activo (por defecto) o pasivo (alquileres, intereses, dividendos). ' +
    'categoria y subcategoria: ELIGE SIEMPRE la que mejor encaje de estas listas, tal como están escritas (no inventes nombres nuevos); null solo si ninguna encaja de verdad. ' +
    'Categorías de gasto: ' + ctx.gastos + '. Categorías de ingreso: ' + ctx.ingresos + '. ' +
    'dividir: si dice que dividió/repartió la cuenta, que pagó por otros o con quién lo compartió ("dividí con Ana y Luis", "salimos con Pedro y yo pagué", "Ana me debe 8"), pon en personas a las OTRAS personas (sin contar al usuario). ' +
    'modo "igual" = partes iguales entre el usuario y esas personas; "porcentaje" si da porcentajes (pct); "monto" si dice cuánto debe devolver cada una (monto). "monto" siempre es lo que PAGÓ el usuario en total. Personas conocidas: [' + ctx.personas + '] (si el nombre se parece a uno de ellos, usa ese). ' +
    'Si dice "entre 4" o "somos 4" sin nombres, pregunta los nombres. ' +
    'FALTANTES: nunca inventes monto, ni si es gasto o ingreso. Si falta algo imprescindible (monto; si es gasto o ingreso; de qué fue si no se puede deducir nada; los nombres cuando quiere dividir y no dice con quién) devuelve en "acciones" lo que sí entendiste (con null en lo que falta) y escribe en "pregunta" UNA sola pregunta corta, amable y concreta. Si no falta nada imprescindible, "pregunta" es null. Lo no imprescindible (moneda, medio de pago, fecha, categoría) se completa solo, no preguntes por eso. ' +
    'Si el mensaje trae varios movimientos, devuelve varias acciones. Si no es un movimiento ni una consulta, devuelve {"acciones":[],"pregunta":null,"no_entendi":true}.';
  let usuario = texto;
  if (pendiente) {
    usuario = 'Mensaje original del usuario: ' + pendiente.original + '\nLo que ya entendiste (JSON): ' + JSON.stringify(pendiente.json) + '\nTu pregunta fue: ' + pendiente.pregunta +
      '\nRespuesta del usuario: ' + texto + '\nCompleta la información con su respuesta y devuelve el JSON final con la misma estructura (si ya no falta nada, "pregunta": null). Si la respuesta no tiene relación con tu pregunta y es un mensaje nuevo, interprétalo como un mensaje nuevo.';
  }
  return extraerJSON_(iaTexto_(system, usuario, 900));
}

var CLASES_GASTO_ = { personal: 'Gasto Personal', fijo: 'Gasto Fijo', inversion: 'Gasto Inversion' };
var CLASES_INGRESO_ = { activo: 'Ingreso Activo', pasivo: 'Ingreso Pasivo' };

/**
 * Registra (o programa, si la fecha es futura) un movimiento. Compartido por Telegram/IA y la app.
 * d = { tipo ('Gasto Personal'...), monto, moneda, descripcion, medio ('Tarjeta'|'Efectivo'|'Debito'), cuentaId, fecha, categoria, subcategoria, alertas, notas }
 * Devuelve { programado: bool, id, idMov, texto }
 */
function registrarOProgramar_(d) {
  const fecha = d.fecha ? fecha_(d.fecha) : null;
  const esGasto = esGasto_(d.tipo);
  if (d.programar && !fecha) throw new Error('Para programar indica la fecha (y hora) en que ocurrirá.');
  if (fecha && (d.programar || esFuturo_(fecha))) {
    const id = crearProgramado({ tipo: d.tipo, monto: d.monto, moneda: d.moneda, medio: d.medio, cuentaId: d.cuentaId, descripcion: d.descripcion, categoria: d.categoria, subcategoria: d.subcategoria, fecha: d.fecha, alertas: d.alertas, notas: d.notas });
    return { programado: true, id: id, idMov: '' };
  }
  if (esGasto) {
    const medio = ['Tarjeta', 'Efectivo', 'Debito'].indexOf(d.medio) !== -1 ? d.medio : medioGastoPorDefecto_(d.moneda);
    if (medio === 'Tarjeta') {
      const r = registrarConsumoTarjetaCuotas_({ tarjetaId: d.tarjetaId, fecha: d.fecha, tipo: d.tipo, monto: d.monto, moneda: d.moneda, descripcion: d.descripcion, categoria: d.categoria, subcategoria: d.subcategoria, notas: d.notas, cuotas: d.cuotas, interes: d.interes, tipoInteres: d.tipoInteres });
      return { programado: false, id: r.idTarjeta, idMov: r.idMov, cuotas: r.cuotas, cuota: r.cuota, totalIntereses: r.totalIntereses };
    }
    const cuentaId = d.cuentaId !== undefined && d.cuentaId !== '' ? d.cuentaId : cuentaPorDefecto_(d.moneda, medio);
    const id = registrarMovimiento({ fecha: d.fecha, tipo: d.tipo, monto: d.monto, moneda: d.moneda, cuentaId: cuentaId, descripcion: d.descripcion, categoria: d.categoria, subcategoria: d.subcategoria, notas: d.notas });
    return { programado: false, id: id, idMov: id };
  }
  const medioI = d.medio === 'Efectivo' ? 'Efectivo' : (d.medio === 'Debito' ? 'Debito' : medioIngresoPorDefecto_(d.moneda));
  const cuentaI = d.cuentaId !== undefined && d.cuentaId !== '' ? d.cuentaId : cuentaPorDefecto_(d.moneda, medioI);
  const idI = registrarMovimiento({ fecha: d.fecha, tipo: d.tipo, monto: d.monto, moneda: d.moneda, cuentaId: cuentaI, descripcion: d.descripcion, categoria: d.categoria, subcategoria: d.subcategoria, notas: d.notas, categoriaSUNAT: d.sunat });
  return { programado: false, id: idI, idMov: idI };
}

function cuentaPorNombre_(txt, moneda) {
  const q = normalizarTxt_(txt);
  if (!q) return null;
  const l = obtenerCuentasActivas_();
  const cand = l.filter(function (c) { const b = normalizarTxt_(c.banco); return b === q || b.indexOf(q) !== -1 || q.indexOf(b) !== -1; });
  return cand.filter(function (c) { return c.moneda === moneda; })[0] || cand[0] || null;
}

/** Devuelve el nombre real de la categoría (sin importar tildes/mayúsculas) o '' si no existe */
function categoriaCanon_(cat, esGastoTipo) {
  const q = normalizarTxt_(cat);
  if (!q) return '';
  const m = categoriasPorTipo_()[esGastoTipo ? 'Gasto' : 'Ingreso'];
  return Object.keys(m).filter(function (c) { return normalizarTxt_(c) === q; })[0] || '';
}
function subcategoriaCanon_(cat, sub, esGastoTipo) {
  const q = normalizarTxt_(sub);
  if (!q || !cat) return '';
  const m = categoriasPorTipo_()[esGastoTipo ? 'Gasto' : 'Ingreso'];
  return (m[cat] || []).filter(function (x) { return normalizarTxt_(x) === q; })[0] || '';
}

/** Lista fija (alfabética) de categorías: los botones de Telegram apuntan a su posición aquí */
function categoriasOrdenadas_(esGastoTipo) {
  return Object.keys(categoriasPorTipo_()[esGastoTipo ? 'Gasto' : 'Ingreso']).filter(function (c) { return c !== 'Intereses y comisiones'; }).sort();
}

/** Categorías más usadas (para ofrecerlas como botones cuando la IA no logró clasificar) */
function categoriasFrecuentes_(esGastoTipo, n) {
  const cuenta = {};
  leerMovimientos_().slice(-300).forEach(function (m) {
    if (esGastoTipo ? !esGasto_(m.tipo) : !esIngreso_(m.tipo)) return;
    const c = String(m.categoria || '').trim();
    if (c && c !== 'Intereses y comisiones') cuenta[c] = (cuenta[c] || 0) + 1;
  });
  const lista = Object.keys(cuenta).sort(function (a, b) { return cuenta[b] - cuenta[a]; });
  Object.keys(categoriasPorTipo_()[esGastoTipo ? 'Gasto' : 'Ingreso']).forEach(function (c) { if (lista.indexOf(c) === -1 && c !== 'Intereses y comisiones') lista.push(c); });
  return lista.slice(0, n);
}

/** Ejecuta las acciones que devolvió la IA. Devuelve { lineas: [texto], deshacer: [{k,id}], sinCategoria: [{id, gasto}] } */
function ejecutarAccionesIA_(acciones) {
  const lineas = [], deshacer = [], sinCategoria = [];
  acciones.forEach(function (a) {
    try {
      if (a.accion === 'consulta') {
        if (a.tema === 'saldo') lineas.push(construirMensajeSaldos_());
        else if (a.tema === 'metas') lineas.push(construirMensajeMetas_());
        else if (a.tema === 'tarjeta') lineas.push(construirMensajeTarjeta_());
        else lineas.push(construirMensajeMes_());
        return;
      }
      if (a.accion !== 'registrar') return;
      const gasto = a.tipo === 'gasto';
      const tipo = gasto ? (CLASES_GASTO_[a.clase] || 'Gasto Personal') : (CLASES_INGRESO_[a.clase] || 'Ingreso Activo');
      let medio = a.medio === 'tarjeta' ? 'Tarjeta' : (a.medio === 'efectivo' ? 'Efectivo' : (a.medio === 'debito' ? 'Debito' : ''));
      const monto = validarMonto_(a.monto);
      const moneda = validarMoneda_(a.moneda || 'PEN');
      let tj = a.tarjeta ? tarjetaPorNombre_(a.tarjeta) : null;
      let cuenta = null;
      if (a.cuenta && !tj) {
        tj = tarjetaPorNombre_(a.cuenta);
        if (!tj) { cuenta = cuentaPorNombre_(a.cuenta, moneda); if (cuenta) { medio = cuenta.tipo === 'Efectivo' ? 'Efectivo' : 'Debito'; } }
      }
      const desc = String(a.descripcion || '').trim();
      // categoría: la que eligió la IA (si existe en tu lista) o, si no, tus reglas guardadas
      let cat = categoriaCanon_(a.categoria, gasto), sub = cat ? subcategoriaCanon_(cat, a.subcategoria, gasto) : '';
      if (!cat) { const auto = autoCategorizar_(desc); cat = auto.categoria || ''; sub = auto.subcategoria || ''; }
      const r = registrarOProgramar_({ tipo: tipo, monto: monto, moneda: moneda, descripcion: desc, medio: medio || (tj ? 'Tarjeta' : ''), tarjetaId: tj ? tj.id : undefined, cuentaId: cuenta ? cuenta.id : undefined, fecha: a.fecha || null, categoria: cat, subcategoria: sub, notas: a.notas || undefined });
      deshacer.push({ k: r.programado ? 'prog' : 'mov', id: r.programado ? r.id : r.idMov });
      const medioTxt = gasto ? ' · ' + (tj ? 'tarjeta ' + escaparMd_(tj.banco) : (cuenta ? escaparMd_(cuenta.banco) : (medio === 'Efectivo' ? 'efectivo' : (medio === 'Debito' ? 'débito' : (medio === 'Tarjeta' ? 'tarjeta' : medioGastoPorDefecto_(moneda).toLowerCase()))))) : (cuenta ? ' · ' + escaparMd_(cuenta.banco) : '');
      let linea = (r.programado ? '📅 Programado' : '✅ Registrado') + ': ' + (gasto ? 'gasto' : 'ingreso') + ' de ' + formatoMonto_(monto, moneda) +
        ' — "' + escaparMd_(desc) + '"' + (cat ? ' (' + escaparMd_(cat) + (sub ? ' / ' + escaparMd_(sub) : '') + ')' : '') +
        (r.programado ? ' para el ' + escaparMd_(String(a.fecha).replace('T', ' ')) : medioTxt);
      // dividir con otras personas
      if (a.dividir && a.dividir.personas && a.dividir.personas.length) {
        if (!gasto || r.programado || !r.idMov) linea += '\n⚠️ Solo se pueden dividir gastos que ya ocurrieron; lo registré sin dividir.';
        else {
          try {
            const d = dividirMovimientoInterno_(r.idMov, { modo: a.dividir.modo, personas: a.dividir.personas });
            linea += '\n➗ Dividido: tu parte ' + formatoMonto_(d.yo, moneda) + ' · ' + d.personas.map(function (x) { return escaparMd_(x.nombre) + ' ' + formatoMonto_(x.monto, moneda); }).join(', ') + ' (por cobrar en Deudas)';
          } catch (e2) { linea += '\n⚠️ Lo registré completo pero no pude dividirlo: ' + escaparMd_(e2.message); }
        }
      }
      lineas.push(linea);
      if (!cat && !r.programado && r.idMov) sinCategoria.push({ id: r.idMov, gasto: gasto });
    } catch (e) {
      lineas.push('❌ No pude registrar "' + escaparMd_(a.descripcion || '') + '": ' + escaparMd_(e.message));
    }
  });
  return { lineas: lineas, deshacer: deshacer, sinCategoria: sinCategoria };
}

function construirMensajeMes_() {
  const movs = leerMovimientos_();
  const r = resumenPeriodo_('Mensual', new Date(), movs, obtenerUltimaTasaReferencial_());
  let t = '📊 *' + nombreMes_(new Date()) + '* (equiv. en soles)\nIngresos: ' + formatoMonto_(r.ingresos, 'PEN') + '\nEgresos: ' + formatoMonto_(r.egresos, 'PEN') + '\nGanancia: ' + formatoMonto_(r.ganancia, 'PEN');
  r.categorias.slice(0, 5).forEach(function (c) { t += '\n• ' + escaparMd_(c.categoria) + ': ' + formatoMonto_(c.monto, 'PEN') + ' (' + c.porcentaje.toFixed(0) + '%)'; });
  return t;
}

// Conversación pendiente: si al bot le falta un dato, lo pregunta y recuerda tu mensaje 20 minutos
function guardarPendienteTg_(p) { try { CacheService.getScriptCache().put('TG_PENDIENTE', JSON.stringify(p), 1200); } catch (e) { /* sin caché */ } }
function leerPendienteTg_() { try { const t = CacheService.getScriptCache().get('TG_PENDIENTE'); return t ? JSON.parse(t) : null; } catch (e) { return null; } }
function borrarPendienteTg_() { try { CacheService.getScriptCache().remove('TG_PENDIENTE'); } catch (e) { /* ignorar */ } }

/** Lo mínimo que hay que saber de un movimiento; devuelve la pregunta que falta o '' */
function preguntaFaltante_(acciones) {
  for (let i = 0; i < acciones.length; i++) {
    const a = acciones[i];
    if (a.accion !== 'registrar') continue;
    const que = a.descripcion ? ' "' + a.descripcion + '"' : '';
    if (a.tipo !== 'gasto' && a.tipo !== 'ingreso') return '¿Fue un gasto o un ingreso' + (a.monto ? ' de ' + a.monto : '') + que + '?';
    if (!(Number(a.monto) > 0)) return '¿De cuánto fue' + (que ? ' el ' + (a.tipo === 'gasto' ? 'gasto' : 'ingreso') + que : '') + '?';
    if (!String(a.descripcion || '').trim() && !a.categoria) return '¿En qué fue ese ' + (a.tipo === 'gasto' ? 'gasto' : 'ingreso') + ' de ' + a.monto + '?';
    if (a.dividir && a.dividir.personas && !a.dividir.personas.length) return '¿Con quién lo dividiste? Dime los nombres.';
  }
  return '';
}

function procesarMensajeLibreTelegram_(texto) {
  if (!iaProveedor_()) {
    return enviarMensajeTelegram('Para que entienda mensajes libres necesito una IA: agrega ANTHROPIC_API_KEY (o GEMINI_API_KEY) en las propiedades del script. Mientras tanto usa /gasto, /efectivo, /ingreso o escribe /ayuda.');
  }
  const pend = leerPendienteTg_();
  if (pend && /^\s*(cancela|cancelar|olvídalo|olvidalo|no importa|déjalo|dejalo)\s*[.!]*\s*$/i.test(texto)) {
    borrarPendienteTg_();
    return enviarMensajeTelegram('Listo, lo olvidé. No registré nada.');
  }
  let j, aviso = '';
  try { j = interpretarMensajeIA_(texto, pend && pend.ronda < 3 ? pend : null); }
  catch (e) {
    j = interpretarLocal_(texto);
    if (!j) return enviarMensajeTelegram('❌ La IA no pudo leer tu mensaje: ' + escaparMd_(e.message) + '\nPuedes escribirlo simple ("gasté 5 soles en huevos") o usar /gasto 5 PEN huevos.');
    aviso = '\n_(La IA no respondió; lo entendí en modo básico. Revísalo.)_';
  }
  const acciones = (j.acciones || []).slice(0, 8);
  if (!acciones.length) {
    if (pend) borrarPendienteTg_();
    return enviarMensajeTelegram('No entendí qué quieres registrar. Prueba: "gasté 5 soles en huevos" o "me pagaron 600 soles del alquiler". /ayuda para ver comandos.');
  }

  // ¿falta algo imprescindible? se pregunta y no se registra nada todavía
  const pregunta = preguntaFaltante_(acciones) || (j.pregunta ? String(j.pregunta) : '');
  if (pregunta && !j.local) {
    const ronda = pend ? (pend.ronda || 0) + 1 : 1;
    if (ronda <= 3) {
      guardarPendienteTg_({ original: pend ? pend.original : texto, json: { acciones: acciones }, pregunta: pregunta, ronda: ronda });
      return enviarMensajeTelegram('🤔 ' + escaparMd_(pregunta) + '\n_(Responde con el dato, o escribe "cancela")_');
    }
  }
  borrarPendienteTg_();

  // si ya se registró solo desde tu correo, no lo duplica: te pregunta
  const nuevas = [];
  acciones.forEach(function (a) {
    let dup = null;
    try { if (a.accion === 'registrar' && Number(a.monto) > 0) dup = posibleDuplicado_(Number(a.monto), a.moneda || 'PEN', a.fecha ? fecha_(a.fecha) : new Date(), a.tipo === 'gasto', true); } catch (e) { dup = null; }
    if (!dup) { nuevas.push(a); return; }
    const k = guardarAccionPendiente_(a);
    enviarMensajeTelegram('📧 Eso ya lo registré desde tu correo: "' + escaparMd_(dup.descripcion) + '" del ' + fechaISO_(dup.fecha) + ' por ' + formatoMonto_(Number(a.monto), a.moneda || 'PEN') + '. ¿Es el mismo?',
      [[{ text: '✅ Es el mismo', data: 'pend:no:' + k }, { text: '➕ Registrar aparte', data: 'pend:si:' + k }]]);
  });
  if (!nuevas.length) return;

  const r = ejecutarAccionesIA_(nuevas);
  const lote = r.deshacer.length ? guardarDeshacer_(r.deshacer) : '';
  enviarMensajeTelegram(r.lineas.join('\n') + aviso, lote ? [[{ text: '↩️ Deshacer', data: 'undo:' + lote }]] : undefined);
  // lo que la IA no pudo clasificar: te lo pregunta con botones
  r.sinCategoria.slice(0, 2).forEach(function (x) {
    const cats = categoriasFrecuentes_(x.gasto, 8);
    if (!cats.length) return;
    const filas = [];
    const orden = categoriasOrdenadas_(x.gasto);
    cats.forEach(function (c, i) { if (i % 2 === 0) filas.push([]); filas[filas.length - 1].push({ text: c, data: 'cat:' + x.id + ':' + (x.gasto ? 'g' : 'i') + ':' + orden.indexOf(c) }); });
    enviarMensajeTelegram('🏷️ No supe en qué categoría poner ese ' + (x.gasto ? 'gasto' : 'ingreso') + '. ¿Cuál es?', filas);
  });
}

function procesarAudioTelegram_(fileId) {
  if (!iaProveedor_() || !prop_('GEMINI_API_KEY')) {
    return enviarMensajeTelegram('Para entender audios agrega GEMINI_API_KEY en las propiedades del script (gratis en aistudio.google.com/apikey).');
  }
  const info = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/getFile?file_id=' + encodeURIComponent(fileId), { muteHttpExceptions: true }).getContentText());
  if (!info.ok) throw new Error('No pude descargar el audio.');
  const blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + tgToken_() + '/' + info.result.file_path).getBlob();
  const texto = transcribirAudio_(blob);
  if (!texto) return enviarMensajeTelegram('No logré escuchar nada en el audio.');
  enviarMensajeTelegram('🎙 Entendí: _' + escaparMd_(texto) + '_');
  procesarMensajeLibreTelegram_(texto);
}

/** Foto de un recibo, boleta, voucher o captura de Yape/transferencia → la IA la lee y sigue el mismo camino que un mensaje */
function procesarFotoTelegram_(fileId, caption) {
  if (!prop_('GEMINI_API_KEY')) {
    return enviarMensajeTelegram('Para leer fotos de recibos agrega tu clave de Gemini (Más → Inteligencia artificial en la app, o GEMINI_API_KEY en las propiedades del script).');
  }
  const info = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + tgToken_() + '/getFile?file_id=' + encodeURIComponent(fileId), { muteHttpExceptions: true }).getContentText());
  if (!info.ok) throw new Error('No pude descargar la foto.');
  const blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + tgToken_() + '/' + info.result.file_path).getBlob();
  const texto = leerReciboFoto_(blob, caption);
  if (!texto || /^NO_ES_RECIBO/i.test(texto)) return enviarMensajeTelegram('🧾 No encontré un pago en esa imagen. Mándame la foto del recibo, boleta o captura del Yape/transferencia' + (caption ? '' : ', o escríbeme el gasto') + '.');
  enviarMensajeTelegram('🧾 Leí: _' + escaparMd_(texto) + '_');
  procesarMensajeLibreTelegram_(texto);
}

/** Devuelve una frase como si la escribiera el usuario ("gasté 45.90 soles en Plaza Vea con tarjeta BCP el 2026-11-11 ...") */
function leerReciboFoto_(blob, caption) {
  let mime = blob.getContentType() || 'image/jpeg';
  if (mime.indexOf('image/') !== 0) mime = 'image/jpeg';
  const system = 'Lees fotos de recibos, boletas, facturas, vouchers de tarjeta y capturas de Yape, Plin o transferencias de un usuario en Perú. ' +
    'Responde SOLO una frase en español, en primera persona, como si el usuario te contara el movimiento, con: si fue gasto o ingreso, el TOTAL pagado (no subtotales ni IGV), ' +
    'la moneda (S/ = soles), el comercio o la persona, la fecha que aparece (AAAA-MM-DD), y el medio de pago si se ve (tarjeta y banco, efectivo, Yape...). ' +
    'Ejemplo: "Gasté 45.90 soles en Plaza Vea el 2026-11-10 con tarjeta de crédito BCP; compré abarrotes". ' +
    'Si la imagen no es un pago o no se lee el total, responde exactamente NO_ES_RECIBO.';
  const partes = [{ text: 'Lee este comprobante.' + (caption ? ' El usuario añadió: "' + caption + '"' : '') }, { inline_data: { mime_type: mime, data: Utilities.base64Encode(blob.getBytes()) } }];
  let t = String(llamarGemini_(system, partes, 400) || '').trim();
  if (caption && !/^NO_ES_RECIBO/i.test(t)) t += '. ' + caption;
  return t;
}

/** ¿Ya hay un movimiento con el mismo monto y moneda, ±2 días? Devuelve el movimiento o null */
function posibleDuplicado_(monto, moneda, fecha, esGastoTipo, soloDeCorreo) {
  const f = fecha || new Date();
  const lista = leerMovimientos_();
  for (let i = lista.length - 1; i >= 0 && i >= lista.length - 400; i--) {
    const m = lista[i];
    if (String(m.moneda) !== String(moneda) || Math.abs(Number(m.monto) - monto) > 0.009) continue;
    if (esGastoTipo ? !esGasto_(m.tipo) : !esIngreso_(m.tipo)) continue;
    const mf = m.fecha instanceof Date ? m.fecha : fecha_(m.fecha);
    if (!mf || Math.abs(mf - f) > 2.5 * 86400000) continue;
    if (soloDeCorreo && !/Desde correo/.test(String(m.notas))) continue;
    return { id: m.id, descripcion: m.descripcion || m.categoria || '', fecha: mf };
  }
  // también contra consumos de tarjeta (no siempre generan movimiento el mismo día)
  return null;
}

/** Guarda una acción para confirmarla luego con un botón de Telegram (7 días) */
function guardarAccionPendiente_(a) {
  const k = Utilities.getUuid().replace(/-/g, '').substring(0, 10);
  PropertiesService.getScriptProperties().setProperty('PEND_' + k, JSON.stringify({ a: a, t: Date.now() }));
  return k;
}
function tomarAccionPendiente_(k) {
  const P = PropertiesService.getScriptProperties();
  const v = P.getProperty('PEND_' + k);
  if (!v) return null;
  P.deleteProperty('PEND_' + k);
  try { return JSON.parse(v).a; } catch (e) { return null; }
}

/** Botones de categoría para lo que la IA no pudo clasificar */
function preguntarCategorias_(sinCategoria) {
  (sinCategoria || []).slice(0, 2).forEach(function (x) {
    const cats = categoriasFrecuentes_(x.gasto, 8);
    if (!cats.length) return;
    const filas = [], orden = categoriasOrdenadas_(x.gasto);
    cats.forEach(function (c, i) { if (i % 2 === 0) filas.push([]); filas[filas.length - 1].push({ text: c, data: 'cat:' + x.id + ':' + (x.gasto ? 'g' : 'i') + ':' + orden.indexOf(c) }); });
    enviarMensajeTelegram('🏷️ No supe en qué categoría poner ese ' + (x.gasto ? 'gasto' : 'ingreso') + '. ¿Cuál es?', filas);
  });
}

/** Botón de Telegram: pend:si:<k> registra la acción guardada; pend:no:<k> la descarta */
function resolverPendienteTelegram_(cb, si, k) {
  const a = tomarAccionPendiente_(k);
  if (!a) { responderCallbackTelegram_(cb.id, 'Ya estaba resuelto'); return; }
  if (!si) { responderCallbackTelegram_(cb.id, 'Listo, no lo dupliqué'); enviarMensajeTelegram('👍 Perfecto, no lo registré de nuevo.'); return; }
  const r = ejecutarAccionesIA_([a]);
  const lote = r.deshacer.length ? guardarDeshacer_(r.deshacer) : '';
  responderCallbackTelegram_(cb.id, 'Registrado');
  enviarMensajeTelegram(r.lineas.join('\n'), lote ? [[{ text: '↩️ Deshacer', data: 'undo:' + lote }]] : undefined);
  preguntarCategorias_(r.sinCategoria);
}

/** Pila de lotes deshacibles (los últimos 15). Cada lote = lo que se registró con un solo mensaje. */
function leerPilaDeshacer_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('DESHACER_PILA') || '[]'); } catch (e) { return []; }
}
function guardarDeshacer_(items) {
  const lote = { id: Date.now().toString(36), items: items.filter(function (x) { return x && x.id !== '' && x.id !== undefined; }) };
  if (!lote.items.length) return lote.id;
  const pila = leerPilaDeshacer_();
  pila.push(lote);
  PropertiesService.getScriptProperties().setProperty('DESHACER_PILA', JSON.stringify(pila.slice(-15)));
  return lote.id;
}
/** Deshace un lote concreto (botón ↩️) o el último (/deshacer). Devuelve false si ya no existía. */
function deshacerLote_(id) {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let lote = null;
  try {
    const pila = leerPilaDeshacer_();
    const idx = id ? pila.map(function (x) { return x.id; }).indexOf(id) : pila.length - 1;
    if (idx < 0) return false;
    lote = pila.splice(idx, 1)[0];
    props.setProperty('DESHACER_PILA', JSON.stringify(pila));
  } finally { lock.releaseLock(); }
  let n = 0;
  lote.items.forEach(function (x) {
    try { if (x.k === 'prog') eliminarProgramado(x.id); else eliminarMovimiento(x.id); n++; }
    catch (e) { enviarMensajeTelegram('⚠️ No pude deshacer #' + x.id + ': ' + escaparMd_(e.message)); }
  });
  enviarMensajeTelegram('↩️ Deshice ' + n + (n === 1 ? ' registro.' : ' registros.') + (leerPilaDeshacer_().length ? ' Quedan ' + leerPilaDeshacer_().length + ' más que puedes deshacer con /deshacer.' : ''));
  return true;
}
/** /deshacer: borra lo último que registré por Telegram; repítelo para seguir hacia atrás */
function deshacerUltimoBot_() {
  if (!deshacerLote_('')) enviarMensajeTelegram('No hay nada reciente para deshacer.');
}

/**
 * Modo básico SIN IA, para cuando Gemini no responde: entiende frases simples con un solo monto
 * ("gasté 5 soles en huevos", "me pagaron 600 del alquiler"). Devuelve null si no está seguro.
 */
function interpretarLocal_(texto) {
  const t = String(texto || '').toLowerCase().trim();
  if (/\b(ayer|anteayer|mañana|pasado mañana|lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado|domingo|pr[oó]xim[oa]|el \d{1,2} de)\b/.test(t)) return null;
  const nums = t.match(/\d+(?:[.,]\d+)?/g) || [];
  if (nums.length !== 1) return null;
  const ingreso = /(me pagaron|me pag[oó]|cobr[eé]|me depositaron|me abonaron|recib[ií]|me transfirieron|me yape|me plin|ingres[oó])/.test(t);
  const gasto = /(gast[eé]|pagu[eé]|compr[eé]|me cobraron|invert[ií]|pago de|pagar)/.test(t);
  if (ingreso === gasto) return null;
  const moneda = /(d[oó]lar|usd|\$)/.test(t) ? 'USD' : (/(euro|€|\beur\b)/.test(t) ? 'EUR' : 'PEN');
  const medio = /(tarjeta|cr[eé]dito)/.test(t) ? 'tarjeta' : (/(efectivo|cash|en mano)/.test(t) ? 'efectivo' : (/(yape|plin|transferencia|d[eé]bito|dep[oó]sito)/.test(t) ? 'debito' : null));
  let desc = '';
  const m = t.match(/\d+(?:[.,]\d+)?\s*(?:soles?|s\/\.?|d[oó]lares?|usd|euros?|€|\$)?\s*(?:de|en|por|para|del)?\s*(.*)$/);
  if (m && m[1]) desc = m[1];
  desc = desc.replace(/\b(con (la )?tarjeta|en efectivo|con yape|por yape|con plin)\b/g, '').trim();
  if (!desc) desc = t.replace(/\d+(?:[.,]\d+)?/, '').replace(/(gast[eé]|pagu[eé]|compr[eé]|me pagaron|me pag[oó]|cobr[eé]|recib[ií]|soles?|d[oó]lares?|euros?)/g, '').replace(/\s+/g, ' ').trim();
  if (!desc) desc = ingreso ? 'ingreso' : 'gasto';
  return { acciones: [{ accion: 'registrar', tipo: gasto ? 'gasto' : 'ingreso', monto: numero_(nums[0]), moneda: moneda, descripcion: desc, medio: medio, fecha: null, clase: ingreso && /(alquiler|inter[eé]s|dividendo)/.test(t) ? 'pasivo' : null, categoria: null }], no_entendi: false, local: true };
}


/* ==== 20_Configuracion.gs ==== */
/**
 * ============================================================
 * CONFIGURACIÓN DESDE LA APP — metas, presupuestos, recurrentes, categorías, ajustes, préstamos
 * ============================================================
 * Todo lo que antes se editaba a mano en el Excel ahora se puede hacer desde la app (pestaña Más).
 */

function guardarFilaPorId_(nombreHoja, id, resto, conservarUltimas) {
  const hoja = getHoja_(nombreHoja);
  if (id !== undefined && id !== null && id !== '') {
    const f = buscarFilaPorId_(nombreHoja, id);
    if (!f) throw new Error('No encontré ese registro.');
    const fila = [f.fila[0]].concat(resto);
    for (let k = 0; k < (conservarUltimas || 0); k++) fila.push(f.fila[fila.length] === undefined ? '' : f.fila[fila.length]);
    hoja.getRange(f.numeroFila, 1, 1, fila.length).setValues([fila]);
    return f.fila[0];
  }
  const nuevo = generarID_(nombreHoja);
  hoja.appendRow([nuevo].concat(resto).concat(conservarUltimas ? new Array(conservarUltimas).fill('') : []));
  return nuevo;
}

function eliminarFilaPorId_(nombreHoja, id) {
  const f = buscarFilaPorId_(nombreHoja, id);
  if (!f) throw new Error('No encontré ese registro.');
  getHoja_(nombreHoja).deleteRow(f.numeroFila);
  return true;
}

// ---------------- METAS ----------------

function api_listarMetas() {
  const datos = leerHoja_('Metas');
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    out.push({ id: datos[i][0], periodo: String(datos[i][1]), moneda: String(datos[i][2]), monto: numero_(datos[i][3]), inicio: fechaISO_(datos[i][4]), fin: fechaISO_(datos[i][5]), crecimiento: numero_(datos[i][6]) });
  }
  return out;
}

function api_guardarMeta(a) {
  const periodo = ['Diaria', 'Semanal', 'Mensual', 'Anual'].indexOf(a.periodo) !== -1 ? a.periodo : null;
  if (!periodo) throw new Error('Elige el periodo de la meta.');
  const crecimiento = numero_(a.crecimiento);
  if (crecimiento < 0 || crecimiento > 1000) throw new Error('El crecimiento debe estar entre 0 y 1000%.');
  const monto = crecimiento > 0 && !String(a.monto || '').trim() ? 0 : validarMonto_(a.monto, 'La meta');
  const inicio = a.inicio ? fecha_(a.inicio) : '', fin = a.fin ? fecha_(a.fin) : '';
  if (inicio && fin && fin < inicio) throw new Error('La fecha final es anterior a la inicial.');
  return guardarFilaPorId_('Metas', a.id, [periodo, validarMoneda_(a.moneda), monto, inicio, fin, crecimiento > 0 ? crecimiento : '']);
}
function api_eliminarMeta(a) { return eliminarFilaPorId_('Metas', a.id); }

// ---------------- PRESUPUESTOS ----------------

function api_listarPresupuestos() {
  const datos = leerHoja_('Presupuestos');
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    out.push({ id: datos[i][0], categoria: String(datos[i][1]), subcategoria: String(datos[i][2]), moneda: String(datos[i][3]), monto: numero_(datos[i][4]), mes: datos[i][5] instanceof Date ? Utilities.formatDate(datos[i][5], zona_(), 'yyyy-MM') : String(datos[i][5]) });
  }
  return out;
}

function api_guardarPresupuesto(a) {
  const categoria = String(a.categoria || '').trim();
  if (!categoria) throw new Error('Elige o escribe la categoría.');
  return guardarFilaPorId_('Presupuestos', a.id, [categoria, String(a.subcategoria || '').trim(), validarMoneda_(a.moneda), validarMonto_(a.monto, 'El límite'), String(a.mes || '')]);
}
function api_eliminarPresupuesto(a) { return eliminarFilaPorId_('Presupuestos', a.id); }

// ---------------- RECURRENTES ----------------

function api_listarRecurrentes() {
  const datos = leerHoja_('Recurrentes');
  const nombres = mapaNombresCuentas_();
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    const f = datos[i];
    out.push({ id: f[0], descripcion: String(f[1]), tipo: String(f[2]), monto: numero_(f[3]), moneda: String(f[4]), dia: numero_(f[5]), categoria: String(f[6]), subcategoria: String(f[7]),
      centro: String(f[8]), cuentaId: f[9], cuentaNombre: f[9] === '' ? 'Tarjeta de crédito' : (nombres[String(f[9])] || ''), activo: String(f[10]).trim().toUpperCase() !== 'NO', ultima: fechaISO_(f[11]), confirmar: String(f[12] || '').trim().toUpperCase() === 'SI' });
  }
  return out;
}

function api_guardarRecurrente(a) {
  const tipo = a.tipo === 'Ingreso' ? 'Ingreso' : 'Gasto';
  const dia = Math.round(numero_(a.dia));
  if (!(dia >= 1 && dia <= 31)) throw new Error('El día del mes debe ser entre 1 y 31.');
  if (!String(a.descripcion || '').trim()) throw new Error('Ponle un nombre.');
  const moneda = validarMoneda_(a.moneda);
  let cuentaId = a.cuentaId === undefined ? '' : a.cuentaId;
  if (tipo === 'Ingreso' && cuentaId === '') throw new Error('Un ingreso recurrente necesita la cuenta donde llega.');
  if (cuentaId !== '') validarCuentaMoneda_(cuentaId, moneda);
  const idRec = guardarFilaPorId_('Recurrentes', a.id, [String(a.descripcion).trim(), tipo, validarMonto_(a.monto), moneda, dia, a.categoria || '', a.subcategoria || '',
    limpiarCentroCosto_(a.centro) || 'Personal', cuentaId, a.activo === false ? 'NO' : 'SI'], 1);
  const hojaR = getHoja_('Recurrentes');
  if (hojaR.getLastColumn() < 13) hojaR.getRange(1, 13).setValue('Confirmar');
  const filaR = buscarFilaPorId_('Recurrentes', idRec);
  if (filaR) hojaR.getRange(filaR.numeroFila, 13).setValue(a.confirmar ? 'SI' : 'NO');
  // Si ya lo pagaste este mes (lo registraste hoy desde "Nuevo gasto"), que no se vuelva a cobrar hasta el próximo
  const ult = a.ultima ? fecha_(a.ultima) : null;
  if (filaR && ult) hojaR.getRange(filaR.numeroFila, 12).setValue(ult);
  return idRec;
}
function api_eliminarRecurrente(a) { return eliminarFilaPorId_('Recurrentes', a.id); }

// ---------------- CATEGORÍAS (reglas de autocategorización) ----------------

function api_listarReglas() {
  const datos = leerHoja_('ReglasCategorizacion');
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    out.push({ id: datos[i][0], palabra: String(datos[i][1]), categoria: String(datos[i][2]), subcategoria: String(datos[i][3]), centro: String(datos[i][4]), prioridad: numero_(datos[i][5]) || 1 });
  }
  return out;
}

function api_guardarRegla(a) {
  const palabra = String(a.palabra || '').trim();
  const categoria = String(a.categoria || '').trim();
  if (!palabra || !categoria) throw new Error('Escribe la palabra clave y la categoría.');
  return guardarFilaPorId_('ReglasCategorizacion', a.id, [palabra, categoria, String(a.subcategoria || '').trim(), limpiarCentroCosto_(a.centro), Math.max(1, Math.round(numero_(a.prioridad)) || 1)]);
}
function api_eliminarRegla(a) { return eliminarFilaPorId_('ReglasCategorizacion', a.id); }

// ---------------- AJUSTES (valores predeterminados) ----------------

function ajustesParaApp_() {
  const r = { medioGasto: {}, medioIngreso: {}, cuentaDefault: {}, alertasDefault: parsearAlertas_(getAjuste_('alertas.default', '1440')), notif: {}, accUmbral: numero_(getAjuste_('acc.umbral', 5)) || 5 };
  ['meta_diaria', 'meta_semanal', 'tarjeta', 'intereses', 'acciones'].forEach(function (k) { r.notif[k] = String(getAjuste_('notif.' + k, 'SI')) !== 'NO'; });
  CONFIG.MONEDAS.forEach(function (m) {
    r.medioGasto[m] = medioGastoPorDefecto_(m);
    r.medioIngreso[m] = medioIngresoPorDefecto_(m);
    r.cuentaDefault[m] = { Efectivo: String(getAjuste_('cuenta.' + m + '.Efectivo', '')), Debito: String(getAjuste_('cuenta.' + m + '.Debito', '')) };
  });
  return r;
}

function api_obtenerAjustes() { return ajustesParaApp_(); }

function api_guardarAjustes(a) {
  CONFIG.MONEDAS.forEach(function (m) {
    if (a.medioGasto && a.medioGasto[m]) setAjuste_('medio.gasto.' + m, ['Tarjeta', 'Efectivo', 'Debito'].indexOf(a.medioGasto[m]) !== -1 ? a.medioGasto[m] : 'Tarjeta');
    if (a.medioIngreso && a.medioIngreso[m]) setAjuste_('medio.ingreso.' + m, a.medioIngreso[m] === 'Efectivo' ? 'Efectivo' : 'Debito');
    if (a.cuentaDefault && a.cuentaDefault[m]) {
      ['Efectivo', 'Debito'].forEach(function (medio) {
        const v = a.cuentaDefault[m][medio];
        if (v === undefined) return;
        if (v !== '') {
          const c = validarCuentaMoneda_(v, m);
          if ((medio === 'Efectivo') !== (c.tipo === 'Efectivo')) throw new Error('La cuenta predeterminada de ' + (medio === 'Efectivo' ? 'efectivo' : 'banco') + ' en ' + m + ' no corresponde al tipo.');
        }
        setAjuste_('cuenta.' + m + '.' + medio, v);
      });
    }
  });
  if (a.alertasDefault) setAjuste_('alertas.default', parsearAlertas_(a.alertasDefault).join(','));
  if (a.notif) Object.keys(a.notif).forEach(function (k) { if (['meta_diaria', 'meta_semanal', 'tarjeta', 'intereses', 'acciones'].indexOf(k) !== -1) setAjuste_('notif.' + k, a.notif[k] ? 'SI' : 'NO'); });
  if (a.accUmbral !== undefined && a.accUmbral !== '') { const u = numero_(a.accUmbral); if (!(u >= 1 && u <= 50)) throw new Error('El umbral de alerta de acciones debe estar entre 1% y 50%.'); setAjuste_('acc.umbral', u); }
  return ajustesParaApp_();
}

// ---------------- PRÉSTAMOS Y DEUDAS ----------------

function api_prestamos() {
  const otorgados = [];
  const po = leerHoja_('PrestamosOtorgados');
  for (let i = 1; i < po.length; i++) {
    const f = po[i];
    if (f[0] === '') continue;
    otorgados.push({ id: f[0], deudor: String(f[1]), moneda: String(f[2]), monto: numero_(f[3]), tasa: numero_(f[4]), tipoTasa: String(f[5]), frecuencia: String(f[6]), inicio: fechaISO_(f[7]),
      plazo: numero_(f[8]), estado: String(f[9]), saldo: numero_(f[10]), interesMensual: interesMensualPrestamo_(f), concepto: String(f[12] || ''), cuentaOrigenId: f[11] });
  }
  const cuotas = leerHoja_('CronogramaCuotas');
  const deudas = [];
  const de = leerHoja_('Deudas');
  for (let i = 1; i < de.length; i++) {
    const f = de[i];
    if (f[0] === '') continue;
    const cs = [];
    for (let j = 1; j < cuotas.length; j++) {
      if (cuotas[j][2] === 'Recibido' && String(cuotas[j][1]) === String(f[0])) {
        cs.push({ id: cuotas[j][0], numero: numero_(cuotas[j][3]), fecha: fechaISO_(cuotas[j][4]), capital: numero_(cuotas[j][5]), interes: numero_(cuotas[j][6]), estado: String(cuotas[j][7]) });
      }
    }
    deudas.push({ id: f[0], acreedor: String(f[1]), moneda: String(f[2]), monto: numero_(f[3]), tasa: numero_(f[4]), tipoTasa: String(f[5]), inicio: fechaISO_(f[6]), plazo: numero_(f[7]), estado: String(f[8]), saldo: numero_(f[9]), cuotas: cs });
  }
  return { otorgados: otorgados, deudas: deudas };
}

function api_crearPrestamo(a) { a.deudor = canonPersona_(a.deudor); return registrarPrestamoOtorgado(a); }
function api_cobrarInteres(a) { return registrarCobroInteresPrestamo(a.id, a.monto, a.cuentaId, a.fecha); }
function api_devolverCapital(a) {
  const partes = partesCobro_(a);
  if (partes.length > 1) {
    const p = buscarFilaPorId_('PrestamosOtorgados', a.id);
    if (!p) throw new Error('Préstamo no encontrado.');
    const total = redondear_(partes.reduce(function (s, x) { return s + x.monto; }, 0));
    if (total > numero_(p.fila[10]) + 0.005) throw new Error('Entre todas las cuentas suman ' + total + ', más del saldo pendiente (' + redondear_(numero_(p.fila[10])) + ').');
  }
  partes.forEach(function (x) { registrarDevolucionCapitalPrestamo(a.id, x.monto, x.cuentaId, a.fecha); });
  return true;
}
function api_crearDeuda(a) { return registrarDeudaRecibida(a); }
function api_pagarCuota(a) { registrarPagoCuotaDeuda(a.id, a.cuentaId, a.fecha); return true; }



/* ==== 21_Mejoras.gs ==== */
/**
 * ============================================================
 * MEJORAS — tarjeta de crédito (línea, cuotas, intereses), objetivos de reparto del ingreso, categorías editables,
 * próximos recurrentes, chat con el asesor, ideas de acciones, historial de cartera y alertas de acciones
 * ============================================================
 */

function hash_(s) {
  s = String(s);
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return String(h);
}

// (La tarjeta de crédito vive ahora en 22_Tarjetas.gs: varias tarjetas, línea compartida o por moneda)

/** Compra con tarjeta en cuotas (sistema francés) y/o con intereses. d.cuotas (1 = sin cuotas), d.interes (%), d.tipoInteres ('Mensual'|'Anual') */
function registrarConsumoTarjetaCuotas_(d) {
  const n = Math.max(1, Math.min(60, Math.round(numero_(d.cuotas)) || 1));
  const tasaPct = Math.max(0, numero_(d.interes));
  if (n === 1 && !(tasaPct > 0)) return registrarConsumoTarjetaDetallado_(d);
  const P = validarMonto_(d.monto);
  validarMoneda_(d.moneda);
  const i = tasaPct > 0 ? (d.tipoInteres === 'Anual' ? Math.pow(1 + tasaPct / 100, 1 / 12) - 1 : tasaPct / 100) : 0;
  const cuota = i > 0 ? P * i / (1 - Math.pow(1 + i, -n)) : P / n;
  const fecha0 = fechaConHora_(d.fecha);
  const grupo = Date.now().toString(36);
  const base = String(d.descripcion || '').trim();
  let saldo = P, primero = null, totalInteres = 0;
  for (let k = 1; k <= n; k++) {
    const interK = i > 0 ? redondear_(saldo * i) : 0;
    const capK = k === n ? redondear_(saldo) : redondear_(i > 0 ? cuota - interK : P / n);
    saldo = redondear_(saldo - capK);
    const fechaK = sumarMeses_(fecha0, k - 1);
    const nota = 'Cuotas #' + grupo + ' (' + k + '/' + n + ')' + (d.notas ? ' · ' + d.notas : '');
    const r = registrarConsumoTarjetaDetallado_({ tarjetaId: d.tarjetaId, fecha: fechaK, tipo: d.tipo, monto: capK, moneda: d.moneda, descripcion: base + ' (cuota ' + k + '/' + n + ')', categoria: d.categoria, subcategoria: d.subcategoria, centroCosto: d.centroCosto, notas: nota });
    if (!primero) primero = r;
    if (interK > 0) {
      totalInteres += interK;
      registrarConsumoTarjetaDetallado_({ tarjetaId: d.tarjetaId, fecha: fechaK, tipo: 'Gasto Personal', monto: interK, moneda: d.moneda, descripcion: 'Interés cuota ' + k + '/' + n + ' — ' + base, categoria: 'Intereses y comisiones', subcategoria: 'Tarjeta', centroCosto: 'Personal', notas: nota });
    }
  }
  return { idTarjeta: primero.idTarjeta, idMov: primero.idMov, cuotas: n, cuota: redondear_(cuota), totalIntereses: redondear_(totalInteres) };
}

/** Borra todas las cuotas de una compra (si alguna ya se pagó, no se puede) */
function eliminarGrupoCuotas_(grupo) {
  const movs = leerMovimientos_().filter(function (x) { return x.notas.indexOf('Cuotas #' + grupo + ' ') !== -1; });
  const tarjetas = movs.map(function (x) { return extraerRef_(x.notas, /Ref\. TarjetaCredito #(\d+)/); });
  tarjetas.forEach(function (tid) {
    const t = tid ? buscarFilaPorId_('TarjetaCredito', tid) : null;
    if (t && String(t.fila[7]).trim() === 'Pagado') throw new Error('Algunas cuotas de esta compra ya se pagaron con un estado de cuenta; no se puede borrar toda la compra.');
  });
  movs.forEach(function (x) { borrarFilaMovimiento_(x.id); });
  tarjetas.forEach(function (tid) { const t = tid ? buscarFilaPorId_('TarjetaCredito', tid) : null; if (t) getHoja_('TarjetaCredito').deleteRow(t.numeroFila); });
  return movs.length;
}

// ---------------------------------------------------------------
// OBJETIVOS: cómo quieres repartir tus ingresos (ahorro / inversión / gasto personal)
// ---------------------------------------------------------------

function objetivosIngreso_() {
  return { ahorro: numero_(getAjuste_('obj.ahorro', 20)), inversion: numero_(getAjuste_('obj.inversion', 10)), gasto: numero_(getAjuste_('obj.gasto', 70)) };
}

function api_objetivos() { return objetivosIngreso_(); }

function api_guardarObjetivos(a) {
  ['ahorro', 'inversion', 'gasto'].forEach(function (k) {
    const v = numero_(a[k]);
    if (!(v >= 0 && v <= 100)) throw new Error('Cada porcentaje debe estar entre 0 y 100.');
    setAjuste_('obj.' + k, v);
  });
  return objetivosIngreso_();
}

function distribucionIngresos_(periodo, fechaRef, movs, tasa) {
  const rango = obtenerRangoFecha_(periodo, fechaRef);
  let ingresos = 0, personal = 0, gastoInv = 0, bolsaNeta = 0;
  movs.forEach(function (m) {
    if (m.efectiva < rango.inicio || m.efectiva > rango.fin) return;
    const eq = convertirAPEN_(m.monto, m.moneda, tasa);
    if (esIngreso_(m.tipo)) ingresos += eq;
    else if (m.tipo === 'Gasto Personal' || m.tipo === 'Gasto Fijo') personal += eq;
    else if (m.tipo === 'Gasto Inversion') gastoInv += eq;
    else if (m.tipo === 'Transferencia' && m.categoria === 'Inversiones') bolsaNeta += -eq;
  });
  const bolsa = Math.max(0, bolsaNeta);
  const inversion = gastoInv + bolsa;
  // Ahorro = ingresos - todos los gastos (igual que "Ganancia" y la tasa de ahorro del Resumen).
  // Las compras de bolsa no son gasto: salen de ese ahorro. Lo que queda después de invertir es el ahorro en efectivo.
  const ahorro = ingresos - personal - gastoInv;
  const efectivo = ahorro - bolsa;
  const pc = function (x) { return ingresos > 0 ? x / ingresos * 100 : 0; };
  return {
    ingresos: redondear_(ingresos), personal: redondear_(personal), inversion: redondear_(inversion), ahorro: redondear_(ahorro),
    efectivo: redondear_(efectivo), pctPersonal: pc(personal), pctInversion: pc(inversion), pctAhorro: pc(ahorro),
    pctEfectivo: pc(efectivo), objetivos: objetivosIngreso_()
  };
}

// ---------------------------------------------------------------
// CATEGORÍAS Y SUBCATEGORÍAS (editables desde la app)
// ---------------------------------------------------------------

function crearHojaCategorias(ss) {
  crearHojaBase_(ss, 'Categorias', ['Categoria', 'Subcategoria', 'Tipo']);
}

var CATEGORIAS_BASE_ = {
  Gasto: ['Alimentacion', 'Transporte', 'Salud', 'Servicios', 'Entretenimiento', 'Educacion', 'Ropa', 'Hogar', 'Viajes', 'Regalos', 'Intereses y comisiones'],
  Ingreso: ['Sueldo', 'Alquileres', 'Ventas', 'Servicios', 'Dividendos', 'Intereses', 'Otros ingresos']
};

function tipoCat_(t) { return String(t || '').trim() === 'Ingreso' ? 'Ingreso' : 'Gasto'; }
function tipoDeMov_(tipoMov) { return esIngreso_(String(tipoMov || '')) ? 'Ingreso' : 'Gasto'; }

/** Qué categorías/subcategorías ya usas, separadas por tipo: { Gasto: {cat:{sub:true}}, Ingreso: {...} } */
function categoriasDerivadas_() {
  const out = { Gasto: {}, Ingreso: {} };
  const add = function (t, c, s) { c = String(c || '').trim(); s = String(s || '').trim(); if (!c) return; out[t][c] = out[t][c] || {}; if (s) out[t][c][s] = true; };
  const reglas = leerHoja_('ReglasCategorizacion');
  for (let i = 1; i < reglas.length; i++) add('Gasto', reglas[i][2], reglas[i][3]);
  leerMovimientos_().forEach(function (m) {
    if (esIngreso_(m.tipo)) add('Ingreso', m.categoria, m.subcategoria);
    else if (esGasto_(m.tipo)) add('Gasto', m.categoria, m.subcategoria);
  });
  Object.keys(CATEGORIAS_BASE_).forEach(function (t) { CATEGORIAS_BASE_[t].forEach(function (c) { add(t, c, ''); }); });
  return out;
}

function escribirFilasCategorias_(filas) {
  const hoja = getHoja_('Categorias');
  hoja.clear();
  hoja.getRange(1, 1, 1, 3).setValues([['Categoria', 'Subcategoria', 'Tipo']]).setFontWeight('bold').setBackground('#1c1c1e').setFontColor('#ffffff');
  if (filas.length) hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  hoja.setFrozenRows(1);
}

/**
 * Lista de categorías por tipo. Primera vez: se llena con lo que ya usas. Si la hoja es de la versión anterior (sin columna Tipo),
 * cada categoría pasa al tipo en que la has usado (si la usaste en ambos, queda en ambos).
 */
function categoriasPorTipo_() {
  const ss = getSS_();
  if (!ss.getSheetByName('Categorias')) crearHojaCategorias(ss);
  let datos = leerHoja_('Categorias');
  const sinTipo = datos.length > 0 && String(datos[0][2] || '').trim() !== 'Tipo';
  if (datos.length < 2 || sinTipo) {
    const usado = categoriasDerivadas_();
    const filas = [];
    if (datos.length < 2) {
      ['Gasto', 'Ingreso'].forEach(function (t) {
        Object.keys(usado[t]).sort().forEach(function (c) {
          filas.push([c, '', t]);
          Object.keys(usado[t][c]).sort().forEach(function (sb) { filas.push([c, sb, t]); });
        });
      });
    } else {
      for (let i = 1; i < datos.length; i++) {
        const c = String(datos[i][0] || '').trim(), sb = String(datos[i][1] || '').trim();
        if (!c) continue;
        let tipos = ['Gasto', 'Ingreso'].filter(function (t) { return usado[t][c]; });
        const base = ['Gasto', 'Ingreso'].filter(function (t) { return CATEGORIAS_BASE_[t].indexOf(c) !== -1; });
        if (!tipos.length) tipos = base.length ? base : ['Gasto'];
        // una categoría de gastos con "Alquileres/Sueldo/Ventas..." solo cuenta como ingreso si realmente la usaste en ingresos
        tipos.forEach(function (t) { filas.push([c, sb, t]); });
      }
    }
    escribirFilasCategorias_(filas);
    datos = leerHoja_('Categorias');
  }
  const out = { Gasto: {}, Ingreso: {} };
  for (let i = 1; i < datos.length; i++) {
    const c = String(datos[i][0] || '').trim(), sb = String(datos[i][1] || '').trim(), t = tipoCat_(datos[i][2]);
    if (!c) continue;
    out[t][c] = out[t][c] || [];
    if (sb && out[t][c].indexOf(sb) === -1) out[t][c].push(sb);
  }
  return out;
}

/** { categoria: [subcategorías] } del tipo pedido ('Gasto' | 'Ingreso') */
function leerCategorias_(tipo) { return categoriasPorTipo_()[tipoCat_(tipo)]; }

function catalogoCategorias_() {
  const p = categoriasPorTipo_();
  const lista = function (m, ing) {
    return Object.keys(m).sort().map(function (c) {
      const o = { categoria: c, subcategorias: m[c].slice().sort() };
      if (ing) o.naturaleza = naturalezaCategoria_(c);
      return o;
    });
  };
  return { Gasto: lista(p.Gasto, false), Ingreso: lista(p.Ingreso, true) };
}

// ---------------------------------------------------------------
// NATURALEZA DE LOS INGRESOS (activo / pasivo): la decide la categoría
// ---------------------------------------------------------------

function naturalezaDefault_(cat) {
  const c = sinTildes_(String(cat || '')).toLowerCase();
  return /alquiler|dividend|interes|renta|rendimiento|regal/.test(c) ? 'Pasivo' : 'Activo';
}
function naturalezaCategoria_(cat) {
  cat = String(cat || '').trim();
  const v = String(getAjuste_('nat.' + cat, '')).trim();
  return v === 'Pasivo' || v === 'Activo' ? v : naturalezaDefault_(cat);
}
/** Si es un ingreso con categoría, el tipo (Activo/Pasivo) sale de la categoría; si no, queda como está */
function tipoIngresoPorCategoria_(tipo, categoria) {
  if (!esIngreso_(String(tipo || '')) || !String(categoria || '').trim()) return tipo;
  return 'Ingreso ' + naturalezaCategoria_(categoria);
}

/** Cambia la naturaleza de una categoría de ingresos y (opcional) la aplica a lo ya registrado */
function api_guardarNaturaleza(a) {
  const cat = String(a.categoria || '').trim();
  const nat = String(a.naturaleza || '').trim() === 'Pasivo' ? 'Pasivo' : 'Activo';
  if (!cat) throw new Error('Falta la categoría.');
  setAjuste_('nat.' + cat, nat);
  let n = 0;
  if (a.aplicar !== false) {
    const nuevo = 'Ingreso ' + nat;
    const op = function (hoja, iTipo, iCat) {
      const h = getSS_().getSheetByName(hoja);
      if (!h) return;
      const d = h.getDataRange().getValues();
      if (d.length < 2) return;
      const col = [];
      let cambios = 0;
      for (let i = 1; i < d.length; i++) {
        let t = d[i][iTipo];
        if (esIngreso_(String(t)) && String(d[i][iCat]).trim() === cat && t !== nuevo) { t = nuevo; cambios++; }
        col.push([t]);
      }
      if (cambios) { h.getRange(2, iTipo + 1, col.length, 1).setValues(col); n += cambios; }
    };
    conLock_(function () { op('Movimientos', M.TIPO, M.CAT); op('Programados', P.TIPO, P.CAT); });
  }
  return { categoria: cat, naturaleza: nat, actualizados: n };
}

/** Si una regla o la IA usa una categoría que no está en tu lista, se agrega sola (en el tipo del movimiento) */
function registrarCategoriaSiNueva_(cat, sub, tipoMov) {
  cat = String(cat || '').trim(); sub = String(sub || '').trim();
  if (!cat) return;
  try {
    const t = tipoDeMov_(tipoMov);
    const m = leerCategorias_(t);
    const hoja = getHoja_('Categorias');
    if (!m[cat]) hoja.appendRow([cat, '', t]);
    if (sub && (!m[cat] || m[cat].indexOf(sub) === -1)) hoja.appendRow([cat, sub, t]);
  } catch (e) { /* la lista es solo una ayuda */ }
}

function api_listarCategorias() { return catalogoCategorias_(); }

/** Cambia el valor de una columna en toda una hoja. filtroCol/filtroVal y pred(fila) son opcionales */
function reemplazarEnHoja_(nombreHoja, col, viejo, nuevo, filtroCol, filtroVal, pred) {
  const hoja = getSS_().getSheetByName(nombreHoja);
  if (!hoja) return 0;
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return 0;
  let n = 0;
  const salida = [];
  for (let i = 1; i < datos.length; i++) {
    let v = datos[i][col];
    if (String(v).trim() === viejo && (filtroCol === undefined || filtroCol === null || String(datos[i][filtroCol]).trim() === filtroVal) && (!pred || pred(datos[i]))) { v = nuevo; n++; }
    salida.push([v]);
  }
  if (n) hoja.getRange(2, col + 1, salida.length, 1).setValues(salida);
  return n;
}

function api_guardarCategoria(a) {
  const tipo = tipoCat_(a.tipo);
  const nuevo = String(a.nuevo || '').trim();
  if (!nuevo) throw new Error('Escribe el nombre de la categoría.');
  const mapa = leerCategorias_(tipo);
  const viejo = String(a.viejo || '').trim();
  if (!viejo) {
    if (mapa[nuevo]) throw new Error('Esa categoría ya existe en ' + (tipo === 'Gasto' ? 'gastos' : 'ingresos') + '.');
    getHoja_('Categorias').appendRow([nuevo, '', tipo]);
    return true;
  }
  if (viejo === nuevo) return true;
  if (mapa[nuevo]) throw new Error('Ya existe una categoría con ese nombre.');
  if (tipo === 'Ingreso') setAjuste_('nat.' + nuevo, naturalezaCategoria_(viejo));
  const hojaC = getHoja_('Categorias');
  const dc = hojaC.getDataRange().getValues();
  for (let i = 1; i < dc.length; i++) if (String(dc[i][0]).trim() === viejo && tipoCat_(dc[i][2]) === tipo) hojaC.getRange(i + 1, 1).setValue(nuevo);
  const esDelTipo = function (idx) { return function (f) { return tipo === 'Ingreso' ? esIngreso_(String(f[idx])) : esGasto_(String(f[idx])); }; };
  reemplazarEnHoja_('Movimientos', M.CAT, viejo, nuevo, null, undefined, esDelTipo(M.TIPO));
  reemplazarEnHoja_('Programados', P.CAT, viejo, nuevo, null, undefined, esDelTipo(P.TIPO));
  reemplazarEnHoja_('Recurrentes', 6, viejo, nuevo, 2, tipo);
  if (tipo === 'Gasto') {
    reemplazarEnHoja_('ReglasCategorizacion', 2, viejo, nuevo);
    reemplazarEnHoja_('Presupuestos', 1, viejo, nuevo);
    reemplazarEnHoja_('TarjetaCredito', 11, viejo, nuevo);
  }
  return true;
}

function api_eliminarCategoria(a) {
  const c = String(a.categoria || '').trim(), tipo = tipoCat_(a.tipo);
  const hoja = getHoja_('Categorias');
  const datos = hoja.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) if (String(datos[i][0]).trim() === c && tipoCat_(datos[i][2]) === tipo) hoja.deleteRow(i + 1);
  return true;
}

function api_guardarSubcategoria(a) {
  const tipo = tipoCat_(a.tipo);
  const cat = String(a.categoria || '').trim(), nuevo = String(a.nuevo || '').trim(), viejo = String(a.viejo || '').trim();
  if (!cat || !nuevo) throw new Error('Escribe el nombre de la subcategoría.');
  const mapa = leerCategorias_(tipo);
  if (!mapa[cat]) throw new Error('La categoría no existe.');
  if (!viejo) {
    if (mapa[cat].indexOf(nuevo) !== -1) throw new Error('Esa subcategoría ya existe.');
    getHoja_('Categorias').appendRow([cat, nuevo, tipo]);
    return true;
  }
  if (viejo === nuevo) return true;
  if (mapa[cat].indexOf(nuevo) !== -1) throw new Error('Ya existe una subcategoría con ese nombre.');
  const hojaC = getHoja_('Categorias');
  const dc = hojaC.getDataRange().getValues();
  for (let i = 1; i < dc.length; i++) if (String(dc[i][0]).trim() === cat && String(dc[i][1]).trim() === viejo && tipoCat_(dc[i][2]) === tipo) hojaC.getRange(i + 1, 2).setValue(nuevo);
  const esDelTipo = function (idx) { return function (f) { return tipo === 'Ingreso' ? esIngreso_(String(f[idx])) : esGasto_(String(f[idx])); }; };
  reemplazarEnHoja_('Movimientos', M.SUB, viejo, nuevo, M.CAT, cat, esDelTipo(M.TIPO));
  reemplazarEnHoja_('Programados', P.SUB, viejo, nuevo, P.CAT, cat, esDelTipo(P.TIPO));
  reemplazarEnHoja_('Recurrentes', 7, viejo, nuevo, 6, cat, function (f) { return String(f[2]).trim() === tipo; });
  if (tipo === 'Gasto') {
    reemplazarEnHoja_('ReglasCategorizacion', 3, viejo, nuevo, 2, cat);
    reemplazarEnHoja_('Presupuestos', 2, viejo, nuevo, 1, cat);
    reemplazarEnHoja_('TarjetaCredito', 12, viejo, nuevo, 11, cat);
  }
  return true;
}

function api_eliminarSubcategoria(a) {
  const cat = String(a.categoria || '').trim(), sub = String(a.sub || '').trim(), tipo = tipoCat_(a.tipo);
  const hoja = getHoja_('Categorias');
  const datos = hoja.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) if (String(datos[i][0]).trim() === cat && String(datos[i][1]).trim() === sub && tipoCat_(datos[i][2]) === tipo) hoja.deleteRow(i + 1);
  return true;
}

// ---------------------------------------------------------------
// PRÓXIMOS RECURRENTES (aparecen junto a los programados)
// ---------------------------------------------------------------

function proximosRecurrentes_() {
  const datos = leerHoja_('Recurrentes');
  const nombres = mapaNombresCuentas_();
  const hoy = new Date();
  const hoy0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (f[0] === '' || String(f[10]).trim().toUpperCase() !== 'SI') continue;
    const dia = Math.min(Math.max(Math.round(numero_(f[5])), 1), 31);
    const tipo = String(f[2]).trim();
    let ref = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    let fecha = new Date(ref.getFullYear(), ref.getMonth(), Math.min(dia, ultimoDiaMes_(ref)), 12, 0);
    const hechoMes = yaSeEjecutoEsteMes_(f[11], hoy);
    if (fecha < hoy0 || (hechoMes && (tipo === 'Gasto' || String(f[12] || '').trim().toUpperCase() === 'SI'))) {
      ref = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 1);
      fecha = new Date(ref.getFullYear(), ref.getMonth(), Math.min(dia, ultimoDiaMes_(ref)), 12, 0);
    }
    const sinCuenta = f[9] === '' || f[9] === null;
    out.push({
      id: 'r' + f[0], recurrente: true, recId: f[0], tipo: tipo === 'Ingreso' ? tipoIngresoPorCategoria_('Ingreso Activo', f[6]) : 'Gasto Fijo', monto: numero_(f[3]), moneda: String(f[4]).trim(),
      medio: sinCuenta ? 'Tarjeta' : 'Debito', cuentaId: f[9], cuentaNombre: sinCuenta ? '' : (nombres[String(f[9])] || ''), descripcion: String(f[1]),
      categoria: String(f[6] || ''), subcategoria: String(f[7] || ''), fecha: fechaHoraISO_(fecha), alertas: [], estado: 'Pendiente', fechaRealizada: '', movimientoId: '', notas: 'Recurrente',
      diasDiferencia: null, atrasado: false, diasAtraso: 0, confirmar: String(f[12] || '').trim().toUpperCase() === 'SI'
    });
  }
  return out;
}

// ---------------------------------------------------------------
// ASESOR (chat con la IA)
// ---------------------------------------------------------------

var SYSTEM_ASESOR_ = 'Eres el asesor financiero personal de un usuario que vive en Trujillo, Perú, y gana en soles, dólares y euros. ' +
  'Actúas a la vez como economista, administrador y contador. NO eres su amigo ni su animador: eres un profesional que cobra por decir la verdad. ' +
  'Habla en español directo, sin rodeos, sin halagos vacíos y sin suavizar. Si una decisión es mala, dilo claramente y explica por qué con cifras de sus datos. ' +
  'Si algo que quiere comprar compromete su flujo, su línea de tarjeta, su fondo de emergencia o su meta de ahorro, dilo sin miedo. ' +
  'Responde corto y accionable (máximo 180 palabras salvo que pida más). Termina con una recomendación concreta (comprar / esperar / no comprar, o el siguiente paso). ' +
  'Usa solo los datos que se te entregan; si falta un dato importante, pídelo en una pregunta breve. No inventes cifras. ' +
    'No des garantías de rentabilidad; para inversiones habla de probabilidades, riesgos y horizonte.';

function contextoAsesor_() {
  const r = construirResumenFinanciero_();
  try { r.tarjetas = api_tarjeta().tarjetas.map(function (t) { return { moneda: t.moneda, linea: t.linea, usada: t.usada, utilizacionPct: Math.round(t.utilizacion), proximoCobro: t.estados[0] || null }; }); } catch (e) { /* opcional */ }
  try { r.objetivos = objetivosIngreso_(); r.repartoMes = distribucionIngresos_('Mensual', new Date(), leerMovimientos_(), obtenerUltimaTasaReferencial_()); } catch (e) { /* opcional */ }
  try { r.proximosMovimientos = resumenProgramados_(45).lista.slice(0, 12).map(function (p) { return { fecha: p.fecha, tipo: p.tipo, monto: p.monto, moneda: p.moneda, descripcion: p.descripcion }; }); } catch (e) { /* opcional */ }
  try { r.posicionesBolsa = resumenInversiones_(false).posiciones.map(function (p) { return { ticker: p.ticker, valor: p.valor, moneda: p.moneda, plPct: Math.round(p.plPct * 10) / 10, cambioDiaPct: Math.round(p.cambioDiaPct * 10) / 10 }; }); } catch (e) { /* opcional */ }
  return r;
}

function api_chat(a) {
  const prov = iaProveedor_();
  if (!prov) throw new Error('Para hablar con el asesor agrega GEMINI_API_KEY (gratis en aistudio.google.com/apikey) en las propiedades del script.');
  const mensajes = (Array.isArray(a.mensajes) ? a.mensajes : []).slice(-12).map(function (m) {
    return { rol: m.rol === 'asesor' ? 'asesor' : 'user', texto: String(m.texto || '').substring(0, 2000) };
  }).filter(function (m) { return m.texto.trim(); });
  if (!mensajes.length || mensajes[mensajes.length - 1].rol !== 'user') throw new Error('Escribe tu pregunta.');
  const system = SYSTEM_ASESOR_ + '\n\nDATOS ACTUALES (JSON, importes en soles equivalentes salvo donde dice moneda; fecha de hoy ' + fechaISO_(new Date()) + '):\n' + JSON.stringify(contextoAsesor_());
  let texto;
  if (prov === 'anthropic') {
    const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': prop_('ANTHROPIC_API_KEY'), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: prop_('AI_MODEL', 'claude-haiku-5-5'), max_tokens: 900, system: system, messages: mensajes.map(function (m) { return { role: m.rol === 'asesor' ? 'assistant' : 'user', content: m.texto }; }) })
    });
    if (r.getResponseCode() !== 200) throw new Error('Anthropic respondió ' + r.getResponseCode());
    texto = (JSON.parse(r.getContentText()).content || []).map(function (b) { return b.text || ''; }).join('');
  } else {
    const contents = mensajes.map(function (m) { return { role: m.rol === 'asesor' ? 'model' : 'user', parts: [{ text: m.texto }] }; });
    texto = geminiGenerar_({ systemInstruction: { parts: [{ text: system }] }, contents: contents, generationConfig: { maxOutputTokens: 2500, temperature: 0.5 } }, prop_('GEMINI_BUSQUEDA') !== 'NO');
  }
  return { respuesta: String(texto).trim(), proveedor: prov };
}

// ---------------------------------------------------------------
// IDEAS DE ACCIONES
// ---------------------------------------------------------------

var UNIVERSO_ACCIONES_ = ['VOO', 'QQQ', 'VTI', 'MSFT', 'AAPL', 'NVDA', 'GOOGL', 'AMZN', 'META', 'BRK-B', 'JPM', 'V', 'KO', 'COST', 'TSLA', 'AVGO'];

function watchlist_() {
  return String(getAjuste_('watchlist', '')).split(',').map(function (t) { return t.trim().toUpperCase(); }).filter(function (t) { return t; });
}
function api_watchlist() { return watchlist_(); }
function api_guardarWatchlist(a) {
  const l = (Array.isArray(a.tickers) ? a.tickers : String(a.tickers || '').split(/[,\s]+/)).map(function (t) { return limpiarTicker_(t); }).filter(function (t) { return t; });
  setAjuste_('watchlist', l.slice(0, 20).join(','));
  CacheService.getScriptCache().remove('ideas_mk');
  return watchlist_();
}

/** Cierres diarios (≈6 meses) de un ticker: Yahoo (con respaldo query2) y, si falla, Stooq. { cierres, nombre, moneda, fuente } o { error } */
function cierres6m_(t) {
  const y = yahooChart_(t, '?range=6mo&interval=1d');
  if (y.res) {
    const c = y.res.indicators.quote[0].close.filter(function (x) { return x !== null && x !== undefined; });
    if (c.length >= 30) return { cierres: c, nombre: y.res.meta.shortName || '', moneda: y.res.meta.currency || 'USD', fuente: 'Yahoo' };
  }
  const s = stooqCierres_(t, 130);
  if (s && s.length >= 30) return { cierres: s.map(function (p) { return p[1]; }), nombre: '', moneda: 'USD', fuente: 'Stooq' };
  return { error: y.error || 'sin datos' };
}

/** Rendimiento y riesgo por ticker. Pide de a pocos para no ser bloqueado. Devuelve { tabla, fallidos: [{ticker, error}], fuentes } */
function metricasMercado_(tickers) {
  const out = [], fallidos = [], fuentes = {};
  tickers.forEach(function (t, i) {
    if (i > 0 && i % 4 === 0) Utilities.sleep(500);
    const d = cierres6m_(t);
    if (d.error) { fallidos.push({ ticker: t, error: d.error }); return; }
    fuentes[d.fuente] = (fuentes[d.fuente] || 0) + 1;
    const c = d.cierres, n = c.length, ult = c[n - 1];
    const ret = function (k) { const b = c[Math.max(0, n - 1 - k)]; return b > 0 ? (ult / b - 1) * 100 : 0; };
    const rets = []; for (let k = 1; k < n; k++) rets.push(c[k] / c[k - 1] - 1);
    const media = rets.reduce(function (s, x) { return s + x; }, 0) / rets.length;
    const vol = Math.sqrt(rets.reduce(function (s, x) { return s + Math.pow(x - media, 2); }, 0) / rets.length) * Math.sqrt(252) * 100;
    const max = Math.max.apply(null, c);
    const sma50 = c.slice(-50).reduce(function (s, x) { return s + x; }, 0) / Math.min(50, n);
    out.push({ ticker: t, nombre: d.nombre, moneda: d.moneda, precio: Math.round(ult * 100) / 100, dia: Math.round(ret(1) * 10) / 10, m1: Math.round(ret(21) * 10) / 10, m3: Math.round(ret(63) * 10) / 10, m6: Math.round(ret(n - 1) * 10) / 10, caidaDesdeMax: Math.round((ult / max - 1) * 1000) / 10, volatilidad: Math.round(vol), sobreMedia50: ult > sma50 });
  });
  return { tabla: out, fallidos: fallidos, fuentes: fuentes };
}

function api_ideasAcciones(a) {
  const cache = CacheService.getScriptCache();
  if (!a.forzar) { const g = cache.get('ideas_mk'); if (g) return JSON.parse(g); }
  const propios = Object.keys(calcularPosiciones_(leerOperacionesInversion_(), false));
  const tickers = [];
  UNIVERSO_ACCIONES_.concat(watchlist_()).forEach(function (t) { if (tickers.indexOf(t) === -1) tickers.push(t); });
  const mk = metricasMercado_(tickers);
  const tabla = mk.tabla;
  tabla.forEach(function (x) {
    x.score = Math.round(x.m6 * 0.35 + x.m3 * 0.4 + x.m1 * 0.25 - Math.max(0, -x.caidaDesdeMax - 10) * 0.6 - Math.max(0, x.volatilidad - 35) * 0.2 + (x.sobreMedia50 ? 3 : -3));
    x.propia = propios.indexOf(x.ticker) !== -1;
  });
  tabla.sort(function (p, q) { return q.score - p.score; });
  const salida = { tabla: tabla, ia: null, error: '', generado: fechaHoraISO_(new Date()), iaDisponible: !!iaProveedor_(), fallidos: mk.fallidos, fuentes: mk.fuentes, total: tickers.length };
  if (!tabla.length) salida.error = 'No pude descargar precios de mercado ahora (Yahoo/Stooq no respondieron). Intenta de nuevo en unos minutos.';
  if (salida.iaDisponible && tabla.length) {
    try {
      const system = SYSTEM_ASESOR_ + ' Ahora analiza esta tabla de mercado (precios reales de Yahoo Finance, rendimientos en %). Elige entre 3 y 5 ideas de las que mejor se ven, y marca también 1 o 2 que EVITARÍA ahora. ' +
        'Para cada una da veredicto (comprar|esperar|evitar), la razón con cifras de la tabla y el principal riesgo. Considera que sus posiciones actuales ya son ' + propios.join(', ') + ' (evita sobreconcentrar). ' +
        'Responde SOLO con JSON: {"resumen":"2 frases sobre el momento del mercado","ideas":[{"ticker":"","veredicto":"comprar|esperar|evitar","texto":"","riesgo":""}]}';
      const j = extraerJSON_(iaTexto_(system, JSON.stringify(tabla.slice(0, 20)), 1800));
      salida.ia = { resumen: String(j.resumen || ''), ideas: (j.ideas || []).slice(0, 7).map(function (x) { return { ticker: String(x.ticker || '').toUpperCase(), veredicto: ['comprar', 'esperar', 'evitar'].indexOf(x.veredicto) !== -1 ? x.veredicto : 'esperar', texto: String(x.texto || ''), riesgo: String(x.riesgo || '') }; }) };
    } catch (e) { salida.error = e.message; }
  }
  if (tabla.length >= 5) { try { cache.put('ideas_mk', JSON.stringify(salida), 10800); } catch (e) { /* demasiado grande */ } }
  return salida;
}

// ---------------------------------------------------------------
// HISTORIAL DE LA CARTERA (gráfico tipo Revolut)
// ---------------------------------------------------------------

function api_historialPortafolio(a) {
  const rango = RANGOS_YAHOO_[a.rango] ? a.rango : '1M';
  const ops = leerOperacionesInversion_();
  const pos = calcularPosiciones_(ops, false);
  const tickers = Object.keys(pos).filter(function (t) { return pos[t].cantidad > 0; });
  const tasa = obtenerUltimaTasaReferencial_();
  const cacheP = CacheService.getScriptCache();
  const firma = 'hp_' + rango + '_' + tickers.slice().sort().join(',') + '_' + ops.length;
  try { const g = cacheP.get(firma); if (g) return JSON.parse(g); } catch (e) { /* sin caché */ }
  const hp = historialesParalelo_(tickers, rango);
  const series = hp.series, fallidos = hp.fallidos;
  const usados = Object.keys(series).filter(function (t) { return series[t].length > 0; });
  if (!usados.length) return { rango: rango, puntos: [], spark: {}, cambioPct: 0, ganancia: 0, fallidos: fallidos, sinPosiciones: !tickers.length, error: tickers.length ? 'No pude descargar el historial de precios ahora. Intenta de nuevo en unos minutos.' : '' };

  let tiempos = [];
  usados.forEach(function (t) { series[t].forEach(function (p) { tiempos.push(p[0]); }); });
  tiempos = tiempos.filter(function (x, i, arr) { return arr.indexOf(x) === i; }).sort(function (x, y) { return x - y; });
  if (tiempos.length > 220) { const paso = tiempos.length / 220; tiempos = tiempos.filter(function (x, i) { return Math.floor(i % paso) === 0 || i === tiempos.length - 1; }); }

  const idx = {}; usados.forEach(function (t) { idx[t] = 0; });
  const puntos = [];
  tiempos.forEach(function (ts) {
    let valor = 0, invertido = 0;
    usados.forEach(function (t) {
      const s = series[t];
      while (idx[t] + 1 < s.length && s[idx[t] + 1][0] <= ts) idx[t]++;
      const precio = s[idx[t]][1];
      let q = 0;
      ops.forEach(function (o) { if (o.ticker !== t || o.fecha.getTime() > ts) return; if (o.operacion === 'Compra') { q += o.cantidad; invertido += convertirAPEN_(o.cantidad * o.precio, o.moneda, tasa); } else if (o.operacion === 'Venta') { q -= o.cantidad; invertido -= convertirAPEN_(o.cantidad * o.precio, o.moneda, tasa); } });
      valor += convertirAPEN_(Math.max(0, q) * precio, pos[t].moneda, tasa);
    });
    puntos.push([ts, redondear_(valor), redondear_(invertido)]);
  });
  // recortar antes de la primera compra (valor 0 e invertido 0 no aportan)
  let ini = 0; while (ini < puntos.length - 1 && puntos[ini][1] <= 0 && puntos[ini][2] <= 0) ini++;
  if (ini > 0) puntos.splice(0, ini);
  const spark = {};
  usados.forEach(function (t) { const s = series[t]; const paso = Math.max(1, Math.floor(s.length / 40)); spark[t] = s.filter(function (p, i) { return i % paso === 0 || i === s.length - 1; }).map(function (p) { return p[1]; }); });
  const p0 = puntos[0], p1 = puntos[puntos.length - 1];
  const ganancia = (p1[1] - p0[1]) - (p1[2] - p0[2]);
  const base = p0[1] > 0 ? p0[1] : Math.max(1, p1[2]);
  const salida = { rango: rango, puntos: puntos, spark: spark, cambioPct: ganancia / base * 100, ganancia: redondear_(ganancia), fallidos: fallidos };
  if (!fallidos.length) { try { cacheP.put(firma, JSON.stringify(salida), rango === '1D' ? 90 : 600); } catch (e) { /* demasiado grande */ } }
  return salida;
}

/** Descarga el historial de todas las acciones a la vez (una sola ronda de red). Las que fallen se reintentan una a una. */
function historialesParalelo_(tickers, rango) {
  const series = {}, fallidos = [];
  const cache = CacheService.getScriptCache();
  const r = RANGOS_YAHOO_[rango] || RANGOS_YAHOO_['1M'];
  const pendientes = [];
  tickers.forEach(function (t) {
    let g = null;
    try { g = cache.get('hist_' + limpiarTicker_(t) + '_' + rango); } catch (e) { g = null; }
    if (g) series[t] = JSON.parse(g).puntos; else pendientes.push(t);
  });
  let resp = [];
  if (pendientes.length) {
    try {
      resp = UrlFetchApp.fetchAll(pendientes.map(function (t) {
        return Object.assign({ url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(limpiarTicker_(t)) + '?range=' + r[0] + '&interval=' + r[1] }, OPCIONES_YAHOO_);
      }));
    } catch (e) { resp = []; }
  }
  pendientes.forEach(function (t, i) {
    let puntos = [];
    try {
      if (resp[i] && resp[i].getResponseCode() === 200) {
        const res = JSON.parse(resp[i].getContentText()).chart.result[0];
        const cierres = res.indicators.quote[0].close;
        res.timestamp.forEach(function (ts, k) { if (cierres[k] !== null && cierres[k] !== undefined) puntos.push([ts * 1000, Math.round(cierres[k] * 10000) / 10000]); });
      }
    } catch (e) { puntos = []; }
    if (puntos.length) {
      if (puntos.length > 220) { const paso = puntos.length / 220; puntos = puntos.filter(function (p, k) { return Math.floor(k % paso) === 0 || k === puntos.length - 1; }); }
      series[t] = puntos;
      try { cache.put('hist_' + limpiarTicker_(t) + '_' + rango, JSON.stringify({ ticker: t, rango: rango, puntos: puntos, moneda: '', fuente: 'Yahoo' }), rango === '1D' ? 120 : 600); } catch (e) { /* sin caché */ }
    } else {
      try { series[t] = historialAccion_(t, rango).puntos; } catch (e) { fallidos.push({ ticker: t, error: e.message }); }
    }
  });
  return { series: series, fallidos: fallidos };
}

// ---------------------------------------------------------------
// ALERTAS DE ACCIONES (suben o bajan mucho en el día)
// ---------------------------------------------------------------

function revisarAccionesAlertas() {
  ejecutarSeguro_('revisarAccionesAlertas', function () {
    if (String(getAjuste_('notif.acciones', 'SI')) === 'NO') return;
    const ahora = new Date();
    const dow = Number(Utilities.formatDate(ahora, zona_(), 'u')), hora = Number(Utilities.formatDate(ahora, zona_(), 'H'));
    if (dow > 5 || hora < 8 || hora > 16) return; // solo en horario de mercado (hora de Lima)
    const pos = calcularPosiciones_(leerOperacionesInversion_(), false);
    const tickers = Object.keys(pos).filter(function (t) { return pos[t].cantidad > 0; });
    if (!tickers.length) return;
    const cot = actualizarCotizaciones_(tickers, 10);
    const umbral = numero_(getAjuste_('acc.umbral', 5)) || 5;
    const props = PropertiesService.getScriptProperties();
    const hoyISO = fechaISO_(ahora);
    let hechas = {};
    try { hechas = JSON.parse(props.getProperty('ACC_ALERTAS') || '{}'); } catch (e) { hechas = {}; }
    Object.keys(hechas).forEach(function (k) { if (String(hechas[k]).indexOf(hoyISO) !== 0) delete hechas[k]; });

    tickers.forEach(function (t) {
      const c = cot[t];
      if (!c || !(c.precio > 0) || Math.abs(c.cambioPct) < umbral) return;
      const dir = c.cambioPct > 0 ? 'sube' : 'baja';
      if (hechas[t] === hoyISO + '|' + dir) return;
      hechas[t] = hoyISO + '|' + dir;
      const p = pos[t];
      const valor = c.precio * p.cantidad, pl = valor - p.costo;
      let texto = (c.cambioPct > 0 ? '📈 ' : '📉 ') + '*' + escaparMd_(t) + '* ' + dir + ' ' + c.cambioPct.toFixed(1) + '% hoy (' + formatoMonto_(c.precio, p.moneda) + ').\nTu posición: ' + formatoMonto_(valor, p.moneda) + ' (' + (pl >= 0 ? '+' : '') + formatoMonto_(pl, p.moneda) + ' en total).';
      if (iaProveedor_()) {
        try {
          const h1 = historialAccion_(t, '1M').puntos, h6 = historialAccion_(t, '6M').puntos;
          const ret = function (h) { return h.length > 1 ? (h[h.length - 1][1] / h[0][1] - 1) * 100 : 0; };
          const datos = { ticker: t, hoyPct: c.cambioPct, precio: c.precio, mes1Pct: Math.round(ret(h1) * 10) / 10, meses6Pct: Math.round(ret(h6) * 10) / 10, miCostoMedio: p.costo / p.cantidad, miPlPct: p.costo > 0 ? Math.round(pl / p.costo * 1000) / 10 : 0 };
          const op = iaTexto_(SYSTEM_ASESOR_ + ' Responde en MÁXIMO 50 palabras: ¿conviene comprar más, mantener o reducir esta posición tras el movimiento de hoy? Empieza con el veredicto en mayúsculas.', JSON.stringify(datos), 500);
          texto += '\n\n🧠 ' + escaparMd_(op.trim()) + '\n_No es garantía; decide con tu horizonte y riesgo._';
        } catch (e) { /* sin opinión de IA */ }
      }
      enviarMensajeTelegram(texto);
    });
    props.setProperty('ACC_ALERTAS', JSON.stringify(hechas));
  });
}


/* ==== 22_Tarjetas.gs ==== */
/**
 * ============================================================
 * TARJETAS DE CRÉDITO — varias tarjetas, cada una con su línea, cierre y cobro.
 * ============================================================
 * - Una tarjeta es como una cuenta más: aparece en Cuentas con su línea, lo usado y lo disponible.
 * - La línea puede ser COMPARTIDA (una sola línea en la moneda base; los consumos en otra moneda
 *   la consumen convertidos al tipo de cambio) o SEPARADA (una línea por moneda: soles, dólares, euros).
 * - Puedes consumir en cualquier moneda con cualquier tarjeta. Cada moneda tiene su propio estado de cuenta.
 * - Hoja "Tarjetas": ID, Banco, MonedaBase, Linea, ModoLinea, LineaUSD, LineaEUR, DiaCierre, DiaCobro, Color, Activa, Principal.
 * - TarjetaCredito lleva una columna nueva (TarjetaID, col. 14). Las filas antiguas sin ese dato son de la tarjeta principal.
 */

var TARJETAS_ENC_ = ['ID', 'Banco', 'MonedaBase', 'Linea', 'ModoLinea', 'LineaUSD', 'LineaEUR', 'DiaCierre', 'DiaCobro', 'Color', 'Activa', 'Principal'];
var TARJETAS_MEM_ = null;
var FILAS_TARJETA_MEM_ = null;

function hojaTarjetas_() {
  const ss = getSS_();
  let h = ss.getSheetByName('Tarjetas');
  if (!h) {
    h = crearHojaBase_(ss, 'Tarjetas', TARJETAS_ENC_);
    migrarTarjetasLegado_(h);
  }
  const t = ss.getSheetByName('TarjetaCredito');
  if (t && String(t.getRange(1, 1, 1, 14).getValues()[0][13]) !== 'TarjetaID') {
    t.getRange(1, 14).setValue('TarjetaID').setFontWeight('bold').setBackground('#1c1c1e').setFontColor('#ffffff');
  }
  return h;
}

/** Primera vez: crea tu tarjeta a partir de lo que ya tenías configurado (por moneda) */
function migrarTarjetasLegado_(h) {
  const c = CONFIG.TARJETA_CREDITO;
  const conf = CONFIG.MONEDAS.map(function (m) {
    return {
      m: m, linea: numero_(getAjuste_('tarjeta.' + m + '.linea', 0)), banco: String(getAjuste_('tarjeta.' + m + '.banco', c.BANCO)),
      cierre: Math.round(numero_(getAjuste_('tarjeta.' + m + '.cierre', c.DIA_CIERRE))), cobro: Math.round(numero_(getAjuste_('tarjeta.' + m + '.cobro', c.DIA_COBRO)))
    };
  });
  const base = conf.filter(function (x) { return x.linea > 0 && x.m === 'PEN'; })[0] || conf.filter(function (x) { return x.linea > 0; })[0] || conf[0];
  const conLinea = conf.filter(function (x) { return x.linea > 0; }).length;
  const usd = conf.filter(function (x) { return x.m === 'USD'; })[0], eur = conf.filter(function (x) { return x.m === 'EUR'; })[0];
  h.appendRow([
    1, base.banco, base.m, base.linea, conLinea > 1 ? 'Separada' : 'Compartida', base.m === 'USD' ? 0 : usd.linea, base.m === 'EUR' ? 0 : eur.linea,
    base.cierre >= 1 && base.cierre <= 31 ? base.cierre : c.DIA_CIERRE, base.cobro >= 1 && base.cobro <= 31 ? base.cobro : c.DIA_COBRO, '#ffb347', 'SI', 'SI'
  ]);
}

function leerTarjetas_() {
  if (TARJETAS_MEM_) return TARJETAS_MEM_;
  const datos = hojaTarjetas_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    const cierre = Math.round(numero_(datos[i][7])), cobro = Math.round(numero_(datos[i][8]));
    out.push({
      id: datos[i][0], fila: i + 1, banco: String(datos[i][1] || '').trim() || CONFIG.TARJETA_CREDITO.BANCO,
      moneda: String(datos[i][2] || 'PEN').trim().toUpperCase(), linea: numero_(datos[i][3]),
      modo: String(datos[i][4]).trim() === 'Separada' ? 'Separada' : 'Compartida',
      lineaUSD: numero_(datos[i][5]), lineaEUR: numero_(datos[i][6]),
      cierre: cierre >= 1 && cierre <= 31 ? cierre : CONFIG.TARJETA_CREDITO.DIA_CIERRE, cobro: cobro >= 1 && cobro <= 31 ? cobro : CONFIG.TARJETA_CREDITO.DIA_COBRO,
      color: String(datos[i][9] || '#ffb347'), activa: String(datos[i][10]).trim().toUpperCase() !== 'NO', principal: String(datos[i][11]).trim().toUpperCase() === 'SI'
    });
  }
  TARJETAS_MEM_ = out;
  return out;
}

function invalidarTarjetas_() { TARJETAS_MEM_ = null; FILAS_TARJETA_MEM_ = null; }

function tarjetasActivas_() { return leerTarjetas_().filter(function (t) { return t.activa; }); }

function tarjetaPrincipal_() {
  const l = leerTarjetas_();
  return l.filter(function (t) { return t.principal && t.activa; })[0] || l.filter(function (t) { return t.activa; })[0] || l[0];
}

/** Tarjeta por ID; si no se indica (o no existe) usa la principal */
function tarjetaPorId_(id) {
  if (id !== undefined && id !== null && String(id) !== '') {
    const t = leerTarjetas_().filter(function (x) { return String(x.id) === String(id); })[0];
    if (t) return t;
  }
  return tarjetaPrincipal_();
}

function normalizarTxt_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }

/** Para Telegram/IA: "con la bcp" → la tarjeta cuyo banco coincide */
function tarjetaPorNombre_(txt) {
  const q = normalizarTxt_(txt);
  if (!q) return null;
  return tarjetasActivas_().filter(function (t) { const b = normalizarTxt_(t.banco); return b === q || b.indexOf(q) !== -1 || q.indexOf(b) !== -1; })[0] || null;
}

/** Compatibilidad: configuración de la tarjeta principal vista "por moneda" */
function tarjetaCfg_(moneda) {
  const t = tarjetaPrincipal_();
  return { moneda: moneda || t.moneda, banco: t.banco, cierre: t.cierre, cobro: t.cobro, linea: lineaDeMoneda_(t, moneda || t.moneda), activa: true };
}

function lineaDeMoneda_(t, moneda) {
  if (moneda === t.moneda) return t.linea;
  if (t.modo !== 'Separada') return 0;
  return moneda === 'USD' ? t.lineaUSD : (moneda === 'EUR' ? t.lineaEUR : 0);
}

/** A qué tarjeta pertenece cada fila de TarjetaCredito (las antiguas, sin dato, son de la principal) */
function mapaFilaTarjeta_() {
  if (FILAS_TARJETA_MEM_) return FILAS_TARJETA_MEM_;
  const datos = leerHoja_('TarjetaCredito');
  const principal = tarjetaPrincipal_();
  const ids = {};
  leerTarjetas_().forEach(function (t) { ids[String(t.id)] = true; });
  const m = {};
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '') continue;
    const tid = datos[i][13] !== undefined ? String(datos[i][13]).trim() : '';
    m[String(datos[i][0])] = ids[tid] ? tid : String(principal.id);
  }
  FILAS_TARJETA_MEM_ = m;
  return m;
}

/** Nombre de la tarjeta de un movimiento (por su referencia a TarjetaCredito) */
function nombreTarjetaDeMov_(m) {
  const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
  const mapa = mapaFilaTarjeta_();
  const t = tarjetaPorId_(tid ? mapa[String(tid)] : '');
  return t ? t.banco : 'Tarjeta';
}

// ---------------------------------------------------------------
// USO DE LA LÍNEA
// ---------------------------------------------------------------

/** Capital pendiente de una tarjeta, por moneda (los intereses futuros no consumen línea) */
function usoTarjeta_(t) {
  const datos = leerHoja_('TarjetaCredito');
  const mapa = mapaFilaTarjeta_();
  const por = { PEN: 0, USD: 0, EUR: 0 };
  let n = 0;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === '' || String(datos[i][7]).trim() === 'Pagado') continue;
    if (mapa[String(datos[i][0])] !== String(t.id)) continue;
    if (String(datos[i][11]) === 'Intereses y comisiones') continue;
    const m = String(datos[i][4]).trim().toUpperCase();
    por[m] = (por[m] || 0) + numero_(datos[i][3]);
    n++;
  }
  Object.keys(por).forEach(function (k) { por[k] = redondear_(por[k]); });
  return { porMoneda: por, consumos: n };
}

function aMonedaBase_(monto, moneda, base, tasa) {
  if (moneda === base) return monto;
  const pen = convertirAPEN_(monto, moneda, tasa);
  return base === 'PEN' ? pen : pen / convertirAPEN_(1, base, tasa);
}

/** Resumen de una tarjeta: líneas (una compartida o una por moneda), usado, disponible, estados de cuenta */
function resumenTarjeta_(t, estadosTodos) {
  const u = usoTarjeta_(t), tasa = obtenerUltimaTasaReferencial_(), hoy = new Date();
  const lineas = [];
  if (t.modo === 'Separada') {
    CONFIG.MONEDAS.forEach(function (m) {
      const linea = lineaDeMoneda_(t, m), usada = u.porMoneda[m] || 0;
      if (linea > 0 || usada > 0 || m === t.moneda) lineas.push({ moneda: m, linea: linea, usada: usada, disponible: linea > 0 ? redondear_(linea - usada) : 0, utilizacion: linea > 0 ? usada / linea * 100 : 0, compartida: false });
    });
  } else {
    let usada = 0;
    const otras = [];
    CONFIG.MONEDAS.forEach(function (m) {
      const x = u.porMoneda[m] || 0;
      usada += aMonedaBase_(x, m, t.moneda, tasa);
      if (m !== t.moneda && x > 0) otras.push({ moneda: m, monto: x, equivalente: redondear_(aMonedaBase_(x, m, t.moneda, tasa)) });
    });
    usada = redondear_(usada);
    lineas.push({ moneda: t.moneda, linea: t.linea, usada: usada, disponible: t.linea > 0 ? redondear_(t.linea - usada) : 0, utilizacion: t.linea > 0 ? usada / t.linea * 100 : 0, compartida: true, otras: otras });
  }
  const estados = (estadosTodos || obtenerEstadosTarjetaPendientes_()).filter(function (e) { return String(e.tarjetaId) === String(t.id); });
  const principal = lineas[0];
  return {
    id: t.id, banco: t.banco, moneda: t.moneda, modo: t.modo, color: t.color, cierre: t.cierre, cobro: t.cobro, principal: t.principal, activa: t.activa,
    linea: t.linea, lineaUSD: t.lineaUSD, lineaEUR: t.lineaEUR,
    lineas: lineas, usada: principal.usada, disponible: principal.disponible, utilizacion: principal.utilizacion,
    porMoneda: u.porMoneda, consumos: u.consumos,
    estados: estados.map(function (e) { return { fechaCobro: fechaISO_(e.fechaCobro), moneda: e.moneda, total: e.total, consumos: e.consumos, cerrado: !!e.fechaCierre && e.fechaCierre <= hoy }; })
  };
}

function tarjetasResumen_() {
  const estados = obtenerEstadosTarjetaPendientes_();
  return tarjetasActivas_().map(function (t) { return resumenTarjeta_(t, estados); });
}

function consejosTarjeta_(t) {
  const out = [];
  const f = function (n, m) { return formatoMonto_(n, m); };
  t.lineas.forEach(function (L) {
    const etiqueta = t.lineas.length > 1 && !L.compartida ? ' en ' + L.moneda : '';
    if (L.linea > 0) {
      const u = L.utilizacion;
      if (u >= 90) out.push({ nivel: 'alerta', titulo: 'Línea casi agotada' + etiqueta + ' (' + u.toFixed(0) + '%)', texto: 'Usas ' + f(L.usada, L.moneda) + ' de ' + f(L.linea, L.moneda) + '. Por encima de 30% tu historial crediticio se resiente y no tendrás margen para una emergencia. Prioriza pagar antes del cierre.' });
      else if (u >= 50) out.push({ nivel: 'alerta', titulo: 'Utilización alta' + etiqueta + ' (' + u.toFixed(0) + '%)', texto: 'Usas ' + f(L.usada, L.moneda) + ' de ' + f(L.linea, L.moneda) + '. Lo sano es mantenerla bajo 30%: paga antes de la fecha de cierre para que el banco reporte menos deuda.' });
      else if (u > 30) out.push({ nivel: 'info', titulo: 'Utilización moderada' + etiqueta + ' (' + u.toFixed(0) + '%)', texto: 'Vas por encima del 30% recomendado. Intenta bajar a ' + f(L.linea * 0.3, L.moneda) + ' pagando antes del cierre.' });
      else out.push({ nivel: 'ok', titulo: 'Línea bajo control' + etiqueta + ' (' + u.toFixed(0) + '%)', texto: 'Usas ' + f(L.usada, L.moneda) + ' de ' + f(L.linea, L.moneda) + '. Sigue así.' });
    } else if (L.moneda === t.moneda) {
      out.push({ nivel: 'info', titulo: 'Falta tu línea de crédito', texto: 'Configúrala para calcular tu utilización y cuánto te queda disponible.' });
    }
    if (L.compartida && L.otras && L.otras.length) {
      out.push({ nivel: 'info', titulo: 'Consumos en otra moneda', texto: L.otras.map(function (o) { return f(o.monto, o.moneda) + ' (≈ ' + f(o.equivalente, t.moneda) + ')'; }).join(' · ') + ' consumen tu línea en ' + t.moneda + ' al tipo de cambio de hoy. Si tu banco tiene línea aparte en esa moneda, cámbiala a "línea separada".' });
    }
  });
  const prox = t.estados[0];
  if (prox) {
    const liquido = obtenerCuentasActivas_().filter(function (c) { return c.moneda === prox.moneda && c.tipo !== 'Inversion'; }).reduce(function (s, c) { return s + c.saldo; }, 0);
    if (liquido < prox.total) out.push({ nivel: 'alerta', titulo: 'No te alcanza para el próximo cobro', texto: 'El ' + prox.fechaCobro + ' se cobra ' + f(prox.total, prox.moneda) + ' y hoy tienes ' + f(liquido, prox.moneda) + ' en cuentas de ' + prox.moneda + '. Si no cubres el total pagarás intereses moratorios y compensatorios. Si tienes plata en otra moneda, usa Cambio de divisa.' });
    else out.push({ nivel: 'info', titulo: 'Próximo cobro: ' + f(prox.total, prox.moneda), texto: 'Es el ' + prox.fechaCobro + '. Paga el total, nunca el mínimo: el interés de tarjetas en Perú suele superar 40% anual.' });
  }
  const futuras = t.estados.slice(1).reduce(function (s, e) { return s + aMonedaBase_(e.total, e.moneda, 'PEN', obtenerUltimaTasaReferencial_()); }, 0);
  if (futuras > 0) out.push({ nivel: 'info', titulo: 'Cuotas ya comprometidas', texto: 'Tienes ≈ ' + f(futuras, 'PEN') + ' en estados posteriores (cuotas y consumos recientes). Esa plata de tus próximos ingresos ya está gastada.' });
  const hace12 = new Date(); hace12.setMonth(hace12.getMonth() - 12);
  const mapa = mapaFilaTarjeta_();
  const interes = leerMovimientos_().filter(function (m) {
    if (!(esGasto_(m.tipo) && m.categoria === 'Intereses y comisiones' && m.efectiva >= hace12 && m.efectiva <= new Date())) return false;
    const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
    return tid && mapa[String(tid)] === String(t.id);
  }).reduce(function (s, m) { return s + aMonedaBase_(m.monto, m.moneda, 'PEN', obtenerUltimaTasaReferencial_()); }, 0);
  if (interes > 0) out.push({ nivel: 'alerta', titulo: 'Intereses pagados: ' + f(interes, 'PEN'), texto: 'En 12 meses has regalado ≈ ' + f(interes, 'PEN') + ' al banco. Antes de financiar en cuotas con interés, compara con ahorrar y pagar al contado.' });
  return out;
}

/** Cuotas vigentes de una tarjeta, agrupadas por compra */
function cuotasVigentes_(t) {
  const mapa = mapaFilaTarjeta_(), filas = leerHoja_('TarjetaCredito');
  const pendiente = {};
  for (let i = 1; i < filas.length; i++) if (filas[i][0] !== '' && String(filas[i][7]).trim() !== 'Pagado') pendiente[String(filas[i][0])] = true;
  const grupos = {};
  leerMovimientos_().forEach(function (m) {
    const g = String(m.notas || '').match(/Cuotas #(\w+) \((\d+)\/(\d+)\)/);
    if (!g || m.categoria === 'Intereses y comisiones') return;
    const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
    if (!tid || mapa[String(tid)] !== String(t.id) || !pendiente[String(tid)]) return;
    const x = grupos[g[1]] = grupos[g[1]] || { grupo: g[1], descripcion: String(m.descripcion).replace(/ \(cuota \d+\/\d+\)$/, ''), moneda: m.moneda, total: Number(g[3]), restantes: 0, capitalPendiente: 0, proxima: null };
    x.restantes++; x.capitalPendiente = redondear_(x.capitalPendiente + m.monto);
    if (!x.proxima || m.efectiva < x.proxima) x.proxima = m.efectiva;
  });
  return Object.keys(grupos).map(function (k) { const x = grupos[k]; x.proxima = x.proxima ? fechaISO_(x.proxima) : ''; return x; })
    .sort(function (a, b) { return a.proxima < b.proxima ? -1 : 1; });
}

// ---------------------------------------------------------------
// API
// ---------------------------------------------------------------

function api_tarjetas() { return tarjetasResumen_(); }

function api_tarjeta(a) {
  const t = tarjetaPorId_(a && a.id);
  const r = resumenTarjeta_(t);
  r.consejos = consejosTarjeta_(r);
  r.cuotas = cuotasVigentes_(t);
  const mapa = mapaFilaTarjeta_(), ahora = new Date();
  r.movimientos = leerMovimientos_().filter(function (x) {
    if (x.medio !== 'Tarjeta de Credito' || !esGasto_(x.tipo) || x.fecha > ahora) return false;
    const tid = extraerRef_(x.notas, /Ref\. TarjetaCredito #(\d+)/);
    return tid && mapa[String(tid)] === String(t.id);
  }).sort(function (p, q) { return (q.fecha - p.fecha) || (numero_(q.id) - numero_(p.id)); }).slice(0, 30)
    .map(function (x) { return { id: x.id, fecha: fechaHoraISO_(x.fecha), efectiva: fechaISO_(x.efectiva), monto: x.monto, moneda: x.moneda, descripcion: x.descripcion, categoria: String(x.categoria || ''), tipo: x.tipo, cuenta: t.banco }; });
  return r;
}

function api_guardarTarjeta(a) {
  const banco = String(a.banco || '').trim();
  if (!banco) throw new Error('Escribe el nombre de la tarjeta.');
  const moneda = validarMoneda_(a.moneda || 'PEN');
  const cierre = Math.round(numero_(a.cierre)), cobro = Math.round(numero_(a.cobro));
  if (!(cierre >= 1 && cierre <= 31) || !(cobro >= 1 && cobro <= 31)) throw new Error('El día de cierre y el día de cobro deben estar entre 1 y 31.');
  const linea = numero_(a.linea), lUSD = numero_(a.lineaUSD), lEUR = numero_(a.lineaEUR);
  if (linea < 0 || lUSD < 0 || lEUR < 0) throw new Error('La línea de crédito no puede ser negativa.');
  const modo = a.modo === 'Separada' ? 'Separada' : 'Compartida';
  const color = String(a.color || '#ffb347');
  return conLock_(function () {
    const hoja = hojaTarjetas_();
    invalidarTarjetas_();
    const existentes = leerTarjetas_();
    const actual = a.id !== undefined && a.id !== '' ? existentes.filter(function (t) { return String(t.id) === String(a.id); })[0] : null;
    if (a.id !== undefined && a.id !== '' && !actual) throw new Error('Tarjeta no encontrada.');
    const principal = actual ? (a.principal === undefined ? actual.principal : !!a.principal) : (a.principal === true || existentes.length === 0);
    const fila = [actual ? actual.id : generarID_('Tarjetas'), banco, moneda, linea, modo, modo === 'Separada' && moneda !== 'USD' ? lUSD : 0, modo === 'Separada' && moneda !== 'EUR' ? lEUR : 0, cierre, cobro, color, actual ? (actual.activa ? 'SI' : 'NO') : 'SI', principal ? 'SI' : 'NO'];
    if (actual) hoja.getRange(actual.fila, 1, 1, fila.length).setValues([fila]); else hoja.appendRow(fila);
    if (principal) {
      invalidarTarjetas_();
      leerTarjetas_().forEach(function (t) { if (String(t.id) !== String(fila[0]) && t.principal) hoja.getRange(t.fila, 12).setValue('NO'); });
    }
    invalidarTarjetas_();
    return resumenTarjeta_(tarjetaPorId_(fila[0]));
  });
}

/** Compatibilidad con la versión anterior: actualiza la tarjeta principal */
function api_configurarTarjeta(a) {
  const p = tarjetaPrincipal_();
  return api_guardarTarjeta({ id: p.id, banco: a.banco || p.banco, moneda: p.moneda, linea: a.linea, modo: p.modo, lineaUSD: p.lineaUSD, lineaEUR: p.lineaEUR, cierre: a.cierre, cobro: a.cobro, color: p.color });
}

/** Elimina la tarjeta si nunca se usó; si tiene historial la archiva (no aparece, pero se conservan sus consumos) */
function api_eliminarTarjeta(a) {
  return conLock_(function () {
    invalidarTarjetas_();
    const t = leerTarjetas_().filter(function (x) { return String(x.id) === String(a.id); })[0];
    if (!t) throw new Error('Tarjeta no encontrada.');
    if (tarjetasActivas_().length <= 1 && t.activa) throw new Error('Debe quedar al menos una tarjeta.');
    const mapa = mapaFilaTarjeta_();
    const filas = leerHoja_('TarjetaCredito');
    let usos = 0, pendientes = 0;
    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0] === '' || mapa[String(filas[i][0])] !== String(t.id)) continue;
      usos++;
      if (String(filas[i][7]).trim() !== 'Pagado') pendientes++;
    }
    if (pendientes > 0) throw new Error('Esta tarjeta aún tiene ' + pendientes + ' consumos sin pagar. Págalos primero.');
    const hoja = hojaTarjetas_();
    if (usos === 0) hoja.deleteRow(t.fila); else hoja.getRange(t.fila, 11).setValue('NO');
    if (t.principal) { invalidarTarjetas_(); const o = tarjetasActivas_()[0]; if (o) hoja.getRange(o.fila, 12).setValue('SI'); }
    invalidarTarjetas_();
    return { archivada: usos > 0 };
  });
}


/* ==== 24_Personas.gs ==== */
/**
 * ============================================================
 * PERSONAS Y COMPRAS PARA OTROS
 * ============================================================
 * Unifica "compras para otra persona" (mayormente tu mamá) con Préstamos y deudas.
 *  - Una persona existe una sola vez (se compara sin tildes ni mayúsculas).
 *  - Compra que te devuelven  => queda "por cobrar" (PrestamosOtorgados, sin interés, con concepto).
 *  - Compra que es un regalo  => queda como gasto normal con la nota "Regalo para X".
 *  - Cuando te pagan, el dinero entra a la cuenta que elijas (es devolución de capital, no ingreso).
 */

function normNombre_(n) { return sinTildes_(String(n || '').trim()).replace(/\s+/g, ' ').toLowerCase(); }

function asegurarColumnaConcepto_() {
  const h = getHoja_('PrestamosOtorgados');
  if (h.getLastColumn() < 13) h.getRange(1, 13).setValue('Concepto');
}

function hojaPersonas_() {
  const ss = getSS_();
  let h = ss.getSheetByName('Personas');
  if (!h) { h = ss.insertSheet('Personas'); h.getRange(1, 1).setValue('Nombre'); h.setFrozenRows(1); }
  return h;
}

/** Nombres guardados (hoja Personas) + deudores de préstamos. Devuelve { clave: nombreCanónico } */
function mapaPersonas_() {
  const m = {};
  const hp = hojaPersonas_().getDataRange().getValues();
  for (let i = 1; i < hp.length; i++) { const n = String(hp[i][0] || '').trim(); if (n && !m[normNombre_(n)]) m[normNombre_(n)] = n; }
  const po = leerHoja_('PrestamosOtorgados');
  for (let i = 1; i < po.length; i++) { const n = String(po[i][1] || '').trim(); if (n && !m[normNombre_(n)]) m[normNombre_(n)] = n; }
  return m;
}

/** Devuelve el nombre ya existente (misma persona) o agrega uno nuevo. Nunca duplica perfiles */
function canonPersona_(nombre) {
  const n = String(nombre || '').trim().replace(/\s+/g, ' ');
  if (!n) throw new Error('Escribe para quién fue la compra.');
  const m = mapaPersonas_();
  if (m[normNombre_(n)]) return m[normNombre_(n)];
  hojaPersonas_().appendRow([n]);
  return n;
}

function api_personas() {
  const m = mapaPersonas_();
  const deben = {};
  const po = leerHoja_('PrestamosOtorgados');
  for (let i = 1; i < po.length; i++) {
    const f = po[i];
    if (f[0] === '' || String(f[9]).trim() === 'Pagado') continue;
    const k = normNombre_(f[1]);
    const d = deben[k] || (deben[k] = { PEN: 0, USD: 0, EUR: 0 });
    d[String(f[2])] = redondear_((d[String(f[2])] || 0) + numero_(f[10]));
  }
  return Object.keys(m).map(function (k) { return { nombre: m[k], debe: deben[k] || { PEN: 0, USD: 0, EUR: 0 } }; })
    .sort(function (a, b) { const sa = a.debe.PEN + a.debe.USD + a.debe.EUR, sb = b.debe.PEN + b.debe.USD + b.debe.EUR; return (sb > 0) - (sa > 0) || a.nombre.localeCompare(b.nombre); });
}

/**
 * Gasto hecho para otra persona. a = datos de api_registrar + { persona, modo: 'devuelve'|'regalo' }
 */
function api_compraParaOtro(a) {
  const persona = canonPersona_(a.persona);
  const modo = a.modo === 'regalo' ? 'regalo' : 'devuelve';
  const monto = validarMonto_(a.monto, 'El monto');
  const moneda = validarMoneda_(a.moneda);
  const desc = String(a.descripcion || '').trim() || 'Compra';
  const medio = a.medio;
  if (modo === 'regalo' || medio === 'Tarjeta') {
    const r = api_registrar({
      clase: 'gasto', tipo: a.tipo, monto: monto, moneda: moneda, descripcion: desc + ' (para ' + persona + ')', medio: medio, cuentaId: a.cuentaId, fecha: a.fecha,
      categoria: a.categoria, subcategoria: a.subcategoria, notas: (modo === 'regalo' ? 'Regalo para ' : 'Compra para ') + persona, tarjetaId: a.tarjetaId,
      cuotas: a.cuotas, interes: a.interes, tipoInteres: a.tipoInteres
    });
    if (modo === 'devuelve') registrarPrestamoOtorgado({ deudor: persona, moneda: moneda, montoPrestado: monto, tasaInteres: 0, plazoMeses: 1, fechaInicio: a.fecha, concepto: desc, cuentaOrigenId: '' });
    return { persona: persona, modo: modo, resultado: r };
  }
  const cuentaId = a.cuentaId !== undefined && a.cuentaId !== '' ? a.cuentaId : cuentaPorDefecto_(moneda, medio === 'Efectivo' ? 'Efectivo' : 'Debito');
  const id = registrarPrestamoOtorgado({ deudor: persona, moneda: moneda, montoPrestado: monto, tasaInteres: 0, plazoMeses: 1, fechaInicio: a.fecha, concepto: desc, cuentaOrigenId: cuentaId });
  return { persona: persona, modo: modo, prestamoId: id };
}

/**
 * Un gasto YA registrado resultó ser para otra persona. args = { id, persona, modo: 'devuelve'|'regalo' }
 *  - devuelve: deja de ser gasto tuyo y pasa a "por cobrar" en Préstamos y deudas (la plata que salió de tu cuenta es la misma).
 *  - regalo: sigue siendo gasto tuyo, con la nota "Regalo para X".
 */
function api_convertirParaOtro(a) {
  const persona = String(a.persona || '').trim();
  if (!persona) throw new Error('Escribe para quién fue.');
  const m = movimientoPorId_(a.id);
  if (!esGasto_(m.tipo)) throw new Error('Solo un gasto puede ser para otra persona.');
  if (RE_DIVIDIDO_.test(m.notas)) throw new Error('Este gasto está dividido: primero deshaz la división.');
  if (/Compra para |Regalo para /.test(String(m.notas))) throw new Error('Este gasto ya está marcado para otra persona.');
  const clase = clasificarMovimiento_(m);
  if (clase !== 'normal' && clase !== 'tarjeta') throw new Error('Este movimiento no se puede cambiar así.');
  if (/Cuotas #\w+ \(\d+\/\d+\)/.test(m.notas)) throw new Error('Un gasto en cuotas no se puede pasar a otra persona.');
  const nombre = canonPersona_(persona);
  if (a.modo === 'regalo') {
    editarMovimiento(m.id, { notas: (m.notas ? m.notas + ' · ' : '') + 'Regalo para ' + nombre });
    return { persona: nombre, modo: 'regalo' };
  }
  // devuelve: se reemplaza el gasto por la compra para la persona (mismo dinero, misma cuenta y fecha)
  const datos = { tipo: m.tipo, monto: m.monto, moneda: m.moneda, descripcion: m.descripcion, medio: clase === 'tarjeta' ? 'Tarjeta' : (m.medio || ''), cuentaId: m.cuentaId,
    fecha: m.fecha, categoria: m.categoria, subcategoria: m.subcategoria, persona: nombre, modo: 'devuelve' };
  if (clase === 'tarjeta') {
    // en tarjeta el consumo se queda (lo cobra el banco) y se suma lo que la persona te debe, igual que al registrar con tarjeta «para otra persona»
    const id = registrarPrestamoOtorgado({ deudor: nombre, moneda: m.moneda, montoPrestado: m.monto, tasaInteres: 0, plazoMeses: 1, fechaInicio: m.fecha, concepto: m.descripcion || 'Compra', cuentaOrigenId: '' });
    editarMovimiento(m.id, { notas: (m.notas ? m.notas + ' · ' : '') + 'Compra para ' + nombre });
    return { persona: nombre, modo: 'devuelve', prestamoId: id };
  }
  eliminarMovimiento(m.id);
  return api_compraParaOtro(datos);
}

/** Me pagó (total o parcial): se aplica a sus deudas más antiguas de esa moneda y entra a una cuenta */
/**
 * Cobro a una persona. Si te pagó en varias cuentas (ej: 10 en efectivo y 90 por banco) manda partes: [{ monto, cuentaId }].
 * Se valida el total antes de mover nada, así nunca queda un cobro a medias.
 */
function api_cobrarPersona(a) {
  const partes = partesCobro_(a);
  if (partes.length === 1) return cobrarPersonaUna_(Object.assign({}, a, partes[0]));
  const total = redondear_(partes.reduce(function (s, x) { return s + x.monto; }, 0));
  const debe = deudaPersona_(String(a.persona || '').trim(), validarMoneda_(a.moneda));
  if (total > debe + 0.005) throw new Error('Entre todas las cuentas suman ' + total + ', más de lo que te debe (' + redondear_(debe) + ').');
  partes.forEach(function (x) { cobrarPersonaUna_(Object.assign({}, a, x)); });
  return true;
}

/** Normaliza { monto, cuentaId } o partes:[{monto,cuentaId}] a una lista válida */
function partesCobro_(a) {
  const lista = Array.isArray(a.partes) && a.partes.length ? a.partes : [{ monto: a.monto, cuentaId: a.cuentaId }];
  const out = lista.map(function (x) { return { monto: validarMonto_(x.monto, 'Cada monto'), cuentaId: x.cuentaId }; });
  if (out.length > 1 && out.some(function (x) { return x.cuentaId === '' || x.cuentaId === undefined; })) throw new Error('Elige la cuenta de cada parte.');
  return out;
}

/** Cuánto te debe una persona en una moneda (préstamos abiertos) */
function deudaPersona_(persona, moneda) {
  const po = leerHoja_('PrestamosOtorgados');
  let t = 0;
  for (let i = 1; i < po.length; i++) {
    const f = po[i];
    if (f[0] === '' || String(f[9]).trim() === 'Pagado' || String(f[2]) !== moneda || normNombre_(f[1]) !== normNombre_(persona)) continue;
    t += numero_(f[10]);
  }
  return t;
}

function cobrarPersonaUna_(a) {
  const persona = String(a.persona || '').trim();
  const moneda = validarMoneda_(a.moneda);
  let resto = validarMonto_(a.monto, 'El monto');
  const abiertos = [];
  const po = leerHoja_('PrestamosOtorgados');
  for (let i = 1; i < po.length; i++) {
    const f = po[i];
    if (f[0] === '' || String(f[9]).trim() === 'Pagado' || String(f[2]) !== moneda || normNombre_(f[1]) !== normNombre_(persona)) continue;
    abiertos.push({ id: f[0], saldo: numero_(f[10]), fecha: fecha_(f[7]) || new Date(0) });
  }
  abiertos.sort(function (x, y) { return x.fecha - y.fecha; });
  const total = abiertos.reduce(function (s, x) { return s + x.saldo; }, 0);
  if (!abiertos.length) throw new Error(persona + ' no te debe nada en ' + moneda + '.');
  if (resto > total + 0.005) throw new Error('Ese monto supera lo que te debe (' + redondear_(total) + ' ' + moneda + ').');
  for (let i = 0; i < abiertos.length && resto > 0.004; i++) {
    const parte = Math.min(resto, abiertos[i].saldo);
    registrarDevolucionCapitalPrestamo(abiertos[i].id, parte, a.cuentaId, a.fecha);
    resto = redondear_(resto - parte);
  }
  return true;
}

/** Editar datos de un préstamo sin tocar el dinero: nombre, concepto, interés, frecuencia, plazo */
function api_editarPrestamo(a) {
  return conLock_(function () {
    const p = buscarFilaPorId_('PrestamosOtorgados', a.id);
    if (!p) throw new Error('Préstamo no encontrado.');
    asegurarColumnaConcepto_();
    const h = getHoja_('PrestamosOtorgados');
    if (a.deudor !== undefined && String(a.deudor).trim()) h.getRange(p.numeroFila, 2).setValue(canonPersona_(a.deudor));
    if (a.tasaInteres !== undefined && a.tasaInteres !== '') h.getRange(p.numeroFila, 5).setValue(normalizarTasa_(a.tasaInteres));
    if (a.tipoTasa) h.getRange(p.numeroFila, 6).setValue(a.tipoTasa === 'Anual' ? 'Anual' : 'Mensual');
    if (a.frecuencia) h.getRange(p.numeroFila, 7).setValue(a.frecuencia === 'Al vencimiento' ? 'Al vencimiento' : 'Mensual');
    if (a.plazoMeses) h.getRange(p.numeroFila, 9).setValue(Math.max(1, Math.round(numero_(a.plazoMeses)) || 1));
    if (a.concepto !== undefined) h.getRange(p.numeroFila, 13).setValue(String(a.concepto));
    return true;
  });
}

/** "Dar por pagado / perdonar": cierra la deuda sin mover dinero (por ejemplo, si decides regalarlo) */
function api_cerrarPrestamo(a) {
  return conLock_(function () {
    const p = buscarFilaPorId_('PrestamosOtorgados', a.id);
    if (!p) throw new Error('Préstamo no encontrado.');
    const h = getHoja_('PrestamosOtorgados');
    h.getRange(p.numeroFila, 11).setValue(0);
    h.getRange(p.numeroFila, 10).setValue('Pagado');
    registrarLog_('PrestamoCerrado', 'ID ' + a.id + ' cerrado sin movimiento de dinero');
    return true;
  });
}

/** Renombra o fusiona personas (si el nuevo nombre ya existe, se unen sus deudas bajo un solo perfil) */
function api_renombrarPersona(a) {
  return conLock_(function () {
    const viejo = String(a.viejo || '').trim(), nuevo = String(a.nuevo || '').trim().replace(/\s+/g, ' ');
    if (!viejo || !nuevo) throw new Error('Escribe el nombre nuevo.');
    const m = mapaPersonas_();
    const destino = m[normNombre_(nuevo)] && normNombre_(nuevo) !== normNombre_(viejo) ? m[normNombre_(nuevo)] : nuevo;
    const h = getHoja_('PrestamosOtorgados'), d = h.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (normNombre_(d[i][1]) === normNombre_(viejo)) h.getRange(i + 1, 2).setValue(destino);
    const hp = hojaPersonas_(), dp = hp.getDataRange().getValues();
    for (let i = dp.length - 1; i >= 1; i--) if (normNombre_(dp[i][0]) === normNombre_(viejo)) hp.deleteRow(i + 1);
    if (!Object.keys(mapaPersonas_()).some(function (k) { return k === normNombre_(destino); })) hp.appendRow([destino]);
    return destino;
  });
}

/** Historial de pagos de un préstamo otorgado (devoluciones e intereses ya registrados) */
function api_prestamoPagos(a) {
  const re = new RegExp('PrestamoOtorgado #' + String(a.id).replace(/[^0-9A-Za-z-]/g, '') + '(\\D|$)');
  const nombres = mapaNombresCuentas_();
  return leerMovimientos_().filter(function (m) { return re.test(m.notas || ''); })
    .sort(function (x, y) { return y.fecha - x.fecha; }).map(function (m) { return movimientoParaApp_(m, nombres); });
}


/* ==== 25_Proyectos.gs ==== */
/**
 * ============================================================
 * PROYECTOS / NEGOCIOS — rentabilidad de cada inversión o negocio independiente
 * ============================================================
 * Un proyecto agrupa movimientos: lo que invertiste (gastos de tipo "Gasto Inversion"), otros costos
 * (cualquier otro gasto asignado) y lo que ingresó. Se calcula en soles al tipo de cambio actual.
 *   Ganancia = ingresos − invertido − costos      ROI = ganancia / (invertido + costos)
 * La columna 16 ("Proyecto") de Movimientos guarda el ID del proyecto.
 */

function asegurarColumnaProyecto_() {
  const h = getHoja_('Movimientos');
  if (h.getLastColumn() < 16) h.getRange(1, 16).setValue('Proyecto');
}

function hojaProyectos_() {
  const ss = getSS_();
  let h = ss.getSheetByName('Proyectos');
  if (!h) { h = ss.insertSheet('Proyectos'); h.getRange(1, 1, 1, 6).setValues([['ID', 'Nombre', 'Tipo', 'Estado', 'Notas', 'Creado']]); h.setFrozenRows(1); }
  return h;
}

function leerProyectos_() {
  const d = hojaProyectos_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === '') continue;
    out.push({ id: d[i][0], nombre: String(d[i][1]), tipo: String(d[i][2] || 'Negocio'), estado: String(d[i][3] || 'Activo'), notas: String(d[i][4] || ''), creado: fechaISO_(d[i][5]) });
  }
  return out;
}

/** Asigna (o quita, con proyectoId vacío) el proyecto de varios movimientos */
function asignarProyecto_(ids, proyectoId) {
  asegurarColumnaProyecto_();
  const set = {};
  (ids || []).forEach(function (x) { set[String(x)] = true; });
  const hoja = getHoja_('Movimientos');
  let n = 0;
  leerMovimientos_().forEach(function (m) {
    if (set[String(m.id)]) { hoja.getRange(m.fila, 16).setValue(proyectoId === '' || proyectoId === undefined ? '' : proyectoId); n++; }
  });
  return n;
}

function mesDe_(d) { return Utilities.formatDate(d, zona_(), 'yyyy-MM'); }

function resumenProyecto_(p, movs, tasa) {
  const mios = movs.filter(function (m) { return String(m.proyecto || '') === String(p.id) && (esIngreso_(m.tipo) || esGasto_(m.tipo)) && m.fecha <= new Date(); });
  const meses = {};
  let invertido = 0, costos = 0, ingresos = 0;
  mios.forEach(function (m) {
    const pen = convertirAPEN_(m.monto, m.moneda, tasa);
    const k = mesDe_(m.efectiva);
    const g = meses[k] || (meses[k] = { mes: k, invertido: 0, costos: 0, ingresos: 0 });
    if (esIngreso_(m.tipo)) { ingresos += pen; g.ingresos += pen; }
    else if (m.tipo === 'Gasto Inversion') { invertido += pen; g.invertido += pen; }
    else { costos += pen; g.costos += pen; }
  });
  const claves = Object.keys(meses).sort();
  let acum = 0, gastoAcum = 0, mesRecup = '';
  const serie = claves.map(function (k) {
    const g = meses[k];
    acum += g.ingresos - g.invertido - g.costos;
    gastoAcum += g.invertido + g.costos;
    if (gastoAcum > 0 && acum >= -0.005 && !mesRecup) mesRecup = k;
    if (acum < -0.005) mesRecup = '';
    return { mes: k, invertido: redondear_(g.invertido), costos: redondear_(g.costos), ingresos: redondear_(g.ingresos), acumulado: redondear_(acum) };
  });
  const gasto = invertido + costos;
  const ganancia = ingresos - gasto;
  // Estimación: ritmo neto mensual de los últimos 3 meses calendario (ingresos − costos operativos)
  let estimado = null;
  if (ganancia < 0 && serie.length) {
    const hoy = new Date();
    let ing3 = 0, cos3 = 0;
    for (let k = 0; k < 3; k++) {
      const d = new Date(hoy.getFullYear(), hoy.getMonth() - k, 1), g = meses[mesDe_(d)];
      if (g) { ing3 += g.ingresos; cos3 += g.costos; }
    }
    const neto = (ing3 - cos3) / 3;
    if (neto > 0) {
      const n = Math.ceil(-ganancia / neto);
      const f = new Date(hoy.getFullYear(), hoy.getMonth() + n, 1);
      estimado = { meses: n, mes: mesDe_(f), netoMensual: redondear_(neto) };
    }
  }
  return {
    id: p.id, nombre: p.nombre, tipo: p.tipo, estado: p.estado, notas: p.notas, n: mios.length,
    invertido: redondear_(invertido), costos: redondear_(costos), ingresos: redondear_(ingresos), ganancia: redondear_(ganancia),
    roi: gasto > 0 ? ganancia / gasto * 100 : 0, recuperado: gasto > 0 ? Math.min(999, ingresos / gasto * 100) : 0,
    mesRecuperacion: ganancia >= -0.005 && gasto > 0 ? mesRecup : '', estimado: estimado, serie: serie
  };
}

function api_proyectos() {
  const movs = leerMovimientos_();
  const tasa = obtenerUltimaTasaReferencial_();
  const lista = leerProyectos_().map(function (p) { const r = resumenProyecto_(p, movs, tasa); delete r.serie; return r; });
  const sinProy = movs.filter(function (m) { return m.tipo === 'Gasto Inversion' && !String(m.proyecto || '') && m.fecha <= new Date(); }).length;
  let bolsa = null;
  try {
    const ops = leerOperacionesInversion_();
    if (ops.length) {
      const r = resumenInversiones_(false);
      const ganancia = r.plPEN + r.dividendosPEN + r.realizadoPEN;
      bolsa = { invertido: r.costoPEN, valor: r.valorPEN, ganancia: redondear_(ganancia), roi: r.costoPEN > 0 ? ganancia / r.costoPEN * 100 : 0, dividendos: r.dividendosPEN };
    }
  } catch (e) { bolsa = null; }
  const tot = lista.reduce(function (t, x) { t.invertido += x.invertido + x.costos; t.ingresos += x.ingresos; return t; }, { invertido: 0, ingresos: 0 });
  return { proyectos: lista, bolsa: bolsa, sinProyecto: sinProy, totales: { invertido: redondear_(tot.invertido), ingresos: redondear_(tot.ingresos), ganancia: redondear_(tot.ingresos - tot.invertido) } };
}

function api_proyecto(a) {
  const p = leerProyectos_().filter(function (x) { return String(x.id) === String(a.id); })[0];
  if (!p) throw new Error('Proyecto no encontrado.');
  const movs = leerMovimientos_();
  const r = resumenProyecto_(p, movs, obtenerUltimaTasaReferencial_());
  const nombres = mapaNombresCuentas_();
  r.movimientos = movimientosReales_(movs).filter(function (m) { return String(m.proyecto || '') === String(p.id); })
    .slice(0, 60).map(function (m) { return movimientoParaApp_(m, nombres); });
  return r;
}

function api_guardarProyecto(a) {
  const nombre = String(a.nombre || '').trim();
  if (!nombre) throw new Error('Ponle un nombre al proyecto.');
  const existentes = leerProyectos_();
  if (existentes.some(function (x) { return normNombre_(x.nombre) === normNombre_(nombre) && String(x.id) !== String(a.id); })) throw new Error('Ya tienes un proyecto con ese nombre.');
  const tipo = a.tipo === 'Inversión' ? 'Inversión' : 'Negocio';
  const estado = a.estado === 'Cerrado' ? 'Cerrado' : 'Activo';
  const h = hojaProyectos_();
  if (a.id !== undefined && a.id !== '') {
    const d = h.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) if (String(d[i][0]) === String(a.id)) { h.getRange(i + 1, 2, 1, 4).setValues([[nombre, tipo, estado, String(a.notas || '')]]); return a.id; }
    throw new Error('Proyecto no encontrado.');
  }
  const id = generarID_('Proyectos');
  h.appendRow([id, nombre, tipo, estado, String(a.notas || ''), new Date()]);
  return id;
}

function api_eliminarProyecto(a) {
  const h = hojaProyectos_(), d = h.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]) === String(a.id)) h.deleteRow(i + 1);
  const ids = leerMovimientos_().filter(function (m) { return String(m.proyecto || '') === String(a.id); }).map(function (m) { return m.id; });
  asignarProyecto_(ids, '');
  return true;
}

function api_asignarProyecto(a) {
  const ids = Array.isArray(a.ids) ? a.ids : [];
  if (!ids.length) throw new Error('Elige al menos un movimiento.');
  if (a.proyecto !== '' && a.proyecto !== undefined && !leerProyectos_().some(function (x) { return String(x.id) === String(a.proyecto); })) throw new Error('Proyecto no encontrado.');
  return asignarProyecto_(ids, a.proyecto);
}

// ---------------------------------------------------------------
// EXPORTAR MOVIMIENTOS (para tu contador, Excel, etc.)
// ---------------------------------------------------------------

function csvCelda_(v) {
  const s = String(v === null || v === undefined ? '' : v);
  return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function generarCsvMovimientos_(desde, hasta) {
  const d = fecha_(desde), h = fecha_(hasta);
  const hFin = h ? new Date(h.getFullYear(), h.getMonth(), h.getDate(), 23, 59, 59, 999) : null;
  const nombres = mapaNombresCuentas_();
  const proy = {}; leerProyectos_().forEach(function (p) { proy[String(p.id)] = p.nombre; });
  const filas = movimientosReales_(leerMovimientos_()).filter(function (m) { return (!d || m.efectiva >= d) && (!hFin || m.efectiva <= hFin); })
    .sort(function (a, b) { return a.efectiva - b.efectiva; });
  const cab = ['Fecha', 'Fecha efectiva', 'Tipo', 'Monto', 'Moneda', 'Cuenta', 'Categoria', 'Subcategoria', 'Descripcion', 'Medio de pago', 'Centro de costo', 'Proyecto', 'Notas'];
  const lineas = [cab.join(',')];
  filas.forEach(function (m) {
    lineas.push([fechaISO_(m.fecha), fechaISO_(m.efectiva), m.tipo, m.monto, m.moneda, m.cuentaId !== '' ? (nombres[String(m.cuentaId)] || '') : (m.medio === 'Tarjeta de Credito' ? 'Tarjeta de crédito' : ''),
      m.categoria, m.subcategoria, m.descripcion, m.medio, m.centro, proy[String(m.proyecto || '')] || '', m.notas].map(csvCelda_).join(','));
  });
  return { n: filas.length, csv: '﻿' + lineas.join('\n') };
}

function api_exportarMovimientos(a) {
  const hoy = new Date();
  const desde = a.desde || (hoy.getFullYear() + '-01-01'), hasta = a.hasta || fechaISO_(hoy);
  const r = generarCsvMovimientos_(desde, hasta);
  return { nombre: 'Movimientos_' + desde + '_a_' + hasta + '.csv', n: r.n, csv: r.csv };
}

function api_enviarExportacion(a) {
  const hoy = new Date();
  const desde = a.desde || (hoy.getFullYear() + '-01-01'), hasta = a.hasta || fechaISO_(hoy);
  const r = generarCsvMovimientos_(desde, hasta);
  const ok = enviarDocumentoTelegram(Utilities.newBlob(r.csv, 'text/csv', 'Movimientos_' + desde + '_a_' + hasta + '.csv'), '📤 ' + r.n + ' movimientos del ' + desde + ' al ' + hasta);
  if (!ok) throw new Error('No pude enviarlo por Telegram. Revisa el estado del sistema.');
  return r.n;
}

// ---------------------------------------------------------------
// FONDO DE EMERGENCIA
// ---------------------------------------------------------------

function fondoEmergencia_(movs, tasa) {
  const objetivo = Math.max(1, numero_(getAjuste_('fondo.meses', '6')) || 6);
  const hoy = new Date();
  let gastos = 0;
  const ini = new Date(hoy.getFullYear(), hoy.getMonth() - 3, 1), fin = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  movs.forEach(function (m) { if (esGasto_(m.tipo) && m.efectiva >= ini && m.efectiva < fin) gastos += convertirAPEN_(m.monto, m.moneda, tasa); });
  const mensual = gastos / 3;
  let liquido = 0;
  const cu = leerHoja_('Cuentas');
  for (let i = 1; i < cu.length; i++) {
    if (cu[i][0] === '' || String(cu[i][3]).trim() === 'Inversion' || String(cu[i][6]).trim().toUpperCase() === 'NO') continue;
    liquido += convertirAPEN_(numero_(cu[i][5]), String(cu[i][1]).trim(), tasa);
  }
  return { objetivoMeses: objetivo, gastoMensual: redondear_(mensual), liquido: redondear_(liquido), meses: mensual > 0 ? Math.round(liquido / mensual * 10) / 10 : 0, meta: redondear_(mensual * objetivo) };
}

function api_guardarFondo(a) {
  const m = numero_(a.meses);
  if (!(m >= 1 && m <= 36)) throw new Error('Elige entre 1 y 36 meses.');
  setAjuste_('fondo.meses', m);
  return true;
}


/* ==== 26_Respaldo.gs ==== */
/**
 * ============================================================
 * COPIA DE SEGURIDAD SEMANAL + RESUMEN SEMANAL
 * ============================================================
 * Cada domingo (y cuando tú lo pidas) se guarda una copia completa de tu hoja de cálculo en la carpeta
 * "Respaldos Finanzas" de tu Google Drive. Se conservan las últimas 8 copias.
 */

var CARPETA_RESPALDOS_ = 'Respaldos Finanzas';
var MAX_RESPALDOS_ = 8;

function carpetaRespaldos_() {
  const it = DriveApp.getFoldersByName(CARPETA_RESPALDOS_);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RESPALDOS_);
}

function listarRespaldos_() {
  const out = [];
  const it = carpetaRespaldos_().getFiles();
  while (it.hasNext()) { const f = it.next(); out.push({ id: f.getId(), nombre: f.getName(), fecha: fechaHoraISO_(f.getDateCreated()), url: f.getUrl() }); }
  return out.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
}

function hacerRespaldo_() {
  const nombre = 'Finanzas respaldo ' + Utilities.formatDate(new Date(), zona_(), 'yyyy-MM-dd HHmm');
  const copia = DriveApp.getFileById(getSS_().getId()).makeCopy(nombre, carpetaRespaldos_());
  const todas = listarRespaldos_();
  todas.slice(MAX_RESPALDOS_).forEach(function (f) { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) { /* ya no existe */ } });
  registrarLog_('Respaldo', 'Copia creada: ' + nombre);
  return { id: copia.getId(), nombre: nombre, url: copia.getUrl() };
}

/** Trigger semanal: domingo 3am */
function respaldoSemanal() {
  ejecutarSeguro_('respaldoSemanal', function () { hacerRespaldo_(); });
}

function api_respaldos() { return listarRespaldos_(); }
function api_respaldoAhora() { return hacerRespaldo_(); }

// ---------------------------------------------------------------
// RESUMEN SEMANAL (Telegram, domingo 9pm junto a presupuestos)
// ---------------------------------------------------------------

function resumenSemanalTexto_() {
  const tasa = obtenerUltimaTasaReferencial_();
  const movs = leerMovimientos_();
  const ahora = new Date();
  const ini = new Date(ahora.getTime() - 7 * 86400000), iniPrev = new Date(ahora.getTime() - 14 * 86400000);
  let ing = 0, gas = 0, gasPrev = 0;
  const cats = {};
  movs.forEach(function (m) {
    if (m.efectiva > ahora) return;
    const pen = convertirAPEN_(m.monto, m.moneda, tasa);
    if (m.efectiva >= ini) {
      if (esIngreso_(m.tipo)) ing += pen;
      else if (esGasto_(m.tipo)) { gas += pen; const c = String(m.categoria || '').trim() || 'Sin categoría'; cats[c] = (cats[c] || 0) + pen; }
    } else if (m.efectiva >= iniPrev && esGasto_(m.tipo)) gasPrev += pen;
  });
  if (!ing && !gas) return '';
  const top = Object.keys(cats).sort(function (a, b) { return cats[b] - cats[a]; }).slice(0, 3);
  const dif = gasPrev > 0 ? (gas - gasPrev) / gasPrev * 100 : null;
  let t = '📊 *Tu semana*\n💰 Ingresos: ' + formatoMonto_(redondear_(ing), 'PEN') + '\n💸 Gastos: ' + formatoMonto_(redondear_(gas), 'PEN') +
    (dif === null ? '' : ' (' + (dif >= 0 ? '+' : '') + Math.round(dif) + '% vs semana anterior)') + '\n';
  if (top.length) t += '🏷️ Donde más gastaste: ' + top.map(function (c) { return escaparMd_(c) + ' ' + formatoMonto_(redondear_(cats[c]), 'PEN'); }).join(' · ') + '\n';
  const prox = listarProgramados_('Pendiente').concat(proximosRecurrentes_()).filter(function (p) { const f = fecha_(p.fecha); return f && f <= new Date(ahora.getTime() + 7 * 86400000); });
  if (prox.length) {
    t += '📅 Próximos 7 días: ' + prox.length + ' movimiento' + (prox.length === 1 ? '' : 's') + ' (' + prox.slice(0, 3).map(function (p) { return escaparMd_(p.descripcion || p.tipo); }).join(', ') + (prox.length > 3 ? '…' : '') + ')\n';
  }
  const f = fondoEmergencia_(movs, tasa);
  if (f.gastoMensual > 0) t += '🛟 Fondo de emergencia: ' + f.meses + ' de ' + f.objetivoMeses + ' meses cubiertos.';
  return t;
}


/* ==== 27_Plan.gs ==== */
/**
 * ============================================================
 * PLAN DEL MES y METAS DE AHORRO
 * ============================================================
 * Plan: reparte tus ingresos entre gasto personal, inversión y ahorro (los % se editan en la app) y muestra
 * cuánto llevas gastado y cuánto te queda, en total y por categoría (con el presupuesto de cada una si lo definiste).
 * Base del plan = lo mayor entre lo que ya ingresó este mes y tu promedio de los últimos 3 meses
 * (así el plan es útil desde el día 1 del mes, antes de cobrar).
 *
 * Metas de ahorro: dinero que apartas para algo (viaje, laptop, fondo...). Opcionalmente ligadas a una categoría:
 * lo que gastes en esa categoría desde que creaste la meta se descuenta de lo ahorrado.
 */

function deAPEN_(pen, moneda, tasa) {
  if (moneda === 'USD') return pen / tasa.usdPen;
  if (moneda === 'EUR') return pen / tasa.eurPen;
  return pen;
}

function api_plan() {
  const movs = leerMovimientos_();
  const tasa = obtenerUltimaTasaReferencial_();
  const hoy = new Date();
  const o = objetivosIngreso_();
  const rango = obtenerRangoFecha_('Mensual', hoy);
  const d = distribucionIngresos_('Mensual', hoy, movs, tasa);

  // promedio de ingresos de los 3 meses anteriores (solo meses con ingresos)
  let suma = 0, n = 0;
  for (let k = 1; k <= 3; k++) {
    const x = distribucionIngresos_('Mensual', new Date(hoy.getFullYear(), hoy.getMonth() - k, 15), movs, tasa).ingresos;
    if (x > 0) { suma += x; n++; }
  }
  const promedio = n ? suma / n : 0;
  const base = Math.max(d.ingresos, promedio);

  // gasto personal y de inversión por categoría (en soles)
  const pers = {}, inv = {};
  let bolsaNeta = 0;
  movs.forEach(function (m) {
    if (m.efectiva < rango.inicio || m.efectiva > rango.fin) return;
    const eq = convertirAPEN_(m.monto, m.moneda, tasa);
    const cat = String(m.categoria || '').trim() || 'Sin categoría';
    if (m.tipo === 'Gasto Personal' || m.tipo === 'Gasto Fijo') pers[cat] = (pers[cat] || 0) + eq;
    else if (m.tipo === 'Gasto Inversion') inv[cat] = (inv[cat] || 0) + eq;
    else if (m.tipo === 'Transferencia' && m.categoria === 'Inversiones') bolsaNeta += -eq;
  });
  if (bolsaNeta > 0) inv['Compras en bolsa'] = (inv['Compras en bolsa'] || 0) + bolsaNeta;

  // presupuestos del mes por categoría (en soles)
  const pres = {};
  obtenerEstadoPresupuestos_(hoy, movs).forEach(function (p) {
    pres[p.categoria] = (pres[p.categoria] || 0) + convertirAPEN_(p.limite, p.moneda, tasa);
  });

  const filas = function (mapa, conPres) {
    const claves = {};
    Object.keys(mapa).forEach(function (c) { claves[c] = true; });
    if (conPres) Object.keys(pres).forEach(function (c) { claves[c] = true; });
    return Object.keys(claves).map(function (c) {
      const g = mapa[c] || 0, p = conPres ? (pres[c] || 0) : 0;
      return { categoria: c, gastado: redondear_(g), presupuesto: redondear_(p), queda: p > 0 ? redondear_(p - g) : null };
    }).sort(function (a, b) { return b.gastado - a.gastado; });
  };

  const asig = function (pct) { return redondear_(base * pct / 100); };
  const diasRest = Math.max(1, Math.round((rango.fin.getTime() - hoy.getTime()) / 86400000) + 1);
  const topeGasto = asig(o.gasto);
  const efectivo = Math.max(0, d.efectivo);

  return {
    base: redondear_(base), baseTipo: d.ingresos >= promedio ? 'mes' : 'promedio', ingresosMes: d.ingresos, promedio: redondear_(promedio),
    objetivos: o, diasRestantes: diasRest,
    gasto: { asignado: topeGasto, gastado: d.personal, queda: redondear_(topeGasto - d.personal), porDia: redondear_(Math.max(0, topeGasto - d.personal) / diasRest), categorias: filas(pers, true) },
    inversion: { asignado: asig(o.inversion), gastado: d.inversion, queda: redondear_(asig(o.inversion) - d.inversion), categorias: filas(inv, false) },
    ahorro: { asignado: asig(o.ahorro), logrado: redondear_(efectivo), falta: redondear_(Math.max(0, asig(o.ahorro) - efectivo)), total: d.ahorro, pctTotal: d.pctAhorro, invertido: d.inversion, efectivo: d.efectivo }
  };
}

// ---------------------------------------------------------------
// METAS DE AHORRO
// ---------------------------------------------------------------

function hojaMetasAhorro_() {
  const ss = getSS_();
  let h = ss.getSheetByName('MetasAhorro');
  if (!h) { h = ss.insertSheet('MetasAhorro'); h.getRange(1, 1, 1, 8).setValues([['ID', 'Nombre', 'Moneda', 'Objetivo', 'Ahorrado', 'Categoria', 'FechaLimite', 'Creada']]); h.setFrozenRows(1); }
  return h;
}

function api_metasAhorro() {
  const d = hojaMetasAhorro_().getDataRange().getValues();
  const movs = leerMovimientos_();
  const tasa = obtenerUltimaTasaReferencial_();
  const hoy = new Date();
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const f = d[i];
    if (f[0] === '') continue;
    const moneda = String(f[2] || 'PEN').trim();
    const objetivo = numero_(f[3]), ahorrado = numero_(f[4]);
    const cat = String(f[5] || '').trim();
    const creada = fecha_(f[7]);
    const limite = fecha_(f[6]);
    let gastado = 0;
    if (cat) {
      movs.forEach(function (m) {
        if (esGasto_(m.tipo) && String(m.categoria || '').trim() === cat && (!creada || m.efectiva >= creada)) gastado += deAPEN_(convertirAPEN_(m.monto, m.moneda, tasa), moneda, tasa);
      });
    }
    gastado = redondear_(gastado);
    const neto = redondear_(ahorrado - gastado);
    const falta = Math.max(0, redondear_(objetivo - neto));
    let aporteMensual = 0, meses = 0;
    if (limite && limite > hoy && falta > 0) {
      meses = Math.max(1, Math.ceil((limite.getTime() - hoy.getTime()) / (86400000 * 30.4)));
      aporteMensual = redondear_(falta / meses);
    }
    out.push({
      id: f[0], nombre: String(f[1]), moneda: moneda, objetivo: objetivo, ahorrado: redondear_(ahorrado), categoria: cat, fechaLimite: fechaISO_(f[6]),
      gastado: gastado, neto: neto, falta: falta, porcentaje: objetivo > 0 ? Math.max(0, neto) / objetivo * 100 : 0, aporteMensual: aporteMensual, meses: meses
    });
  }
  return out;
}

function api_guardarMetaAhorro(a) {
  const nombre = String(a.nombre || '').trim();
  if (!nombre) throw new Error('Escribe el nombre de la meta.');
  const moneda = validarMoneda_(a.moneda || 'PEN');
  const objetivo = validarMonto_(a.objetivo);
  const cat = String(a.categoria || '').trim();
  const limite = a.fechaLimite ? fecha_(a.fechaLimite) : '';
  if (a.fechaLimite && !limite) throw new Error('Fecha límite inválida.');
  return conLock_(function () {
    const h = hojaMetasAhorro_();
    if (a.id) {
      const d = h.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (String(d[i][0]) === String(a.id)) {
          h.getRange(i + 1, 2, 1, 3).setValues([[nombre, moneda, objetivo]]);
          h.getRange(i + 1, 6, 1, 2).setValues([[cat, limite]]);
          return true;
        }
      }
      throw new Error('Meta no encontrada.');
    }
    const id = generarID_('MetasAhorro');
    const ini = a.ahorrado === undefined || a.ahorrado === '' ? 0 : Math.max(0, numero_(a.ahorrado));
    h.appendRow([id, nombre, moneda, objetivo, ini, cat, limite, new Date()]);
    return id;
  });
}

function api_eliminarMetaAhorro(a) {
  return conLock_(function () {
    const h = hojaMetasAhorro_();
    const d = h.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]) === String(a.id)) { h.deleteRow(i + 1); return true; }
    throw new Error('Meta no encontrada.');
  });
}

/** Suma (aportar) o resta (retirar) dinero apartado. monto con signo */
function api_moverMetaAhorro(a) {
  const monto = numero_(a.monto);
  if (!monto) throw new Error('Escribe un monto.');
  return conLock_(function () {
    const h = hojaMetasAhorro_();
    const d = h.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]) === String(a.id)) {
        const nuevo = redondear_(numero_(d[i][4]) + monto);
        if (nuevo < 0) throw new Error('No puedes retirar más de lo ahorrado.');
        h.getRange(i + 1, 5).setValue(nuevo);
        return nuevo;
      }
    }
    throw new Error('Meta no encontrada.');
  });
}

/* ==== 30_DividirGastos.gs ==== */
/**
 * ============================================================
 * DIVIDIR GASTOS
 * ============================================================
 * Cuando pagas tú una cuenta compartida (ej: salen 4 y pagas 20 soles), solo TU parte es gasto tuyo.
 * La parte de los demás queda como "por cobrar" en Deudas (Préstamos otorgados, sin interés) a nombre de cada persona.
 * Así no parece que gastaste de más, y cuando te devuelven la plata entra como devolución (no como ingreso).
 *
 *  - Efectivo/débito: el movimiento baja a tu parte y cada persona queda como deuda (el dinero que salió de tu cuenta
 *    sigue siendo el total: una parte es gasto, el resto es "por cobrar").
 *  - Tarjeta: el estado de cuenta sigue cobrando el total (así lo cobra el banco); el gasto reconocido es tu parte y el
 *    resto queda por cobrar.
 *
 * Marcas internas: el movimiento lleva en Notas "Dividido #ID (total X)" y cada deuda guarda el ID del movimiento
 * en la columna "Division" (col. 14 de PrestamosOtorgados).
 *
 * a = { modo: 'igual'|'porcentaje'|'monto', personas: [{ nombre, pct?, monto? }] }
 *   igual: el total se reparte en partes iguales entre tú y las personas (centavos sobrantes para ti).
 *   porcentaje: cada persona paga pct % del total; tú te quedas con el resto.
 *   monto: cada persona devuelve exactamente ese monto; tú te quedas con el resto.
 */

const RE_DIVIDIDO_ = /Dividido #(\w+) \(total ([\d.]+)\)/;

function asegurarColumnaDivision_() {
  asegurarColumnaConcepto_();
  const h = getHoja_('PrestamosOtorgados');
  if (h.getLastColumn() < 14) h.getRange(1, 14).setValue('Division');
}

/** "ana  perez" → "Ana Perez" (solo si lo escribieron todo en minúsculas; si ya trae mayúsculas se respeta) */
function nombrePropio_(n) {
  const t = String(n || '').trim().replace(/\s+/g, ' ');
  if (t !== t.toLowerCase()) return t;
  return t.replace(/(^|\s)(\S)/g, function (m, a, b) { return a + b.toUpperCase(); });
}

/** Calcula cuánto le toca a cada uno. Devuelve { yo, otros, personas: [{ nombre, monto }] } */
function calcularDivision_(total, a) {
  const personas = ((a && a.personas) || []).map(function (p) {
    return { nombre: nombrePropio_(p.nombre), pct: p.pct, monto: p.monto };
  });
  if (!personas.length) throw new Error('Agrega al menos una persona para dividir el gasto.');
  const vistos = {};
  personas.forEach(function (p) {
    if (!p.nombre) throw new Error('Escribe el nombre de cada persona.');
    const k = normNombre_(p.nombre);
    if (vistos[k]) throw new Error('"' + p.nombre + '" aparece dos veces.');
    vistos[k] = true;
  });
  const modo = a.modo === 'porcentaje' || a.modo === 'monto' ? a.modo : 'igual';
  let otros = 0;
  if (modo === 'igual') {
    const base = Math.floor(total / (personas.length + 1) * 100 + 1e-9) / 100;
    personas.forEach(function (p) { p.monto = redondear_(base); otros += p.monto; });
  } else {
    personas.forEach(function (p) {
      if (modo === 'porcentaje') {
        const pc = numero_(p.pct);
        if (!(pc > 0) || pc >= 100) throw new Error('El porcentaje de ' + p.nombre + ' debe estar entre 0 y 100.');
        p.monto = redondear_(total * pc / 100);
      } else {
        p.monto = redondear_(numero_(p.monto));
      }
      if (!(p.monto > 0)) throw new Error('Escribe cuánto te debe devolver ' + p.nombre + '.');
      otros += p.monto;
    });
  }
  otros = redondear_(otros);
  const yo = redondear_(total - otros);
  if (yo < 0) throw new Error('Lo que te deben (' + otros + ') supera el total del gasto (' + total + ').');
  if (yo === 0) throw new Error('Con esos montos tu parte queda en 0. Si el gasto fue entero para otra persona usa "Es para otra persona".');
  return { yo: yo, otros: otros, personas: personas.map(function (p) { return { nombre: p.nombre, monto: p.monto }; }) };
}

function movimientoPorId_(id) {
  const m = leerMovimientos_().filter(function (x) { return String(x.id) === String(id); })[0];
  if (!m) throw new Error('Movimiento no encontrado.');
  return m;
}

/** Aplica la división a un gasto ya registrado. Devuelve el resumen */
function dividirMovimientoInterno_(movId, a) {
  return conLock_(function () {
    const m = movimientoPorId_(movId);
    const clase = clasificarMovimiento_(m);
    if (!esGasto_(m.tipo) || (clase !== 'normal' && clase !== 'tarjeta')) throw new Error('Solo se pueden dividir gastos normales o con tarjeta.');
    if (RE_DIVIDIDO_.test(m.notas)) throw new Error('Este gasto ya está dividido.');
    if (/Cuotas #\w+ \(\d+\/\d+\)/.test(m.notas)) throw new Error('Un gasto en cuotas no se puede dividir.');
    if (clase === 'tarjeta') {
      const tid = extraerRef_(m.notas, /Ref\. TarjetaCredito #(\d+)/);
      const tarj = tid ? buscarFilaPorId_('TarjetaCredito', tid) : null;
      if (tarj && String(tarj.fila[7]).trim() === 'Pagado') throw new Error('Ese consumo ya se pagó con un estado de cuenta: no se puede dividir.');
    }
    const total = m.monto;
    const plan = calcularDivision_(total, a);
    asegurarColumnaDivision_();
    const hoja = getHoja_('Movimientos');
    const marca = 'Dividido #' + m.id + ' (total ' + total + ')';

    // 1) Mi parte: el gasto baja de total a lo mío. En efectivo/débito se devuelve a la cuenta lo ajeno,
    //    que enseguida sale de nuevo como deuda de cada persona (el saldo de la cuenta no cambia).
    if (clase === 'normal') actualizarSaldoCuenta_(m.cuentaId, plan.otros);
    hoja.getRange(m.fila, M.MONTO + 1).setValue(plan.yo);
    hoja.getRange(m.fila, M.NOTAS + 1).setValue(m.notas ? m.notas + ' · ' + marca : marca);

    // 2) Lo ajeno: una deuda por persona
    plan.personas.forEach(function (p) {
      const persona = canonPersona_(p.nombre);
      const id = registrarPrestamoOtorgado({
        deudor: persona, moneda: m.moneda, montoPrestado: p.monto, tasaInteres: 0, plazoMeses: 1, fechaInicio: m.fecha,
        concepto: m.descripcion || 'Gasto compartido', cuentaOrigenId: clase === 'normal' ? m.cuentaId : ''
      });
      const fila = buscarFilaPorId_('PrestamosOtorgados', id);
      if (fila) getHoja_('PrestamosOtorgados').getRange(fila.numeroFila, 14).setValue(String(m.id));
      p.nombre = persona;
    });
    registrarLog_('DividirGasto', '#' + m.id + ': total ' + total + ', mi parte ' + plan.yo + ', ' + plan.personas.length + ' persona(s)');
    return { idMov: m.id, total: total, yo: plan.yo, personas: plan.personas };
  });
}

/** Divide un gasto que ya está registrado. args = { id, modo, personas } */
function api_registrarDivision(a) {
  return dividirMovimientoInterno_(a.id, a);
}

/**
 * Registra un gasto y lo divide en un solo paso. args = los de api_registrar + { division: { modo, personas } }
 * Si algo falla al dividir, el gasto recién creado se deshace (no queda a medias).
 */
function api_registrarDividido(a) {
  if (a.programar) throw new Error('Un gasto programado no se puede dividir: divídelo cuando se realice.');
  if (a.clase && a.clase !== 'gasto') throw new Error('Solo se pueden dividir gastos.');
  if (Number(a.cuotas) > 1) throw new Error('Un gasto en cuotas no se puede dividir.');
  const div = a.division || {};
  const total = validarMonto_(a.monto, 'El monto');
  calcularDivision_(total, div); // valida antes de registrar nada
  const args = {};
  Object.keys(a).forEach(function (k) { args[k] = a[k]; });
  args.clase = 'gasto';
  const r = api_registrar(args);
  if (r.programado || r.idMov === '' || r.idMov === undefined) throw new Error('Este gasto no se puede dividir.');
  try {
    const d = dividirMovimientoInterno_(r.idMov, div);
    return { idMov: r.idMov, total: d.total, yo: d.yo, personas: d.personas };
  } catch (e) {
    try { eliminarMovimiento(r.idMov); } catch (x) { /* ignorar */ }
    throw e;
  }
}

/** Deudas (préstamos otorgados) que nacieron de la división de un movimiento */
function deudasDeDivision_(movId) {
  const out = [];
  const po = leerHoja_('PrestamosOtorgados');
  for (let i = 1; i < po.length; i++) {
    if (po[i][0] === '' || po[i].length < 14 || String(po[i][13]) !== String(movId)) continue;
    out.push({ id: po[i][0], nombre: String(po[i][1]), moneda: String(po[i][2]), monto: numero_(po[i][3]), saldo: numero_(po[i][10]), estado: String(po[i][9]).trim() });
  }
  return out;
}

/** Quita las deudas de una división (solo si nadie ha devuelto nada todavía) */
function quitarDeudasDivision_(movId) {
  const deudas = deudasDeDivision_(movId);
  deudas.forEach(function (d) {
    if (d.estado === 'Pagado' || Math.abs(d.saldo - d.monto) > 0.005) throw new Error(d.nombre + ' ya te devolvió parte de este gasto. Primero revierte ese cobro o cierra la deuda.');
  });
  deudas.forEach(function (d) {
    leerMovimientos_().filter(function (x) { return new RegExp('Ref\\. PrestamoOtorgado #' + String(d.id) + '(\\D|$)').test(x.notas); })
      .forEach(function (x) { borrarFilaMovimiento_(x.id); });
    const f = buscarFilaPorId_('PrestamosOtorgados', d.id);
    if (f) getHoja_('PrestamosOtorgados').deleteRow(f.numeroFila);
  });
  return deudas.length;
}

/** Deshace la división: el gasto vuelve a ser del total y se borran las deudas. args = { id } */
function api_eliminarDivision(a) {
  return conLock_(function () {
    const m = movimientoPorId_(a.id);
    const dv = String(m.notas).match(RE_DIVIDIDO_);
    if (!dv) throw new Error('Este gasto no está dividido.');
    const clase = clasificarMovimiento_(m);
    const total = numero_(dv[2]);
    const otros = redondear_(total - m.monto);
    quitarDeudasDivision_(m.id);
    if (clase === 'normal') actualizarSaldoCuenta_(m.cuentaId, -otros);
    const hoja = getHoja_('Movimientos');
    hoja.getRange(m.fila, M.MONTO + 1).setValue(total);
    hoja.getRange(m.fila, M.NOTAS + 1).setValue(String(m.notas).replace(/ · Dividido #\w+ \(total [\d.]+\)/, '').replace(RE_DIVIDIDO_, '').trim());
    registrarLog_('DeshacerDivision', '#' + m.id);
    return true;
  });
}

/** Para el detalle del movimiento: total, mi parte y cada persona con su estado */
function detalleDivision_(m) {
  const dv = String(m.notas).match(RE_DIVIDIDO_);
  if (!dv) return null;
  return { total: numero_(dv[2]), yo: m.monto, moneda: m.moneda, personas: deudasDeDivision_(m.id) };
}

/* ==== 31_AnalistaIA.gs ==== */
/**
 * ============================================================
 * ANALISTA FINANCIERO IA  +  PERFIL FINANCIERO
 * ============================================================
 * Arquitectura (la IA NO hace las cuentas):
 *   1) La IA lee tu pregunta y decide qué herramientas de cálculo necesita (JSON).
 *   2) Las herramientas (este archivo) calculan con tus movimientos reales: sumas, promedios, comparaciones,
 *      anomalías, recurrentes, proyecciones y simulaciones. Son exactas y reproducibles.
 *   3) La IA recibe SOLO esos resultados y los explica en lenguaje claro, sin inventar cifras.
 * Son 2 llamadas a la IA por pregunta (cabe en la cuota gratuita de Gemini).
 *
 * Importes: todo se compara en soles equivalentes (tipo de cambio referencial actual), salvo que la herramienta diga otra cosa.
 * El perfil financiero (api_perfil) usa reglas transparentes —no un puntaje inventado— y la IA solo lo redacta.
 * Este módulo solo LEE datos: nunca crea, edita ni borra nada.
 */

// ---------------------------------------------------------------
// BASE DE DATOS DE ANÁLISIS
// ---------------------------------------------------------------

var ANA_ = null;

/** Movimientos reales (ya ocurridos) de ingreso/gasto, normalizados y en soles equivalentes */
function anaBase_() {
  if (ANA_) return ANA_;
  const tasa = obtenerUltimaTasaReferencial_();
  const ahora = new Date();
  const nombres = {};
  try { obtenerCuentasActivas_().concat(cuentasArchivadas_()).forEach(function (c) { nombres[String(c.id)] = c.banco; }); } catch (e) { /* opcional */ }
  const lista = [];
  leerMovimientos_().forEach(function (m) {
    if (m.efectiva > ahora) return;
    const g = esGasto_(m.tipo), i = esIngreso_(m.tipo);
    if (!g && !i) return;
    lista.push({
      id: m.id, f: m.efectiva, ing: i, tipo: m.tipo, cat: String(m.categoria || 'Sin categoría'), sub: String(m.subcategoria || ''),
      desc: String(m.descripcion || ''), medio: m.medio || '', cuenta: m.cuentaId !== '' ? (nombres[String(m.cuentaId)] || '') : (m.medio === 'Tarjeta de Credito' ? 'Tarjeta' : ''),
      pen: convertirAPEN_(m.monto, m.moneda, tasa), monto: m.monto, mon: m.moneda, dividido: RE_DIVIDIDO_.test(m.notas)
    });
  });
  ANA_ = { lista: lista, tasa: tasa, ahora: ahora, nombres: nombres };
  return ANA_;
}

function anaN_(n) { return Math.round(Number(n) * 100) / 100; }
function anaIso_(d) { return fechaISO_(d); }
function anaInicioMes_(d, delta) { return new Date(d.getFullYear(), d.getMonth() + (delta || 0), 1); }
function anaFinMes_(d, delta) { return new Date(d.getFullYear(), d.getMonth() + (delta || 0) + 1, 0, 23, 59, 59); }
function anaClaveMes_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }
function anaNorm_(t) { return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim(); }

/** Rango de fechas desde args: periodo (atajo) o desde/hasta (AAAA-MM-DD). Por defecto, últimos 90 días. */
function anaRango_(a) {
  const hoy = new Date();
  const fin0 = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate(), 23, 59, 59);
  const p = String((a && a.periodo) || '').toLowerCase();
  const dias = function (n) { return { desde: new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - n + 1), hasta: fin0, etiqueta: 'últimos ' + n + ' días' }; };
  if (a && a.desde) {
    const d = fecha_(a.desde), h = a.hasta ? fecha_(a.hasta) : null;
    if (d) {
      const hasta = h ? new Date(h.getFullYear(), h.getMonth(), h.getDate(), 23, 59, 59) : fin0;
      return { desde: new Date(d.getFullYear(), d.getMonth(), d.getDate()), hasta: hasta, etiqueta: anaIso_(d) + ' a ' + anaIso_(hasta) };
    }
  }
  if (p === 'este_mes') return { desde: anaInicioMes_(hoy), hasta: fin0, etiqueta: 'este mes' };
  if (p === 'mes_pasado') return { desde: anaInicioMes_(hoy, -1), hasta: anaFinMes_(hoy, -1), etiqueta: 'mes pasado' };
  if (p === 'ultimos_7') return dias(7);
  if (p === 'ultimos_30') return dias(30);
  if (p === 'ultimos_180') return dias(180);
  if (p === 'ultimos_365') return dias(365);
  if (p === 'este_ano') return { desde: new Date(hoy.getFullYear(), 0, 1), hasta: fin0, etiqueta: 'este año' };
  if (p === 'ano_pasado') return { desde: new Date(hoy.getFullYear() - 1, 0, 1), hasta: new Date(hoy.getFullYear() - 1, 11, 31, 23, 59, 59), etiqueta: 'año pasado' };
  if (p === 'todo') return { desde: new Date(2000, 0, 1), hasta: fin0, etiqueta: 'todo el historial' };
  return dias(90);
}

function anaFiltrar_(a, rango) {
  const base = anaBase_().lista;
  const tipo = String((a && a.tipo) || '').toLowerCase();
  const cat = a && a.categoria ? anaNorm_(a.categoria) : '';
  const sub = a && a.subcategoria ? anaNorm_(a.subcategoria) : '';
  const txt = a && a.texto ? anaNorm_(a.texto) : '';
  const cuenta = a && a.cuenta ? anaNorm_(a.cuenta) : '';
  const mn = a && a.monto_min !== undefined && a.monto_min !== '' ? Number(a.monto_min) : null;
  const mx = a && a.monto_max !== undefined && a.monto_max !== '' ? Number(a.monto_max) : null;
  return base.filter(function (m) {
    if (m.f < rango.desde || m.f > rango.hasta) return false;
    if (tipo === 'gasto' && m.ing) return false;
    if (tipo === 'ingreso' && !m.ing) return false;
    if (cat && anaNorm_(m.cat).indexOf(cat) === -1) return false;
    if (sub && anaNorm_(m.sub).indexOf(sub) === -1) return false;
    if (txt && (anaNorm_(m.desc) + ' ' + anaNorm_(m.cat) + ' ' + anaNorm_(m.sub)).indexOf(txt) === -1) return false;
    if (cuenta && anaNorm_(m.cuenta).indexOf(cuenta) === -1) return false;
    if (mn !== null && m.pen < mn) return false;
    if (mx !== null && m.pen > mx) return false;
    return true;
  });
}

function anaSuma_(l) { return l.reduce(function (s, m) { return s + m.pen; }, 0); }
function anaMediana_(v) { if (!v.length) return 0; const s = v.slice().sort(function (a, b) { return a - b; }); const k = Math.floor(s.length / 2); return s.length % 2 ? s[k] : (s[k - 1] + s[k]) / 2; }
function anaMedia_(v) { return v.length ? v.reduce(function (s, x) { return s + x; }, 0) / v.length : 0; }
function anaDesv_(v) { if (v.length < 2) return 0; const m = anaMedia_(v); return Math.sqrt(v.reduce(function (s, x) { return s + Math.pow(x - m, 2); }, 0) / v.length); }

var DIAS_ES_ = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

function anaClave_(m, por) {
  if (por === 'subcategoria') return m.cat + ' › ' + (m.sub || '(sin subcategoría)');
  if (por === 'cuenta') return m.cuenta || '(sin cuenta)';
  if (por === 'medio') return m.medio || '(sin medio)';
  if (por === 'dia_semana') return DIAS_ES_[m.f.getDay()];
  if (por === 'mes') return anaClaveMes_(m.f);
  if (por === 'descripcion') return anaNorm_(m.desc) || '(sin descripción)';
  if (por === 'tipo') return m.tipo;
  return m.cat;
}

/** Totales por mes de un filtro: { 'AAAA-MM': {ingresos, gastos} } para los últimos n meses completos + el actual */
function anaPorMes_(n) {
  const hoy = new Date(), out = [], idx = {};
  for (let k = n - 1; k >= 0; k--) {
    const d = anaInicioMes_(hoy, -k);
    const o = { mes: anaClaveMes_(d), ingresos: 0, gastos: 0, parcial: k === 0 };
    idx[o.mes] = o; out.push(o);
  }
  anaBase_().lista.forEach(function (m) {
    const o = idx[anaClaveMes_(m.f)];
    if (!o) return;
    if (m.ing) o.ingresos += m.pen; else o.gastos += m.pen;
  });
  out.forEach(function (o) { o.ingresos = anaN_(o.ingresos); o.gastos = anaN_(o.gastos); o.ahorro = anaN_(o.ingresos - o.gastos); });
  return out;
}

// ---------------------------------------------------------------
// HERRAMIENTAS (todas devuelven JSON pequeño y exacto)
// ---------------------------------------------------------------

function toolResumen_(a) {
  const r = anaRango_(a), l = anaFiltrar_(a, r);
  const ing = l.filter(function (m) { return m.ing; }), gas = l.filter(function (m) { return !m.ing; });
  const I = anaSuma_(ing), G = anaSuma_(gas);
  const dias = Math.max(1, Math.round((Math.min(r.hasta, new Date()) - r.desde) / 86400000) + 1);
  return {
    periodo: r.etiqueta, desde: anaIso_(r.desde), hasta: anaIso_(r.hasta < new Date() ? r.hasta : new Date()),
    ingresos: anaN_(I), gastos: anaN_(G), ahorro: anaN_(I - G), tasaAhorroPct: I > 0 ? anaN_((I - G) / I * 100) : null,
    movimientosIngreso: ing.length, movimientosGasto: gas.length,
    gastoPromedioDiario: anaN_(G / dias), gastoPromedioPorMovimiento: gas.length ? anaN_(G / gas.length) : 0,
    ingresoPasivoPct: I > 0 ? anaN_(anaSuma_(ing.filter(function (m) { return m.tipo === 'Ingreso Pasivo'; })) / I * 100) : null,
    nota: 'Importes en soles equivalentes.'
  };
}

function toolAgrupar_(a) {
  const r = anaRango_(a), l = anaFiltrar_(a, r);
  const por = String(a.agrupar || 'categoria');
  const g = {};
  l.forEach(function (m) { const k = anaClave_(m, por); if (!g[k]) g[k] = { total: 0, n: 0, max: 0 }; g[k].total += m.pen; g[k].n++; g[k].max = Math.max(g[k].max, m.pen); });
  const total = anaSuma_(l);
  let filas = Object.keys(g).map(function (k) { return { grupo: k, total: anaN_(g[k].total), movimientos: g[k].n, promedio: anaN_(g[k].total / g[k].n), mayor: anaN_(g[k].max), pct: total > 0 ? anaN_(g[k].total / total * 100) : 0 }; });
  if (por === 'mes') filas.sort(function (x, y) { return x.grupo < y.grupo ? -1 : 1; }); else filas.sort(function (x, y) { return y.total - x.total; });
  const top = Math.min(Math.max(Number(a.top) || 12, 1), 30);
  return { periodo: r.etiqueta, agrupadoPor: por, tipo: a.tipo || 'todos', totalGeneral: anaN_(total), movimientos: l.length, filas: filas.slice(0, top), gruposOmitidos: Math.max(0, filas.length - top) };
}

function toolSerieMensual_(a) {
  const n = Math.min(Math.max(Number(a.meses) || 12, 2), 36);
  const s = anaPorMes_(n);
  const completos = s.filter(function (x) { return !x.parcial; });
  return {
    meses: s, nota: 'El último mes es parcial (aún en curso).',
    promedioIngresos: anaN_(anaMedia_(completos.map(function (x) { return x.ingresos; }))),
    promedioGastos: anaN_(anaMedia_(completos.map(function (x) { return x.gastos; }))),
    mejorMes: completos.length ? completos.slice().sort(function (x, y) { return y.ahorro - x.ahorro; })[0].mes : null,
    peorMes: completos.length ? completos.slice().sort(function (x, y) { return x.ahorro - y.ahorro; })[0].mes : null
  };
}

function toolComparar_(a) {
  const ra = anaRango_({ periodo: a.periodo_a, desde: a.a_desde, hasta: a.a_hasta });
  const rb = anaRango_({ periodo: a.periodo_b, desde: a.b_desde, hasta: a.b_hasta });
  const por = String(a.agrupar || 'categoria');
  const tipo = a.tipo || 'gasto';
  const f = function (r) { const l = anaFiltrar_({ tipo: tipo, categoria: a.categoria, texto: a.texto, cuenta: a.cuenta }, r); const g = {}; l.forEach(function (m) { const k = anaClave_(m, por); g[k] = (g[k] || 0) + m.pen; }); return { total: anaSuma_(l), g: g, n: l.length }; };
  const A = f(ra), B = f(rb);
  const claves = {}; Object.keys(A.g).concat(Object.keys(B.g)).forEach(function (k) { claves[k] = 1; });
  const filas = Object.keys(claves).map(function (k) {
    const x = A.g[k] || 0, y = B.g[k] || 0;
    return { grupo: k, a: anaN_(x), b: anaN_(y), diferencia: anaN_(x - y), cambioPct: y > 0 ? anaN_((x / y - 1) * 100) : null };
  }).sort(function (p, q) { return Math.abs(q.diferencia) - Math.abs(p.diferencia); }).slice(0, 12);
  return {
    periodoA: ra.etiqueta + ' (' + anaIso_(ra.desde) + ' a ' + anaIso_(ra.hasta) + ')', periodoB: rb.etiqueta + ' (' + anaIso_(rb.desde) + ' a ' + anaIso_(rb.hasta) + ')',
    tipo: tipo, totalA: anaN_(A.total), totalB: anaN_(B.total), diferencia: anaN_(A.total - B.total), cambioPct: B.total > 0 ? anaN_((A.total / B.total - 1) * 100) : null,
    agrupadoPor: por, mayoresCambios: filas, nota: 'A = periodo más reciente a comparar, B = periodo de referencia. Si un periodo es más corto, la comparación no es pareja.'
  };
}

function toolBuscar_(a) {
  const r = anaRango_(a.desde || a.periodo ? a : { periodo: 'todo' });
  const l = anaFiltrar_(a, r).sort(function (x, y) { return y.f - x.f; });
  const lim = Math.min(Math.max(Number(a.limite) || 15, 1), 40);
  return {
    coincidencias: l.length, total: anaN_(anaSuma_(l)), periodo: r.etiqueta,
    movimientos: l.slice(0, lim).map(function (m) { return { fecha: anaIso_(m.f), tipo: m.ing ? 'ingreso' : 'gasto', monto: m.monto, moneda: m.mon, equivalentePEN: anaN_(m.pen), categoria: m.cat, subcategoria: m.sub, descripcion: m.desc, cuenta: m.cuenta }; })
  };
}

/** Pagos que se repiten (suscripciones, servicios, cuotas): 3+ meses distintos en los últimos 8 */
function recurrentes_() {
  const desde = anaInicioMes_(new Date(), -7);
  const g = {};
  anaBase_().lista.forEach(function (m) {
    if (m.ing || m.f < desde) return;
    const k = anaNorm_(m.desc) || anaNorm_(m.cat + ' ' + m.sub);
    if (!k) return;
    if (!g[k]) g[k] = { nombre: m.desc || (m.cat + ' ' + m.sub), cat: m.cat, meses: {}, montos: [], ult: m.f, pu: m.pen };
    const mes = anaClaveMes_(m.f);
    if (!g[k].meses[mes]) { g[k].meses[mes] = 0; g[k].montos.push({ f: m.f, v: m.pen }); }
    g[k].meses[mes] += m.pen;
    if (m.f >= g[k].ult) { g[k].ult = m.f; g[k].pu = m.pen; }
  });
  const out = [];
  Object.keys(g).forEach(function (k) {
    const x = g[k], meses = Object.keys(x.meses);
    if (meses.length < 3) return;
    const vals = meses.map(function (mm) { return x.meses[mm]; });
    const media = anaMedia_(vals), cv = media > 0 ? anaDesv_(vals) / media : 0;
    if (cv > 0.15) return;   // importes muy distintos cada mes: no es un pago fijo
    const ord = x.montos.slice().sort(function (p, q) { return p.f - q.f; });
    const primero = ord[0].v, ultimo = ord[ord.length - 1].v;
    out.push({ pago: x.nombre, categoria: x.cat, mesesConPago: meses.length, promedioMensual: anaN_(media), ultimoPago: anaN_(x.pu), ultimaFecha: anaIso_(x.ult), subioPct: primero > 0 && ultimo > primero * 1.05 ? anaN_((ultimo / primero - 1) * 100) : 0, costoAnualEstimado: anaN_(media * 12) });
  });
  out.sort(function (p, q) { return q.promedioMensual - p.promedioMensual; });
  return out;
}

function toolRecurrentes_() {
  const l = recurrentes_();
  return { pagos: l.slice(0, 20), totalMensualEstimado: anaN_(l.reduce(function (s, x) { return s + x.promedioMensual; }, 0)), totalAnualEstimado: anaN_(l.reduce(function (s, x) { return s + x.costoAnualEstimado; }, 0)), nota: 'Detectados por repetición en 3+ meses con importe parecido; revisa que sean realmente pagos fijos.' };
}

/** Señales para revisar (no son acusaciones): montos atípicos, posibles duplicados, categorías disparadas */
function anomalias_(rango) {
  const base = anaBase_().lista.filter(function (m) { return !m.ing; });
  const out = [];
  const porCat = {};
  base.forEach(function (m) { (porCat[m.cat] = porCat[m.cat] || []).push(m); });
  Object.keys(porCat).forEach(function (c) {
    const l = porCat[c];
    if (l.length < 6) return;
    const v = l.map(function (m) { return m.pen; });
    const med = anaMediana_(v), media = anaMedia_(v), sd = anaDesv_(v);
    l.forEach(function (m) {
      if (m.f < rango.desde || m.f > rango.hasta) return;
      if (m.pen >= 50 && m.pen > Math.max(med * 3, media + 2.5 * sd)) out.push({ tipo: 'monto atípico', fecha: anaIso_(m.f), descripcion: m.desc || m.cat, categoria: c, monto: anaN_(m.pen), tipicoEnCategoria: anaN_(med), id: m.id });
    });
  });
  const vistos = {};
  base.forEach(function (m) {
    if (m.f < rango.desde || m.f > rango.hasta) return;
    const k = anaNorm_(m.desc) + '|' + anaN_(m.pen);
    (vistos[k] = vistos[k] || []).push(m);
  });
  Object.keys(vistos).forEach(function (k) {
    const l = vistos[k].sort(function (p, q) { return p.f - q.f; });
    for (let i = 1; i < l.length; i++) {
      if (l[i].pen >= 10 && (l[i].f - l[i - 1].f) <= 2 * 86400000) { out.push({ tipo: 'posible duplicado', fecha: anaIso_(l[i].f), descripcion: l[i].desc || l[i].cat, monto: anaN_(l[i].pen), fechaAnterior: anaIso_(l[i - 1].f), id: l[i].id }); break; }
    }
  });
  // categorías que se dispararon este mes frente a su promedio de los 3 meses anteriores
  const hoy = new Date(), mesAct = anaClaveMes_(hoy), prev = [1, 2, 3].map(function (k) { return anaClaveMes_(anaInicioMes_(hoy, -k)); });
  const tab = {};
  base.forEach(function (m) { const mm = anaClaveMes_(m.f); if (mm !== mesAct && prev.indexOf(mm) === -1) return; tab[m.cat] = tab[m.cat] || { act: 0, ant: 0 }; if (mm === mesAct) tab[m.cat].act += m.pen; else tab[m.cat].ant += m.pen / 3; });
  const dia = hoy.getDate() / ultimoDiaMes_(hoy);
  Object.keys(tab).forEach(function (c) {
    const t = tab[c];
    if (t.ant >= 80 && t.act > t.ant * 1.5 && t.act - t.ant >= 100 && dia >= 0.3) out.push({ tipo: 'categoría disparada', categoria: c, esteMes: anaN_(t.act), promedio3Meses: anaN_(t.ant), subePct: anaN_((t.act / t.ant - 1) * 100) });
  });
  return out;
}

function toolAnomalias_(a) {
  const r = anaRango_(a.desde || a.periodo ? a : { periodo: 'ultimos_90' });
  const l = anomalias_(r);
  return { periodo: r.etiqueta, senales: l.slice(0, 15), nota: 'Son señales para REVISAR, no prueba de error ni fraude.' };
}

function toolEstado_() {
  const tasa = anaBase_().tasa;
  const cuentas = obtenerCuentasActivas_().map(function (c) { return { cuenta: c.banco, tipo: c.tipo, moneda: c.moneda, saldo: c.saldo, equivalentePEN: anaN_(convertirAPEN_(c.saldo, c.moneda, tasa)) }; });
  const out = { tipoCambio: { usdPen: tasa.usdPen, eurPen: tasa.eurPen, respaldo: !!tasa.respaldo }, cuentas: cuentas, patrimonioDetallePEN: calcularPatrimonioDetalle_() };
  const liquidoPorMoneda = {};
  cuentas.forEach(function (c) { if (c.tipo !== 'Inversion') liquidoPorMoneda[c.moneda] = anaN_((liquidoPorMoneda[c.moneda] || 0) + c.saldo); });
  out.liquidoPorMoneda = liquidoPorMoneda;
  try { const c = contextoAsesor_(); out.tarjetas = c.tarjetas; out.tarjetaPendiente = c.tarjetaPendiente; out.presupuestos = c.presupuestos; out.metas = c.metas; out.proximosMovimientos = c.proximosMovimientos; out.posicionesBolsa = c.posicionesBolsa; out.inversiones = c.inversiones; } catch (e) { /* opcional */ }
  try {
    const po = leerHoja_('PrestamosOtorgados'), porCobrar = [];
    for (let i = 1; i < po.length; i++) if (po[i][0] !== '' && String(po[i][9]).trim() !== 'Pagado') porCobrar.push({ persona: String(po[i][1]), moneda: String(po[i][2]), saldo: numero_(po[i][10]), desde: String(po[i][5] ? fechaISO_(fecha_(po[i][5]) || new Date()) : '') });
    out.meDeben = porCobrar.slice(0, 30);
  } catch (e) { /* opcional */ }
  return out;
}

/** Simula cuánto efectivo tendrías mes a mes. Parte de tu promedio real de los últimos 3 meses completos */
function toolSimular_(a) {
  const serie = anaPorMes_(7).filter(function (x) { return !x.parcial; }).slice(-3);
  const ing0 = anaMedia_(serie.map(function (x) { return x.ingresos; })), gas0 = anaMedia_(serie.map(function (x) { return x.gastos; }));
  const tasa = anaBase_().tasa;
  let liquido = 0;
  obtenerCuentasActivas_().forEach(function (c) { if (c.tipo !== 'Inversion') liquido += convertirAPEN_(c.saldo, c.moneda, tasa); });
  const n = Math.min(Math.max(Number(a.meses) || 12, 1), 36);
  const dI = Number(a.ingresos_pct) || 0, mI = Number(a.ingresos_monto) || 0, dG = Number(a.gastos_pct) || 0, mG = Number(a.gastos_monto) || 0;
  const sinIng = Number(a.sin_ingresos_meses) || 0, unico = Number(a.gasto_unico) || 0, mesUnico = Math.min(Math.max(Number(a.mes_gasto_unico) || 1, 1), n);
  const rend = (Number(a.rendimiento_anual_pct) || 0) / 100 / 12;
  const ing1 = ing0 * (1 + dI / 100) + mI, gas1 = gas0 * (1 + dG / 100) + mG;
  const meses = [];
  let sBase = liquido, sEsc = liquido, minEsc = liquido, mesNeg = null;
  for (let k = 1; k <= n; k++) {
    sBase += ing0 - gas0;
    sEsc = sEsc * (1 + rend) + (k <= sinIng ? 0 : ing1) - gas1 - (k === mesUnico ? unico : 0);
    if (sEsc < minEsc) minEsc = sEsc;
    if (sEsc < 0 && mesNeg === null) mesNeg = k;
    if (k <= 12 || k === n) meses.push({ mes: k, escenario: anaN_(sEsc), sinCambios: anaN_(sBase) });
  }
  return {
    base: { ingresoMensualPromedio: anaN_(ing0), gastoMensualPromedio: anaN_(gas0), ahorroMensualPromedio: anaN_(ing0 - gas0), efectivoActualPEN: anaN_(liquido), mesesUsados: serie.map(function (x) { return x.mes; }) },
    escenario: { ingresoMensual: anaN_(ing1), gastoMensual: anaN_(gas1), ahorroMensual: anaN_(ing1 - gas1), sinIngresosMeses: sinIng, gastoUnico: unico, mesGastoUnico: unico ? mesUnico : null, rendimientoAnualPct: Number(a.rendimiento_anual_pct) || 0 },
    resultado: { efectivoAlFinal: anaN_(sEsc), efectivoAlFinalSinCambios: anaN_(sBase), diferencia: anaN_(sEsc - sBase), saldoMinimo: anaN_(minEsc), primerMesEnNegativo: mesNeg, meses: n },
    proyeccion: meses, nota: 'Estimación lineal con tus promedios; no incluye estacionalidad, ni pagos de tarjeta pendientes, ni programados. El efectivo cuenta cuentas no-inversión en soles equivalentes.'
  };
}

function toolMeta_(a) {
  const meta = Number(a.meta), ahorro = Number(a.ahorro_mensual);
  const serie = anaPorMes_(7).filter(function (x) { return !x.parcial; }).slice(-3);
  const prom = anaMedia_(serie.map(function (x) { return x.ahorro; }));
  const aporte = ahorro > 0 ? ahorro : prom;
  const actual = Number(a.actual) || 0;
  if (!(meta > 0)) return { error: 'Falta el monto de la meta.' };
  if (!(aporte > 0)) return { error: 'Con tu ahorro mensual actual (' + anaN_(prom) + ') no se llega a la meta: no hay ahorro positivo.', ahorroMensualPromedio: anaN_(prom) };
  const meses = Math.ceil(Math.max(meta - actual, 0) / aporte);
  return { meta: meta, yaTienes: actual, aporteMensualUsado: anaN_(aporte), ahorroMensualRealPromedio: anaN_(prom), mesesParaLlegar: meses, anios: anaN_(meses / 12), nota: 'Sin rendimiento ni inflación.' };
}

var HERRAMIENTAS_ = {
  resumen: { f: toolResumen_, d: 'Ingresos, gastos, ahorro, tasa de ahorro, promedios de un periodo (acepta filtros categoria, texto, cuenta, tipo). args: periodo|desde,hasta' },
  agrupar: { f: toolAgrupar_, d: 'Totales agrupados. args: agrupar = categoria|subcategoria|cuenta|medio|dia_semana|mes|descripcion|tipo, tipo = gasto|ingreso, periodo|desde,hasta, categoria, texto, cuenta, monto_min, monto_max, top' },
  serie_mensual: { f: toolSerieMensual_, d: 'Ingresos, gastos y ahorro mes a mes, mejor y peor mes. args: meses (2-36)' },
  comparar: { f: toolComparar_, d: 'Compara dos periodos y muestra qué cambió más. args: periodo_a|a_desde,a_hasta, periodo_b|b_desde,b_hasta, tipo = gasto|ingreso, agrupar, categoria, texto' },
  buscar: { f: toolBuscar_, d: 'Busca movimientos concretos y suma lo que encuentre (por comercio, persona, viaje, palabra). args: texto, categoria, subcategoria, cuenta, tipo, periodo|desde,hasta, monto_min, monto_max, limite' },
  recurrentes: { f: toolRecurrentes_, d: 'Pagos fijos y suscripciones detectados, su costo mensual/anual y si subieron de precio. Sin args' },
  anomalias: { f: toolAnomalias_, d: 'Montos atípicos, posibles duplicados, categorías disparadas. args: periodo|desde,hasta' },
  estado_actual: { f: toolEstado_, d: 'FOTO DE HOY: saldos por cuenta y moneda, patrimonio, tarjetas, presupuestos, metas, próximos pagos, inversiones, quién me debe. Sin args' },
  simular: { f: toolSimular_, d: 'Escenarios hipotéticos y proyección de efectivo. args: meses, ingresos_pct, ingresos_monto, gastos_pct, gastos_monto, sin_ingresos_meses, gasto_unico, mes_gasto_unico, rendimiento_anual_pct' },
  meta: { f: toolMeta_, d: 'Cuánto tardas en juntar un monto. args: meta, actual, ahorro_mensual (opcional: usa tu ahorro real)' }
};

function ejecutarHerramienta_(nombre, args) {
  const h = HERRAMIENTAS_[nombre];
  if (!h) return { error: 'Herramienta desconocida: ' + nombre };
  try { return h.f(args || {}); } catch (e) { return { error: e.message }; }
}

// ---------------------------------------------------------------
// ANALISTA (chat con herramientas)
// ---------------------------------------------------------------

function anaCatalogoTexto_() {
  return Object.keys(HERRAMIENTAS_).map(function (k) { return '- ' + k + ': ' + HERRAMIENTAS_[k].d; }).join('\n');
}

function anaCategoriasTexto_() {
  try {
    const top = {};
    anaBase_().lista.forEach(function (m) { if (!m.ing) top[m.cat] = (top[m.cat] || 0) + m.pen; });
    return Object.keys(top).sort(function (a, b) { return top[b] - top[a]; }).slice(0, 25).join(', ');
  } catch (e) { return ''; }
}

function anaPlan_(mensajes) {
  const hoy = new Date();
  const system = 'Eres el planificador de un analista financiero personal. Decides qué herramientas de cálculo hacen falta para responder la última pregunta del usuario con SUS datos reales. ' +
    'Hoy es ' + fechaISO_(hoy) + ' (' + DIAS_ES_[hoy.getDay()] + '). Las fechas van como AAAA-MM-DD. Si no dice periodo, usa periodo "ultimos_90" para análisis y "ultimos_365" para tendencias.\n' +
    'Herramientas:\n' + anaCatalogoTexto_() + '\n' +
    'Periodos válidos: este_mes, mes_pasado, ultimos_7, ultimos_30, ultimos_90, ultimos_180, ultimos_365, este_ano, ano_pasado, todo.\n' +
    'Categorías de gasto del usuario (por monto): ' + anaCategoriasTexto_() + '\n' +
    'Reglas: usa de 1 a 4 herramientas; para decisiones de compra o viaje combina estado_actual + serie_mensual + simular (gasto_unico) + recurrentes. ' +
    'Para preguntas sobre un comercio, persona o viaje usa buscar. Usa el contexto de la conversación para entender preguntas cortas ("¿y el año pasado?"). ' +
    'Si es un saludo o algo que no necesita datos, devuelve la lista vacía. No inventes herramientas. ' +
    'Responde SOLO con JSON: {"herramientas":[{"nombre":"","args":{}}]}';
  const conv = mensajes.slice(-6).map(function (m) { return (m.rol === 'asesor' ? 'ASISTENTE: ' : 'USUARIO: ') + m.texto; }).join('\n');
  const j = extraerJSON_(iaTexto_(system, conv, 700));
  return (Array.isArray(j.herramientas) ? j.herramientas : []).filter(function (h) { return h && HERRAMIENTAS_[h.nombre]; }).slice(0, 4);
}

var SYSTEM_ANALISTA_ = 'Eres el analista financiero personal del usuario (vive en Trujillo, Perú; soles, dólares y euros). ' +
  'Recibes RESULTADOS DE CÁLCULO hechos sobre sus movimientos reales. Habla en español directo y profesional: sin halagos, sin suavizar, sin relleno. ' +
  'REGLAS: (1) Usa únicamente cifras que aparezcan en los resultados; no inventes ni recalcules de memoria; si algo no está, di qué dato falta. ' +
  '(2) Indica el periodo y la moneda de cada cifra (los importes están en soles equivalentes salvo que se diga otra cosa). ' +
  '(3) Distingue siempre entre DATO (lo registrado), ESTIMACIÓN/PROYECCIÓN (cálculo con supuestos) y OPINIÓN o recomendación. Las proyecciones no son garantías. ' +
  '(4) Una anomalía es una señal para revisar, no una acusación; un patrón observado no prueba una causa. ' +
  '(5) Si hay pocos datos, dilo. Si el tipo de cambio es de respaldo, adviértelo. ' +
  '(6) Respuesta corta y accionable: máximo 200 palabras salvo que pida un informe; usa **negritas** para cifras clave y viñetas "- " solo si ayudan. Cierra con el siguiente paso concreto cuando aplique. ' +
  '(7) Para inversiones y bolsa habla de riesgos y horizonte, sin prometer rentabilidad. No ejecutas operaciones: solo analizas y sugieres.';

function api_analista(a) {
  const prov = iaProveedor_();
  if (!prov) throw new Error('Para hablar con el analista agrega GEMINI_API_KEY (gratis en aistudio.google.com/apikey) o ANTHROPIC_API_KEY en las propiedades del script.');
  const mensajes = (Array.isArray(a.mensajes) ? a.mensajes : []).slice(-10).map(function (m) {
    return { rol: m.rol === 'asesor' ? 'asesor' : 'user', texto: String(m.texto || '').substring(0, 1500) };
  }).filter(function (m) { return m.texto.trim(); });
  if (!mensajes.length || mensajes[mensajes.length - 1].rol !== 'user') throw new Error('Escribe tu pregunta.');
  ANA_ = null;

  let plan = [], planError = '';
  try { plan = anaPlan_(mensajes); } catch (e) { planError = e.message; }
  const resultados = plan.map(function (h) { return { herramienta: h.nombre, args: h.args || {}, resultado: ejecutarHerramienta_(h.nombre, h.args) }; });
  if (!resultados.length && planError) resultados.push({ herramienta: 'estado_actual', args: {}, resultado: ejecutarHerramienta_('estado_actual', {}) });   // plan B: foto de hoy

  const system = SYSTEM_ANALISTA_ + '\n\nFecha de hoy: ' + fechaISO_(new Date()) + '.' + (anaBase_().tasa.respaldo ? ' AVISO: el tipo de cambio es de respaldo.' : '') +
    '\n\nRESULTADOS DE CÁLCULO (JSON):\n' + (resultados.length ? JSON.stringify(resultados).substring(0, 24000) : '(ninguno: la pregunta no necesitó datos)');
  const conv = mensajes.map(function (m) { return (m.rol === 'asesor' ? 'ASISTENTE: ' : 'USUARIO: ') + m.texto; }).join('\n') + '\nASISTENTE:';
  const texto = iaTexto_(system, conv, 1100);
  ANA_ = null;
  return {
    respuesta: String(texto).trim().replace(/^ASISTENTE:\s*/i, ''), proveedor: prov,
    herramientas: resultados.map(function (r) { return { nombre: r.herramienta, args: r.args }; }), avisoPlan: planError
  };
}

// ---------------------------------------------------------------
// PERFIL FINANCIERO
// ---------------------------------------------------------------

/** Indicadores con semáforo y fórmula explicada. nivel: ok | atencion | alerta | info */
function indicadoresPerfil_() {
  ANA_ = null;
  const base = anaBase_(), tasa = base.tasa, hoy = new Date();
  const serie = anaPorMes_(7).filter(function (x) { return !x.parcial; }).slice(-6).filter(function (x) { return x.ingresos > 0 || x.gastos > 0; });
  const out = { meses: serie.length, indicadores: [] };
  const add = function (id, titulo, valor, texto, nivel, formula) { out.indicadores.push({ id: id, titulo: titulo, valor: valor, texto: texto, nivel: nivel, formula: formula }); };
  if (serie.length < 2) { out.pocos = true; return out; }

  const I = serie.map(function (x) { return x.ingresos; }), G = serie.map(function (x) { return x.gastos; });
  const iT = anaSuma_([]) + I.reduce(function (s, x) { return s + x; }, 0), gT = G.reduce(function (s, x) { return s + x; }, 0);
  const tAh = iT > 0 ? (iT - gT) / iT * 100 : null;
  out.tasaAhorro = tAh;
  if (tAh !== null) add('ahorro', 'Tasa de ahorro', anaN_(tAh) + '%', 'Conservas ' + anaN_(tAh) + '% de lo que ingresa (promedio de ' + serie.length + ' meses completos).', tAh >= 20 ? 'ok' : tAh >= 8 ? 'atencion' : 'alerta', '(ingresos − gastos) / ingresos');
  const consist = serie.filter(function (x) { return x.ahorro > 0; }).length;
  out.mesesConAhorro = consist;
  add('constancia', 'Constancia de ahorro', consist + ' de ' + serie.length + ' meses', 'Meses en los que gastaste menos de lo que ingresó.', consist / serie.length >= 0.8 ? 'ok' : consist / serie.length >= 0.5 ? 'atencion' : 'alerta', 'meses con ahorro > 0');
  const mI = anaMedia_(I), cvI = mI > 0 ? anaDesv_(I) / mI : null;
  out.cvIngresos = cvI;
  if (cvI !== null) add('estabilidad', 'Estabilidad de ingresos', anaN_(cvI * 100) + '% de variación', cvI <= 0.2 ? 'Tus ingresos son bastante parejos mes a mes.' : cvI <= 0.5 ? 'Tus ingresos varían de forma moderada.' : 'Tus ingresos cambian mucho de un mes a otro.', cvI <= 0.2 ? 'ok' : cvI <= 0.5 ? 'atencion' : 'alerta', 'desviación estándar / promedio de ingresos mensuales');

  const desde6 = anaInicioMes_(hoy, -serie.length);
  const l6 = base.lista.filter(function (m) { return m.f >= desde6 && m.f < anaInicioMes_(hoy); });
  const ing6 = l6.filter(function (m) { return m.ing; }), gas6 = l6.filter(function (m) { return !m.ing; });
  const fuentes = {}; ing6.forEach(function (m) { fuentes[m.cat] = (fuentes[m.cat] || 0) + m.pen; });
  const fk = Object.keys(fuentes).sort(function (a, b) { return fuentes[b] - fuentes[a]; });
  const iTot = anaSuma_(ing6);
  if (fk.length && iTot > 0) {
    const p = fuentes[fk[0]] / iTot * 100; out.fuentePrincipal = { nombre: fk[0], pct: p };
    add('dependencia', 'Dependencia de tu fuente principal', anaN_(p) + '%', '"' + fk[0] + '" aporta ' + anaN_(p) + '% de tus ingresos' + (fk.length > 1 ? ' (tienes ' + fk.length + ' fuentes).' : ' y es tu única fuente.'), p <= 60 ? 'ok' : p <= 85 ? 'atencion' : 'alerta', 'ingreso de la mayor fuente / ingresos totales');
  }
  const pasivo = iTot > 0 ? anaSuma_(ing6.filter(function (m) { return m.tipo === 'Ingreso Pasivo'; })) / iTot * 100 : 0;
  add('pasivo', 'Ingresos pasivos', anaN_(pasivo) + '%', pasivo >= 20 ? 'Una parte relevante de tus ingresos no depende de tu tiempo.' : 'Casi todo tu ingreso depende de que trabajes activamente.', pasivo >= 20 ? 'ok' : 'info', 'ingresos marcados como pasivos / ingresos totales');

  const gTot = anaSuma_(gas6);
  const fijos = recurrentes_().reduce(function (s, x) { return s + x.promedioMensual; }, 0);
  const gMes = gTot / Math.max(serie.length, 1);
  if (gMes > 0) add('fijos', 'Pagos fijos detectados', anaN_(fijos / gMes * 100) + '% de tu gasto', 'Pagos que se repiten cada mes suman unos ' + anaN_(fijos) + ' al mes (estimación).', fijos / gMes <= 0.5 ? 'ok' : 'atencion', 'recurrentes detectados / gasto mensual');
  const peq = gas6.filter(function (m) { return m.pen < 20; });
  if (gTot > 0) add('pequenos', 'Gastos pequeños (< S/ 20)', anaN_(anaSuma_(peq) / gTot * 100) + '%', peq.length + ' compras chicas suman ' + anaN_(anaSuma_(peq)) + ' en ' + serie.length + ' meses.', anaSuma_(peq) / gTot <= 0.1 ? 'ok' : 'atencion', 'gastos < S/ 20 / gasto total');
  const tj = gas6.filter(function (m) { return m.medio === 'Tarjeta de Credito'; });
  if (gTot > 0) add('tarjeta', 'Uso de tarjeta de crédito', anaN_(anaSuma_(tj) / gTot * 100) + '% de tus gastos', 'Parte de tu gasto pasa por tarjeta.', 'info', 'gastos con tarjeta / gasto total');
  // ¿gasta más justo después de cobrar? (5 días tras un ingreso vs resto)
  const cobros = ing6.filter(function (m) { return m.pen >= 0.3 * (mI || 1); }).map(function (m) { return m.f.getTime(); });
  if (cobros.length >= 2 && gas6.length >= 20) {
    // se excluyen los pagos fijos (alquiler, suscripciones): caen justo después de cobrar por calendario, no por impulso
    const fijosNom = {}; recurrentes_().forEach(function (x) { fijosNom[anaNorm_(x.pago)] = 1; });
    const gasV = gas6.filter(function (m) { return !fijosNom[anaNorm_(m.desc) || anaNorm_(m.cat + ' ' + m.sub)]; });
    const gVar = anaSuma_(gasV);
    let postG = 0;
    const ventana = {};
    cobros.forEach(function (t) { for (let k = 0; k < 5; k++) ventana[anaIso_(new Date(t + k * 86400000))] = 1; });
    const diasTot = Math.max(1, Math.round((anaInicioMes_(hoy) - desde6) / 86400000));
    const dPost = Math.min(Object.keys(ventana).length, diasTot - 1);
    gasV.forEach(function (m) { if (ventana[anaIso_(m.f)]) postG += m.pen; });
    const ratio = (postG / Math.max(dPost, 1)) / (((gVar - postG) / Math.max(diasTot - dPost, 1)) || 1);
    if (isFinite(ratio)) { out.ratioPostCobro = ratio; add('postcobro', 'Gasto tras cobrar', 'x' + anaN_(ratio) + ' vs. un día normal', ratio >= 1.5 ? 'Gastas bastante más en los 5 días posteriores a un ingreso grande.' : 'No se nota un salto de gasto después de cobrar.', ratio >= 1.5 ? 'atencion' : 'ok', 'gasto diario en los 5 días tras un ingreso / gasto diario del resto'); }
  }
  const dias = {}; gas6.forEach(function (m) { const k = DIAS_ES_[m.f.getDay()]; dias[k] = (dias[k] || 0) + m.pen; });
  const dk = Object.keys(dias).sort(function (a, b) { return dias[b] - dias[a]; });
  if (dk.length) { out.diaMayorGasto = dk[0]; add('dia', 'Día de mayor gasto', dk[0], 'El ' + dk[0] + ' concentra ' + anaN_(dias[dk[0]] / gTot * 100) + '% de tu gasto.', 'info', 'gasto por día de la semana'); }

  // liquidez y patrimonio
  let liquido = 0, inv = 0;
  obtenerCuentasActivas_().forEach(function (c) { const v = convertirAPEN_(c.saldo, c.moneda, tasa); if (c.tipo === 'Inversion') inv += v; else liquido += v; });
  const pat = calcularPatrimonioDetalle_();
  const meses = gMes > 0 ? liquido / gMes : null;
  out.mesesLiquidez = meses;
  if (meses !== null) add('liquidez', 'Colchón de liquidez', anaN_(meses) + ' meses', 'Tu efectivo cubre ' + anaN_(meses) + ' meses de gastos al ritmo actual.', meses >= 6 ? 'ok' : meses >= 3 ? 'atencion' : 'alerta', 'efectivo disponible / gasto mensual promedio');
  const patTotal = (pat && (pat.total !== undefined ? pat.total : pat.neto)) || null;
  const invTotal = Number((pat && pat.inversiones) || 0) + inv;
  const activos = liquido + invTotal;
  out.pctInvertido = activos > 0 ? invTotal / activos * 100 : 0;
  add('invertido', 'Parte invertida', anaN_(out.pctInvertido) + '%', invTotal > 0 ? 'Tienes ' + anaN_(out.pctInvertido) + '% de tu dinero en inversiones.' : 'Todavía no tienes dinero en inversiones registradas.', out.pctInvertido >= 20 ? 'ok' : 'info', 'inversiones / (efectivo + inversiones)');
  const deuda = Number((pat && pat.deudas) || 0) + Number((pat && pat.tarjeta) || 0);
  if (deuda > 0 && iTot > 0) add('deuda', 'Deudas y tarjeta pendientes', anaN_(deuda), 'Equivale a ' + anaN_(deuda / (iTot / serie.length)) + ' meses de ingreso.', deuda / (iTot / serie.length) <= 1 ? 'atencion' : 'alerta', 'deudas + consumos de tarjeta no pagados');
  out.patrimonio = pat; out.mensual = serie;
  return out;
}

/** Arquetipo por reglas explícitas: se muestran los datos que lo disparan */
function arquetipoPerfil_(o) {
  if (o.pocos) return { nombre: 'Perfil en construcción', explicacion: 'Aún hay menos de 2 meses completos de movimientos. Sigue registrando y el perfil se afina solo.', reglas: [] };
  const r = [];
  const tA = o.tasaAhorro === null || o.tasaAhorro === undefined ? 0 : o.tasaAhorro;
  const inv = o.pctInvertido || 0, cv = o.cvIngresos === null || o.cvIngresos === undefined ? 0 : o.cvIngresos, cons = o.mesesConAhorro / Math.max(o.meses, 1);
  let nombre, expl;
  if (tA < 0) { nombre = 'En déficit'; expl = 'Estás gastando más de lo que ingresa. Es la prioridad número uno.'; r.push('tasa de ahorro ' + anaN_(tA) + '% (negativa)'); }
  else if (tA >= 20 && inv >= 20) { nombre = 'Constructor de patrimonio'; expl = 'Ahorras con firmeza y ya haces trabajar parte de ese dinero en inversiones.'; r.push('tasa de ahorro ' + anaN_(tA) + '% ≥ 20%', 'invertido ' + anaN_(inv) + '% ≥ 20%'); }
  else if (tA >= 20) { nombre = 'Ahorrador conservador'; expl = 'Ahorras bien, pero casi todo se queda parado como efectivo: el siguiente paso es decidir qué parte invertir.'; r.push('tasa de ahorro ' + anaN_(tA) + '% ≥ 20%', 'invertido ' + anaN_(inv) + '% < 20%'); }
  else if (cv > 0.5) { nombre = 'Ingresos irregulares'; expl = 'Tu mayor reto no es cuánto gastas sino lo cambiantes que son tus ingresos; necesitas un colchón mayor y un presupuesto basado en tu mes mínimo.'; r.push('variación de ingresos ' + anaN_(cv * 100) + '% > 50%'); }
  else if (tA >= 8) { nombre = 'Equilibrado'; expl = 'Cubres tus gastos y ahorras algo, aunque con margen para mejorar.'; r.push('tasa de ahorro ' + anaN_(tA) + '% entre 8% y 20%'); }
  else { nombre = 'Gasto al límite'; expl = 'Casi todo lo que ingresa se va en gastos; cualquier imprevisto te pega fuerte.'; r.push('tasa de ahorro ' + anaN_(tA) + '% < 8%'); }
  if (cons < 0.5 && tA >= 0) r.push('ahorras solo en ' + o.mesesConAhorro + ' de ' + o.meses + ' meses: el ahorro depende de lo que sobra');
  return { nombre: nombre, explicacion: expl, reglas: r, aviso: 'Etiqueta orientativa basada en reglas simples, no un diagnóstico ni un puntaje.' };
}

function perfilNarrativa_(o, arq, hallazgos) {
  const system = SYSTEM_ANALISTA_ + ' Con los indicadores y hallazgos entregados redacta un perfil breve del comportamiento financiero del usuario. ' +
    'Responde SOLO con JSON: {"resumen":"2-3 frases","fortalezas":["..."],"debilidades":["..."],"habitos":["patrones observados, sin asumir causas"],"siguiente_paso":"una acción concreta"}. ' +
    'Máximo 3 elementos por lista, cada uno de una línea con cifras de los datos.';
  const datos = { arquetipo: arq, indicadores: o.indicadores.map(function (x) { return { titulo: x.titulo, valor: x.valor, nivel: x.nivel, texto: x.texto }; }), mensual: o.mensual, hallazgos: hallazgos };
  return extraerJSON_(iaTexto_(system, JSON.stringify(datos), 900));
}

function api_perfil(a) {
  ANA_ = null;
  const o = indicadoresPerfil_();
  const arq = arquetipoPerfil_(o);
  let hallazgos = [];
  try {
    const rec = recurrentes_().filter(function (x) { return x.subioPct > 0; }).slice(0, 3).map(function (x) { return { tipo: 'subió de precio', texto: x.pago + ' subió ' + x.subioPct + '% (ahora ' + x.ultimoPago + ')' }; });
    const an = anomalias_(anaRango_({ periodo: 'ultimos_30' })).slice(0, 4).map(function (x) { return { tipo: x.tipo, texto: x.tipo + ': ' + (x.descripcion || x.categoria) + (x.monto ? ' (' + x.monto + ')' : (x.subePct ? ' (+' + x.subePct + '%)' : '')) }; });
    hallazgos = rec.concat(an);
  } catch (e) { /* opcional */ }
  const sal = { arquetipo: arq, indicadores: o.indicadores, mesesAnalizados: o.meses, hallazgos: hallazgos, narrativa: null, iaDisponible: !!iaProveedor_(), generado: '', error: '' };
  if (!sal.iaDisponible || o.pocos) { ANA_ = null; return sal; }
  const cache = CacheService.getScriptCache();
  const g = cache.get('ia_perfil');
  if (g && !(a && a.forzar)) { try { const x = JSON.parse(g); sal.narrativa = x.n; sal.generado = x.g; ANA_ = null; return sal; } catch (e) { /* recalcular */ } }
  try {
    sal.narrativa = perfilNarrativa_(o, arq, hallazgos);
    sal.generado = fechaHoraISO_(new Date());
    cache.put('ia_perfil', JSON.stringify({ n: sal.narrativa, g: sal.generado }), 21600);
  } catch (e) { sal.error = e.message; }
  ANA_ = null;
  return sal;
}

/* ==== 28_ExportarLocal.gs ==== */
/**
 * ============================================================
 * PASAR A LA APP LOCAL
 * ============================================================
 * Genera un archivo con todos tus datos en el formato de copia de seguridad de la app local (la que corre en el teléfono).
 * Se descarga desde Más > Pasar a la app local, y se carga en la app local con "Ya tengo una copia de seguridad".
 */

const HOJAS_LOCAL_RESPALDO_ = ['Cuentas', 'Movimientos', 'CambiosDivisa', 'TarjetaCredito', 'Tarjetas', 'PrestamosOtorgados', 'Deudas', 'CronogramaCuotas', 'Metas', 'MetasAhorro',
  'Presupuestos', 'Recurrentes', 'ReglasCategorizacion', 'TipoCambioReferencial', 'AjustesSaldo', 'Programados', 'Ajustes', 'Inversiones', 'Cotizaciones', 'Categorias', 'Proyectos', 'Personas'];

function api_exportarParaLocal() {
  const ss = getSS_();
  const hojas = [];
  let id = 1;
  HOJAS_LOCAL_RESPALDO_.forEach(function (nombre) {
    const h = ss.getSheetByName(nombre);
    if (!h || h.getLastRow() < 1) return;
    const filas = h.getRange(1, 1, h.getLastRow(), h.getLastColumn()).getValues().map(function (f) {
      return f.map(function (c) {
        if (c instanceof Date) return { $d: c.getTime() };
        return c === null || c === undefined ? '' : c;
      });
    });
    hojas.push({ name: nombre, id: id++, d: filas });
  });
  const fecha = Utilities.formatDate(new Date(), zona_(), 'yyyy-MM-dd');
  return { nombre: 'Finanzas_copia_' + fecha + '.json', json: JSON.stringify({ v: 1, nextId: id, props: {}, sheets: hojas }), hojas: hojas.length };
}

/* ==== 29_ApiLocal.gs ==== */
/**
 * ============================================================
 * PUERTA DE LA APP LOCAL (PWA) HACIA ESTA HOJA
 * ============================================================
 * Permite que la app del celular (la versión local) lea y escriba en TU hoja, para que todo quede sincronizado
 * con lo que registras por Telegram. La hoja sigue siendo la fuente de verdad.
 *
 * Seguridad: nada se entrega sin el token que se obtiene con tu DASHBOARD_PIN (igual que el dashboard de siempre).
 * Solo se activa cuando la petición trae ?api=1 (ver el cambio en doPost de 09_Telegram.gs).
 * Después de pegar este archivo: Implementar > Administrar implementaciones > editar > Nueva versión.
 */

/** Recibe { nombre, args, token } en JSON y responde lo mismo que apiLlamar. */
function apiHttp_(e) {
  const salida = function (o) {
    return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  };
  try {
    const cuerpo = JSON.parse((e.postData && e.postData.contents) || '{}');
    return salida(apiLlamar(String(cuerpo.nombre || ''), cuerpo.args || {}, String(cuerpo.token || '')));
  } catch (err) {
    return salida({ __error: err.message });
  }
}

/**
 * "Huella" barata de la hoja: cambia cuando alguien (tú, Telegram, un disparador) modifica algo.
 * La app la consulta cada rato y solo descarga los datos completos cuando cambió.
 */
function api_versionHoja() {
  const ss = getSS_();
  let t = 0;
  try { t = DriveApp.getFileById(ss.getId()).getLastUpdated().getTime(); } catch (e) { /* sin permiso de Drive: queda solo la cuenta de filas */ }
  let filas = 0;
  ['Movimientos', 'Programados', 'Cuentas'].forEach(function (n) {
    const h = ss.getSheetByName(n);
    if (h) filas += h.getLastRow();
  });
  return t + '-' + filas;
}

/* ==== 32_InstalarCliente.gs ==== */
/**
 * ============================================================
 * INSTALACIÓN EN LA HOJA DE UN CLIENTE
 * ============================================================
 * Cada persona usa SU PROPIA hoja de Google: la app (en el celular) se conecta a ella con un enlace.
 * Flujo (lo guía la app en Más > Conectar con mi hoja):
 *   1) Crear una hoja nueva en Google Sheets.
 *   2) Extensiones > Apps Script, pegar el código que copia la app (ya trae la clave de conexión).
 *   3) Ejecutar "instalarTodo" una vez (pide permisos).
 *   4) Implementar > Nueva implementación > Aplicación web > "Cualquier persona" > copiar la URL /exec y pegarla en la app.
 * Nada de esto pasa por servidores de nadie más: los datos viven en la hoja del cliente.
 */

/** La app reemplaza este texto por la clave de conexión del cliente al copiar el código. */
const CLAVE_APP_ = '__CLAVE_APP__';

function instalarTodo() {
  if (!CLAVE_APP_ || CLAVE_APP_.indexOf('__') === 0) {
    throw new Error('Copia el código desde la app (Más > Conectar con mi hoja > Copiar código): ahí ya viene tu clave de conexión.');
  }
  const P = PropertiesService.getScriptProperties();
  crearEstructuraCompleta();
  P.setProperty('DASHBOARD_PIN', CLAVE_APP_);
  if (!P.getProperty('SESSION_SECRET')) P.setProperty('SESSION_SECRET', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''));
  try { actualizarTipoCambioReferencial(); } catch (e) { /* sin red: se actualiza luego */ }
  try { instalarTriggers(); } catch (e) { /* opcional */ }
  registrarLog_('Instalación', 'Hoja preparada desde la app');
  avisar_('✅ Tu hoja está lista. Ahora: Implementar > Nueva implementación > Aplicación web (acceso: Cualquier persona) y copia la URL que termina en /exec.');
}

/** Cuántos datos tiene la hoja (la app lo usa para decidir si sube los datos del celular o baja los de la hoja) */
function api_estadoHoja() {
  const mov = getSS_().getSheetByName('Movimientos');
  const cu = getSS_().getSheetByName('Cuentas');
  return { movimientos: mov ? Math.max(mov.getLastRow() - 1, 0) : 0, cuentas: cu ? Math.max(cu.getLastRow() - 1, 0) : 0 };
}

/**
 * Sube a la hoja una copia de los datos del celular (formato de la copia de seguridad local).
 * Solo si la hoja no tiene movimientos, salvo que se pida forzar. Reemplaza el contenido de las pestañas de datos.
 */
function api_importarDesdeLocal(a) {
  return conLock_(function () {
    const est = api_estadoHoja();
    if (est.movimientos > 0 && !a.forzar) throw new Error('Tu hoja ya tiene movimientos: no se reemplazó nada.');
    let o;
    try { o = JSON.parse(a.json); } catch (e) { throw new Error('Los datos del celular no son válidos.'); }
    if (!o || !Array.isArray(o.sheets)) throw new Error('Los datos del celular no son válidos.');
    const ss = getSS_();
    const revivir = function (c) { return c && typeof c === 'object' && c.$d !== undefined ? new Date(c.$d) : c; };
    let n = 0;
    o.sheets.forEach(function (s) {
      if (HOJAS_LOCAL_RESPALDO_.indexOf(s.name) === -1 || !Array.isArray(s.d) || !s.d.length) return;
      let h = ss.getSheetByName(s.name);
      if (!h) h = ss.insertSheet(s.name);
      const filas = s.d.map(function (f) { return f.map(revivir); });
      const ancho = filas.reduce(function (m, f) { return Math.max(m, f.length); }, 1);
      filas.forEach(function (f) { while (f.length < ancho) f.push(''); });
      if (h.getMaxColumns() < ancho) h.insertColumnsAfter(h.getMaxColumns(), ancho - h.getMaxColumns());
      if (h.getMaxRows() < filas.length) h.insertRowsAfter(h.getMaxRows(), filas.length - h.getMaxRows());
      h.clearContents();
      h.getRange(1, 1, filas.length, ancho).setValues(filas);
      n++;
    });
    SpreadsheetApp.flush();
    registrarLog_('ImportarLocal', n + ' pestañas subidas desde la app');
    return { hojas: n };
  });
}

/** Guarda la clave de IA del cliente en SU script (la usa el bot de Telegram). args = { clave, proveedor: 'gemini'|'anthropic' } */
function api_guardarClaveIA(a) {
  const k = String(a.clave || '').trim();
  const nombre = a.proveedor === 'anthropic' ? 'ANTHROPIC_API_KEY' : 'GEMINI_API_KEY';
  const P = PropertiesService.getScriptProperties();
  if (k) P.setProperty(nombre, k); else P.deleteProperty(nombre);
  return true;
}

/** Conecta el bot de Telegram del cliente sin tocar las propiedades a mano. args = { token, chatId, url } (url = la /exec de su implementación) */
function api_guardarTelegram(a) {
  const token = String(a.token || '').trim(), chat = String(a.chatId || '').trim(), base = String(a.url || '').trim();
  if (!/^\d+:[\w-]{20,}$/.test(token)) throw new Error('El token del bot no parece correcto (se ve como 123456:ABC…).');
  if (!/^-?\d+$/.test(chat)) throw new Error('El ID de chat son solo números.');
  if (base.indexOf('https://script.google.com/') !== 0) throw new Error('Falta la URL de tu implementación.');
  const v = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getMe', { muteHttpExceptions: true });
  if (v.getResponseCode() !== 200) throw new Error('Telegram no reconoce ese token. Cópialo de nuevo desde @BotFather.');
  const P = PropertiesService.getScriptProperties();
  P.setProperty('TELEGRAM_BOT_TOKEN', token);
  P.setProperty('TELEGRAM_CHAT_ID', chat);
  if (!P.getProperty('TELEGRAM_WEBHOOK_SECRET')) P.setProperty('TELEGRAM_WEBHOOK_SECRET', Utilities.getUuid().replace(/-/g, ''));
  const url = base.split('?')[0] + '?tg=' + encodeURIComponent(P.getProperty('TELEGRAM_WEBHOOK_SECRET'));
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook?drop_pending_updates=true&max_connections=1&allowed_updates=' +
    encodeURIComponent('["message","callback_query"]') + '&url=' + encodeURIComponent(url), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Telegram respondió: ' + r.getContentText().substring(0, 160));
  registrarLog_('Telegram', 'Webhook conectado desde la app');
  return true;
}

/* ==== 33_Correo.gs ==== */
/**
 * ============================================================
 * LECTURA AUTOMÁTICA DE TU GMAIL (alertas de los bancos)
 * ============================================================
 * Cada 15 minutos revisa los correos de alertas de tus bancos (consumos, Yape, transferencias),
 * la IA saca monto, comercio, cuenta y categoría, y:
 *   - si no parece repetido → lo registra solo y te avisa por Telegram (con ↩️ Deshacer);
 *   - si parece que ya lo registraste (por Telegram o a mano) → NO lo registra y te pregunta.
 * Nunca toca correos de BBVA ni de Banca di Asti (no son tus cuentas).
 * Solo lee: no borra ni responde correos. Marca los procesados con la etiqueta "Finanzas/Procesado".
 *
 * Activar una vez: ejecuta activarLecturaCorreo() desde el editor (Google pedirá permiso para leer Gmail).
 * Desactivar: desactivarLecturaCorreo().
 */

var CORREO_REMITENTES_ = ['notificacionesbcp.com.pe', 'yape.pe', 'interbank.pe', 'netinterbank.com.pe', 'intercorp.com.pe',
  'scotiabank.com.pe', 'bancofalabella.pe', 'bancoripley.com.pe', 'plin.pe', 'revolut.com', 'tunki.com.pe', 'mibanco.com.pe', 'pichincha.pe'];
/** Bancos que NO son del usuario: se ignoran siempre (BBVA = dinero de su hermano; Banca di Asti = cuentas familiares en Italia) */
var CORREO_EXCLUIR_ = /bbva|banca\s*di\s*asti|bancadiasti|banco\s*di\s*asti|cr\s*asti/i;
/** Correos de los bancos que no son movimientos */
var CORREO_RUIDO_ = /c[oó]digo de verificaci|clave|contraseña|encuesta|opini[oó]n|bienvenid|t[eé]rminos|promoci|cr[eé]dito listo|activaci[oó]n|activ[oó]|solicitud|estado de cuenta|extracto|ventajas|recibe eur|te quedan|reorganizaci|configuraci[oó]n|l[ií]mite|actualiza tus datos|seguridad|newsletter|beneficio/i;
var CORREO_ETIQUETA_ = 'Finanzas/Procesado';

function activarLecturaCorreo() {
  desactivarLecturaCorreo();
  GmailApp.getInboxUnreadCount(); // fuerza el permiso de Gmail
  if (!GmailApp.getUserLabelByName(CORREO_ETIQUETA_)) GmailApp.createLabel(CORREO_ETIQUETA_);
  PropertiesService.getScriptProperties().setProperty('CORREO_DESDE', String(Date.now()));
  ScriptApp.newTrigger('revisarCorreosBancarios').timeBased().everyMinutes(15).create();
  registrarLog_('Correo', 'Lectura automática de Gmail activada');
  return 'Listo: desde ahora reviso tus alertas bancarias cada 15 minutos.';
}

function desactivarLecturaCorreo() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'revisarCorreosBancarios') ScriptApp.deleteTrigger(t); });
}

function api_estadoCorreo() {
  const activo = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'revisarCorreosBancarios'; });
  return { activo: activo, desde: Number(prop_('CORREO_DESDE') || 0) };
}

function correoYaProcesado_(id) {
  const l = String(prop_('CORREO_HECHOS') || '').split(',');
  return l.indexOf(id) !== -1;
}
function marcarCorreoProcesado_(id) {
  const l = String(prop_('CORREO_HECHOS') || '').split(',').filter(Boolean);
  l.push(id);
  PropertiesService.getScriptProperties().setProperty('CORREO_HECHOS', l.slice(-400).join(','));
}

/** Lo que corre cada 15 minutos */
function revisarCorreosBancarios() {
  if (!iaProveedor_()) return;
  const desde = Number(prop_('CORREO_DESDE') || 0);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const etiqueta = GmailApp.getUserLabelByName(CORREO_ETIQUETA_) || GmailApp.createLabel(CORREO_ETIQUETA_);
    const q = 'from:(' + CORREO_REMITENTES_.join(' OR ') + ') newer_than:3d -label:' + CORREO_ETIQUETA_.replace('/', '-');
    const hilos = GmailApp.search(q, 0, 15);
    let usadas = 0;
    hilos.forEach(function (h) {
      h.getMessages().forEach(function (m) {
        const id = m.getId();
        if (correoYaProcesado_(id) || m.getDate().getTime() < desde || usadas >= 8) return;
        try { if (procesarCorreoBancario_(m)) usadas++; }
        catch (e) { registrarLog_('ERROR Correo', m.getSubject() + ': ' + e.message); }
        marcarCorreoProcesado_(id);
      });
      if (h.getMessages().every(function (m) { return correoYaProcesado_(m.getId()); })) h.addLabel(etiqueta);
    });
  } finally { lock.releaseLock(); }
}

/** Limpia el cuerpo: quita líneas legales y espacios para que la IA lea solo lo importante */
function textoCorreo_(m) {
  const t = String(m.getPlainBody() || '').replace(/\r/g, '').split('\n').map(function (x) { return x.trim(); }).filter(function (x) {
    return x && !/^(descargar| )$/i.test(x) && !/nunca te solicitar|por tu seguridad|comun[ií]cate|si no reconoces|no respondas|políticas|privacidad|desuscrib|copyright/i.test(x);
  });
  return t.join('\n').substring(0, 1800);
}

/** Devuelve true si gastó IA en este correo */
function procesarCorreoBancario_(m) {
  const de = m.getFrom(), asunto = m.getSubject(), cuerpo = textoCorreo_(m);
  if (CORREO_EXCLUIR_.test(de + ' ' + asunto) || CORREO_EXCLUIR_.test(cuerpo.substring(0, 600))) return false;
  if (CORREO_RUIDO_.test(asunto) && !/consumo|pago|yape|transferencia|constancia|operaci/i.test(asunto)) return false;

  const j = interpretarCorreoIA_(de, asunto, cuerpo, m.getDate());
  if (!j || j.ignorar || !Array.isArray(j.acciones) || !j.acciones.length) {
    if (j && j.propia) enviarMensajeTelegram('📧 Vi un movimiento entre tus propias cuentas (' + escaparMd_(asunto) + '). No lo registré: si quieres anotarlo, usa Traspaso en la app.');
    return true;
  }
  const banco = escaparMd_(String(de).replace(/<.*>/, '').trim() || 'tu banco');
  j.acciones.slice(0, 3).forEach(function (a) {
    if (a.accion !== 'registrar' || !(Number(a.monto) > 0) || (a.tipo !== 'gasto' && a.tipo !== 'ingreso')) return;
    a.notas = 'Desde correo · ' + asunto.substring(0, 60);
    const dup = posibleDuplicado_(Number(a.monto), a.moneda || 'PEN', a.fecha ? fecha_(a.fecha) : m.getDate(), a.tipo === 'gasto');
    if (dup) {
      const k = guardarAccionPendiente_(a);
      enviarMensajeTelegram('📧 *' + banco + '*: ' + (a.tipo === 'gasto' ? 'gasto' : 'ingreso') + ' de ' + formatoMonto_(Number(a.monto), a.moneda || 'PEN') + ' — "' + escaparMd_(a.descripcion || '') + '"\n' +
        'Parece que ya lo tienes: "' + escaparMd_(dup.descripcion) + '" del ' + fechaISO_(dup.fecha) + '. ¿Es el mismo?',
        [[{ text: '✅ Es el mismo', data: 'pend:no:' + k }, { text: '➕ Registrar aparte', data: 'pend:si:' + k }]]);
      return;
    }
    const r = ejecutarAccionesIA_([a]);
    const lote = r.deshacer.length ? guardarDeshacer_(r.deshacer) : '';
    enviarMensajeTelegram('📧 Desde tu correo (' + banco + '):\n' + r.lineas.join('\n'), lote ? [[{ text: '↩️ Deshacer', data: 'undo:' + lote }]] : undefined);
    preguntarCategorias_(r.sinCategoria);
  });
  return true;
}

function interpretarCorreoIA_(de, asunto, cuerpo, fechaCorreo) {
  const extra = '\n\nESTE MENSAJE NO LO ESCRIBIÓ EL USUARIO: es una notificación automática de su banco que llegó a su correo el ' +
    Utilities.formatDate(fechaCorreo, zona_(), "yyyy-MM-dd'T'HH:mm") + '. Reglas extra:\n' +
    '- Nunca preguntes nada: "pregunta" siempre null. Si no es un movimiento de dinero ya realizado (publicidad, aviso, código, solicitud), devuelve acciones vacías.\n' +
    '- Usa la fecha y hora de la operación que dice el correo (o la del correo).\n' +
    '- "Realizaste un consumo con tu Tarjeta de Crédito X" → gasto, medio tarjeta, tarjeta = X. "Tarjeta de Débito X" → gasto, medio debito, cuenta = X.\n' +
    '- "Recepción de Yapeo/Plin/transferencia recibida" → ingreso a la cuenta del banco que envía el correo (Yape es del BCP). "Pago exitoso" de Yape o "yapeo enviado" → gasto, cuenta BCP.\n' +
    '- descripcion: el comercio o la persona, corto y legible (ej. "Mercado Pago", "Yape de Juan Pérez"), sin números de operación.\n' +
    '- Si el destinatario o el que envía es el propio usuario (Piero Gallo) o es "pago de tarjeta de crédito propia", es un movimiento entre sus cuentas: devuelve acciones vacías y agrega "propia": true.\n' +
    '- Elige categoría y subcategoría de su lista según el comercio.';
  const texto = 'Remitente: ' + de + '\nAsunto: ' + asunto + '\n' + cuerpo + extra;
  return interpretarMensajeIA_(texto, null);
}


