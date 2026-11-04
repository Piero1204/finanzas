/* Ajustes para la versión local (se ejecuta DESPUÉS del código del sistema) */
CONFIG.CUENTAS_POR_MONEDA = { PEN: ['Efectivo'], USD: [], EUR: [] };
CONFIG.CENTROS_COSTO = ['Personal'];
CONFIG.API_TIPO_CAMBIO_URL = 'https://open.er-api.com/v6/latest/PEN';

function poblarReglasCategorizacionIniciales(ss) {
  ss = ss || getSS_();
  const hoja = ss.getSheetByName('ReglasCategorizacion');
  if (hoja.getLastRow() > 1) return;
  const reglas = [
    ['taxi', 'Transporte', 'Taxi', 'Personal', 2], ['uber', 'Transporte', 'Uber', 'Personal', 2], ['bus', 'Transporte', 'Bus', 'Personal', 2],
    ['gasolina', 'Transporte', 'Combustible', 'Personal', 2], ['almuerzo', 'Alimentacion', 'Almuerzo', 'Personal', 2], ['cena', 'Alimentacion', 'Cena', 'Personal', 2],
    ['desayuno', 'Alimentacion', 'Desayuno', 'Personal', 2], ['mercado', 'Alimentacion', 'Mercado', 'Personal', 2], ['supermercado', 'Alimentacion', 'Supermercado', 'Personal', 2],
    ['delivery', 'Alimentacion', 'Delivery', 'Personal', 1], ['pedidosya', 'Alimentacion', 'Delivery', 'Personal', 1], ['rappi', 'Alimentacion', 'Delivery', 'Personal', 1],
    ['netflix', 'Entretenimiento', 'Netflix', 'Personal', 1], ['spotify', 'Entretenimiento', 'Spotify', 'Personal', 1], ['luz', 'Servicios', 'Luz', 'Personal', 2],
    ['agua', 'Servicios', 'Agua', 'Personal', 3], ['internet', 'Servicios', 'Internet', 'Personal', 2], ['celular', 'Servicios', 'Celular', 'Personal', 2],
    ['farmacia', 'Salud', 'Farmacia', 'Personal', 2], ['doctor', 'Salud', 'Consulta', 'Personal', 2], ['gym', 'Salud', 'Gimnasio', 'Personal', 1],
    ['sueldo', 'Sueldo', '', 'Personal', 1]
  ];
  hoja.getRange(2, 1, reglas.length, 6).setValues(reglas.map(function (r, i) { return [i + 1].concat(r); }));
}

function migrarTarjetasLegado_(h) { /* sin tarjeta por defecto: cada persona crea la suya */ }

SYSTEM_ASESOR_ = 'Eres el asesor financiero personal de un usuario que vive en Perú y maneja soles, dólares y euros. ' + SYSTEM_ASESOR_.replace(/^[^.]*\.[^.]*\.\s*/, '');

(function () {
  const original = actualizarTipoCambioReferencial;
  const urls = ['https://open.er-api.com/v6/latest/PEN', 'https://api.exchangerate-api.com/v4/latest/PEN'];
  actualizarTipoCambioReferencial = function () {
    for (let i = 0; i < urls.length; i++) {
      CONFIG.API_TIPO_CAMBIO_URL = urls[i];
      const antes = getHoja_('TipoCambioReferencial').getLastRow();
      original();
      if (getHoja_('TipoCambioReferencial').getLastRow() > antes) return true;
    }
    return false;
  };
})();

/** Tareas del día al abrir la app (en la nube las hacían los disparadores automáticos) */
function tareasLocales_() {
  asegurarEstructura_();
  const hoy = Utilities.formatDate(new Date(), '', 'yyyy-MM-dd');
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('LOCAL_TAREAS_DIA') === hoy) return false;
  ejecutarSeguro_('Recurrentes', procesarRecurrentesDelDia);
  ejecutarSeguro_('TipoCambio', actualizarTipoCambioReferencial);
  p.setProperty('LOCAL_TAREAS_DIA', hoy);
  return true;
}

function medioGastoPorDefecto_(moneda) {
  const m = String(getAjuste_('medio.gasto.' + moneda, 'Efectivo'));
  return ['Tarjeta', 'Efectivo', 'Debito'].indexOf(m) !== -1 ? m : 'Efectivo';
}
function medioIngresoPorDefecto_(moneda) {
  return getAjuste_('medio.ingreso.' + moneda, 'Efectivo') === 'Debito' ? 'Debito' : 'Efectivo';
}

// ---------------- datos para reportes, avisos y exportación a Google Sheets ----------------
function dmy_(d) { return d.getDate() + '/' + (d.getMonth() + 1); }

function api_datosReporte(a) {
  const ref = a.fecha ? fecha_(a.fecha) : new Date();
  const movs = leerMovimientos_();
  const anual = a.tipo === 'anual';
  const rango = obtenerRangoFecha_(anual ? 'Anual' : 'Mensual', ref);
  const an = analizarPeriodo_(rango, movs);
  const capIni = capitalAntesDe_(rango.inicio, movs), capFin = capitalAntesDe_(new Date(rango.fin.getTime() + 1), movs);
  const out = { tipo: anual ? 'anual' : 'mensual', titulo: anual ? 'Año ' + ref.getFullYear() : nombreMes_(ref) + ' ' + ref.getFullYear(), capIni: capIni, capFin: capFin, ingresos: an.ingresos, egresos: an.egresos, totales: an.totales, porCategoria: an.porCategoria };
  const tasa = obtenerUltimaTasaReferencial_();
  if (!anual) {
    out.metas = {}; out.faltan = {};
    CONFIG.MONEDAS.forEach(function (k) { const p = calcularProgresoMeta('Mensual', k, ref, movs); out.metas[k] = p.meta; out.faltan[k] = p.meta > 0 ? p.falta : 0; });
    out.secciones = {};
    Object.keys(an.secciones).forEach(function (s) {
      out.secciones[s] = an.secciones[s].map(function (m) { return { f: dmy_(m.efectiva), d: m.descripcion || '(sin descripción)', m: m.moneda, v: m.monto }; });
    });
  } else {
    out.meses = [];
    for (let mes = 0; mes < 12; mes++) {
      const x = analizarPeriodo_(obtenerRangoFecha_('Mensual', new Date(ref.getFullYear(), mes, 15)), movs);
      out.meses.push({ mes: MESES_ES[mes].charAt(0).toUpperCase() + MESES_ES[mes].slice(1), ing: x.ingresos, egr: x.egresos });
    }
    out.distribucion = ['Gastos personales', 'Inversión', 'Gastos fijos'].map(function (s) {
      return { s: s, v: redondear_(CONFIG.MONEDAS.reduce(function (acc, k) { return acc + convertirAPEN_(an.totales[s][k], k, tasa); }, 0)) };
    });
  }
  return out;
}

function api_datosAvisos() {
  const prog = listarProgramados_('Pendiente').concat(proximosRecurrentes_()).map(function (p) {
    return { fecha: fechaISO_(p.fecha), descripcion: p.descripcion || p.categoria || '', monto: p.monto, moneda: p.moneda, tipo: p.tipo, atrasado: !!p.atrasado };
  });
  const tj = leerTarjetas_().filter(function (t) { return t.activa; }).map(function (t) { return { banco: t.banco, cierre: t.cierre, cobro: t.cobro }; });
  return { programados: prog, tarjetas: tj, hoy: fechaISO_(new Date()) };
}

const HOJAS_EXPORTAR_ = ['Movimientos', 'Cuentas', 'Presupuestos', 'MetasAhorro', 'Programados', 'Recurrentes', 'PrestamosOtorgados', 'Deudas', 'CronogramaCuotas', 'Inversiones', 'Tarjetas', 'Categorias'];
function api_datosHoja() {
  const ss = getSS_(), out = [];
  HOJAS_EXPORTAR_.forEach(function (n) {
    const h = ss.getSheetByName(n);
    if (!h || h.getLastRow() < 1) return;
    const w = h.getLastColumn();
    const v = h.getRange(1, 1, h.getLastRow(), w).getValues().map(function (f) {
      return f.map(function (c) {
        if (c instanceof Date) return Utilities.formatDate(c, '', 'yyyy-MM-dd HH:mm');
        return c === null || c === undefined ? '' : c;
      });
    });
    out.push({ nombre: n, valores: v });
  });
  return out;
}
