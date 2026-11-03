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
