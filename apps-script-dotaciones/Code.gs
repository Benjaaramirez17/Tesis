/**
 * Reporte Dotaciones L120 + Indicadores
 * Web App con 2 pestañas: "Dotación" e "Indicadores".
 */

// ===================== CONFIGURACIÓN =====================
const CONFIG = {
  // Planilla con los indicadores (Resumen Diario / Histórico)
  SOURCE_ID: '1SpQH5HsR_Ghp6UqCg1Q2xsfgaD0PsHeRiAjU8sgn0H0',
  HOJA_DIARIO: 'Resumen Diario',
  HOJA_HISTORICO: 'Historico',

  // Local a filtrar
  LOCAL: '120',

  // Filas que se leen por bloque al buscar desde el final de la hoja
  BLOQUE_LECTURA: 3000,

  // Planilla donde se guardan los reportes de dotación.
  // Por defecto la misma planilla de indicadores; cámbialo si quieres otra.
  REGISTRO_ID: '1SpQH5HsR_Ghp6UqCg1Q2xsfgaD0PsHeRiAjU8sgn0H0',
  HOJA_REGISTRO: 'Registro Dotación',

  COORDINADORES: [
    'Monica Flores',
    'Miguel Rodríguez',
    'Bárbara Valdebenito',
    'Pamela Figueroa'
  ],

  DOTACIONES: ['Shoppers', 'Pickers Crowd', 'Pickers Out']
};

const TZ = 'America/Santiago';
const DIAS = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];

// ===================== WEB APP =====================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Reporte Dotaciones L' + CONFIG.LOCAL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getConfig() {
  return {
    local: CONFIG.LOCAL,
    coordinadores: CONFIG.COORDINADORES,
    dotaciones: CONFIG.DOTACIONES,
    hoy: Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy')
  };
}

// ===================== DOTACIÓN =====================
/**
 * data = { coordinadores: [..], dotacion: { Shoppers: {am, pm}, ... }, novedades: '' }
 */
function guardarDotacion(data) {
  const ss = SpreadsheetApp.openById(CONFIG.REGISTRO_ID);
  let sh = ss.getSheetByName(CONFIG.HOJA_REGISTRO);
  const headers = ['Timestamp', 'Fecha', 'Local', 'Coordinadores'];
  CONFIG.DOTACIONES.forEach(d => headers.push(d + ' AM', d + ' PM'));
  headers.push('Total AM', 'Total PM', 'Novedades', 'Usuario');

  if (!sh) {
    sh = ss.insertSheet(CONFIG.HOJA_REGISTRO);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#0b1f4b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }

  const now = new Date();
  const row = [now, Utilities.formatDate(now, TZ, 'dd-MM-yyyy'), CONFIG.LOCAL,
    (data.coordinadores || []).join(', ')];
  let totAm = 0, totPm = 0;
  CONFIG.DOTACIONES.forEach(d => {
    const v = (data.dotacion && data.dotacion[d]) || {};
    const am = Number(v.am) || 0, pm = Number(v.pm) || 0;
    totAm += am; totPm += pm;
    row.push(am, pm);
  });
  row.push(totAm, totPm, data.novedades || '', Session.getActiveUser().getEmail());
  sh.appendRow(row);
  return { ok: true, fecha: row[1] };
}

// ===================== INDICADORES =====================
/** Promedio del día (Resumen Diario) para el local configurado. */
function getIndicadoresDia() {
  const filas = leerHoja_(CONFIG.HOJA_DIARIO, dayKey_(addDays_(new Date(), -7)));
  const hoyKey = dayKey_(new Date());

  // Si hoy aún no hay datos, mostrar la última fecha disponible
  let fechaKey = hoyKey;
  if (!filas.some(f => f.key === hoyKey)) {
    fechaKey = filas.reduce((m, f) => (f.key > m ? f.key : m), '');
  }
  const delDia = filas.filter(f => f.key === fechaKey);

  const porInd = {};
  const horas = new Set();
  delDia.forEach(f => {
    const ind = porInd[f.indicador] || (porInd[f.indicador] = {
      indicador: f.indicador, valores: [], porHora: {}, pct: false
    });
    ind.valores.push(f.valor);
    ind.pct = ind.pct || f.pct;
    if (f.hora) {
      horas.add(f.hora);
      (ind.porHora[f.hora] = ind.porHora[f.hora] || []).push(f.valor);
    }
  });

  const horasOrd = Array.from(horas).sort();
  const indicadores = Object.keys(porInd).sort().map(k => {
    const i = porInd[k];
    const ultimaHora = Object.keys(i.porHora).sort().pop();
    return {
      indicador: i.indicador,
      pct: i.pct,
      promedio: avg_(i.valores),
      min: Math.min.apply(null, i.valores),
      max: Math.max.apply(null, i.valores),
      registros: i.valores.length,
      ultimo: ultimaHora ? avg_(i.porHora[ultimaHora]) : null,
      ultimaHora: ultimaHora || '',
      porHora: horasOrd.map(h => (i.porHora[h] ? avg_(i.porHora[h]) : null))
    };
  });

  return {
    fecha: fechaKey ? keyToLabel_(fechaKey) : '',
    esHoy: fechaKey === hoyKey,
    horas: horasOrd,
    indicadores: indicadores,
    actualizado: Utilities.formatDate(new Date(), TZ, 'HH:mm')
  };
}

/** Consolidado semana anterior completa (Lun-Dom) + semana en curso (Histórico). */
function getConsolidadoSemanal() {
  const hoy = startOfDay_(new Date());
  const lunesActual = addDays_(hoy, -((hoy.getDay() + 6) % 7));
  const lunesAnterior = addDays_(lunesActual, -7);
  const filas = leerHoja_(CONFIG.HOJA_HISTORICO, dayKey_(lunesAnterior));

  return {
    anterior: consolidarSemana_(filas, lunesAnterior, addDays_(lunesAnterior, 6)),
    actual: consolidarSemana_(filas, lunesActual, hoy)
  };
}

function consolidarSemana_(filas, desde, hasta) {
  const dias = [];
  for (let d = new Date(desde); d <= hasta; d = addDays_(d, 1)) dias.push(dayKey_(d));
  const fin = addDays_(desde, 6);
  const keys = new Set(dias);

  const porInd = {};
  filas.forEach(f => {
    if (!keys.has(f.key)) return;
    const ind = porInd[f.indicador] || (porInd[f.indicador] = { valores: [], porDia: {}, pct: false });
    ind.valores.push(f.valor);
    ind.pct = ind.pct || f.pct;
    (ind.porDia[f.key] = ind.porDia[f.key] || []).push(f.valor);
  });

  return {
    titulo: 'Semana ' + isoWeek_(desde) + ' · ' + keyToLabel_(dayKey_(desde)) +
      ' al ' + keyToLabel_(dayKey_(fin)),
    dias: dias.map((k, i) => DIAS[i] + ' ' + k.slice(8, 10)),
    indicadores: Object.keys(porInd).sort().map(k => ({
      indicador: k,
      pct: porInd[k].pct,
      porDia: dias.map(d => (porInd[k].porDia[d] ? avg_(porInd[k].porDia[d]) : null)),
      promedio: avg_(porInd[k].valores)
    }))
  };
}

// ===================== LECTURA DE DATOS =====================
/**
 * Lee una hoja y devuelve las filas del local configurado:
 * { local, key: 'yyyy-MM-dd', indicador, hora, valor, pct }
 * Las columnas se buscan por nombre (Fecha, Local, Indicador, Hora, Valor).
 * Para no leer hojas enormes completas: primero lee solo la columna Fecha,
 * ubica el bloque de filas con fecha >= desdeKey y lee únicamente ese bloque.
 */
function leerHoja_(nombre, desdeKey) {
  const sh = SpreadsheetApp.openById(CONFIG.SOURCE_ID).getSheetByName(nombre);
  if (!sh) throw new Error('No se encontró la hoja "' + nombre + '"');
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  // Encabezados (primeras 10 filas)
  const cab = sh.getRange(1, 1, Math.min(10, lastRow), lastCol).getValues();
  let hRow = -1, col = {};
  for (let r = 0; r < cab.length; r++) {
    const norm = cab[r].map(normaliza_);
    if (norm.indexOf('fecha') > -1 && norm.indexOf('valor') > -1) {
      hRow = r;
      ['fecha', 'local', 'indicador', 'hora', 'valor'].forEach(c => { col[c] = norm.indexOf(c); });
      break;
    }
  }
  if (hRow < 0) throw new Error('No se encontraron encabezados Fecha/Valor en "' + nombre + '"');
  if (col.local < 0 || col.indicador < 0) {
    throw new Error('Faltan columnas Local/Indicador en "' + nombre + '"');
  }

  // 1) Columna Fecha leída desde el final hacia arriba, en bloques: como los datos
  //    más nuevos están abajo, se detiene al encontrar una fecha anterior a desdeKey.
  const primeraFila = hRow + 2;
  if (lastRow < primeraFila) return [];
  let keys = [], inicio = lastRow + 1;
  while (inicio > primeraFila) {
    const desde = Math.max(primeraFila, inicio - CONFIG.BLOQUE_LECTURA);
    const fechas = sh.getRange(desde, col.fecha + 1, inicio - desde, 1).getValues();
    const bloque = fechas.map(r => { const f = parseFecha_(r[0]); return f ? dayKey_(f) : null; });
    keys = bloque.concat(keys);
    inicio = desde;
    if (desdeKey && bloque.some(k => k && k < desdeKey)) break;
  }
  const n = keys.length;
  let primera = -1, ultima = -1;
  for (let i = 0; i < n; i++) {
    if (keys[i] && (!desdeKey || keys[i] >= desdeKey)) {
      if (primera < 0) primera = i;
      ultima = i;
    }
  }
  if (primera < 0) return [];

  // 2) Solo el bloque de filas y columnas necesarias
  const usadas = [col.local, col.indicador, col.hora, col.valor].filter(c => c > -1);
  const c0 = Math.min.apply(null, usadas), c1 = Math.max.apply(null, usadas);
  const filas = ultima - primera + 1;
  const values = sh.getRange(inicio + primera, c0 + 1, filas, c1 - c0 + 1).getValues();
  const dispValor = sh.getRange(inicio + primera, col.valor + 1, filas, 1).getDisplayValues();
  const dispHora = col.hora > -1 ? sh.getRange(inicio + primera, col.hora + 1, filas, 1).getDisplayValues() : null;

  const locales = new Set([String(Number(CONFIG.LOCAL))]);
  const out = [];
  for (let i = 0; i < filas; i++) {
    const key = keys[primera + i];
    if (!key || (desdeKey && key < desdeKey)) continue;
    const v = values[i];
    const local = localDe_(v[col.local - c0]);
    if (!locales.has(local)) continue;
    const valor = parseNum_(v[col.valor - c0]);
    if (valor === null) continue;
    out.push({
      local: local,
      key: key,
      indicador: String(v[col.indicador - c0] || 'Sin indicador').trim(),
      hora: col.hora > -1 ? parseHora_(v[col.hora - c0], dispHora[i][0]) : '',
      valor: valor,
      pct: String(dispValor[i][0]).indexOf('%') > -1
    });
  }
  return out;
}

/** "L120", "120 - Nombre", 120 → "120" */
function localDe_(v) {
  const m = String(v).match(/\d+/);
  return m ? String(Number(m[0])) : '';
}

function normaliza_(s) {
  return String(s).trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function parseFecha_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/); // dd-mm-yyyy
  if (m) {
    const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
    return new Date(y, Number(m[2]) - 1, Number(m[1]));
  }
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); // yyyy-mm-dd
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

function parseHora_(v, disp) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  if (typeof v === 'number' && v < 1) { // fracción de día
    const min = Math.round(v * 1440);
    return ('0' + Math.floor(min / 60)).slice(-2) + ':' + ('0' + (min % 60)).slice(-2);
  }
  if (typeof v === 'number') return ('0' + v).slice(-2) + ':00'; // hora entera
  const s = String(disp || v).trim();
  const m = s.match(/^(\d{1,2})(?::(\d{2}))?/);
  return m ? ('0' + m[1]).slice(-2) + ':' + (m[2] || '00') : s;
}

function parseNum_(v) {
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  if (!s) return null;
  const pct = s.indexOf('%') > -1;
  s = s.replace(/[%\s$]/g, '');
  if (s.indexOf(',') > -1) s = s.replace(/\./g, '').replace(',', '.'); // formato 1.234,5
  const n = Number(s);
  if (isNaN(n)) return null;
  return pct ? n / 100 : n;
}

// ===================== UTILIDADES =====================
function avg_(arr) {
  return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
}
function startOfDay_(d) {
  const k = dayKey_(d).split('-');
  return new Date(Number(k[0]), Number(k[1]) - 1, Number(k[2]));
}
function addDays_(d, n) {
  const x = new Date(d); x.setDate(x.getDate() + n); return x;
}
function dayKey_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function keyToLabel_(k) {
  const p = k.split('-'); return p[2] + '-' + p[1] + '-' + p[0];
}
function isoWeek_(d) {
  const t = startOfDay_(d);
  t.setDate(t.getDate() + 3 - ((t.getDay() + 6) % 7)); // jueves de esa semana
  const ene4 = new Date(t.getFullYear(), 0, 4);
  return 1 + Math.round(((t - ene4) / 864e5 - 3 + ((ene4.getDay() + 6) % 7)) / 7);
}
