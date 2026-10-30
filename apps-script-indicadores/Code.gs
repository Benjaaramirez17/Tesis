/**
 * Indicadores multi-tienda (L94, L120, L657)
 * Web App: promedio del día (Resumen Diario) + consolidado semanal (Historico),
 * una columna por tienda.
 */

// ===================== CONFIGURACIÓN =====================
const CONFIG = {
  SOURCE_ID: '1SpQH5HsR_Ghp6UqCg1Q2xsfgaD0PsHeRiAjU8sgn0H0',
  HOJA_DIARIO: 'Resumen Diario',
  HOJA_HISTORICO: 'Historico',
  LOCALES: ['94', '120', '657'],

  // Correo diario. Vacío = se envía a la cuenta dueña del script.
  // Para varios destinatarios: 'correo1@x.cl, correo2@x.cl'
  EMAIL_DESTINO: '',
  HORA_CORREO: 8
};

const TZ = 'America/Santiago';
const DIAS = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];

// ===================== WEB APP =====================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Indicadores Tiendas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Devuelve todo lo necesario para pintar las 3 columnas en una sola llamada. */
function getIndicadores() {
  const diario = leerHoja_(CONFIG.HOJA_DIARIO);
  const historico = leerHoja_(CONFIG.HOJA_HISTORICO);

  const hoy = startOfDay_(new Date());
  const lunesActual = addDays_(hoy, -((hoy.getDay() + 6) % 7));
  const lunesAnterior = addDays_(lunesActual, -7);

  return {
    actualizado: Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm'),
    tiendas: CONFIG.LOCALES.map(local => {
      const fd = diario.filter(f => f.local === local);
      const fh = historico.filter(f => f.local === local);
      return {
        local: local,
        dia: promedioDia_(fd),
        anterior: consolidarSemana_(fh, lunesAnterior, addDays_(lunesAnterior, 6)),
        actual: consolidarSemana_(fh, lunesActual, hoy)
      };
    })
  };
}

// ===================== CÁLCULOS =====================
function promedioDia_(filas) {
  const hoyKey = dayKey_(new Date());
  // Si hoy aún no hay datos, se usa la última fecha disponible
  let fechaKey = hoyKey;
  if (!filas.some(f => f.key === hoyKey)) {
    fechaKey = filas.reduce((m, f) => (f.key > m ? f.key : m), '');
  }

  const porInd = {};
  filas.filter(f => f.key === fechaKey).forEach(f => {
    const ind = porInd[f.indicador] || (porInd[f.indicador] = { valores: [], porHora: {}, pct: false });
    ind.valores.push(f.valor);
    ind.pct = ind.pct || f.pct;
    if (f.hora) (ind.porHora[f.hora] = ind.porHora[f.hora] || []).push(f.valor);
  });

  return {
    fecha: fechaKey ? keyToLabel_(fechaKey) : '',
    esHoy: fechaKey === hoyKey,
    indicadores: Object.keys(porInd).sort().map(k => {
      const i = porInd[k];
      const ultimaHora = Object.keys(i.porHora).sort().pop();
      return {
        indicador: k,
        pct: i.pct,
        promedio: avg_(i.valores),
        ultimo: ultimaHora ? avg_(i.porHora[ultimaHora]) : null,
        ultimaHora: ultimaHora || ''
      };
    })
  };
}

function consolidarSemana_(filas, desde, hasta) {
  const dias = [];
  for (let d = new Date(desde); d <= hasta; d = addDays_(d, 1)) dias.push(dayKey_(d));
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
    semana: isoWeek_(desde),
    rango: keyToLabel_(dayKey_(desde)).slice(0, 5) + ' al ' + keyToLabel_(dayKey_(addDays_(desde, 6))).slice(0, 5),
    dias: dias.map((k, i) => DIAS[i] + ' ' + k.slice(8, 10)),
    indicadores: Object.keys(porInd).sort().map(k => ({
      indicador: k,
      pct: porInd[k].pct,
      porDia: dias.map(d => (porInd[k].porDia[d] ? avg_(porInd[k].porDia[d]) : null)),
      promedio: avg_(porInd[k].valores)
    }))
  };
}

// ===================== CORREO DIARIO =====================
/**
 * Ejecutar UNA vez a mano (menú Ejecutar) para dejar programado el correo diario.
 * Google lo dispara dentro de la ventana 8:00–9:00.
 */
function crearTriggerDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarResumenDiario')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarResumenDiario')
    .timeBased().everyDays(1).atHour(CONFIG.HORA_CORREO).inTimezone(TZ).create();
}

/** Envía por correo los indicadores del día anterior de cada tienda. */
function enviarResumenDiario() {
  const hoy = startOfDay_(new Date());
  const ayer = addDays_(hoy, -1);
  const semAnt = addDays_(ayer, -7);
  const kAyer = dayKey_(ayer), kSemAnt = dayKey_(semAnt);

  // El día anterior se busca en Historico; si aún no está, en Resumen Diario
  const historico = leerHoja_(CONFIG.HOJA_HISTORICO);
  let filas = historico;
  if (!historico.some(f => f.key === kAyer)) {
    filas = historico.concat(leerHoja_(CONFIG.HOJA_DIARIO).filter(f => f.key === kAyer));
  }

  // datos[indicador][local] = { prom, base, pct }
  const datos = {};
  filas.forEach(f => {
    if (f.key !== kAyer && f.key !== kSemAnt) return;
    const d = (datos[f.indicador] = datos[f.indicador] || {});
    const c = (d[f.local] = d[f.local] || { ayer: [], base: [], pct: false });
    (f.key === kAyer ? c.ayer : c.base).push(f.valor);
    c.pct = c.pct || f.pct;
  });

  const locales = CONFIG.LOCALES.map(l => String(Number(l)));
  const indicadores = Object.keys(datos).sort()
    .filter(i => locales.some(l => datos[i][l] && datos[i][l].ayer.length));

  const fechaTxt = Utilities.formatDate(ayer, TZ, 'dd-MM-yyyy');
  const diaSem = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'][ayer.getDay()];

  const th = 'background:#0b1f4b;color:#fff;padding:8px 10px;font-size:13px;text-align:center;';
  const td = 'padding:8px 10px;border-bottom:1px solid #e3e7f0;font-size:13px;text-align:center;';
  let html = '<div style="font-family:Arial,sans-serif;color:#1c2340;max-width:680px">' +
    '<div style="background:#0b1f4b;color:#fff;padding:14px 18px;font-weight:bold;font-size:17px">' +
    'INDICADORES TIENDAS · ' + diaSem + ' ' + fechaTxt + '</div>' +
    '<div style="background:#d63f8c;color:#fff;padding:8px;text-align:center;font-weight:bold;font-size:13px">' +
    'PROMEDIO DEL DÍA ANTERIOR</div>';

  if (!indicadores.length) {
    html += '<p style="padding:14px">No se encontraron datos para el ' + fechaTxt + '.</p>';
  } else {
    html += '<table style="border-collapse:collapse;width:100%"><tr><th style="' + th + 'text-align:left">Indicador</th>' +
      locales.map(l => '<th style="' + th + '">L' + l + '</th>').join('') + '</tr>';
    indicadores.forEach((ind, n) => {
      const bg = n % 2 ? 'background:#f7f9fd;' : '';
      html += '<tr><td style="' + td + bg + 'text-align:left;font-weight:bold">' + ind + '</td>';
      locales.forEach(l => {
        const c = datos[ind][l];
        if (!c || !c.ayer.length) { html += '<td style="' + td + bg + '">–</td>'; return; }
        const prom = avg_(c.ayer), base = avg_(c.base);
        html += '<td style="' + td + bg + '"><b>' + fmtNum_(prom, c.pct) + '</b>' +
          varTxt_(prom, base, c.pct) + '</td>';
      });
      html += '</tr>';
    });
    html += '</table><p style="font-size:11px;color:#7a829b;padding:8px 2px">Variación vs. mismo día de la ' +
      'semana anterior (' + Utilities.formatDate(semAnt, TZ, 'dd-MM') + '). En porcentajes se expresa en puntos (pp).</p>';
  }
  html += '</div>';

  const destino = CONFIG.EMAIL_DESTINO || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: destino,
    subject: 'Indicadores L' + locales.join(' · L') + ' – ' + fechaTxt,
    htmlBody: html
  });
}

function fmtNum_(v, pct) {
  const n = pct ? v * 100 : v;
  const txt = (Math.round(n * (pct ? 10 : 100)) / (pct ? 10 : 100)).toString().replace('.', ',');
  return txt + (pct ? '%' : '');
}

function varTxt_(actual, base, pct) {
  if (base === null) return '';
  const diff = actual - base;
  const n = pct ? diff * 100 : (base ? diff / Math.abs(base) * 100 : diff);
  const r = Math.round(n * 10) / 10;
  const unidad = pct ? ' pp' : (base ? '%' : '');
  if (r === 0) return '<br><span style="font-size:11px;color:#7a829b">= 0' + unidad + '</span>';
  const color = r > 0 ? '#178a4a' : '#c0392b';
  return '<br><span style="font-size:11px;color:' + color + '">' + (r > 0 ? '▲ +' : '▼ ') +
    String(r).replace('.', ',') + unidad + '</span>';
}

// ===================== LECTURA DE DATOS =====================
/**
 * Lee una hoja y devuelve las filas de los locales configurados:
 * { local, key: 'yyyy-MM-dd', indicador, hora, valor, pct }
 * Las columnas se buscan por nombre (Fecha, Local, Indicador, Hora, Valor).
 */
function leerHoja_(nombre) {
  const sh = SpreadsheetApp.openById(CONFIG.SOURCE_ID).getSheetByName(nombre);
  if (!sh) throw new Error('No se encontró la hoja "' + nombre + '"');
  const rango = sh.getDataRange();
  const values = rango.getValues();
  const display = rango.getDisplayValues();

  let hRow = -1, col = {};
  for (let r = 0; r < Math.min(10, values.length); r++) {
    const norm = values[r].map(normaliza_);
    if (norm.indexOf('fecha') > -1 && norm.indexOf('valor') > -1) {
      hRow = r;
      ['fecha', 'local', 'indicador', 'hora', 'valor'].forEach(c => { col[c] = norm.indexOf(c); });
      break;
    }
  }
  if (hRow < 0) throw new Error('No se encontraron encabezados Fecha/Valor en "' + nombre + '"');

  const locales = new Set(CONFIG.LOCALES.map(l => String(Number(l))));
  const out = [];
  for (let r = hRow + 1; r < values.length; r++) {
    const v = values[r];
    const local = localDe_(v[col.local]);
    if (!locales.has(local)) continue;
    const fecha = parseFecha_(v[col.fecha]);
    const valor = parseNum_(v[col.valor]);
    if (!fecha || valor === null) continue;
    out.push({
      local: local,
      key: dayKey_(fecha),
      indicador: String(v[col.indicador] || 'Sin indicador').trim(),
      hora: col.hora > -1 ? parseHora_(v[col.hora], display[r][col.hora]) : '',
      valor: valor,
      pct: String(display[r][col.valor]).indexOf('%') > -1
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
