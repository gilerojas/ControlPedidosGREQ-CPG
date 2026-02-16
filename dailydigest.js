/**
 * DAILY DIGEST GREQ – Versión mejorada con formato WhatsApp
 * Resumen diario con negritas estratégicas y mejor legibilidad
 * 
 * Autor: Gilberto Rojas (GREQ)
 * Fecha: Noviembre 2025
 * Versión: 2.0
 */

// ─────────────────────────────── CONFIG ───────────────────────────────

const TZ = 'America/Santo_Domingo';
const clean = str => String(str || '').replace(/\s+/g, ' ').trim();

// ─────────────────────────────── ENV HELPERS ───────────────────────────────

/**
 * Cambia rápidamente el entorno sin abrir Script Properties.
 * Ej.: setEnv('PROD')  o  setEnv('TEST')
 */
function setEnv(newEnv) {
  PropertiesService.getScriptProperties()
    .setProperty('ENV', newEnv.toUpperCase());
  Logger.log('ENV cambiado a: ' + newEnv.toUpperCase());
}

function debugDigestConfig() {
  const p        = PropertiesService.getScriptProperties();
  const env      = (p.getProperty('ENV') || '—vacío—').toUpperCase();
  const testJID  = p.getProperty('GROUP_ID_TEST')      || '—vacío—';
  const prodJID  = p.getProperty('GROUP_GREQ_MAIN')    || '—vacío—';
  const chosen   = env === 'TEST' ? testJID : prodJID;

  Logger.log('ENV              : ' + env);
  Logger.log('GROUP_ID_TEST     : ' + testJID);
  Logger.log('GROUP_GREQ_MAIN   : ' + prodJID);
  Logger.log('JID que se enviará: ' + chosen);
}

// ─────────────────────────────── BUILDERS ───────────────────────────────

function buildGreeting() {
  return '☀️ *¡Buenos días, equipo GREQ!*';
}

function buildClosing() {
  return '💪 *¡Buen inicio de jornada—manos a la obra!*';
}

/**
 * Genera el bloque "Resumen de Pedidos" con formato mejorado para WhatsApp.
 * Usa negritas (*texto*) en información clave: clientes, cantidades, IDs.
 */
/**
 * Genera el bloque "Resumen de Pedidos" con TODOS los estados operativos.
 * Estados incluidos: PENDIENTE, EN PRODUCCION, ESPERA DE CALIDAD, 
 *                    LISTO P/ ENVASAR, LISTO P/ DESPACHAR
 * Estados excluidos: DESPACHADO, CANCELADO
 */
function generarResumenPedidos() {
  const ss   = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName('Pedidos');
  const data = hoja.getDataRange().getValues().slice(1);   // sin encabezados

  const estadosMap = {
    'LISTO P/ DESPACHAR':  { emoji:'🟢', nombre:'LISTO P/ DESPACHAR', items:[], orden:5 },
    'LISTO P/ ENVASAR':    { emoji:'🟡', nombre:'LISTO P/ ENVASAR', items:[], orden:4 },
    'ESPERA DE CALIDAD':   { emoji:'🔬', nombre:'ESPERA DE CALIDAD', items:[], orden:3 },
    'EN PRODUCCION':       { emoji:'🟣', nombre:'EN PRODUCCIÓN', items:[], orden:2 },
    'PENDIENTE':           { emoji:'🟠', nombre:'PENDIENTE', items:[], orden:1 },
  };

  data.forEach(r => {
    const estado   = r[10];
    if (estado === 'DESPACHADO' || estado === 'CANCELADO') return;

    const urgencia = clean(r[9]).toUpperCase();
    const urgente  = urgencia === 'ALTA' ? '❗*URGENTE*❗ ' : '';

    const cliente  = clean(r[2]);
    const producto = clean(r[3]);
    const color    = clean(r[4]);
    const cantidad = r[5];
    const unidad   = r[6];
    const promesa  = new Date(r[7]);
    const pedId    = r[11];

    const promesaTxt = (promesa instanceof Date && !isNaN(promesa))
      ? ` – Promesa *${Utilities.formatDate(promesa, TZ, 'dd-MMM').toLowerCase()}*`
      : '';

    // Formato con negritas en información clave
    const linea = `• ${urgente}*${cliente}* – *${cantidad}* ${unidad} ${producto} ${color}${promesaTxt} _(ID ${pedId})_`;

    if (estadosMap[estado]) {
      estadosMap[estado].items.push({ promesa, linea });
    }
  });

  // Ordenar items por fecha de promesa dentro de cada estado
  Object.values(estadosMap).forEach(obj => {
    obj.items.sort((a, b) => {
      const av = a.promesa instanceof Date && !isNaN(a.promesa) ? a.promesa : Infinity;
      const bv = b.promesa instanceof Date && !isNaN(b.promesa) ? b.promesa : Infinity;
      return av - bv;
    });
  });

  const hoyTxt = Utilities.formatDate(new Date(), TZ, 'dd-MMM-yyyy');
  let msg = `📅 *Resumen de Pedidos* — ${hoyTxt} (08:00)\n\n`;

  // Ordenar estados por campo 'orden' (inverso: del más avanzado al más atrasado)
  const estadosOrdenados = Object.entries(estadosMap)
    .sort((a, b) => b[1].orden - a[1].orden);

  // Construir secciones solo si tienen items
  estadosOrdenados.forEach(([estado, obj]) => {
    if (obj.items.length) {
      msg += `${obj.emoji} *${obj.nombre}* (${obj.items.length})\n`;
      obj.items.forEach(it => msg += `${it.linea}\n`);
      msg += '\n';
    }
  });

  msg += '⚠️ *Acciones clave*\n';
  msg += '• Priorizar pedidos *URGENTES* y con promesa más cercana\n';
  msg += '• Verificar etiquetas y transporte para "*Listo p/ Despachar*"';

  return msg.trim();
}

// ─────────────────────────────── MAIN DIGEST ───────────────────────────────

/**
 * Envía el resumen diario al grupo de WhatsApp indicado por ENV.
 * No se envía los domingos.
 */
function sendDailyDigest() {
  const hoy = new Date();
  const diaSemana = hoy.getDay(); // 0 = domingo

  if (diaSemana === 0) {
    Logger.log('Hoy es domingo. No se envía resumen.');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const env   = (props.getProperty('ENV') || 'PROD').toUpperCase();
  const token = props.getProperty('WAS_TOKEN');

  const group = env === 'TEST'
    ? props.getProperty('GROUP_ID_TEST')
    : props.getProperty('GROUP_GREQ_MAIN');

  const saludo  = buildGreeting();
  const resumen = generarResumenPedidos();
  const cierre  = buildClosing();
  
  // Mensaje final con separación clara entre bloques
  const text = `${saludo}\n\n${resumen}\n\n${cierre}`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify({ to: group, text }),
    muteHttpExceptions: true,
  };

  const resp = UrlFetchApp.fetch('https://www.wasenderapi.com/api/send-message', options);
  Logger.log('WA response: ' + resp.getContentText());
}

// ─────────────────────────────── TESTING ───────────────────────────────

/**
 * Función de prueba para ver el mensaje antes de enviarlo.
 * Ejecutar directamente desde el editor de Apps Script.
 */
function testDailyDigestFormat() {
  const saludo  = buildGreeting();
  const resumen = generarResumenPedidos();
  const cierre  = buildClosing();
  const mensaje = `${saludo}\n\n${resumen}\n\n${cierre}`;
  
  Logger.log('════════════════════════════════');
  Logger.log('PREVIEW DEL MENSAJE:');
  Logger.log('════════════════════════════════');
  Logger.log(mensaje);
  Logger.log('════════════════════════════════');
}