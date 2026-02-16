// ───────────────────────────────────────────────────────────
// Code.gs – tablero PEDIDOS + WhatsApp (OCT 2025 - Guiones + Negritas)
// Formato: Guiones simples + Negritas (*texto*)
// WhatsApp-friendly, móvil-optimizado
// ───────────────────────────────────────────────────────────

/**
 * Obtener cantidad actual de un ID en inventario
 */
function obtenerStockActual(idInventario) {
  try {
    const inventarioSS = SpreadsheetApp.openById('1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw');
    const inventarioSheet = inventarioSS.getSheetByName('Inventario');
    
    const datos = inventarioSheet.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][5] === idInventario) { // Columna F = ID
        return {
          encontrado: true,
          cantidad: datos[i][9] // Columna J = Cantidad
        };
      }
    }
    return { encontrado: false, cantidad: 0 };
  } catch (error) {
    Logger.log(`Error obteniendo stock de ${idInventario}: ${error}`);
    return { encontrado: false, cantidad: 0 };
  }
}

/**
 * Función para estandarizar nombres de productos
 */
function estandarizarNombresProductos() {
  const productosEstandarizados = [
    "ACRILICA SUPERIOR HP",
    "ACRILICA SUPERIOR TIPO B", 
    "SEMIGLOSS PREMIUM",
    "SEMIGLOSS TIPO B",
    "SATINADA",
    "PROYECTO O CONTRACTOR",
    "PROYECTO P/ TECHOS",
    "ECONOMICA",
    "PRIMER ACRILICO",
    "SELLADOR TECHOS HP",
    "SELLADOR TECHOS TIPO B",
    "ESMALTE SINTETICO",
    "ESMALTE INDUSTRIAL",
    "TEXTURIZADAS",
    "EPOXICA",
    "BARNIZ CLEAR INDUSTRIAL",
    "BARNIZ PORT EPOXI CLEAR",
    "DRY WET",
    "ESMALTE INDUSTRIAL ANTICORROSIVO",
    "ESMALTE TRAFICO",
    "SEALER WATER"
  ];
  
  Logger.log("=== PRODUCTOS ESTANDARIZADOS ===");
  productosEstandarizados.forEach((producto, index) => {
    Logger.log(`${index + 1}. ${producto}`);
  });
  
  return productosEstandarizados;
}

/***** 1. CONFIG *****/
const props = PropertiesService.getScriptProperties();
const WASENDER_TOKEN = props.getProperty('WAS_TOKEN');
// Ahora buscamos el grupo desde Properties
const WASENDER_GROUP = props.getProperty('GROUP_GREQ_PEDIDOS') || '120363418347012464@g.us'; 
const WASENDER_URL   = 'https://www.wasenderapi.com/api/send-message';

/***** 2. ENVÍO & LOG *****/
function sendToWA(text, pedId, estado){
  // Usamos WASENDER_GROUP definido arriba
  const resp = UrlFetchApp.fetch(WASENDER_URL,{
    method      :'post',
    contentType :'application/json',
    payload     :JSON.stringify({ to: WASENDER_GROUP, text: String(text) }),
    headers     :{ Authorization: 'Bearer '+WASENDER_TOKEN },
    muteHttpExceptions:true
  });
  
  const ok = (()=>{ try{ return JSON.parse(resp.getContentText()).success;}catch{ return false; }})();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('LOG') || ss.insertSheet('LOG').appendRow(['Timestamp','PED_ID','Estado','HTTP','Success']);
  log.appendRow([new Date(), pedId, estado, resp.getResponseCode(), ok]);
}

/***** 3. UTILIDADES *****/
/**
 * Normalizar texto (tildes → sin, mayúsculas, trim)
 */
function normalize(text) {
  return text
    .toString()
    .normalize('NFD')                // separa diacríticos
    .replace(/[\u0300-\u036f]/g, '') // quita diacríticos
    .toUpperCase()
    .trim()
    .replace(/\s+/g, ' ');           // colapsa espacios
}

/**
 * Obtener código de tipo de producto (robusto con normalización)
 */
function obtenerCodigoTipo(producto) {
  const mapeoTipos = {
    'ACRILICA SUPERIOR HP': '001',
    'ACRILICA SUPERIOR TIPO B': '002',
    'SEMIGLOSS PREMIUM': '003',
    'SEMIGLOSS TIPO B': '004',
    'SATINADA': '005',
    'PROYECTO O CONTRACTOR': '006',
    'PROYECTO P/ TECHOS': '007',
    'ECONOMICA': '008',
    'PRIMER ACRILICO': '009',
    'SELLADOR TECHOS HP': '010',
    'SELLADOR TECHOS TIPO B': '011',
    'ESMALTE SINTETICO': '012',
    'ESMALTE INDUSTRIAL': '013',
    'TEXTURIZADAS': '014',
    'EPOXICA': '015',
    'BARNIZ CLEAR INDUSTRIAL': '016',
    'BARNIZ PORT EPOXI CLEAR': '017',
    'DRY WET': '018',
    'ESMALTE INDUSTRIAL ANTICORROSIVO': '019',
    'ESMALTE TRAFICO': '020',
    'SEALER WATER': '021'
  };

  const productoNorm = normalize(producto);
  for (const [key, value] of Object.entries(mapeoTipos)) {
    if (normalize(key) === productoNorm) {
      return value;
    }
  }
  
  Logger.log(`⚠️ Producto no reconocido: ${producto}`);
  return '000'; // Código por defecto
}

/**
 * Obtener código de envase (robusto con normalización)
 */
function obtenerCodigoEnvase(envase) {
  const mapeoEnvases = {
    'GALON': '1',
    'CUBETA': '5', 
    'CUARTILLO': '2'
  };

  const envaseNorm = normalize(envase);
  for (const [key, value] of Object.entries(mapeoEnvases)) {
    if (normalize(key) === envaseNorm) {
      return value;
    }
  }
  
  Logger.log(`⚠️ Envase no reconocido: ${envase}`);
  return '0'; // Código por defecto
}

/***** 4. DEBUG *****/
function debugToken(){
  Logger.log('TOKEN → '+(WASENDER_TOKEN ? WASENDER_TOKEN.substring(0,10)+'…' : 'null'));
}

/**
 * ───────────────────────────────────────────────────────────
 * FUNCIONES DE CONSTRUCCIÓN DE MENSAJES WHATSAPP POR ESTADO
 * Formato: Guiones simples + Negritas (*texto*)
 * ───────────────────────────────────────────────────────────
 */

/**
 * Construir mensaje PENDIENTE (con mention a Peter)
 */
function construirMensajePendiente(row, tz, pedId, codBarra) {
  const cliente = row[2];     // C
  const producto = row[3];   // D
  const color = row[4];      // E
  const qty = row[5];        // F
  const unidad = row[6];     // G
  const fechaProm = row[8];  // I
  const urgencia = row[9];   // J

  let fechaPromStr = '';
  if (fechaProm instanceof Date && !isNaN(fechaProm)) {
    fechaPromStr = Utilities.formatDate(fechaProm, tz, 'dd-MMM');
  }

  // ═══════════════════════════════════════════════════════════
  // MENTION A PETER
  // ═══════════════════════════════════════════════════════════
  const peterNumero = "18097234205";

  let msg = `🆕 *NUEVO PEDIDO @${peterNumero}*\n`;
  msg += `.............................\n`;
  msg += `*ID:* ${pedId}\n`;
  msg += `*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${qty} ${unidad}\n`;
  msg += `*Código:* ${codBarra}\n`;
  
  if (fechaPromStr) {
    msg += `*Promesa:* ${fechaPromStr}\n`;
  }
  
  if (urgencia === 'Alta') {
    msg += `*Urgencia:* ALTA\n`;
  }
  
  msg += `.............................`;
  
  return msg;
}

/**
 * Construir mensaje EN PRODUCCIÓN
 */
function construirMensajeEnProduccion(row, tz, pedId, codBarra) {
  const cliente = row[2];     // C
  const producto = row[3];   // D
  const color = row[4];      // E
  const qty = row[5];        // F
  const unidad = row[6];     // G
  const urgencia = row[9];   // J

  let msg = `🏭 *EN PRODUCCIÓN*\n`;
  msg += `.............................\n`;
  msg += `*ID:* ${pedId}\n`;
  msg += `*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${qty} ${unidad}\n`;
  msg += `*Código:* ${codBarra}\n`;
  
  if (urgencia === 'Alta') {
    msg += `*Urgencia:* ALTA\n`;
  }
  
  msg += `.............................`;
  
  return msg;
}

/**
 * Construir mensaje LISTO P/ ENVASAR
 */
function construirMensajeLisEnvasar(row, tz, pedId, codBarra) {
  const cliente = row[2];     // C
  const producto = row[3];   // D
  const color = row[4];      // E
  const qty = row[5];        // F
  const unidad = row[6];     // G

  let msg = `✅ *LISTO P/ ENVASAR*\n`;
  msg += `.............................\n`;
  msg += `*ID:* ${pedId}\n`;
  msg += `*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${qty} ${unidad}\n`;
  msg += `*Código:* ${codBarra}\n`;
  msg += `.............................`;
  
  return msg;
}

/**
 * Construir mensaje LISTO P/ DESPACHAR
 */
/**
 * Construir mensaje LISTO P/ DESPACHAR (con métricas flexibles)
 */
function construirMensajeLisDespachar(row, tz, pedId, codBarra, sheet, rowNum) {
  const cliente = row[2];
  const producto = row[3];
  const color = row[4];
  const qty = row[5];
  const unidad = row[6];

  // Obtener métricas si existen
  let tiempoProd = null, tiempoCal = null, tiempoTotal = null;
  
  try {
    tiempoProd = sheet.getRange(rowNum, 20).getValue();  // T
    tiempoCal = sheet.getRange(rowNum, 21).getValue();   // U
    tiempoTotal = sheet.getRange(rowNum, 22).getValue(); // V
  } catch (error) {
    // Si no existen columnas, no pasa nada
  }

  let msg = `📦 *LISTO P/ DESPACHAR*\n`;
  msg += `.............................\n`;
  msg += `*ID:* ${pedId}\n`;
  msg += `*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${qty} ${unidad}\n`;
  msg += `*Código:* ${codBarra}\n`;
  
  // Agregar métricas SOLO si existen Y no son vacías
  const hayMetricas = (tiempoProd && tiempoProd !== '') || 
                      (tiempoCal && tiempoCal !== '') || 
                      (tiempoTotal && tiempoTotal !== '');
  
  if (hayMetricas) {
    msg += `\n⏱️ *TIEMPOS*\n`;
    if (tiempoProd && tiempoProd !== '') {
      msg += `Producción: *${tiempoProd}*\n`;
    }
    if (tiempoCal && tiempoCal !== '') {
      msg += `Calidad: *${tiempoCal}*\n`;
    }
    if (tiempoTotal && tiempoTotal !== '') {
      msg += `Total: *${tiempoTotal}*\n`;
    }
  }
  
  msg += `.............................`;
  
  return msg;
}

/**
 * Construir mensaje DESPACHADO (bifurcado por origen)
 */
function construirMensajeDespachado(row, tz, pedId, codBarra, sheet, rowNum) {
  const cliente = row[2];         // C
  const producto = row[3];       // D
  const color = row[4];          // E
  const qty = row[5];            // F
  const unidad = row[6];         // G
  
  // Obtener datos de Origen e ID_Inventario (si existen)
  const origen = sheet.getRange(rowNum, 16).getValue();       // P = Origen
  const idInventario = sheet.getRange(rowNum, 17).getValue(); // Q = ID_Inventario

  let msg = `✈️ *DESPACHADO*\n`;
  msg += `.............................\n`;
  msg += `*ID:* ${pedId}\n`;
  msg += `*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${qty} ${unidad}\n`;
  msg += `*Código:* ${codBarra}\n`;

  // CASO: Con inventario (origen = INVENTARIO + ID especificado)
  if (origen === 'INVENTARIO' && idInventario) {
    const stockInfo = obtenerStockActual(idInventario);
    
    msg += `\n*Inventario:* Se debitará ${qty} ${unidad} del ID ${idInventario}\n`;
    
    if (stockInfo.encontrado) {
      const stockDespues = stockInfo.cantidad - qty;
      msg += `*Stock:* ${stockInfo.cantidad} → ${stockDespues}\n`;
    }
    
    msg += `*Débito:* 6PM automático\n`;
  }

  msg += `.............................`;
  
  return msg;
}

/**
 * Construir mensaje CANCELADO
 */
function construirMensajeCancelado(row, tz, pedId, codBarra) {
  const cliente = row[2];     // C
  const producto = row[3];   // D
  const color = row[4];      // E
  const qty = row[5];        // F
  const unidad = row[6];     // G

  let msg = `⚠️ *CANCELADO*\n`;
  msg += `.............................\n`;
  msg += `*ID:* ${pedId}\n`;
  msg += `*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${qty} ${unidad}\n`;
  msg += `*Código:* ${codBarra}\n`;
  msg += `.............................`;
  
  return msg;
}

/**
 * ───────────────────────────────────────────────────────────
 * TRIGGER handleEdit(e)
 * ───────────────────────────────────────────────────────────
 */
function handleEdit(e) {
  const SH_NAME   = 'Pedidos';
  const COL_STATE = 11;   // K  (Estado)
  const COL_FECHA = 1;    // A  (Fecha_pedido)
  const COL_PEDID = 12;   // L  (PED_ID)
  const COL_TIME  = 13;   // M  (Último cambio)
  const COL_CODB  = 15;   // O  (Cod_Barra)

  const rng = e.range, sh = rng.getSheet();
  if (sh.getName() !== SH_NAME) return;
  if (![COL_STATE, COL_FECHA].includes(rng.getColumn()) || rng.getRow() === 1) return;

  const row = sh.getRange(rng.getRow(), 1, 1, COL_CODB).getValues()[0];
  const [fechaPedido,, cliente, producto, color, qty, unidad,, fechaProm, urgencia, estado, pedIdCell] = row;
  const tz = Session.getScriptTimeZone();
  const rowNum = rng.getRow();

  /* ---------- 1) Asignar PED_ID si falta ---------- */
  let pedId = pedIdCell;
  if ((!pedId || String(pedId).trim() === '') && fechaPedido && estado) {
    pedId = getNextPedId();
    sh.getRange(rowNum, COL_PEDID).setValue(pedId);
    row[11] = pedId;
  }

  /* ---------- 1b) Generar Cod_Barra si falta ---------- */
  const codCell = sh.getRange(rowNum, COL_CODB);
  const codBarra = codCell.getValue();
  
  if (!codBarra && producto && unidad) {
    const tipoCode = obtenerCodigoTipo(producto);
    const envaseCode = obtenerCodigoEnvase(unidad);
    const prefijo = `911${tipoCode}${envaseCode}`;

    const catalogo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Catalogo_Codigos');
    if (!catalogo) {
      Logger.log("❌ Falta la hoja Catalogo_Codigos"); 
      return;
    }
    
    const lastRow = catalogo.getLastRow();
    const registros = lastRow > 1
      ? catalogo.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      : [];
    const existentes = registros.filter(c => String(c).startsWith(prefijo));
    const secuencial = ('0' + (existentes.length + 1)).slice(-2);
    const nuevoCod = `${prefijo}${secuencial}`;

    codCell.setValue(nuevoCod);
    catalogo.appendRow([nuevoCod, producto, unidad, Number(secuencial), pedId, new Date()]);
    row[COL_CODB - 1] = nuevoCod;
  }

  /* ---------- 2) Timestamp al cambiar Estado ---------- */
  if (rng.getColumn() === COL_STATE) {
    sh.getRange(rowNum, COL_TIME)
      .setValue(new Date())
      .setNumberFormat('dd/MM/yy HH:mm');
  }

  /* ---------- 3) Mensaje WhatsApp - Por Estado Específico ---------- */
  let msg = '';
  let mentionJID = null;  // ← NUEVO
  const codBarraFinal = codCell.getValue() || row[COL_CODB - 1] || '';

  if (estado === 'PENDIENTE') {
    msg = construirMensajePendiente(row, tz, pedId, codBarraFinal);
    mentionJID = "18097234205@s.whatsapp.net";  // ← PETER
  } 
  else if (estado === 'EN PRODUCCIÓN') {
    msg = construirMensajeEnProduccion(row, tz, pedId, codBarraFinal);
  } 
  else if (estado === 'LISTO P/ ENVASAR') {
    msg = construirMensajeLisEnvasar(row, tz, pedId, codBarraFinal);
  } 
  else if (estado === 'LISTO P/ DESPACHAR') {
    msg = construirMensajeLisDespachar(row, tz, pedId, codBarraFinal);
  } 
  else if (estado === 'DESPACHADO') {
    msg = construirMensajeDespachado(row, tz, pedId, codBarraFinal, sh, rowNum);
  } 
  else if (estado === 'CANCELADO') {
    msg = construirMensajeCancelado(row, tz, pedId, codBarraFinal);
  }

  /* ---------- 4) Enviar mensaje (con o sin mention) ---------- */
  if (pedId && estado && msg) {
    if (mentionJID) {
      sendToWAWithMention(msg, mentionJID);  // ← CON MENTION
    } else {
      sendToWA(msg, pedId, estado);          // ← SIN MENTION
    }
  }
}

/**
 * ───────────────────────────────────────────────────────────
 * GENERADOR DE PED_ID
 * ───────────────────────────────────────────────────────────
 */
function getNextPedId() {
  const props = PropertiesService.getDocumentProperties();
  const tz    = Session.getScriptTimeZone();
  const year  = Utilities.formatDate(new Date(), tz, 'yyyy');

  let lastYear = props.getProperty('PED_YEAR');
  let seq      = Number(props.getProperty('NEXT_PED_SEQ')) || 1;

  // Reset anual
  if (lastYear !== year) {
    seq = 1;
    props.setProperty('PED_YEAR', year);
  }

  const id = `PED-${year}-${('000' + seq).slice(-3)}`;
  props.setProperty('NEXT_PED_SEQ', seq + 1);
  props.setProperty('PED_YEAR', year);
  return id;
}

function asignarPedIdSiVacio(sheet, row) {
  const COL_PED = 12;
  const cell    = sheet.getRange(row, COL_PED);
  if (!cell.getValue()) cell.setValue(getNextPedId());
}

/**
 * Resincronizar todos los PED_ID (mantenimiento)
 */
function resincronizarPedId() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const shPed = ss.getSheetByName("Pedidos");
  const shDes = ss.getSheetByName("Despachados");

  if (!shPed || !shDes) {
    SpreadsheetApp.getUi().alert("❌ Falta la hoja 'Pedidos' o 'Despachados'");
    return;
  }

  const dataPed = shPed.getDataRange().getValues();
  const dataDes = shDes.getDataRange().getValues();

  let registros = [];

  for (let i=1; i<dataPed.length; i++) {
    const row = dataPed[i];
    registros.push({
      origen: "Pedidos",
      row: i+1,
      fecha: row[12] || row[0],
      pedId: row[11]
    });
  }

  for (let i=1; i<dataDes.length; i++) {
    const row = dataDes[i];
    registros.push({
      origen: "Despachados",
      row: i+1,
      fecha: row[12] || row[14] || row[0],
      pedId: row[11]
    });
  }

  registros = registros.filter(r => r.fecha && !isNaN(new Date(r.fecha).getTime()));
  registros.sort((a,b) => new Date(a.fecha) - new Date(b.fecha));

  let yearActual = null, seq = 1;
  registros.forEach(reg => {
    const fecha = new Date(reg.fecha);
    const yr = fecha.getFullYear();

    if (yr !== yearActual) {
      yearActual = yr;
      seq = 1;
    }
    reg.nuevoId = `PED-${yr}-${('000' + seq).slice(-3)}`;
    seq++;
  });

  registros.forEach(reg => {
    const sh = (reg.origen === "Pedidos") ? shPed : shDes;
    sh.getRange(reg.row, 12).setValue(reg.nuevoId);
  });

  const ultimo = registros[registros.length-1];
  const ultimoParts = ultimo.nuevoId.split("-");
  const props = PropertiesService.getDocumentProperties();

  props.setProperty("PED_YEAR", ultimoParts[1]);
  const nextSeq = parseInt(ultimoParts[2], 10) + 1;
  props.setProperty("NEXT_PED_SEQ", String(nextSeq));

  Logger.log(`✅ Reasignados ${registros.length} PED_ID.
Último usado: ${ultimo.nuevoId}.
Siguiente será: PED-${props.getProperty("PED_YEAR")}-${('000' + nextSeq).slice(-3)}`);
}

/**
 * ═══════════════════════════════════════════════════════════
 * SISTEMA DE TRACKING TEMPORAL - ADDON PURO
 * Se agrega AL FINAL de Code.gs sin modificar nada existente
 * ═══════════════════════════════════════════════════════════
 */

// Constantes para tracking (columnas R-V)
const TRACK_COL = {
  INICIO_PROD: 18,    // R: Inicio_Produccion (oculta)
  INICIO_CAL: 19,     // S: Inicio_Calidad (oculta)
  TIEMPO_PROD: 20,    // T: Tiempo_Produccion (oculta)
  TIEMPO_CAL: 21,     // U: Tiempo_Calidad (oculta)
  TIEMPO_TOTAL: 22    // V: Tiempo_Total (oculta)
};

/**
 * Crear columnas de tracking y ocultarlas
 * EJECUTAR UNA VEZ manualmente
 */
function configurarColumnasTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pedidos');
  
  if (!sheet) {
    Logger.log('❌ No existe hoja Pedidos');
    return;
  }
  
  // Verificar si ya existen (checar columna R)
  const headerR = sheet.getRange(1, TRACK_COL.INICIO_PROD).getValue();
  
  if (headerR && headerR.includes('Inicio_Produccion')) {
    Logger.log('⚠️ Las columnas de tracking ya existen');
    sheet.hideColumns(TRACK_COL.INICIO_PROD, 5);
    Logger.log('✅ Columnas R-V confirmadas como ocultas');
    return;
  }
  
  // Crear headers
  const headers = [
    'Inicio_Produccion',   // R
    'Inicio_Calidad',      // S
    'Tiempo_Produccion',   // T
    'Tiempo_Calidad',      // U
    'Tiempo_Total'         // V
  ];
  
  sheet.getRange(1, TRACK_COL.INICIO_PROD, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#FFE0B2');
  
  // OCULTAR columnas R-V
  sheet.hideColumns(TRACK_COL.INICIO_PROD, 5);
  
  Logger.log('✅ Columnas de tracking creadas y ocultadas (R-V)');
}

/**
 * ═══════════════════════════════════════════════════════════
 * TRIGGER: trackingOnEdit(e)
 * Captura timestamps al ENTRAR y SALIR de estados específicos
 * 
 * LÓGICA:
 * - EN PRODUCCIÓN → Marca inicio (R)
 * - Sale de EN PRODUCCIÓN → Calcula duración (T)
 * - EN ESPERA DE CALIDAD → Marca inicio (S)
 * - Sale de EN ESPERA DE CALIDAD → Calcula duración (U) y total (V)
 * ═══════════════════════════════════════════════════════════
 */
function trackingOnEdit(e) {
  const SH_NAME = 'Pedidos';
  const COL_STATE = 11; // K (Estado)
  
  const rng = e.range;
  const sh = rng.getSheet();
  
  // Solo procesar si es cambio de estado en hoja Pedidos
  if (sh.getName() !== SH_NAME || rng.getColumn() !== COL_STATE || rng.getRow() === 1) {
    return;
  }
  
  const rowNum = rng.getRow();
  const estadoNuevo = e.value;
  const estadoAnterior = e.oldValue;
  const ahora = new Date();
  
  // ═══════════════════════════════════════════════════════════
  // CASO 1: ENTRA a EN PRODUCCIÓN
  // ═══════════════════════════════════════════════════════════
  if (estadoNuevo === 'EN PRODUCCIÓN') {
    const inicioProd = sh.getRange(rowNum, 18).getValue(); // R
    if (!inicioProd) {
      sh.getRange(rowNum, 18)
        .setValue(ahora)
        .setNumberFormat('dd/MM/yy HH:mm');
      Logger.log(`⏱️ R - Inicio producción: ${ahora}`);
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // CASO 2: SALE de EN PRODUCCIÓN
  // ═══════════════════════════════════════════════════════════
  if (estadoAnterior === 'EN PRODUCCIÓN' && estadoNuevo !== 'EN PRODUCCIÓN') {
    const inicioProd = sh.getRange(rowNum, 18).getValue(); // R
    const tiempoProd = sh.getRange(rowNum, 20).getValue(); // T
    
    // Solo calcular si NO se ha calculado antes y existe timestamp inicial
    if (inicioProd && !tiempoProd) {
      const diffMs = ahora - new Date(inicioProd);
      const minutos = Math.round(diffMs / (1000 * 60));
      const horas = Math.floor(minutos / 60);
      const mins = minutos % 60;
      
      const tiempoTexto = minutos < 60 
        ? `${minutos}min`
        : `${horas}h ${mins}min`;
      
      sh.getRange(rowNum, 20).setValue(tiempoTexto); // T: Tiempo_Produccion
      
      Logger.log(`⏱️ T - Tiempo producción: ${tiempoTexto} (${minutos} min)`);
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // CASO 3: ENTRA a EN ESPERA DE CALIDAD
  // ═══════════════════════════════════════════════════════════
  if (estadoNuevo === 'EN ESPERA DE CALIDAD') {
    const inicioCal = sh.getRange(rowNum, 19).getValue(); // S
    if (!inicioCal) {
      sh.getRange(rowNum, 19)
        .setValue(ahora)
        .setNumberFormat('dd/MM/yy HH:mm');
      Logger.log(`⏱️ S - Inicio calidad: ${ahora}`);
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // CASO 4: SALE de EN ESPERA DE CALIDAD
  // ═══════════════════════════════════════════════════════════
  if (estadoAnterior === 'EN ESPERA DE CALIDAD' && estadoNuevo !== 'EN ESPERA DE CALIDAD') {
    const inicioCal = sh.getRange(rowNum, 19).getValue(); // S
    const tiempoCal = sh.getRange(rowNum, 21).getValue(); // U
    
    // Solo calcular si NO se ha calculado antes y existe timestamp inicial
    if (inicioCal && !tiempoCal) {
      const diffMs = ahora - new Date(inicioCal);
      const minutos = Math.round(diffMs / (1000 * 60));
      const horas = Math.floor(minutos / 60);
      const mins = minutos % 60;
      
      const tiempoTexto = minutos < 60 
        ? `${minutos}min`
        : `${horas}h ${mins}min`;
      
      sh.getRange(rowNum, 21).setValue(tiempoTexto); // U: Tiempo_Calidad
      
      Logger.log(`⏱️ U - Tiempo calidad: ${tiempoTexto} (${minutos} min)`);
      
      // ═══════════════════════════════════════════════════════
      // CALCULAR TIEMPO TOTAL (V) = T + U
      // ═══════════════════════════════════════════════════════
      const tiempoProd = sh.getRange(rowNum, 20).getValue(); // T
      
      if (tiempoProd) {
        const minutosProd = parsearTiempo(tiempoProd);
        const totalMinutos = minutosProd + minutos;
        const totalHoras = Math.floor(totalMinutos / 60);
        const totalMins = totalMinutos % 60;
        
        const totalTexto = totalMinutos < 60 
          ? `${totalMinutos}min`
          : `${totalHoras}h ${totalMins}min`;
        
        sh.getRange(rowNum, 22).setValue(totalTexto); // V: Tiempo_Total
        
        Logger.log(`⏱️ V - Tiempo total: ${totalTexto} (${totalMinutos} min)`);
      }
    }
  }
}

/**
 * Función auxiliar: parsear texto de tiempo a minutos
 * Formatos aceptados: "45min", "2h 30min"
 */
function parsearTiempo(textoTiempo) {
  if (!textoTiempo) return 0;
  
  const texto = textoTiempo.toString().trim();
  let totalMinutos = 0;
  
  // Extraer horas (ej: "2h")
  const matchHoras = texto.match(/(\d+)h/);
  if (matchHoras) {
    totalMinutos += parseInt(matchHoras[1]) * 60;
  }
  
  // Extraer minutos (ej: "30min")
  const matchMinutos = texto.match(/(\d+)min/);
  if (matchMinutos) {
    totalMinutos += parseInt(matchMinutos[1]);
  }
  
  return totalMinutos;
}

/**
 * Instalar trigger de tracking (ejecutar UNA VEZ)
 */
function instalarTriggerTracking() {
  // Eliminar triggers duplicados si existen
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'trackingOnEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger onEdit
  ScriptApp.newTrigger('trackingOnEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
    
  Logger.log('✅ Trigger trackingOnEdit instalado correctamente');
}

function testMentionOficial() {
  const props = PropertiesService.getScriptProperties();
  const WAS_TOKEN = props.getProperty('WAS_TOKEN');
  const GROUP_ID = props.getProperty('GROUP_GREQ_TECNICO'); // Tu grupo @g.us
  
  // DATOS DE MAURO
  const numeroPuro = "18099530116"; 
  const mauroJID = numeroPuro + "@s.whatsapp.net";
  
  // 1. EL TEXTO (Debe incluir el @numeroPuro para que se vea azul)
  const mensajeVisual = `🔔 Prueba Oficial de Documentación\n\nHola @${numeroPuro}, confirma si te llegó la notificación aunque el grupo esté silenciado.`;

  // 2. EL PAYLOAD (Estructura exacta de la documentación)
  const payload = {
    "to": GROUP_ID,
    "text": mensajeVisual,
    "mentions": [mauroJID] // Array de strings
  };

  Logger.log("📤 Enviando Payload:");
  Logger.log(JSON.stringify(payload, null, 2));

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${WAS_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const url = "https://www.wasenderapi.com/api/send-message";
    const response = UrlFetchApp.fetch(url, options);
    
    Logger.log("════════════════════════════════");
    Logger.log("✅ RESPUESTA DEL SERVIDOR:");
    Logger.log("Code: " + response.getResponseCode());
    Logger.log("Body: " + response.getContentText());
    Logger.log("════════════════════════════════");
    
  } catch (e) {
    Logger.log("❌ Error fatal: " + e);
  }
}

/**
 * Enviar mensaje con mention (Corregido para usar el grupo de pedidos)
 */
function sendToWAWithMention(mensaje, mentionJID) {
  const props = PropertiesService.getScriptProperties();
  const WAS_TOKEN = props.getProperty('WAS_TOKEN');
  
  // USAMOS LA VARIABLE GLOBAL QUE YA TIENE EL ID DEL GRUPO DE PEDIDOS
  const GROUP_ID = WASENDER_GROUP; 
  
  if (!WAS_TOKEN || !GROUP_ID) {
    Logger.log("⚠️ Token o Grupo no configurado en Properties");
    return;
  }
  
  const payload = {
    "to": GROUP_ID,
    "text": mensaje,
    "mentions": [mentionJID]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${WAS_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const url = "https://www.wasenderapi.com/api/send-message";
    const response = UrlFetchApp.fetch(url, options);
    
    Logger.log("════════════════════════════════");
    Logger.log("✅ RESPUESTA MENTION (PEDIDOS):");
    Logger.log("Code: " + response.getResponseCode());
    Logger.log("Body: " + response.getContentText());
    Logger.log("════════════════════════════════");
    
  } catch (e) {
    Logger.log("❌ Error fatal en mention: " + e);
  }
}