/**
 * ═══════════════════════════════════════════════════════════════
 * QCBRIDGE v2.3 - SISTEMA CPG → CCG
 * Flujo: Ventas → Calidad (con validación de QC)
 * 
 * UPDATES v2.3:
 * - Simple trigger onEdit() (multi-usuario)
 * - WhatsApp con mention a Mauro
 * - Formato mejorado (estilo construirMensaje)
 * - Trigger duplicado eliminado
 * ═══════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ══════════════════════════════════════════════════════════════

const CONFIG_CPG = {
  ID_ARCHIVO_CCG: "1knF-ghqVFur9GCgIhaRX9ieWSAB6HaBOhOUA7-CMgOI",
  NOMBRE_HOJA_PEDIDOS: "Pedidos",
  NOMBRE_HOJA_CCG: "CCG",
  NOMBRE_HOJA_METRICAS_CCG: "Metricas_QC",
  MAURO_JID: "18099530116@s.whatsapp.net",
  
  COL: {
    FECHA: 1, DIA_PEDIDO: 2, CLIENTE: 3, PRODUCTO: 4, COLOR: 5,
    CANTIDAD: 6, UNIDAD: 7, FECHA_PAUTADA: 8, DIA_PAUTADO: 9,
    URGENCIA: 10, ESTADO: 11, PED_ID: 12, ULTIMO_CAMBIO: 13,
    OBSERVACIONES: 14, CODIGO_BARRA: 15
  },
  
  COL_CCG: {
    PED_ID: 1, CLIENTE: 2, PRODUCTO: 3, COLOR: 4, CANTIDAD: 5,
    UNIDAD: 6, GLS_TOTALES: 7, ORIGEN: 8, GLS_REALES: 9,
    VISCOSIDAD: 10, PH: 11, DENSIDAD: 12, ESTADO_QC: 13,
    FECHA: 14, RESPONSABLE: 15
  }
};

// ══════════════════════════════════════════════════════════════
// TRIGGER PRINCIPAL (instalable)
// ══════════════════════════════════════════════════════════════

/**
 * ══════════════════════════════════════════════════════════════
 * QCBRIDGE - FUNCIÓN DE CONTROL DE EDICIÓN (ROBUSTA)
 * ══════════════════════════════════════════════════════════════
 */
function onEditQCBridge(e) {
  // 1. Verificación de seguridad inicial
  if (!e || !e.range) return;
  
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // 2. Validar que la edición sea en la hoja 'Pedidos' y no sea el encabezado
  if (sheet.getName() !== CONFIG_CPG.NOMBRE_HOJA_PEDIDOS || row < 2) return;

  // 3. Filtro de columna: Solo procesar si se editó la columna de ESTADO (Col 11)
  if (col !== CONFIG_CPG.COL.ESTADO) return;

  // 4. LECTURA DIRECTA: No confiamos en e.value para evitar fallos en Dropdowns o Pegados
  const estadoNuevo = range.getValue(); 
  const pedId = sheet.getRange(row, CONFIG_CPG.COL.PED_ID).getValue();

  // Log de auditoría (Visible en la consola de Google Apps Script)
  console.log(`[QCBridge] Edit detectado en Fila ${row}. Pedido: ${pedId}, Estado: ${estadoNuevo}`);

  // 5. Validaciones de salida
  if (!pedId || !estadoNuevo) {
    console.warn(`[QCBridge] Abortado: Falta ID de pedido o el estado está vacío.`);
    return;
  }

  // ══════════════════════════════════════════════════════════
  // FLUJO A: MANDAR A CALIDAD (CCG)
  // ══════════════════════════════════════════════════════════
  if (estadoNuevo === "PENDIENTE") {
    // Reducimos el tiempo de espera a 2 segundos para dar tiempo a otros triggers 
    // sin que Google mate la ejecución por exceso de tiempo.
    Utilities.sleep(2000); 
    
    console.log(`[QCBridge] Iniciando envío a CCG para ID: ${pedId}`);
    enviarACCG(pedId, sheet, row);
  }

  // ══════════════════════════════════════════════════════════
  // FLUJO B: VALIDAR ANTES DE DESPACHAR
  // ══════════════════════════════════════════════════════════
  if (estadoNuevo === "DESPACHADO") {
    // e.oldValue suele ser confiable aquí para revertir si QC no ha aprobado
    const estadoAnt = e.oldValue || "LISTO P/ DESPACHAR";
    
    console.log(`[QCBridge] Validando aprobación de QC para ID: ${pedId}`);
    validarAprobacionRemota(e, pedId, estadoAnt, sheet, row);
  }
}

// ══════════════════════════════════════════════════════════════
// ENVIAR A CCG
// ══════════════════════════════════════════════════════════════

function enviarACCG(pedId, sheet, row) {
  try {
    const ssCCG = SpreadsheetApp.openById(CONFIG_CPG.ID_ARCHIVO_CCG);
    const shCCG = ssCCG.getSheetByName(CONFIG_CPG.NOMBRE_HOJA_CCG);
    const shMetricas = ssCCG.getSheetByName(CONFIG_CPG.NOMBRE_HOJA_METRICAS_CCG);
    
    if (!shCCG) {
      SpreadsheetApp.getActive().toast("❌ Hoja CCG no encontrada", "Error");
      return;
    }
    
    const ids = shCCG.getRange("A:A").getValues().flat();
    if (ids.includes(pedId)) {
      Logger.log(`⚠️ ${pedId} ya existe en CCG`);
      return;
    }
    
    const cliente = sheet.getRange(row, CONFIG_CPG.COL.CLIENTE).getValue();
    const producto = sheet.getRange(row, CONFIG_CPG.COL.PRODUCTO).getValue();
    const color = sheet.getRange(row, CONFIG_CPG.COL.COLOR).getValue();
    const cantidad = sheet.getRange(row, CONFIG_CPG.COL.CANTIDAD).getValue();
    const unidad = sheet.getRange(row, CONFIG_CPG.COL.UNIDAD).getValue();
    const glsTotales = calcularGalones(cantidad, unidad);
    
    const rowDataCCG = [
      pedId, cliente, producto, color, cantidad, unidad, glsTotales,
      "PENDIENTE", "", "", "", "", "PENDIENTE", "", ""
    ];
    
    const newRow = shCCG.getLastRow() + 1;
    shCCG.appendRow(rowDataCCG);
    
    const dvOrigen = SpreadsheetApp.newDataValidation()
      .requireValueInList(["PENDIENTE", "PRODUCCION", "STOCK", "MIXTO"])
      .setAllowInvalid(false).build();
    
    const dvEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(["PENDIENTE", "APROBADO"])
      .setAllowInvalid(false).build();
    
    shCCG.getRange(newRow, CONFIG_CPG.COL_CCG.ORIGEN).setDataValidation(dvOrigen);
    shCCG.getRange(newRow, CONFIG_CPG.COL_CCG.ESTADO_QC).setDataValidation(dvEstado);
    
    if (shMetricas) {
      shMetricas.appendRow([pedId, cliente, producto, color, "", new Date(), "", "", "", "", "", "", "", "", "", "", "", ""]);
    }
    
    notificarPedidoEnviadoCCG(pedId, sheet, row);
    
    SpreadsheetApp.getActive().toast("✅ Enviado a Calidad", "GREQ");
    Logger.log(`✅ ${pedId} → CCG`);
    
  } catch (err) {
    Logger.log(`❌ Error: ${err}`);
    SpreadsheetApp.getActive().toast("⚠️ Error enviando a Calidad", "Error");
  }
}

// ══════════════════════════════════════════════════════════════
// VALIDAR QC
// ══════════════════════════════════════════════════════════════

function validarAprobacionRemota(e, pedId, estadoAnt, sheet, row) {
  try {
    const ssCCG = SpreadsheetApp.openById(CONFIG_CPG.ID_ARCHIVO_CCG);
    const shCCG = ssCCG.getSheetByName(CONFIG_CPG.NOMBRE_HOJA_CCG);
    const data = shCCG.getDataRange().getValues();
    const registro = data.find(f => f[0] === pedId);
    
    if (!registro) {
      Logger.log(`ℹ️ ${pedId} - Pedido anterior`);
      return;
    }
    
    const estadoQC = registro[CONFIG_CPG.COL_CCG.ESTADO_QC - 1];
    
    if (estadoQC !== "APROBADO") {
      e.range.setValue(estadoAnt || "LISTO P/ DESPACHAR");
      SpreadsheetApp.getActive().toast(`⛔ ${pedId} no aprobado por QC`, "BLOQUEADO", 10);
      Logger.log(`⛔ ${pedId} bloqueado`);
      return;
    }
    
    SpreadsheetApp.getActive().toast("✅ Despacho Autorizado", "GREQ");
    
  } catch (err) {
    e.range.setValue(estadoAnt || "LISTO P/ DESPACHAR");
    Logger.log(`❌ Error: ${err}`);
  }
}

// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════

function calcularGalones(cantidad, unidad) {
  if (!cantidad || !unidad) return 0;
  const u = unidad.toString().trim().toUpperCase();
  if (u.includes("CUB")) return cantidad * 5;
  if (u.includes("CUART") || u.includes("1/4")) return cantidad * 0.25;
  return cantidad * 1;
}

function notificarPedidoEnviadoCCG(pedId, sheet, row) {
  const tz = 'America/Santo_Domingo';
  const cliente = sheet.getRange(row, CONFIG_CPG.COL.CLIENTE).getValue();
  const producto = sheet.getRange(row, CONFIG_CPG.COL.PRODUCTO).getValue();
  const color = sheet.getRange(row, CONFIG_CPG.COL.COLOR).getValue();
  const cantidad = sheet.getRange(row, CONFIG_CPG.COL.CANTIDAD).getValue();
  const unidad = sheet.getRange(row, CONFIG_CPG.COL.UNIDAD).getValue();
  const codBarra = sheet.getRange(row, CONFIG_CPG.COL.CODIGO_BARRA).getValue();
  const fechaProm = sheet.getRange(row, CONFIG_CPG.COL.FECHA_PAUTADA).getValue();
  const urgencia = sheet.getRange(row, CONFIG_CPG.COL.URGENCIA).getValue();

  let fechaPromStr = '';
  if (fechaProm instanceof Date && !isNaN(fechaProm)) {
    fechaPromStr = Utilities.formatDate(fechaProm, tz, 'dd-MMM');
  }

  // ═══════════════════════════════════════════════════════════
  // FIX: Incluir @numero en el texto
  // ═══════════════════════════════════════════════════════════
  const mauroNumero = "18099530116";
  
  let msg = `🔔 *NUEVO PEDIDO → QC*\n.............................\n`;
  msg += `*ID:* ${pedId}\n*Cliente:* ${cliente}\n`;
  msg += `*Producto:* ${producto} ${color}\n`;
  msg += `*Cantidad:* ${cantidad} ${unidad}\n*Código:* ${codBarra}\n`;
  
  if (fechaPromStr) msg += `*Promesa:* ${fechaPromStr}\n`;
  if (urgencia === 'Alta') msg += `*Urgencia:* ALTA\n`;
  
  // ← AQUÍ ESTÁ EL FIX
  msg += `\n⏱️ *ACCIÓN REQUERIDA @${mauroNumero}:*\n`;
  msg += `Calidad → Llenar *ORIGEN* en CCG\n`;
  msg += `• ¿Salió de STOCK?\n• ¿Viene de PRODUCCIÓN?\n.............................`;
  
  enviarWhatsAppConMention(msg, mauroNumero + "@s.whatsapp.net");
}

function enviarWhatsAppConMention(mensaje, mentionJID) {
  const props = PropertiesService.getScriptProperties();
  const WAS_TOKEN = props.getProperty('WAS_TOKEN');
  const GROUP_ID = props.getProperty('GROUP_GREQ_TECNICO');
  
  if (!WAS_TOKEN || !GROUP_ID) return;
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${WAS_TOKEN}` },
    payload: JSON.stringify({
      to: GROUP_ID,
      text: mensaje,
      mentions: [mentionJID]
    }),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch("https://www.wasenderapi.com/api/send-message", options);
    Logger.log(`📱 WhatsApp: ${response.getResponseCode()}`);
  } catch (error) {
    Logger.log(`❌ WhatsApp: ${error}`);
  }
}



