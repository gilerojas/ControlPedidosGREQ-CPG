/**
 * ═══════════════════════════════════════════════════════════
 * MOVER DESPACHADOS - VERSIÓN CON TRACKING TEMPORAL
 * ═══════════════════════════════════════════════════════════
 * 
 * COLUMNAS PEDIDOS (22 columnas totales):
 * A-Q: Columnas visibles (17 columnas)
 * R-V: Columnas tracking ocultas (5 columnas)
 * 
 * COLUMNAS DESPACHADOS (21 columnas totales):
 * A-O: Datos del pedido (15 columnas - sin Origen ni ID_Inventario)
 * P: Fecha_archivo
 * Q-U: Columnas tracking (5 columnas - trasladadas desde R-V)
 */

// Constantes para mayor claridad y prevención de errores
const COL = {
  FECHA: 0, DIA_PED: 1, CLIENTE: 2, PRODUCTO: 3, COLOR: 4,
  CANTIDAD: 5, UNIDAD: 6, FECHA_PAUT: 7, DIA_PAUT: 8, URGENCIA: 9,
  ESTADO: 10, PED_ID: 11, ULTIMO_CAMBIO: 12, OBSERVACIONES: 13,
  CODIGO_BARRA: 14, ORIGEN: 15, ID_INVENTARIO: 16
};


function moverDespachados() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPedidos = ss.getSheetByName('Pedidos');
  const sheetArch    = ss.getSheetByName('Despachados');
  
  if (!sheetArch) { 
    Logger.log("❌ Falta la hoja 'Despachados'."); 
    return; 
  }

  const data        = sheetPedidos.getDataRange().getValues();
  const headers     = data[0];
  const despachados = [];
  const filasDel    = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const estado       = row[COL.ESTADO];
    const pedId        = row[COL.PED_ID];
    const cantidad     = row[COL.CANTIDAD];
    const origen       = row[COL.ORIGEN];
    const idInventario = row[COL.ID_INVENTARIO];

    // Solo mover si estado es DESPACHADO y hay PED_ID válido
    if (estado === 'DESPACHADO' && pedId && pedId.toString().trim() !== "") {
      
      // CRÍTICO: Debitar inventario AHORA si es de origen INVENTARIO
      if (origen === 'INVENTARIO' && idInventario) {
        debitarInventarioAlDespachar(idInventario, cantidad, pedId);
      }

      // ✅ NUEVA ESTRUCTURA: A-O (datos) + Fecha_archivo + R-V (tracking)
      const nuevaFila = [
        ...row.slice(0, 15),           // A-O: Datos del pedido (sin Origen ni ID_Inventario)
        new Date(),                    // P: Fecha_archivo
        row[TRACK_COL.INICIO_PROD],    // Q: Inicio_Produccion (era R)
        row[TRACK_COL.INICIO_CAL],     // R: Inicio_Calidad (era S)
        row[TRACK_COL.TIEMPO_PROD],    // S: Tiempo_Produccion (era T)
        row[TRACK_COL.TIEMPO_CAL],     // T: Tiempo_Calidad (era U)
        row[TRACK_COL.TIEMPO_TOTAL]    // U: Tiempo_Total (era V)
      ];
      
      despachados.push(nuevaFila);
      filasDel.push(i + 1);
      
      Logger.log(`📦 Preparando para archivar: ${pedId} - ${row[COL.CLIENTE]}`);
    }
  }

  if (despachados.length === 0) {
    Logger.log('ℹ️ No hay pedidos válidos para mover.');
    return;
  }

  // Añadir encabezado si la hoja está vacía
  if (sheetArch.getLastRow() === 0) {
    const newHeaders = [
      ...headers.slice(0, 15),    // A-O: Datos
      'Fecha_archivo',            // P
      'Inicio_Produccion',        // Q
      'Inicio_Calidad',           // R
      'Tiempo_Produccion',        // S
      'Tiempo_Calidad',           // T
      'Tiempo_Total'              // U
    ];
    sheetArch.appendRow(newHeaders);
    
    // Formatear headers
    sheetArch.getRange(1, 1, 1, newHeaders.length)
      .setFontWeight('bold')
      .setBackground('#E8F5E8');
    
    // Ocultar columnas de tracking en Despachados también (Q-U)
    sheetArch.hideColumns(17, 5); // Columnas Q-U (17-21)
    
    Logger.log('📋 Encabezados creados en Despachados con columnas tracking ocultas');
  }

  // Pegar filas en Despachados
  sheetArch
    .getRange(sheetArch.getLastRow() + 1, 1, despachados.length, despachados[0].length)
    .setValues(despachados);

  Logger.log(`✅ ${despachados.length} pedidos copiados a Despachados (con métricas de tiempo)`);

  // Borrar filas de Pedidos (de abajo hacia arriba)
  for (let i = filasDel.length - 1; i >= 0; i--) {
    sheetPedidos.deleteRow(filasDel[i]);
  }

  Logger.log(`🗑️ ${filasDel.length} filas eliminadas de Pedidos`);
  Logger.log(`✅ PROCESO COMPLETO: ${despachados.length} pedidos archivados. Inventario debitado.`);
}

/**
 * Debitar inventario al momento real del despacho (6PM)
 */
function debitarInventarioAlDespachar(idInventario, cantidad, pedId) {
  const inventarioSS = SpreadsheetApp.openById('1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw');
  const inventarioSheet = inventarioSS.getSheetByName('Inventario');
  
  const datos = inventarioSheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][5] === idInventario) { // Columna F = ID
      const cantidadAnterior = datos[i][9]; // Columna J = Cantidad
      const nuevaCantidad = cantidadAnterior - cantidad;
      
      // Actualizar inventario
      inventarioSheet.getRange(i + 1, 10).setValue(nuevaCantidad);
      
      // Obtener datos para el LOG
      const producto = datos[i][2]; // C = Tipo
      const color = datos[i][3];    // D = Color
      const envase = datos[i][8];   // I = Envase
      
      // Crear LOG del débito
      crearLogDebitos(pedId, 'N/A', producto, color, cantidad, envase, idInventario, cantidadAnterior, nuevaCantidad);
      
      Logger.log(`💰 DEBITADO: ${pedId} - ID ${idInventario}: ${cantidad} und. Restante: ${nuevaCantidad}`);
      return;
    }
  }
  Logger.log(`⚠️ ID ${idInventario} no encontrado para debitar (${pedId})`);
}

/**
 * Crear LOG de débitos de inventario
 */
function crearLogDebitos(pedId, cliente, producto, color, cantidad, envase, idInventario, cantidadAnterior, cantidadNueva) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('LOG_Debitos');
  
  // Crear hoja si no existe
  if (!logSheet) {
    logSheet = ss.insertSheet('LOG_Debitos');
    const headers = [
      'Timestamp', 'PED_ID', 'Cliente', 'Producto', 'Color', 
      'Cantidad_Pedida', 'Envase', 'ID_Inventario', 
      'Stock_Anterior', 'Stock_Nuevo', 'Diferencia'
    ];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4CAF50')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  
  const registro = [
    new Date(),
    pedId,
    cliente,
    producto,
    color,
    cantidad,
    envase,
    idInventario,
    cantidadAnterior,
    cantidadNueva,
    cantidadAnterior - cantidadNueva
  ];
  
  logSheet.appendRow(registro);
  Logger.log(`📝 LOG: Débito registrado para ${pedId} - ID ${idInventario}`);
}

/**
 * Crea un disparador para ejecutar moverDespachados() todos los días a las 6:00 PM
 */
function programarMoverDespachados() {
  // Eliminar triggers existentes primero
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'moverDespachados') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger
  ScriptApp.newTrigger('moverDespachados')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
    
  Logger.log('✅ Trigger moverDespachados programado para 6:00 PM diario');
}

/**
 * Función de testing para verificar estructura
 */
function testMoverDespachados() {
  Logger.log('🧪 INICIANDO TEST: moverDespachados()');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPedidos = ss.getSheetByName('Pedidos');
  
  if (!sheetPedidos) {
    Logger.log('❌ No existe hoja Pedidos');
    return;
  }
  
  const data = sheetPedidos.getDataRange().getValues();
  let despachadosEncontrados = 0;
  let conTracking = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.ESTADO] === 'DESPACHADO') {
      despachadosEncontrados++;
      
      // Verificar si tiene datos de tracking
      if (data[i][TRACK_COL.TIEMPO_TOTAL]) {
        conTracking++;
      }
    }
  }
  
  Logger.log(`📊 Pedidos DESPACHADO encontrados: ${despachadosEncontrados}`);
  Logger.log(`⏱️ Pedidos con métricas de tiempo: ${conTracking}`);
  
  if (despachadosEncontrados === 0) {
    Logger.log('ℹ️ No hay pedidos DESPACHADO para testing');
    Logger.log('💡 Marca un pedido como DESPACHADO para probar');
    return;
  }
  
  // Ejecutar moverDespachados
  moverDespachados();
  
  Logger.log('✅ TEST COMPLETADO - Revisar hoja Despachados');
  Logger.log('🔍 Verificar que columnas Q-U estén ocultas en Despachados');
}

/**
 * Función para verificar columnas de tracking en Despachados
 */
function verificarTrackingEnDespachados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const despachados = ss.getSheetByName('Despachados');
  
  if (!despachados || despachados.getLastRow() < 2) {
    Logger.log('⚠️ No hay datos en Despachados para verificar');
    return;
  }
  
  const headers = despachados.getRange(1, 1, 1, despachados.getLastColumn()).getValues()[0];
  
  Logger.log('📋 ESTRUCTURA DESPACHADOS:');
  headers.forEach((header, index) => {
    const columna = String.fromCharCode(65 + index);
    Logger.log(`${columna}: ${header}`);
  });
  
  // Verificar datos de tracking en primera fila
  const primeraFila = despachados.getRange(2, 1, 1, despachados.getLastColumn()).getValues()[0];
  
  Logger.log('\n⏱️ DATOS DE TRACKING (primera fila):');
  Logger.log(`Q (Inicio_Produccion): ${primeraFila[16]}`);
  Logger.log(`R (Inicio_Calidad): ${primeraFila[17]}`);
  Logger.log(`S (Tiempo_Produccion): ${primeraFila[18]}`);
  Logger.log(`T (Tiempo_Calidad): ${primeraFila[19]}`);
  Logger.log(`U (Tiempo_Total): ${primeraFila[20]}`);
}

// ═══════════════════════════════════════════════════════════
// MOVER TERMINADOS - Archivado automático de producciones
// (Mantener código existente sin cambios)
// ═══════════════════════════════════════════════════════════

function moverTerminados() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const produccionSheet = spreadsheet.getSheetByName('Produccion');
  
  if (!produccionSheet) {
    Logger.log('⚠️ No existe hoja Produccion');
    return;
  }
  
  const datos = produccionSheet.getDataRange().getValues();
  const encabezados = datos[0];
  
  const filasTerminadas = [];
  const filasAMover = [];
  
  // Constantes para hoja Producción (asegurar que existan)
  const PROD_COL = {
    PED_ID: 1, CLIENTE: 2, PRODUCTO: 3, COLOR: 4, CANTIDAD: 5, ENVASE: 6,
    MAQUINA: 7, OPERARIO: 8, ESTADO_PROD: 9, FECHA_PROD: 10, HORA_INICIO: 11,
    HORA_FIN: 12, CANT_PRODUCIDA: 13, UNIDAD_ENVASE: 14, REPROCESO: 15, 
    ID_CONSUMIDO: 16, TIEMPO_TOTAL: 17, ID_PRODUCCION: 18
  };
  
  for (let i = 1; i < datos.length; i++) {
    const estadoProd = datos[i][PROD_COL.ESTADO_PROD - 1];
    
    if (estadoProd === "TERMINADO") {
      filasTerminadas.push(i + 1);
      const filaConFecha = [...datos[i], new Date()];
      filasAMover.push(filaConFecha);
    }
  }
  
  if (filasTerminadas.length === 0) {
    Logger.log('✅ No hay producciones TERMINADO para mover');
    return;
  }
  
  Logger.log(`📦 Encontradas ${filasTerminadas.length} producciones TERMINADO para archivar`);
  
  let completadasSheet = spreadsheet.getSheetByName('Producciones_Completadas');
  if (!completadasSheet) {
    completadasSheet = crearHojaProducciones();
  }
  
  if (filasAMover.length > 0) {
    const rangoDestino = completadasSheet.getRange(
      completadasSheet.getLastRow() + 1, 1, 
      filasAMover.length, filasAMover[0].length
    );
    rangoDestino.setValues(filasAMover);
    Logger.log(`📁 ${filasAMover.length} producciones movidas a Producciones_Completadas`);
  }
  
  filasTerminadas.reverse().forEach(fila => {
    produccionSheet.deleteRow(fila);
  });
  
  Logger.log(`🗑️ ${filasTerminadas.length} filas eliminadas de Produccion`);
  
  enviarNotificacionArchivado(filasTerminadas.length, filasAMover);
  registrarMovimientoProduccion(filasTerminadas.length, filasAMover);
  
  Logger.log(`✅ Archivado completado: ${filasTerminadas.length} producciones procesadas`);
}

function crearHojaProducciones() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const completadasSheet = spreadsheet.insertSheet('Producciones_Completadas');
  
  const encabezados = [
    'PED_ID', 'Cliente', 'Producto', 'Color', 'Cantidad', 'Envase',
    'Máquina', 'Operario', 'Estado_Prod', 'Fecha_PROD', 'Hora_Inicio', 
    'Hora_Fin', 'Cant_Producida', 'Unidad_Envase', 'REPROCESO', 
    'ID_CONSUMIDO', 'Tiempo_Total', 'ID_PRODUCCION', 'Fecha_archivo'
  ];
  
  completadasSheet.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  completadasSheet.getRange(1, 1, 1, encabezados.length).setFontWeight('bold');
  completadasSheet.setFrozenRows(1);
  
  completadasSheet.getRange('J:J').setNumberFormat('dd/MM/yyyy');
  completadasSheet.getRange('K:L').setNumberFormat('HH:mm');
  completadasSheet.getRange('S:S').setNumberFormat('dd/MM/yyyy HH:mm');
  completadasSheet.getRange(1, 1, 1, encabezados.length).setBackground('#E8F4FD');
  
  Logger.log('📋 Hoja Producciones_Completadas creada automáticamente');
  return completadasSheet;
}

function enviarNotificacionArchivado(cantidadMovida, datosMovidos) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('WAS_TOKEN');
  const grupo = props.getProperty('GROUP_GREQ_MAIN');
  
  if (!token || !grupo) {
    Logger.log('⚠️ Faltan credenciales WhatsApp para notificación de archivado');
    return;
  }
  
  const ahora = new Date();
  const fechaTexto = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  
  const resumenProducciones = datosMovidos.slice(0, 3).map(fila => {
    const pedId = fila[0];
    const producto = fila[2];
    const color = fila[3];
    const cantProducida = fila[12];
    const unidadEnvase = fila[13];
    const idProduccion = fila[17];
    
    return `• ${pedId}: ${cantProducida} ${unidadEnvase} ${producto} ${color} (${idProduccion})`;
  }).join('\n');
  
  let mensaje = `🗂️ *ARCHIVADO AUTOMÁTICO - PRODUCCIÓN*\n`;
  mensaje += `📅 ${fechaTexto}\n\n`;
  mensaje += `📊 *Producciones completadas archivadas:* ${cantidadMovida}\n\n`;
  
  if (cantidadMovida <= 3) {
    mensaje += `📋 *Detalles:*\n${resumenProducciones}`;
  } else {
    mensaje += `📋 *Primeras 3:*\n${resumenProducciones}\n`;
    mensaje += `... y ${cantidadMovida - 3} más`;
  }
  
  mensaje += `\n\n📍 *Ubicación:* Hoja "Producciones_Completadas"`;
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify({ to: grupo, text: String(mensaje) }),
    muteHttpExceptions: true
  };
  
  try {
    const resp = UrlFetchApp.fetch('https://www.wasenderapi.com/api/send-message', options);
    Logger.log(`📱 Notificación archivado enviada: ${resp.getResponseCode()}`);
  } catch (error) {
    Logger.log(`❌ Error enviando notificación archivado: ${error}`);
  }
}

function registrarMovimientoProduccion(cantidad, datosMovidos) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('LOG_Movimientos_Produccion');
    
    if (!logSheet) {
      logSheet = crearLogMovimientos();
    }
    
    const timestamp = new Date();
    const pedIds = datosMovidos.map(fila => fila[0]).join(', ');
    
    logSheet.appendRow([
      timestamp,
      'ARCHIVADO_AUTOMATICO',
      cantidad,
      pedIds,
      'Sistema automático 6PM',
      `${cantidad} producciones TERMINADO archivadas`
    ]);
    
    Logger.log(`📝 Movimiento registrado en LOG: ${cantidad} producciones archivadas`);
    
  } catch (error) {
    Logger.log(`⚠️ Error registrando movimiento en LOG: ${error}`);
  }
}

function crearLogMovimientos() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.insertSheet('LOG_Movimientos_Produccion');
  
  const headers = [
    'Timestamp', 'Tipo_Movimiento', 'Cantidad_Movida', 
    'PED_IDs_Afectados', 'Ejecutado_Por', 'Detalles'
  ];
  
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  logSheet.getRange(1, 1, 1, headers.length).setBackground('#FFF2CC');
  logSheet.setFrozenRows(1);
  logSheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  
  Logger.log('📋 Hoja LOG_Movimientos_Produccion creada');
  return logSheet;
}

function crearTriggerMoverTerminados() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'moverTerminados') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('moverTerminados')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
    
  Logger.log('✅ Trigger moverTerminados() programado para 6:00 PM diario');
}

function testMoverTerminadosCompleto() {
  Logger.log('🧪 INICIANDO TEST: moverTerminados()');
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const produccionSheet = spreadsheet.getSheetByName('Produccion');
  
  if (!produccionSheet) {
    Logger.log('❌ No existe hoja Produccion');
    return;
  }
  
  const PROD_COL = {
    ESTADO_PROD: 9
  };
  
  const datos = produccionSheet.getDataRange().getValues();
  let terminadosEncontrados = 0;
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][PROD_COL.ESTADO_PROD - 1] === "TERMINADO") {
      terminadosEncontrados++;
    }
  }
  
  Logger.log(`📊 Producciones TERMINADO encontradas: ${terminadosEncontrados}`);
  
  if (terminadosEncontrados === 0) {
    Logger.log('ℹ️ No hay producciones TERMINADO para testing');
    Logger.log('💡 Crea una producción y márcala como TERMINADO para probar');
    return;
  }
  
  moverTerminados();
  Logger.log('✅ TEST COMPLETADO - Revisar hoja Producciones_Completadas');
}