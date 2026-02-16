/**
 * ═══════════════════════════════════════════════════════════
 * BOT WHATSAPP CPG - V2.0 (FIXED)
 * ═══════════════════════════════════════════════════════════
 */

// ⚙️ CONFIGURACIÓN (Rellena estos datos)
const CONFIG = {
  // ID de la hoja de Google Sheets (lo encuentras en la URL: /d/ESTE_CODIGO/edit)
  SPREADSHEET_ID: "1xSIphg3pD6n4ob70jYmHIOfd8BmsOnBA_K5IG9UCB6w", 
  
  // Nombre exacto de la pestaña
  SHEET_NAME: "Pedidos",
  
  // Token de WaSender (o sácalo de Script Properties si prefieres)
  WAS_TOKEN: PropertiesService.getScriptProperties().getProperty('WAS_TOKEN'),
  
  // ID del grupo para respuestas (o sácalo de Script Properties)
  GROUP_ID: PropertiesService.getScriptProperties().getProperty('GROUP_ID_TEST')
};

// ══════════════════════════════════════════════════════════
// ENDPOINT WEBHOOK (doPost)
// ══════════════════════════════════════════════════════════

function doPost(e) {
  // 1. INTENTO DE REGISTRO MANUAL EN LA HOJA 'DEBUG'
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let shDebug = ss.getSheetByName("Debug");
    if (!shDebug) shDebug = ss.insertSheet("Debug"); // La crea si no existe
    
    // Registramos que llegó ALGO
    shDebug.appendRow([new Date(), "Webhook recibido", JSON.stringify(e)]);
    
    // 2. PARSEO DE DATOS
    let postDataString = "Sin datos";
    if (e && e.postData && e.postData.contents) {
      postDataString = e.postData.contents;
      shDebug.appendRow([new Date(), "Contenido", postDataString]);
    } else {
      shDebug.appendRow([new Date(), "Error", "No venía postData"]);
    }

    // 3. TU LÓGICA ORIGINAL (Resumida)
    const data = JSON.parse(postDataString);
    let mensaje = "";
    
    if (data.event === "messages.received" || data.event === "messages-group.received") {
      mensaje = data.data?.messages?.messageBody || "";
    }
    
    if (mensaje.trim().toLowerCase() === "/pendientes") {
      shDebug.appendRow([new Date(), "Comando detectado", "/pendientes"]);
      cmdPendientesCPG(); // Tu función principal
    }

    return ContentService.createTextOutput(JSON.stringify({success: true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Si falla, intentamos escribir el error en la hoja
    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const shDebug = ss.getSheetByName("Debug");
      shDebug.appendRow([new Date(), "ERROR FATAL", error.toString()]);
    } catch (e) {
      // Si falla escribir en la hoja, no podemos hacer nada más
    }
    return ContentService.createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "Online",
    server_time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
// LÓGICA /pendientes
// ══════════════════════════════════════════════════════════

function cmdPendientesCPG() {
  try {
    console.log("🔍 Ejecutando consulta de pendientes...");
    
    // 1. Abrir Spreadsheet por ID (CRÍTICO PARA WEB APPS)
    if (CONFIG.SPREADSHEET_ID === "PON_AQUI_TU_ID_DE_SPREADSHEET") {
      console.error("⚠️ ID de Spreadsheet no configurado");
      enviarWhatsApp("⚠️ Error: Configura el ID del Sheet en el script.");
      return;
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shPedidos = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!shPedidos) {
      console.error(`❌ Hoja '${CONFIG.SHEET_NAME}' no encontrada`);
      enviarWhatsApp(`⚠️ Error: Hoja '${CONFIG.SHEET_NAME}' no encontrada.`);
      return;
    }
    
    // 2. Leer datos
    const data = shPedidos.getDataRange().getValues();
    const pendientes = [];
    
    // 3. Filtrar (Asumiendo encabezados en fila 0, datos desde fila 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Mapeo de columnas (A=0, B=1, etc.)
      const pedId = row[0];     
      const cliente = row[1];   
      const producto = row[2];  
      const color = row[3];     
      const cantidad = row[4];  
      const unidad = row[5];    
      const estado = row[8];    // Columna I
      
      // Lógica de filtrado
      if (estado && estado !== "DESPACHADO" && estado !== "CANCELADO" && pedId) {
        pendientes.push({
          pedId, cliente, producto, color, cantidad, unidad, estado
        });
      }
    }
    
    console.log(`✅ Pendientes encontrados: ${pendientes.length}`);
    
    // 4. Construir Mensaje
    let msg = "";
    
    if (pendientes.length === 0) {
      msg = "✅ *Todo limpio en CPG.*\nNo hay pedidos pendientes activos.";
    } else {
      msg = `📋 *PEDIDOS ACTIVOS (${pendientes.length})*\n`;
      msg += `.............................\n`;
      
      // Agrupar por estado para mejor lectura
      const porEstado = {};
      pendientes.forEach(p => {
        if (!porEstado[p.estado]) porEstado[p.estado] = [];
        porEstado[p.estado].push(p);
      });
      
      const ordenDeseado = ["PENDIENTE", "EN PRODUCCIÓN", "EN ESPERA DE CALIDAD", "LISTO P/ ENVASAR", "LISTO P/ DESPACHAR"];
      
      // Imprimir en orden
      ordenDeseado.forEach(estado => {
        if (porEstado[estado]) {
          msg += `\n🔸 *${estado}* (${porEstado[estado].length})\n`;
          porEstado[estado].forEach(p => {
            msg += `   • ${p.pedId} | ${p.cliente}\n`;
            msg += `     ${p.producto} ${p.color} (${p.cantidad} ${p.unidad})\n`;
          });
        }
      });
      
      // Imprimir estados que no estén en la lista "ordenDeseado" (por si acaso hay nuevos)
      for (const est in porEstado) {
        if (!ordenDeseado.includes(est)) {
          msg += `\n❓ *${est}* (${porEstado[est].length})\n`;
          porEstado[est].forEach(p => {
             msg += `   • ${p.pedId} | ${p.cliente}\n`;
          });
        }
      }
      
      msg += `\n.............................`;
    }
    
    // 5. Enviar
    enviarWhatsApp(msg);
    
  } catch (error) {
    console.error("❌ Error en lógica: " + error);
    enviarWhatsApp("⚠️ Error interno al consultar pendientes.");
  }
}

// ══════════════════════════════════════════════════════════
// UTILIDAD ENVÍO
// ══════════════════════════════════════════════════════════

function enviarWhatsApp(texto) {
  if (!CONFIG.WAS_TOKEN || !CONFIG.GROUP_ID) {
    console.error("⚠️ Faltan credenciales (Token o GroupID)");
    return;
  }

  const url = "https://www.wasenderapi.com/api/send-message";
  const payload = {
    to: CONFIG.GROUP_ID,
    text: texto
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${CONFIG.WAS_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log(`📱 Envío WhatsApp: ${response.getResponseCode()} - ${response.getContentText()}`);
  } catch (e) {
    console.error("❌ Fallo red WhatsApp: " + e);
  }
}