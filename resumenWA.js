// ───────────────────────────────────────────────────────────
// resumenWA.gs – Envío de resumen quincenal y mensual de despachos (CORREGIDO)
// ───────────────────────────────────────────────────────────

/**
 * Envía un mensaje de resumen a WhatsApp usando la API de WASender
 * @param {string} text - Texto del mensaje a enviar
 * @param {string} resumenTipo - Tipo de resumen ('RESUMEN_MENSUAL' o 'RESUMEN_QUINCENAL')
 */
function sendResumenToWA(text, resumenTipo = 'RESUMEN_MENSUAL') {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('WAS_TOKEN');
  const group = props.getProperty('GROUP_GREQ_MAIN');

  const resp = UrlFetchApp.fetch('https://www.wasenderapi.com/api/send-message', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ to: group, text: String(text) }),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const ok = (() => { try { return JSON.parse(resp.getContentText()).success; } catch { return false; } })();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName('LOG') || ss.insertSheet('LOG').appendRow(['Timestamp', 'Tipo', 'Estado', 'HTTP', 'Success']);
  log.appendRow([new Date(), resumenTipo, "ENVIADO", resp.getResponseCode(), ok]);
}

/**
 * Genera el resumen quincenal de despachos
 * @returns {string} Mensaje de resumen formateado
 */
function generarQuincenalResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Despachados');
  if (!sh) return "❌ Falta hoja Despachados";

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return "No hay datos en Despachados";

  const data = sh.getRange(2, 1, lastRow-1, sh.getLastColumn()).getValues();

  // === Determinar quincena PASADA (siempre la inmediatamente anterior) ===
  const hoy = new Date();
  const year = hoy.getFullYear();
  const month = hoy.getMonth(); // 0 = Enero
  const day = hoy.getDate();

  let inicio, fin;
  
  // Si hoy es entre 1-15 → buscar segunda quincena del mes anterior (16-fin)
  if (day <= 15) {
    const prevMonth = month - 1;
    const prevYear = prevMonth < 0 ? year - 1 : year;
    const realMonth = (prevMonth + 12) % 12;
    inicio = new Date(prevYear, realMonth, 16);
    fin    = new Date(prevYear, realMonth + 1, 0); // último día del mes anterior
  } else {
    // Si hoy es 16+ → buscar primera quincena del mes actual (1-15)
    inicio = new Date(year, month, 1);
    fin    = new Date(year, month, 15);
  }

  // Filtrar registros en rango (solo con PED_ID válido)
  const enRango = data.filter(r => {
    const fecha = new Date(r[15]); // col 16 = Fecha_archivo (índice 15)
    const pedId = r[11];           // col 12 = PED_ID
    return (
      fecha >= inicio && fecha <= fin &&
      pedId && pedId.toString().trim() !== ""
    );
  });

  if (enRango.length === 0) return "No hay despachos en la quincena pasada.";

  // === Estadísticas ===
  const totalPedidos = enRango.length;
  let totalGal = 0, totalCub = 0;
  const productos = {};

  enRango.forEach(r => {
    const producto = r[3];    // Producto (col D)
    const cantidad = Number(r[5]) || 0;
    const unidad = (r[6] || "").toString().toLowerCase(); // Unidad (col G)

    if (unidad.includes("gal")) totalGal += cantidad;
    if (unidad.includes("cub")) totalCub += cantidad;

    productos[producto] = (productos[producto] || 0) + cantidad;
  });

  // Top 3 productos
  const topProductos = Object.entries(productos).sort((a,b) => b[1] - a[1]).slice(0, 3);

  // === Construir mensaje ===
  const fechaInicio = Utilities.formatDate(inicio, "America/Santo_Domingo", "dd/MM");
  const fechaFin = Utilities.formatDate(fin, "America/Santo_Domingo", "dd/MM/yyyy");
  
  let msg = `📦 Resumen quincenal de despachos (${fechaInicio}–${fechaFin})\n\n`;
  msg += `• Pedidos despachados: ${totalPedidos}\n`;
  msg += `• Volumen total: ${totalCub} Cubetas / ${totalGal} Galones\n`;
  msg += `• Productos más despachados:\n`;
  topProductos.forEach(p => { msg += `   - ${p[0]} – ${p[1]}\n`; });
  msg += `\n✅ Todos los pedidos están registrados en hoja Despachados.`;

  return msg;
}

/**
 * Envía automáticamente el resumen quincenal (para trigger)
 */
function sendQuincenalResumen() {
  const msg = generarQuincenalResumen();
  if (msg.startsWith("❌") || msg.startsWith("No hay")) {
    Logger.log(msg); 
    return;
  }
  Logger.log("Mensaje generado:\n" + msg);
  sendResumenToWA(msg, 'RESUMEN_QUINCENAL'); // ✅ CORREGIDO
}

/**
 * Genera vista previa del resumen quincenal (para botón manual)
 * @returns {string} Mensaje de resumen
 */
function debugQuincenalResumen() {
  const msg = generarQuincenalResumen();
  
  Logger.log("Vista previa:\n" + msg);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Preview_Resumen");
  if (!sh) sh = ss.insertSheet("Preview_Resumen");
  sh.clear();
  sh.getRange(1,1).setValue("Vista previa del resumen quincenal");
  sh.getRange(3,1).setValue(msg);

  return msg;
}

/**
 * Crea triggers automáticos para envío quincenal (días 1 y 16 de cada mes)
 */
function crearTriggerQuincenalExacto() {
  // Elimina triggers antiguos para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendQuincenalResumen") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Trigger el día 1 de cada mes a las 9:00 AM
  ScriptApp.newTrigger("sendQuincenalResumen")
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  // Trigger el día 16 de cada mes a las 9:00 AM
  ScriptApp.newTrigger("sendQuincenalResumen")
    .timeBased()
    .onMonthDay(16)
    .atHour(9)
    .create();

  Logger.log("✅ Triggers creados: sendQuincenalResumen() el día 1 y 16 de cada mes a las 9:00 AM");
}

/**
 * Genera el resumen mensual de despachos del mes COMPLETO anterior
 * @returns {string} Mensaje de resumen formateado
 */
function generarMensualResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Despachados');
  if (!sh) return "❌ Falta hoja Despachados";

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return "No hay datos en Despachados";

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  // === Determinar mes COMPLETO anterior ===
  const hoy = new Date();
  const mesAnterior = hoy.getMonth() === 0 ? 11 : hoy.getMonth() - 1;
  const año = hoy.getMonth() === 0 ? hoy.getFullYear() - 1 : hoy.getFullYear();
  
  const inicio = new Date(año, mesAnterior, 1);
  const fin = new Date(año, mesAnterior + 1, 0); // último día del mes anterior

  // === Filtrar datos dentro del mes anterior ===
  const enMes = data.filter(r => {
    const fecha = new Date(r[15]); // col 16 = Fecha_archivo (índice 15)
    const pedId = r[11];           // col 12 = PED_ID
    return fecha >= inicio && fecha <= fin && pedId && pedId.toString().trim() !== "";
  });

  if (enMes.length === 0) return "No hay despachos en el mes anterior.";

  // === Inicializar acumuladores ===
  const productos = {};          // { Producto: {gal: xx, cub: xx, pedidos: Set() } }
  const volumenPedidosGal = []; // [{pedId, cliente, producto, gal}]
  const volumenPedidosCub = []; // [{pedId, cliente, producto, cub}]

  let totalCub = 0, totalGal = 0;
  
  enMes.forEach(r => {
    const producto = r[3];         // col D = Producto
    const cantidad = Number(r[5]) || 0; // col F = Cantidad
    const unidad = (r[6] || "").toString().toLowerCase(); // col G = Unidad
    const cliente = r[2];          // col C = Cliente
    const pedId = r[11];           // col 12 = PED_ID

    if (!productos[producto]) {
      productos[producto] = { gal: 0, cub: 0, pedidos: new Set() };
    }
    productos[producto].pedidos.add(pedId);

    if (unidad.includes("gal")) {
      productos[producto].gal += cantidad;
      totalGal += cantidad;
      volumenPedidosGal.push({ pedId, cliente, producto, gal: cantidad });
    }

    if (unidad.includes("cub")) {
      productos[producto].cub += cantidad;
      totalCub += cantidad;
      volumenPedidosCub.push({ pedId, cliente, producto, cub: cantidad });
    }
  });

  // === Productos más frecuentes ===
  const topProductos = Object.entries(productos)
    .map(([prod, val]) => ({
      producto: prod,
      gal: val.gal,
      cub: val.cub,
      pedidos: val.pedidos.size
    }))
    .sort((a, b) => b.pedidos - a.pedidos)
    .slice(0, 3);

  // === Top 3 por volumen (galones) ===
  const topGal = volumenPedidosGal
    .sort((a, b) => b.gal - a.gal)
    .slice(0, 3);

  // === Top 3 por volumen (cubetas) ===
  const topCub = volumenPedidosCub
    .sort((a, b) => b.cub - a.cub)
    .slice(0, 3);

  // === Construcción del mensaje ===
  const mesTexto = Utilities.formatDate(inicio, "America/Santo_Domingo", "MMMM yyyy");

  let msg = `📊 Resumen mensual de despachos – ${capitalizeFirst(mesTexto)}\n\n`;
  msg += `• Pedidos despachados: ${enMes.length}\n`;
  msg += `• Volumen total: ${totalCub} Cubetas / ${totalGal} Galones\n\n`;

  msg += `📈 Productos más frecuentes (por cantidad de pedidos):\n`;
  topProductos.forEach((p, i) => {
    msg += `${i + 1}. ${p.producto} (${p.gal} Gal / ${p.cub} Cub) – en ${p.pedidos} pedidos\n`;
  });

  msg += `\n🏋️ Pedidos con mayor volumen en galones:\n`;
  topGal.forEach((p, i) => {
    msg += `${i + 1}. #${p.pedId} – ${p.producto} (${p.gal} Gal) – cliente: ${p.cliente}\n`;
  });

  msg += `\n🏋️ Pedidos con mayor volumen en cubetas:\n`;
  topCub.forEach((p, i) => {
    msg += `${i + 1}. #${p.pedId} – ${p.producto} (${p.cub} Cub) – cliente: ${p.cliente}\n`;
  });

  msg += `\n✅ Todos los pedidos están registrados en hoja Despachados.`;

  return msg;
}

/**
 * Capitaliza la primera letra de un texto
 * @param {string} text - Texto a capitalizar
 * @returns {string} Texto capitalizado
 */
function capitalizeFirst(text) {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

/**
 * Envía automáticamente el resumen mensual (para trigger)
 */
function sendMensualResumen() {
  const msg = generarMensualResumen();
  if (msg.startsWith("❌") || msg.startsWith("No hay")) {
    Logger.log(msg);
    return;
  }
  Logger.log("Mensaje generado:\n" + msg);
  sendResumenToWA(msg, 'RESUMEN_MENSUAL'); // ✅ CORREGIDO
}

/**
 * Genera vista previa del resumen mensual (para botón manual)
 * @returns {string} Mensaje de resumen
 */
function debugMensualResumen() {
  const msg = generarMensualResumen();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Preview_Resumen_Mensual");
  if (!sh) sh = ss.insertSheet("Preview_Resumen_Mensual");
  sh.clear();
  sh.getRange(1, 1).setValue("Vista previa del resumen mensual");
  sh.getRange(3, 1).setValue(msg);
  return msg;
}

/**
 * Crea trigger automático para envío mensual (día 1 de cada mes)
 */
function crearTriggerMensualExacto() {
  // Elimina cualquier trigger anterior duplicado
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendMensualResumen") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Crear trigger el día 1 de cada mes a las 9:00 AM
  ScriptApp.newTrigger("sendMensualResumen")
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  Logger.log("✅ Trigger creado: sendMensualResumen() el día 1 de cada mes a las 9:00 AM");
}

/**
 * FUNCIÓN DE EMERGENCIA: Envía el resumen de octubre manualmente
 * Útil cuando los triggers fallaron y necesitas enviar un mes específico
 */
function enviarResumenOctubreManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Despachados');
  if (!sh) {
    Logger.log("❌ Falta hoja Despachados");
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay datos en Despachados");
    return;
  }

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  // Forzar octubre 2025 (cambiar año si es necesario)
  const inicio = new Date(2025, 9, 1);  // Mes 9 = Octubre
  const fin = new Date(2025, 9, 31);

  const enMes = data.filter(r => {
    const fecha = new Date(r[15]); // col 16 = Fecha_archivo (índice 15)
    const pedId = r[11];
    return fecha >= inicio && fecha <= fin && pedId && pedId.toString().trim() !== "";
  });

  if (enMes.length === 0) {
    Logger.log("No hay despachos en octubre 2025");
    return;
  }

  // === Procesamiento (igual que generarMensualResumen) ===
  const productos = {};
  const volumenPedidosGal = [];
  const volumenPedidosCub = [];
  let totalCub = 0, totalGal = 0;
  
  enMes.forEach(r => {
    const producto = r[3];
    const cantidad = Number(r[5]) || 0;
    const unidad = (r[6] || "").toString().toLowerCase();
    const cliente = r[2];
    const pedId = r[11];

    if (!productos[producto]) {
      productos[producto] = { gal: 0, cub: 0, pedidos: new Set() };
    }
    productos[producto].pedidos.add(pedId);

    if (unidad.includes("gal")) {
      productos[producto].gal += cantidad;
      totalGal += cantidad;
      volumenPedidosGal.push({ pedId, cliente, producto, gal: cantidad });
    }

    if (unidad.includes("cub")) {
      productos[producto].cub += cantidad;
      totalCub += cantidad;
      volumenPedidosCub.push({ pedId, cliente, producto, cub: cantidad });
    }
  });

  const topProductos = Object.entries(productos)
    .map(([prod, val]) => ({
      producto: prod,
      gal: val.gal,
      cub: val.cub,
      pedidos: val.pedidos.size
    }))
    .sort((a, b) => b.pedidos - a.pedidos)
    .slice(0, 3);

  const topGal = volumenPedidosGal.sort((a, b) => b.gal - a.gal).slice(0, 3);
  const topCub = volumenPedidosCub.sort((a, b) => b.cub - a.cub).slice(0, 3);

  let msg = `📊 Resumen mensual de despachos – Octubre 2025\n\n`;
  msg += `• Pedidos despachados: ${enMes.length}\n`;
  msg += `• Volumen total: ${totalCub} Cubetas / ${totalGal} Galones\n\n`;
  msg += `📈 Productos más frecuentes (por cantidad de pedidos):\n`;
  topProductos.forEach((p, i) => {
    msg += `${i + 1}. ${p.producto} (${p.gal} Gal / ${p.cub} Cub) – en ${p.pedidos} pedidos\n`;
  });
  msg += `\n🏋️ Pedidos con mayor volumen en galones:\n`;
  topGal.forEach((p, i) => {
    msg += `${i + 1}. #${p.pedId} – ${p.producto} (${p.gal} Gal) – cliente: ${p.cliente}\n`;
  });
  msg += `\n🏋️ Pedidos con mayor volumen en cubetas:\n`;
  topCub.forEach((p, i) => {
    msg += `${i + 1}. #${p.pedId} – ${p.producto} (${p.cub} Cub) – cliente: ${p.cliente}\n`;
  });
  msg += `\n✅ Todos los pedidos están registrados en hoja Despachados.`;

  Logger.log("📤 Enviando resumen de octubre...\n" + msg);
  sendResumenToWA(msg, 'RESUMEN_MENSUAL_OCTUBRE');
}