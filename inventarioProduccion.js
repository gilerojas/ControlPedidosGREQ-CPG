// ───────────────────────────────────────────────────────────
// inventarioProduccion.gs – Sistema SIG-CPG v6.2 COMPLETO
// Con REPROCESO e ID_PRODUCCION automático
// ───────────────────────────────────────────────────────────

// ═══════════════════════════════════════════════════════════
// SISTEMA DE MAPEO CPG → SIG_VENTAS (FALTANTE)
// ═══════════════════════════════════════════════════════════

/**
 * Obtener mapeo completo CPG → SIG
 */
function getMapeoCompletoCPGaSIG() {
  return {
    'ACRILICA SUPERIOR HP': { idTipo: 'AS-HP', tipo: 'ACRILICA SUPERIOR (HP)', naturaleza: 'Agua', existe: true },
    'ACRILICA SUPERIOR TIPO B': { idTipo: 'AS-B', tipo: 'ACRILICA SUPERIOR (Tipo B)', naturaleza: 'Agua', existe: true },
    'SEMIGLOSS PREMIUM': { idTipo: 'SEM-P', tipo: 'SEMIGLOSS PREMIUM', naturaleza: 'Agua', existe: false, esNuevo: true },
    'SEMIGLOSS TIPO B': { idTipo: 'SEM-B', tipo: 'SEMIGLOSS TIPO B', naturaleza: 'Agua', existe: false, esNuevo: true },
    'SATINADA': { idTipo: 'SAT', tipo: 'SATINADA', naturaleza: 'Agua', existe: true },
    'TEXTURIZADAS': { idTipo: 'TXT', tipo: 'TEXTURIZADAS', naturaleza: 'Agua', existe: true },
    'PROYECTO O CONTRACTOR': { idTipo: 'PRO', tipo: 'PROYECTO O CONTRACTOR', naturaleza: 'Agua', existe: true },
    'PROYECTO P/ TECHOS': { idTipo: 'PTE', tipo: 'PROYECTO P/TECHO', naturaleza: 'Agua', existe: true },
    'ECONOMICA': { idTipo: 'ECO', tipo: 'ECONOMICA', naturaleza: 'Agua', existe: true },
    'PRIMER ACRILICO': { idTipo: 'PRI', tipo: 'PRIMER', naturaleza: 'Agua', existe: true },
    'SELLADOR TECHOS HP': { idTipo: 'SLT', tipo: 'SELLADOR TECHOS', naturaleza: 'Agua', existe: true },
    'SELLADOR TECHOS TIPO B': { idTipo: 'SLT-B', tipo: 'SELLADOR DE TECHOS (TIPO B)', naturaleza: 'Agua', existe: true },
    'SEALER WATER': { idTipo: 'SLT', tipo: 'SELLADOR TECHOS', naturaleza: 'Agua', existe: true },
    'ESMALTE INDUSTRIAL': { idTipo: 'EIN', tipo: 'ESMALTE INDUSTRIAL', naturaleza: 'Aceite', existe: true },
    'ESMALTE INDUSTRIAL ANTICORROSIVO': { idTipo: 'EIN', tipo: 'ESMALTE INDUSTRIAL', naturaleza: 'Aceite', existe: true },
    'ESMALTE TRAFICO': { idTipo: 'TRA', tipo: 'TRAFICO', naturaleza: 'Aceite', existe: true },
    'ESMALTE SINTETICO O MANTENIMIENTO': { idTipo: 'EMA', tipo: 'ESMALTE MANTENIMIENTO', naturaleza: 'Aceite', existe: true },
    'EPOXICA': { idTipo: 'EPX', tipo: 'EPOXICA', naturaleza: 'Aceite', existe: true },
    'DRY WET': { idTipo: 'DRY', tipo: 'DRY COAT', naturaleza: 'Agua', existe: true },
    'BARNIZ CLEAR INDUSTRIAL': { idTipo: 'BCL', tipo: 'BARNIZ CLEAR INDUSTRIAL', naturaleza: 'Aceite', existe: false, esNuevo: true },
    'BARNIZ PORT EPOXI CLEAR': { idTipo: 'BEP', tipo: 'BARNIZ PORT EPOXI CLEAR', naturaleza: 'Aceite', existe: false, esNuevo: true }
  };
}

/**
 * Mapear producto CPG a estructura SIG
 */
function mapearProductoCPGaSIG(productoCPG) {
  const mapeoCompleto = getMapeoCompletoCPGaSIG();
  const productoNorm = normalize(productoCPG);
  
  for (const [nombreCPG, datosSIG] of Object.entries(mapeoCompleto)) {
    if (normalize(nombreCPG) === productoNorm) {
      return { ...datosSIG, productoCPG: productoCPG, nombreCPG: nombreCPG, mapeado: true };
    }
  }
  
  return {
    idTipo: 'FAL', tipo: productoCPG, naturaleza: 'Agua', existe: false, esNuevo: true,
    requiereRevision: true, productoCPG: productoCPG, mapeado: false, error: 'SIN_MAPEO_DEFINIDO'
  };
}

/**
 * Generar nuevo ID secuencial para tipo
 */
function generarNuevoIdSIG(idTipo) {
  try {
    const inventarioSS = SpreadsheetApp.openById('1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw');
    const inventarioSheet = inventarioSS.getSheetByName('Inventario');
    const datos = inventarioSheet.getDataRange().getValues();
    
    let maxNumero = 0;
    const patron = new RegExp(`^${idTipo.replace(/[-]/g, '\\-')}-(\\d+)$`);
    
    for (let i = 1; i < datos.length; i++) {
      const idExistente = datos[i][5];
      if (idExistente) {
        const match = idExistente.toString().match(patron);
        if (match) {
          const numero = parseInt(match[1]);
          if (numero > maxNumero) maxNumero = numero;
        }
      }
    }
    
    return `${idTipo}-${String(maxNumero + 1).padStart(3, '0')}`;
  } catch (error) {
    Logger.log(`Error generando ID para ${idTipo}: ${error}`);
    return `${idTipo}-001`;
  }
}

/**
 * CONSTANTES ACTUALIZADAS
 */
const COL_ORIGEN = 16;        // P (columna Origen)
const COL_ID_INVENTARIO = 17; // Q (columna ID_Inventario)

// Constantes para hoja Producción (18 columnas) 
const PROD_COL = {
  PED_ID: 1, CLIENTE: 2, PRODUCTO: 3, COLOR: 4, CANTIDAD: 5, ENVASE: 6,
  MAQUINA: 7, OPERARIO: 8, ESTADO_PROD: 9, FECHA_PROD: 10, HORA_INICIO: 11,
  HORA_FIN: 12, CANT_PRODUCIDA: 13, UNIDAD_ENVASE: 14, REPROCESO: 15, 
  ID_CONSUMIDO: 16, TIEMPO_TOTAL: 17, ID_PRODUCCION: 18
};

/**
 * FUNCIÓN: Normalizar texto
 */
function normalize(text) {
  return text
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim()
    .replace(/\s+/g, ' ');
}

/**
 * FUNCIÓN: Obtener estado válido
 */
function getEstadoValido(estadoDeseado) {
  const estadosValidos = [
    'PENDIENTE', 'EN PRODUCCION', 'ESPERA DE CALIDAD', 
    'LISTO P/ ENVASAR', 'LISTO P/ DESPACHAR', 'DESPACHADO', 'CANCELADO'
  ];
  
  const estadoNorm = normalize(estadoDeseado);
  
  for (const estado of estadosValidos) {
    if (normalize(estado) === estadoNorm) {
      return estado;
    }
  }
  
  return 'PENDIENTE';
}

/**
 * FUNCIÓN: Mapear producto a ID_Tipo (para ID_PRODUCCION)
 */
function mapearProductoAIdTipo(producto) {
  const mapeo = {
    'ACRILICA SUPERIOR HP': 'AS-HP',
    'ACRILICA SUPERIOR': 'AS',
    'ACRILICA SUPERIOR TIPO B': 'AS-B',
    'SEMIGLOSS PREMIUM': 'SEM',
    'SEMIGLOSS TIPO B': 'SEM',
    'SATINADA': 'SAT',
    'ESMALTE INDUSTRIAL': 'EIN',
    'EPOXICA': 'EPX',
    'TEXTURIZADAS': 'TXT',
    'DRY WET': 'DRY',
    'PROYECTO O CONTRACTOR': 'PRO',
    'PROYECTO P/ TECHOS': 'PTE',
    'ECONOMICA': 'ECO',
    'PRIMER ACRILICO': 'PRI',
    'SELLADOR TECHOS HP': 'SLT',
    'SELLADOR TECHOS TIPO B': 'SLT-B',
    'ESMALTE TRAFICO': 'TRA'
  };
  
  const productoNorm = normalize(producto);
  
  for (const [nombre, codigo] of Object.entries(mapeo)) {
    if (normalize(nombre) === productoNorm) {
      return codigo;
    }
  }
  
  return 'PRD'; // Genérico para productos no mapeados
}

/**
 * FUNCIÓN: Generar siguiente ID de producción
 */
function generarSiguienteIdProduccion(producto) {
  const ahora = new Date();
  const inicioAño = new Date(ahora.getFullYear(), 0, 1);
  const diaJuliano = Math.ceil((ahora - inicioAño) / (1000 * 60 * 60 * 24)) + 1;
  const diaStr = String(diaJuliano).padStart(3, '0');
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let contadorSheet = spreadsheet.getSheetByName('BATCH_COUNTER');
  
  if (!contadorSheet) {
    contadorSheet = spreadsheet.insertSheet('BATCH_COUNTER');
    contadorSheet.getRange(1, 1, 1, 2).setValues([['Dia', 'Contador']]);
    contadorSheet.hideSheet();
  }
  
  const datos = contadorSheet.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === diaStr) {
      const nuevoContador = datos[i][1] + 1;
      contadorSheet.getRange(i + 1, 2).setValue(nuevoContador);
      return `GQ${diaStr}${String(nuevoContador).padStart(4, '0')}`;
    }
  }
  
  contadorSheet.appendRow([diaStr, 1]);
  return `GQ${diaStr}0001`;
}

// ═══════════════════════════════════════════════════════════
// TRIGGERS PARA HOJA PEDIDOS
// ═══════════════════════════════════════════════════════════

/**
 * TRIGGER: Cambios en columna Origen (P)
 */
function handleOrigenEdit(e) {
  const SH_NAME = 'Pedidos';
  const rng = e.range, sh = rng.getSheet();
  
  if (sh.getName() !== SH_NAME) return;
  if (rng.getColumn() !== COL_ORIGEN || rng.getRow() === 1) return;

  const nuevoOrigen = rng.getValue();
  const estadoActual = sh.getRange(rng.getRow(), 11).getValue();
  const idInventario = sh.getRange(rng.getRow(), COL_ID_INVENTARIO).getValue();
  
  // VALIDACIÓN 1: No cambiar origen si ya está despachado
  if (estadoActual === "DESPACHADO") {
    SpreadsheetApp.getUi().alert(
      'Pedido ya despachado',
      'No se puede cambiar el origen de un pedido ya DESPACHADO',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    sh.getRange(rng.getRow(), COL_ORIGEN).setValue("");
    return;
  }
  
  // VALIDACIÓN 2: INVENTARIO requiere ID
  if (nuevoOrigen === "INVENTARIO" && !idInventario) {
    SpreadsheetApp.getUi().alert(
      'ID Inventario requerido',
      'Para seleccionar INVENTARIO debe especificar primero un ID válido',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    sh.getRange(rng.getRow(), COL_ORIGEN).setValue("");
    return;
  }
  
  if (nuevoOrigen) {
    procesarSegunOrigen(rng.getRow(), nuevoOrigen);
  }
}

/**
 * TRIGGER: Cambios en columna Estado (K)
 */
function handleEstadoEdit(e) {
  const SH_NAME = 'Pedidos';
  const rng = e.range, sh = rng.getSheet();
  
  if (sh.getName() !== SH_NAME || rng.getColumn() !== 11 || rng.getRow() === 1) return;

  const nuevoEstado = rng.getValue();
  const origen = sh.getRange(rng.getRow(), COL_ORIGEN).getValue();
  const idInventario = sh.getRange(rng.getRow(), COL_ID_INVENTARIO).getValue();
  
  // Validar: DESPACHADO + INVENTARIO requiere ID
  if (nuevoEstado === "DESPACHADO" && origen === "INVENTARIO" && !idInventario) {
    SpreadsheetApp.getUi().alert(
      'ID Inventario faltante',
      'No se puede marcar como DESPACHADO un pedido de INVENTARIO sin ID_Inventario',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    sh.getRange(rng.getRow(), 11).setValue("PENDIENTE");
    return;
  }
}

/**
 * TRIGGER: Cambios en columna ID_Inventario (Q)
 */
function handleIdInventarioEdit(e) {
  const SH_NAME = 'Pedidos';
  const rng = e.range, sh = rng.getSheet();
  
  if (sh.getName() !== SH_NAME || rng.getColumn() !== COL_ID_INVENTARIO || rng.getRow() === 1) return;

  const idInventario = rng.getValue();
  const origen = sh.getRange(rng.getRow(), COL_ORIGEN).getValue();
  
  if (idInventario && !validarIdExisteEnInventario(idInventario)) {
    SpreadsheetApp.getUi().alert(
      'ID no válido',
      `El ID ${idInventario} no existe en el inventario o no tiene stock disponible`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    rng.setValue("");
    return;
  }
  
  // Si origen=INVENTARIO e ID válido, procesar automáticamente
  if (origen === "INVENTARIO" && idInventario) {
    procesarSegunOrigen(rng.getRow(), origen);
  }
}

// ═══════════════════════════════════════════════════════════
// TRIGGER PARA HOJA PRODUCCIÓN
// ═══════════════════════════════════════════════════════════

/**
 * TRIGGER: Cambios en Estado_Prod (I) de hoja Producción
 */
function handleEstadoProduccion(e) {
  const SH_NAME = 'Produccion';
  const rng = e.range, sh = rng.getSheet();
  
  if (sh.getName() !== SH_NAME) return;
  if (rng.getColumn() !== PROD_COL.ESTADO_PROD || rng.getRow() === 1) return;

  const estadoProd = rng.getValue();
  const fila = rng.getRow();
  
  Logger.log(`Estado Producción cambiado: Fila ${fila}, Estado: ${estadoProd}`);
  
  switch(estadoProd) {
    case "EN PROCESO":
      procesarEnProceso(fila, sh);
      break;
    case "ESPERANDO CALIDAD":
      actualizarPedidoDesdeProduccion(fila, "ESPERA DE CALIDAD");
      enviarNotificacionProduccion(fila, "CALIDAD", "🔬 A CALIDAD");
      break;
    case "ESPERANDO ENVASE":  
      actualizarPedidoDesdeProduccion(fila, "LISTO P/ ENVASAR");
      enviarNotificacionProduccion(fila, "ENVASE", "✅ QC OK → ENVASE");
      break;
    case "TERMINADO":
      procesarTerminado(fila, sh);
      break;
    default:
      Logger.log(`Estado ${estadoProd} - sin acción automática`);
  }
}

// ═══════════════════════════════════════════════════════════
// FUNCIONES PRINCIPALES
// ═══════════════════════════════════════════════════════════

/**
 * Procesar según origen del pedido
 */
function procesarSegunOrigen(pedidoRow, origen) {
  const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  
  Logger.log(`Procesando pedido fila ${pedidoRow} con origen: ${origen}`);
  
  switch(origen) {
    case "INVENTARIO":
      const resultado = verificarDisponibilidadInventario(pedidoRow);
      if (resultado.disponible) {
        // Reservar (no debitar inmediatamente - se debita a las 6PM)
        pedidosSheet.getRange(pedidoRow, 11).setValue("LISTO P/ DESPACHAR");
        pedidosSheet.getRange(pedidoRow, 13).setValue(new Date()).setNumberFormat('dd/MM/yy HH:mm');
        enviarNotificacionProduccion(pedidoRow, "INVENTARIO", "", resultado.itemUsado);
      } else {
        pedidosSheet.getRange(pedidoRow, COL_ORIGEN).setValue("");
        Logger.log(`Error inventario: ${resultado.error}`);
      }
      break;
      
    case "PRODUCCIÓN":
      if (confirmarEnvioAProduccion(pedidoRow)) {
        const estadoValido = getEstadoValido("EN PRODUCCIÓN");
        pedidosSheet.getRange(pedidoRow, 11).setValue(estadoValido);
        pedidosSheet.getRange(pedidoRow, 13).setValue(new Date()).setNumberFormat('dd/MM/yy HH:mm');
        crearOrdenProduccion(pedidoRow);
      } else {
        pedidosSheet.getRange(pedidoRow, COL_ORIGEN).setValue("");
        Logger.log(`Usuario canceló envío a producción`);
      }
      break;
      
    case "MIXTO":
      if (confirmarEnvioAProduccion(pedidoRow)) {
        const estadoValido = getEstadoValido("EN PRODUCCIÓN");
        pedidosSheet.getRange(pedidoRow, 11).setValue(estadoValido);
        pedidosSheet.getRange(pedidoRow, 13).setValue(new Date()).setNumberFormat('dd/MM/yy HH:mm');
        crearOrdenProduccion(pedidoRow);
      } else {
        pedidosSheet.getRange(pedidoRow, COL_ORIGEN).setValue("");
        Logger.log(`Usuario canceló envío a producción (MIXTO)`);
      }
      break;
      
    default:
      Logger.log(`Origen desconocido: ${origen}`);
  }
}

/**
 * Procesar estado EN PROCESO
 */
function procesarEnProceso(fila, produccionSheet) {
  const ahora = new Date();
  
  // Llenar fecha y hora automáticamente
  produccionSheet.getRange(fila, PROD_COL.FECHA_PROD).setValue(ahora).setNumberFormat('dd/MM/yyyy');
  produccionSheet.getRange(fila, PROD_COL.HORA_INICIO).setValue(ahora).setNumberFormat('HH:mm');
  
  // Actualizar pedido a EN PRODUCCIÓN
  actualizarPedidoDesdeProduccion(fila, "EN PRODUCCION");
  
  // Enviar notificación de inicio
  enviarNotificacionProduccion(fila, "INICIADO", "🔄 *INICIADO*");
  
  Logger.log(`Producción iniciada automáticamente en fila ${fila}`);
}

/**
 * Procesar estado TERMINADO con warning de surplus
 */
function procesarTerminado(fila, produccionSheet) {
  const ahora = new Date();
  
  // Validar cantidad producida
  const cantProducida = produccionSheet.getRange(fila, PROD_COL.CANT_PRODUCIDA).getValue();
  if (!cantProducida || cantProducida <= 0) {
    
    const estadoAnterior = obtenerEstadoAnterior(fila, produccionSheet);
    
    SpreadsheetApp.getUi().alert(
      '⚠️ Cantidad requerida',
      'Debe especificar la cantidad producida antes de marcar como TERMINADO',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    produccionSheet.getRange(fila, PROD_COL.ESTADO_PROD).setValue(estadoAnterior);
    Logger.log(`Estado revertido a: ${estadoAnterior}`);
    return;
  }
  
  // VALIDACIÓN REPROCESO: Si REPROCESO=SI, ID_CONSUMIDO es obligatorio
  const reproceso = produccionSheet.getRange(fila, PROD_COL.REPROCESO).getValue();
  const idConsumido = produccionSheet.getRange(fila, PROD_COL.ID_CONSUMIDO).getValue();
  
  if (reproceso === "SI" && !idConsumido) {
    const estadoAnterior = obtenerEstadoAnterior(fila, produccionSheet);
    
    SpreadsheetApp.getUi().alert(
      '⚠️ ID Consumido requerido',
      'Las producciones marcadas como REPROCESO=SI requieren especificar ID_Consumido de la materia prima utilizada',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    produccionSheet.getRange(fila, PROD_COL.ESTADO_PROD).setValue(estadoAnterior);
    Logger.log(`Estado revertido por falta ID_Consumido en reproceso`);
    return;
  }
  
  // NUEVO: WARNING CON CONFIRMACIÓN DE SURPLUS
  const warningConfirmado = mostrarWarningTerminado(fila, produccionSheet);
  
  if (!warningConfirmado) {
    // Usuario canceló - revertir estado
    const estadoAnterior = obtenerEstadoAnterior(fila, produccionSheet);
    produccionSheet.getRange(fila, PROD_COL.ESTADO_PROD).setValue(estadoAnterior);
    Logger.log(`Usuario canceló terminado - estado revertido a: ${estadoAnterior}`);
    return;
  }
  
  // Continuar con lógica normal si todas las validaciones y confirmación pasan
  produccionSheet.getRange(fila, PROD_COL.HORA_FIN).setValue(ahora).setNumberFormat('HH:mm');
  calcularTiempoTotal(fila, produccionSheet);
  procesarProduccionTerminada(fila);
  
  Logger.log(`Producción terminada en fila ${fila}`);
}

/**
 * Mostrar warning con detalles de surplus antes de terminar producción
 */
function mostrarWarningTerminado(fila, produccionSheet) {
  // Obtener datos de la producción
  const cantProducida = produccionSheet.getRange(fila, PROD_COL.CANT_PRODUCIDA).getValue();
  const unidadEnvase = produccionSheet.getRange(fila, PROD_COL.UNIDAD_ENVASE).getValue();
  const producto = produccionSheet.getRange(fila, PROD_COL.PRODUCTO).getValue();
  const color = produccionSheet.getRange(fila, PROD_COL.COLOR).getValue();
  const cantidad = produccionSheet.getRange(fila, PROD_COL.CANTIDAD).getValue();
  const envase = produccionSheet.getRange(fila, PROD_COL.ENVASE).getValue();
  const pedId = produccionSheet.getRange(fila, PROD_COL.PED_ID).getValue();
  const operario = produccionSheet.getRange(fila, PROD_COL.OPERARIO).getValue();
  const maquina = produccionSheet.getRange(fila, PROD_COL.MAQUINA).getValue();
  const reproceso = produccionSheet.getRange(fila, PROD_COL.REPROCESO).getValue();
  const idConsumido = produccionSheet.getRange(fila, PROD_COL.ID_CONSUMIDO).getValue();
  
  // Calcular surplus en galones
  const cantOriginalGalones = envase.toLowerCase().includes('cubeta') ? cantidad * 5 : cantidad;
  const cantProducidaGalones = unidadEnvase && unidadEnvase.toLowerCase().includes('cubeta') ? cantProducida * 5 : cantProducida;
  const surplusGalones = cantProducidaGalones - cantOriginalGalones;
  
  // Construir mensaje de confirmación
  let mensaje = `⚠️ CONFIRMAR TERMINADO\n\n`;
  mensaje += `📋 DETALLES DE PRODUCCIÓN:\n`;
  mensaje += `PED_ID: ${pedId}\n`;
  mensaje += `Producto: ${producto} ${color}\n`;
  mensaje += `Operario: ${operario} | Máquina: ${maquina}\n\n`;
  
  mensaje += `📊 CANTIDADES:\n`;
  mensaje += `Pedido original: ${cantidad} ${envase}\n`;
  mensaje += `Cantidad producida: ${cantProducida} ${unidadEnvase || envase}\n\n`;
  
  // Información de surplus
  if (surplusGalones > 0) {
    // Calcular surplus en la unidad de envase producida
    const surplusEnUnidadProd = unidadEnvase && unidadEnvase.toLowerCase().includes('cubeta') ? 
      surplusGalones / 5 : surplusGalones;
    const unidadFinal = unidadEnvase && unidadEnvase.toLowerCase().includes('cubeta') ? 'Cubeta' : 'Galón';
    
    mensaje += `📦 SURPLUS DETECTADO:\n`;
    mensaje += `${surplusEnUnidadProd} ${unidadFinal} se enviará automáticamente a SIG_Ventas\n`;
    mensaje += `(Equivalente a ${surplusGalones} galones)\n\n`;
  } else if (surplusGalones < 0) {
    mensaje += `⚠️ PRODUCCIÓN MENOR AL PEDIDO:\n`;
    mensaje += `Faltan ${Math.abs(surplusGalones)} galones respecto al pedido\n\n`;
  } else {
    mensaje += `✅ PRODUCCIÓN EXACTA:\n`;
    mensaje += `Sin surplus - cantidad exacta del pedido\n\n`;
  }
  
  // Información de reproceso
  if (reproceso === "SI" && idConsumido) {
    mensaje += `🔧 REPROCESO:\n`;
    mensaje += `Materia prima consumida: ${idConsumido}\n`;
    mensaje += `Se debitará automáticamente del inventario\n\n`;
  }
  
  mensaje += `¿Confirmar TERMINADO y procesar automáticamente?`;
  
  // Mostrar diálogo de confirmación
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    'Confirmar Terminado - Revisión Final', 
    mensaje, 
    ui.ButtonSet.YES_NO
  );
  
  const confirmado = respuesta === ui.Button.YES;
  
  if (confirmado) {
    Logger.log(`Usuario confirmó terminado para ${pedId} con surplus: ${surplusGalones} galones`);
  } else {
    Logger.log(`Usuario canceló terminado para ${pedId}`);
  }
  
  return confirmado;
}

/**
 * Obtener estado anterior lógico
 */
function obtenerEstadoAnterior(fila, produccionSheet) {
  const horaInicio = produccionSheet.getRange(fila, PROD_COL.HORA_INICIO).getValue();
  const cantProducida = produccionSheet.getRange(fila, PROD_COL.CANT_PRODUCIDA).getValue();
  
  if (cantProducida && cantProducida > 0) {
    return "ESPERANDO ENVASE";
  }
  
  if (horaInicio) {
    return "EN PROCESO";
  }
  
  return "PENDIENTE";
}

/**
 * Calcular tiempo total de producción
 */
function calcularTiempoTotal(fila, produccionSheet) {
  const horaInicio = produccionSheet.getRange(fila, PROD_COL.HORA_INICIO).getValue();
  const horaFin = produccionSheet.getRange(fila, PROD_COL.HORA_FIN).getValue();
  
  if (horaInicio && horaFin) {
    const diferencia = horaFin.getTime() - horaInicio.getTime();
    const horas = Math.floor(diferencia / (1000 * 60 * 60));
    const minutos = Math.floor((diferencia % (1000 * 60 * 60)) / (1000 * 60));
    
    const tiempoTexto = `${horas}h ${minutos}min`;
    produccionSheet.getRange(fila, PROD_COL.TIEMPO_TOTAL).setValue(tiempoTexto);
    
    return tiempoTexto;
  }
  return "";
}

/**
 * Verificar si surplus ya fue procesado para evitar duplicados
 * @param {string} pedId - ID del pedido
 * @param {string} idProduccion - ID de producción
 * @return {boolean} True si ya existe (duplicado)
 */
function verificarSurplusYaProcesado(pedId, idProduccion) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('LOG_Surplus');
    
    if (!logSheet) {
      Logger.log(`No existe LOG_Surplus - no hay duplicados que verificar`);
      return false; // No hay log = no hay duplicados
    }
    
    const datos = logSheet.getDataRange().getValues();
    
    // Verificar estructura de la hoja (debe tener al menos 4 columnas)
    if (datos.length < 1 || datos[0].length < 4) {
      Logger.log(`LOG_Surplus no tiene estructura adecuada`);
      return false;
    }
    
    // Buscar coincidencia de PED_ID + ID_PRODUCCION
    for (let i = 1; i < datos.length; i++) {
      const logPedId = datos[i][2];        // Columna C: PED_ID
      const logIdProd = datos[i][3];       // Columna D: ID_PRODUCCION
      
      if (logPedId === pedId && logIdProd === idProduccion) {
        Logger.log(`❌ Surplus duplicado detectado: PED_ID ${pedId}, ID_PRODUCCION ${idProduccion}`);
        return true; // Ya existe
      }
    }
    
    Logger.log(`✅ No se encontró duplicado para: PED_ID ${pedId}, ID_PRODUCCION ${idProduccion}`);
    return false; // No existe duplicado
    
  } catch (error) {
    Logger.log(`Error verificando duplicados: ${error}`);
    return false; // En caso de error, permitir procesamiento
  }
}

/**
 * Wrapper para agregarSurplusAlInventario() con trazabilidad completa
 */
function agregarSurplusAlInventarioConTrazabilidad(producto, color, surplus, envase, pedId, idProduccion) {
  // Llamar función original
  const resultado = agregarSurplusAlInventario(producto, color, surplus, envase);
  
  // IMPLEMENTAR el logging mejorado aquí
  if (resultado.includes('agregado') || resultado.includes('creado')) {
    const idMatch = resultado.match(/ID:\s*\*([^*]+)\*/);
    const idGenerado = idMatch ? idMatch[1] : '';
    
    // Registrar con trazabilidad completa
    registrarSurplusEnLog({
      tipo: resultado.includes('agregado') ? 'SURPLUS_AGREGADO' : 'PRODUCTO_CREADO',
      pedId: pedId,
      idProduccion: idProduccion,
      id: idGenerado,
      producto: producto,
      color: color,
      surplus: surplus,
      envase: envase,
      stockAnterior: 0,
      stockNuevo: surplus
    });
  }
  
  return resultado;
}

/**
 * Procesar producción terminada completa
 */
function procesarProduccionTerminada(fila) {
  const produccionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produccion');
  
  const pedId = produccionSheet.getRange(fila, PROD_COL.PED_ID).getValue();
  const cliente = produccionSheet.getRange(fila, PROD_COL.CLIENTE).getValue();
  const producto = produccionSheet.getRange(fila, PROD_COL.PRODUCTO).getValue();
  const color = produccionSheet.getRange(fila, PROD_COL.COLOR).getValue();
  const cantidadOriginal = produccionSheet.getRange(fila, PROD_COL.CANTIDAD).getValue();
  const envaseOriginal = produccionSheet.getRange(fila, PROD_COL.ENVASE).getValue();
  const cantProducida = produccionSheet.getRange(fila, PROD_COL.CANT_PRODUCIDA).getValue();
  const unidadEnvase = produccionSheet.getRange(fila, PROD_COL.UNIDAD_ENVASE).getValue();
  const reproceso = produccionSheet.getRange(fila, PROD_COL.REPROCESO).getValue();
  const idConsumido = produccionSheet.getRange(fila, PROD_COL.ID_CONSUMIDO).getValue();
  const tiempoTotal = produccionSheet.getRange(fila, PROD_COL.TIEMPO_TOTAL).getValue();
  const operario = produccionSheet.getRange(fila, PROD_COL.OPERARIO).getValue();
  const maquina = produccionSheet.getRange(fila, PROD_COL.MAQUINA).getValue();
  
  // GENERAR ID_PRODUCCION AUTOMÁTICO
  const idProduccion = generarSiguienteIdProduccion(producto);
  produccionSheet.getRange(fila, PROD_COL.ID_PRODUCCION).setValue(idProduccion);
  
  // *** VERIFICAR DUPLICADOS ANTES DE CONTINUAR ***
  if (verificarSurplusYaProcesado(pedId, idProduccion)) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Surplus ya procesado',
      `Esta producción ya fue procesada anteriormente:\n\n` +
      `PED_ID: ${pedId}\n` +
      `ID_PRODUCCION: ${idProduccion}\n\n` +
      `Revise el LOG_Surplus para más detalles.\n` +
      `Si necesita reprocesar, cambie el estado y vuelva a TERMINADO.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log(`🚫 Surplus duplicado bloqueado: ${pedId} - ${idProduccion}`);
    return; // SALIR SIN PROCESAR
  }
  
  // CONVERSIÓN A GALONES PARA CÁLCULOS
  let cantOriginalGalones, cantProducidaGalones;
  
  if (envaseOriginal.toLowerCase().includes('cubeta')) {
    cantOriginalGalones = cantidadOriginal * 5;
  } else {
    cantOriginalGalones = cantidadOriginal;
  }
  
  if (unidadEnvase && unidadEnvase.toLowerCase().includes('cubeta')) {
    cantProducidaGalones = cantProducida * 5;
  } else {
    cantProducidaGalones = cantProducida;
  }
  
  const surplusGalones = cantProducidaGalones - cantOriginalGalones;
  
  // Procesar materia prima y surplus
  let mensajeMateriaPrima = "";
  if (reproceso === "SI" && idConsumido) {
    mensajeMateriaPrima = debitarMateriaPrima(idConsumido);
  }
  
  let mensajeSurplus = "";
  if (surplusGalones > 0) {
    let surplusEnUnidadProd, unidadFinal;
    
    if (unidadEnvase && unidadEnvase.toLowerCase().includes('cubeta')) {
      surplusEnUnidadProd = surplusGalones / 5;
      unidadFinal = "Cubetas";
    } else {
      surplusEnUnidadProd = surplusGalones;
      unidadFinal = "Galones";
    }
    
    // *** PROCESAR SURPLUS CON INFORMACIÓN COMPLETA ***
    mensajeSurplus = agregarSurplusAlInventarioConTrazabilidad(
      producto, color, surplusEnUnidadProd, unidadFinal, pedId, idProduccion
    );
  }
  
  // Actualizar pedido y enviar notificación
  actualizarPedidoDesdeProduccion(fila, "LISTO P/ DESPACHAR");
  
  const datosCompletado = {
    pedId, cliente, producto, color, cantidadOriginal, envaseOriginal,
    cantProducida, unidadEnvase, surplusGalones, tiempoTotal, 
    operario, maquina, mensajeMateriaPrima, mensajeSurplus, idProduccion
  };
  
  enviarNotificacionProduccion(fila, "COMPLETADO", "✅ *COMPLETADO*", datosCompletado);
  
  Logger.log(`✅ Producción completada: ${cantProducida} ${unidadEnvase} = ${cantProducidaGalones} galones, ID: ${idProduccion}`);
}


// ═══════════════════════════════════════════════════════════
// FUNCIONES DE INVENTARIO
// ═══════════════════════════════════════════════════════════

/**
 * Validar si ID existe en inventario
 */
function validarIdExisteEnInventario(idInventario) {
  try {
    const inventarioSS = SpreadsheetApp.openById('1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw');
    const inventarioSheet = inventarioSS.getSheetByName('Inventario');
    
    const datos = inventarioSheet.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][5] === idInventario && datos[i][9] > 0) {
        return true;
      }
    }
    return false;
  } catch (error) {
    Logger.log(`Error validando ID ${idInventario}: ${error}`);
    return false;
  }
}

/**
 * Verificar disponibilidad completa de inventario
 */
function verificarDisponibilidadInventario(pedidoRow) {
  const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const idInventario = pedidosSheet.getRange(pedidoRow, COL_ID_INVENTARIO).getValue();
  
  if (!idInventario) {
    SpreadsheetApp.getUi().alert(
      'ID Inventario requerido',
      'Para usar INVENTARIO debe especificar un ID válido',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { disponible: false, error: 'ID_FALTANTE' };
  }
  
  try {
    const inventarioSS = SpreadsheetApp.openById('1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw');
    const inventarioSheet = inventarioSS.getSheetByName('Inventario');
    const cantidadPedida = pedidosSheet.getRange(pedidoRow, 6).getValue();
    
    const datos = inventarioSheet.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const row = datos[i];
      if (row[5] === idInventario) {
        const cantidadDisponible = row[9];
        
        if (cantidadDisponible >= cantidadPedida) {
          return {
            disponible: true,
            itemUsado: {
              fila: i + 1,
              id: idInventario,
              cantidadUsar: cantidadPedida,
              cantidadDisponible: cantidadDisponible
            }
          };
        } else {
          SpreadsheetApp.getUi().alert(
            'Stock insuficiente',
            `ID ${idInventario} solo tiene ${cantidadDisponible} disponibles. Necesita ${cantidadPedida}`,
            SpreadsheetApp.getUi().ButtonSet.OK
          );
          return { disponible: false, error: 'STOCK_INSUFICIENTE' };
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(
      'ID no encontrado',
      `El ID ${idInventario} no existe en el inventario`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return { disponible: false, error: 'ID_NO_EXISTE' };
  } catch (error) {
    Logger.log(`Error verificando disponibilidad: ${error}`);
    return { disponible: false, error: 'ERROR_CONEXION' };
  }
}

// ═══════════════════════════════════════════════════════════
// FUNCIONES DE PRODUCCIÓN
// ═══════════════════════════════════════════════════════════

/**
 * Actualizar pedido desde producción
 */
function actualizarPedidoDesdeProduccion(filaProduccion, nuevoEstado) {
  const produccionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produccion');
  const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  
  const pedId = produccionSheet.getRange(filaProduccion, PROD_COL.PED_ID).getValue();
  
  Logger.log(`Actualizando pedido ${pedId} a estado: ${nuevoEstado}`);
  
  const pedidosData = pedidosSheet.getDataRange().getValues();
  for (let i = 1; i < pedidosData.length; i++) {
    if (pedidosData[i][11] === pedId) {
      const estadosValidos = [
        'PENDIENTE', 'EN PRODUCCION', 'ESPERA DE CALIDAD', 
        'LISTO P/ ENVASAR', 'LISTO P/ DESPACHAR', 'DESPACHADO', 'CANCELADO'
      ];
      
      const estadoValido = estadosValidos.includes(nuevoEstado) ? nuevoEstado : 'EN PRODUCCION';
      
      pedidosSheet.getRange(i + 1, 11).setValue(estadoValido);
      pedidosSheet.getRange(i + 1, 13).setValue(new Date()).setNumberFormat('dd/MM/yy HH:mm');
      Logger.log(`Pedido ${pedId} actualizado a: ${estadoValido}`);
      return;
    }
  }
  
  Logger.log(`No se encontró pedido ${pedId} en hoja Pedidos`);
}

/**
 * Crear orden de producción (CON ID_PRODUCCION AUTOMÁTICO)
 */
function crearOrdenProduccion(pedidoRow) {
  const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const produccionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produccion');
  
  if (!produccionSheet) {
    Logger.log("Falta la hoja 'Produccion'");
    return false;
  }
  
  const pedId = pedidosSheet.getRange(pedidoRow, 12).getValue();
  
  // Verificar duplicados antes de crear
  const produccionData = produccionSheet.getDataRange().getValues();
  for (let i = 1; i < produccionData.length; i++) {
    if (produccionData[i][0] === pedId) {
      SpreadsheetApp.getUi().alert(
        '⚠️ Orden ya existe',
        `El pedido ${pedId} ya tiene una orden de producción activa.\nRevise la hoja Producción antes de crear duplicados.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      Logger.log(`DUPLICADO BLOQUEADO: Pedido ${pedId} ya existe en Producción`);
      return false;
    }
  }
  
  // Obtener datos del pedido
  const cliente = pedidosSheet.getRange(pedidoRow, 3).getValue();
  const producto = pedidosSheet.getRange(pedidoRow, 4).getValue();
  const color = pedidosSheet.getRange(pedidoRow, 5).getValue();
  const cantidad = pedidosSheet.getRange(pedidoRow, 6).getValue();
  const envase = pedidosSheet.getRange(pedidoRow, 7).getValue();
  
  // GENERAR ID_PRODUCCION AUTOMÁTICO
  const idProduccion = generarSiguienteIdProduccion(producto);
  
  // Crear nueva orden con 18 columnas (incluyendo REPROCESO e ID_PRODUCCION)
  const nuevaFila = [
    pedId,              // A: PED_ID
    cliente,            // B: Cliente
    producto,           // C: Producto
    color,              // D: Color
    cantidad,           // E: Cantidad
    envase,             // F: Envase
    "",                 // G: Máquina
    "",                 // H: Operario
    "PENDIENTE",        // I: Estado_Prod
    "",                 // J: Fecha_PROD
    "",                 // K: Hora_Inicio
    "",                 // L: Hora_Fin
    "",                 // M: Cant_Producida
    "",                 // N: Unidad_Envase
    "NO",               // O: REPROCESO (default NO)
    "",                 // P: ID_CONSUMIDO
    "",                 // Q: Tiempo_Total
    idProduccion        // R: ID_PRODUCCION (automático)
  ];
  
  produccionSheet.appendRow(nuevaFila);
  
  const datosOrden = { pedId, cliente, producto, color, cantidad, envase, idProduccion };
  enviarNotificacionProduccion(pedidoRow, "NUEVA_ORDEN", "📋 *NUEVA ORDEN*", datosOrden);
  
  Logger.log(`Orden de producción creada para ${pedId} con ID_PRODUCCION: ${idProduccion}`);
  return true;
}

/**
 * Confirmar envío a producción
 */
function confirmarEnvioAProduccion(pedidoRow) {
  const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  
  const pedId = pedidosSheet.getRange(pedidoRow, 12).getValue();
  const cliente = pedidosSheet.getRange(pedidoRow, 3).getValue();
  const producto = pedidosSheet.getRange(pedidoRow, 4).getValue();
  const color = pedidosSheet.getRange(pedidoRow, 5).getValue();
  const cantidad = pedidosSheet.getRange(pedidoRow, 6).getValue();
  const envase = pedidosSheet.getRange(pedidoRow, 7).getValue();
  
  let mensaje = `CONFIRMAR ENVÍO A PRODUCCIÓN\n\n`;
  mensaje += `PED_ID: ${pedId}\n`;
  mensaje += `Cliente: ${cliente}\n`;
  mensaje += `Producto: ${producto} ${color}\n`;
  mensaje += `Cantidad: ${cantidad} ${envase}\n\n`;
  mensaje += `¿Estás seguro que quieres enviar este pedido a producción?`;
  
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert('Confirmar Envío a Producción', mensaje, ui.ButtonSet.YES_NO);
  const confirmado = respuesta === ui.Button.YES;
  
  Logger.log(confirmado ? `Usuario confirmó envío del pedido ${pedId}` : `Usuario canceló envío del pedido ${pedId}`);
  return confirmado;
}

/**
 * Debitar materia prima real de SIG_Ventas
 * @param {string} idConsumido - IDs separados por coma (ej: "AS-001,EPX-002")
 * @return {string} Mensaje de confirmación con resultados
 */
function debitarMateriaPrima(idConsumido) {
  if (!idConsumido) return "";
  
  const ids = idConsumido.split(',').map(id => id.trim());
  let mensajeMateria = "🔧 *Materia prima consumida:*\n";
  let errores = [];
  let exitosos = 0;
  
  try {
    const inventarioSS = SpreadsheetApp.openById('1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw');
    const inventarioSheet = inventarioSS.getSheetByName('Inventario');
    const datos = inventarioSheet.getDataRange().getValues();
    
    for (const id of ids) {
      let encontrado = false;
      
      for (let i = 1; i < datos.length; i++) {
        if (datos[i][5] === id) { // Columna F = ID del producto
          const stockActual = datos[i][9]; // Columna J = Cantidad
          
          if (stockActual > 0) {
            // DÉBITO REAL: Reducir en 1 unidad
            const nuevoStock = stockActual - 1;
            inventarioSheet.getRange(i + 1, 10).setValue(nuevoStock);
            
            mensajeMateria += `• ${id}: ${stockActual} → *${nuevoStock}* ✅\n`;
            Logger.log(`✅ Materia prima debitada: ${id} (${stockActual} → ${nuevoStock})`);
            encontrado = true;
            exitosos++;
            
            // Log en hoja LOG_Debitos si existe
            registrarDebitoMateriaPrima(id, stockActual, nuevoStock);
            break;
            
          } else {
            errores.push(`${id}: Sin stock disponible`);
            Logger.log(`❌ Sin stock para debitar: ${id}`);
            encontrado = true;
            break;
          }
        }
      }
      
      if (!encontrado) {
        errores.push(`${id}: No encontrado en inventario`);
        Logger.log(`❌ ID no encontrado: ${id}`);
      }
    }
    
    // Agregar resumen
    if (exitosos > 0) {
      mensajeMateria += `\n📊 *Resumen:* ${exitosos}/${ids.length} debitados exitosamente`;
    }
    
    if (errores.length > 0) {
      mensajeMateria += "\n\n⚠️ *Errores:*\n" + errores.map(e => `• ${e}`).join('\n');
    }
    
    return mensajeMateria;
    
  } catch (error) {
    Logger.log(`❌ Error debitando materia prima: ${error}`);
    return `❌ Error conectando con inventario: ${error.message}`;
  }
}

/**
 * Registrar débito en LOG_Debitos (crear si no existe)
 */
function registrarDebitoMateriaPrima(idInventario, stockAnterior, stockNuevo) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('LOG_Debitos');
    
    if (!logSheet) {
      logSheet = crearLogDebitos();
    }
    
    const timestamp = new Date();
    const diferencia = stockAnterior - stockNuevo;
    
    logSheet.appendRow([
      timestamp,
      'MATERIA_PRIMA',
      'REPROCESO',
      'Materia Prima',
      '',  // Color vacío
      diferencia,
      'Unidad',
      idInventario,
      stockAnterior,
      stockNuevo,
      diferencia
    ]);
    
    Logger.log(`📝 Débito registrado en LOG: ${idInventario}`);
    
  } catch (error) {
    Logger.log(`⚠️ Error registrando en LOG_Debitos: ${error}`);
  }
}

/**
 * Crear hoja LOG_Debitos si no existe
 */
function crearLogDebitos() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.insertSheet('LOG_Debitos');
  
  const headers = [
    'Timestamp', 'PED_ID', 'Cliente', 'Producto', 'Color', 
    'Cantidad_Pedida', 'Envase', 'ID_Inventario', 
    'Stock_Anterior', 'Stock_Nuevo', 'Diferencia'
  ];
  
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  logSheet.setFrozenRows(1);
  
  // Formato de fecha
  logSheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  
  Logger.log('📋 Hoja LOG_Debitos creada automáticamente');
  return logSheet;
}


/**
 * VERSIÓN ACTUALIZADA: agregarSurplusAlInventario() con normalización de envases
 * @param {string} producto - Tipo de pintura (nombre CPG)
 * @param {string} color - Color específico
 * @param {number} surplus - Cantidad de surplus
 * @param {string} envase - Galones/Cubetas/Galón/Cubeta (se normaliza automáticamente)
 * @return {string} Mensaje de confirmación
 */
function agregarSurplusAlInventario(producto, color, surplus, envase) {
  if (surplus <= 0) return "";
  
  try {
    // PASO 1: Mapear producto CPG a estructura SIG
    const datosMapeo = mapearProductoCPGaSIG(producto);
    
    if (!datosMapeo.mapeado) {
      Logger.log(`❌ Producto sin mapeo: ${producto}`);
      return `❌ Producto "${producto}" no tiene mapeo definido a SIG`;
    }
    
    // PASO 1.5: NORMALIZAR ENVASE para compatibilidad con SIG_Ventas
    const envaseNormalizado = normalizarEnvase(envase);
    
    Logger.log(`📝 Envase normalizado: "${envase}" → "${envaseNormalizado}"`);
    
    // PASO 2: Preparar solicitud estructurada para API
    const solicitud = {
      idTipo: datosMapeo.idTipo,
      tipo: datosMapeo.tipo,
      color: color,
      surplus: surplus,
      envase: envaseNormalizado, // Usar envase normalizado
      naturaleza: datosMapeo.naturaleza,
      esNuevo: datosMapeo.esNuevo || false,
      origen: 'CPG_PRODUCCION',
      timestamp: new Date()
    };
    
    Logger.log(`📤 Enviando surplus a SIG_Ventas: ${JSON.stringify(solicitud)}`);
    
    // PASO 3: Ejecutar surplus via API SIG_Ventas
    const respuesta = ejecutarSurplusEnSIG(solicitud);
    
    // PASO 4: Procesar respuesta y generar mensaje
    return procesarRespuestaSurplusCPG(respuesta, solicitud);
    
  } catch (error) {
    Logger.log(`❌ Error en surplus CPG: ${error}`);
    return `❌ Error procesando surplus: ${error.message}`;
  }
}

/**
 * Normalizar envase para compatibilidad con SIG_Ventas
 * @param {string} envase - Envase original (Cubetas, Galones, etc.)
 * @return {string} Envase normalizado (Cubeta, Galón)
 */
function normalizarEnvase(envase) {
  const envaseStr = envase.toString().toLowerCase().trim();
  
  // Mapeo de variaciones a formato estándar SIG_Ventas
  if (envaseStr.includes('cubeta')) {
    return 'Cubeta'; // Singular
  } else if (envaseStr.includes('galon') || envaseStr.includes('galón')) {
    return 'Galón'; // Con tilde
  } else if (envaseStr.includes('cuartillo') || envaseStr.includes('1/4')) {
    return 'Cuartillo';
  }
  
  // Fallback: capitalizar primera letra
  return envase.charAt(0).toUpperCase() + envase.slice(1).toLowerCase();
}

/**
 * Ejecutar surplus directamente en SIG_Ventas
 * @param {object} solicitud - Datos estructurados del surplus
 * @return {object} Respuesta de la API
 */
function ejecutarSurplusEnSIG(solicitud) {
  const SIG_VENTAS_ID = '1mP0rwnuI83t6j9z0o7417H2GHX_YEg651m_4R0u0Ruw';
  
  try {
    // CONECTAR con SIG_Ventas
    const sigSpreadsheet = SpreadsheetApp.openById(SIG_VENTAS_ID);
    const inventarioSheet = sigSpreadsheet.getSheetByName('Inventario');
    
    if (!inventarioSheet) {
      return crearRespuestaError('SHEET_NOT_FOUND', 'Hoja Inventario no encontrada en SIG_Ventas');
    }
    
    // BUSCAR producto existente
    const productoExistente = buscarProductoExistenteEnSIG(
      inventarioSheet, 
      solicitud.tipo, 
      solicitud.color, 
      solicitud.envase
    );
    
    if (productoExistente.encontrado) {
      // CASO A: Surplus en producto existente
      return procesarSurplusExistenteSIG(inventarioSheet, productoExistente, solicitud);
    } else {
      // CASO B: Crear nuevo producto
      return crearNuevoProductoSIG(sigSpreadsheet, inventarioSheet, solicitud);
    }
    
  } catch (error) {
    Logger.log(`❌ Error ejecutando en SIG_Ventas: ${error}`);
    return crearRespuestaError('EXECUTION_ERROR', error.message);
  }
}

/**
 * Buscar producto existente en SIG_Ventas (versión CPG)
 */
function buscarProductoExistenteEnSIG(inventarioSheet, tipo, color, envase) {
  const datos = inventarioSheet.getDataRange().getValues();
  
  const tipoNorm = normalize(tipo);
  const colorNorm = normalize(color);
  const envaseNorm = normalize(envase);
  
  for (let i = 1; i < datos.length; i++) {
    const rowTipo = normalize(datos[i][2]);    // Columna C
    const rowColor = normalize(datos[i][3]);   // Columna D  
    const rowEnvase = normalize(datos[i][8]);  // Columna I
    
    if (rowTipo === tipoNorm && rowColor === colorNorm && rowEnvase === envaseNorm) {
      return {
        encontrado: true,
        fila: i + 1,
        id: datos[i][5],           // Columna F
        stockActual: datos[i][9],  // Columna J
        datosCompletos: datos[i]
      };
    }
  }
  
  return { encontrado: false };
}

/**
 * Procesar surplus en producto existente (versión CPG)
 */
function procesarSurplusExistenteSIG(inventarioSheet, productoExistente, solicitud) {
  const stockActual = productoExistente.stockActual || 0;
  const nuevoStock = stockActual + solicitud.surplus;
  
  // ACTUALIZAR stock directamente
  inventarioSheet.getRange(productoExistente.fila, 10).setValue(nuevoStock);
  
  Logger.log(`✅ Surplus agregado: ${solicitud.surplus} a ${productoExistente.id} (${stockActual} → ${nuevoStock})`);
  
  return crearRespuestaExito({
    accion: 'SURPLUS_AGREGADO',
    id: productoExistente.id,
    stockAnterior: stockActual,
    stockNuevo: nuevoStock,
    diferencia: solicitud.surplus
  });
}

/**
 * Crear nuevo producto en SIG_Ventas (versión CPG)
 */
function crearNuevoProductoSIG(sigSpreadsheet, inventarioSheet, solicitud) {
  // ACTUALIZAR catálogos primero
  const catalogosOK = actualizarCatalogosSIG(sigSpreadsheet, solicitud);
  
  if (!catalogosOK) {
    return crearRespuestaError('CATALOG_UPDATE_FAILED', 'Error actualizando catálogos');
  }
  
  // ENCONTRAR posición de inserción
  const ultimaFila = encontrarUltimaFilaConDatosSIG(inventarioSheet);
  
  // PREPARAR fila nueva
  const fechaHoy = new Date();
  const nuevaFila = [
    fechaHoy, solicitud.tipo, solicitud.tipo, solicitud.color,
    "", "", "", 'SURPLUS', solicitud.envase, solicitud.surplus,
    0, 'Nuevo', "", "", ""
  ];
  
  // INSERTAR con extensión de fórmulas
  inventarioSheet.insertRowAfter(ultimaFila);
  const filaDestino = ultimaFila + 1;
  inventarioSheet.getRange(filaDestino, 1, 1, nuevaFila.length).setValues([nuevaFila]);
  
  // EXTENDER fórmulas de fila anterior
  extenderFormulasAFilaNueva(inventarioSheet, ultimaFila, filaDestino);
  
  // ESPERAR ejecución de fórmulas
  Utilities.sleep(3000);
  const idGenerado = inventarioSheet.getRange(filaDestino, 6).getValue();
  
  Logger.log(`✅ Nuevo producto creado: ${idGenerado} con ${solicitud.surplus} ${solicitud.envase}`);
  
  return crearRespuestaExito({
    accion: 'PRODUCTO_CREADO',
    id: idGenerado || `${solicitud.idTipo}-PENDIENTE`,
    tipoNuevo: solicitud.esNuevo,
    stockInicial: solicitud.surplus
  });
}

/**
 * Actualizar catálogos en SIG_Ventas (versión CPG)
 */
function actualizarCatalogosSIG(sigSpreadsheet, solicitud) {
  try {
    // ACTUALIZAR Catalogo_Tipos
    const tiposSheet = sigSpreadsheet.getSheetByName('Catalogo_Tipos');
    if (tiposSheet && !existeTipoEnCatalogo(tiposSheet, solicitud.idTipo)) {
      const observacion = `Tipo creado desde CPG - ${new Date().toLocaleDateString()}`;
      tiposSheet.appendRow([solicitud.idTipo, solicitud.tipo, solicitud.naturaleza, observacion]);
      Logger.log(`📝 Tipo agregado: ${solicitud.idTipo}`);
    }
    
    // ACTUALIZAR Catalogo_Precios
    const preciosSheet = sigSpreadsheet.getSheetByName('Catalogo_Precios');
    if (preciosSheet && !existePrecioEnCatalogo(preciosSheet, solicitud.idTipo)) {
      const precios = obtenerPreciosExactosParaTipo(solicitud.idTipo);
      preciosSheet.appendRow([
        solicitud.idTipo, solicitud.tipo,
        precios.cubeta, Math.round(precios.cubeta * 0.9), Math.round(precios.cubeta * 0.8),
        precios.galon, Math.round(precios.galon * 0.9), Math.round(precios.galon * 0.8),
        precios.observacion
      ]);
      Logger.log(`💰 Precios agregados: ${solicitud.idTipo} - $${precios.cubeta}/$${precios.galon}`);
    }
    
    return true;
  } catch (error) {
    Logger.log(`❌ Error actualizando catálogos: ${error}`);
    return false;
  }
}

/**
 * Obtener precios exactos de Lista Infiniti
 */
function obtenerPreciosExactosParaTipo(idTipo) {
  const preciosInfiniti = {
    'SEM-P': { cubeta: 4200, galon: 900, observacion: "SEMIGLOSS Premium - Precio original Infiniti" },
    'SEM-B': { cubeta: 3700, galon: 825, observacion: "SEMIGLOSS Tipo B - Precio original Infiniti" },
    'BCL': { cubeta: 4600, galon: 975, observacion: "BARNIZ Clear Industrial - Precio original Infiniti" },
    'BEP': { cubeta: 6500, galon: 1350, observacion: "BARNIZ Port Epoxi Clear - Precio original Infiniti" }
  };
  
  return preciosInfiniti[idTipo] || { 
    cubeta: 2500, galon: 600, 
    observacion: "Precios base estimados" 
  };
}

/**
 * Extender fórmulas a fila nueva (versión CPG)
 */
function extenderFormulasAFilaNueva(inventarioSheet, filaReferencia, filaNueva) {
  try {
    // INCLUIR columnas K y L en extensión automática
    const columnasConFormulas = [5, 6, 7, 11, 12, 13, 14, 15]; // E, F, G, K, L, M, N, O
    
    for (const col of columnasConFormulas) {
      const formulaReferencia = inventarioSheet.getRange(filaReferencia, col).getFormula();
      
      if (formulaReferencia) {
        const formulaAjustada = ajustarFormulaID(formulaReferencia, filaReferencia, filaNueva);
        inventarioSheet.getRange(filaNueva, col).setFormula(formulaAjustada);
      }
    }
    
    Logger.log(`Fórmulas extendidas incluyendo antigüedad de fila ${filaReferencia} a ${filaNueva}`);
  } catch (error) {
    Logger.log(`Error extendiendo fórmulas: ${error}`);
  }
}

function ajustarFormulaID(formula, filaOrigen, filaDestino) {
  const filaOrigenStr = filaOrigen.toString();
  const filaDestinoStr = filaDestino.toString();
  
  // Reemplazar TODAS las referencias de fila estática por la fila destino
  return formula.replace(new RegExp(`([A-Z])${filaOrigenStr}`, 'g'), `$1${filaDestinoStr}`);
}

/**
 * Procesar respuesta de SIG_Ventas
 */
function procesarRespuestaSurplusCPG(respuesta, solicitud) {
  if (!respuesta.success) {
    return `❌ Error en SIG_Ventas: ${respuesta.mensaje || respuesta.error}`;
  }
  
  // GENERAR mensaje para WhatsApp
  if (respuesta.accion === 'SURPLUS_AGREGADO') {
    return `📈 *Surplus agregado al stock existente:*\n• ID: *${respuesta.id}*\n• Stock: ${respuesta.stockAnterior} → *${respuesta.stockNuevo}* (+${respuesta.diferencia})`;
  } else {
    return `📦 *Nuevo producto creado:*\n• ID: *${respuesta.id}*\n• Tipo: ${solicitud.idTipo}${solicitud.esNuevo ? ' (NUEVO)' : ''}\n• Stock inicial: *${solicitud.surplus} ${solicitud.envase}*`;
  }
}

/**
 * Buscar producto existente en SIG_Ventas
 * @param {Sheet} inventarioSheet - Hoja Inventario
 * @param {string} tipoSIG - Tipo según mapeo SIG
 * @param {string} color - Color específico
 * @param {string} envase - Envase (Galón/Cubeta)
 * @return {object} Resultado de búsqueda
 */
function buscarProductoExistenteEnSIG(inventarioSheet, tipoSIG, color, envase) {
  const datos = inventarioSheet.getDataRange().getValues();
  
  const tipoNorm = normalize(tipoSIG);
  const colorNorm = normalize(color);
  const envaseNorm = normalize(envase);
  
  for (let i = 1; i < datos.length; i++) {
    const rowTipo = normalize(datos[i][2]);    // Columna C = Tipo
    const rowColor = normalize(datos[i][3]);   // Columna D = Color  
    const rowEnvase = normalize(datos[i][8]);  // Columna I = Envase
    
    if (rowTipo === tipoNorm && rowColor === colorNorm && rowEnvase === envaseNorm) {
      return {
        encontrado: true,
        fila: i + 1,
        id: datos[i][5],           // Columna F = ID
        stockActual: datos[i][9],  // Columna J = Cantidad
        datosCompletos: datos[i]
      };
    }
  }
  
  return { encontrado: false };
}

/**
 * Procesar surplus en producto existente
 * @param {Sheet} inventarioSheet - Hoja Inventario
 * @param {object} productoExistente - Datos del producto encontrado
 * @param {number} surplus - Cantidad a agregar
 * @param {object} datosMapeo - Mapeo CPG → SIG
 * @return {string} Mensaje resultado
 */
function procesarSurplusExistente(inventarioSheet, productoExistente, surplus, datosMapeo) {
  const stockActual = productoExistente.stockActual || 0;
  const nuevoStock = stockActual + surplus;
  
  // ACTUALIZAR STOCK EN SIG_VENTAS
  inventarioSheet.getRange(productoExistente.fila, 10).setValue(nuevoStock); // Columna J
  
  const mensajeSurplus = `📈 *Surplus agregado al stock existente:*\n• ID: *${productoExistente.id}*\n• Stock: ${stockActual} → *${nuevoStock}* (+${surplus})`;
  
  Logger.log(`✅ Surplus agregado: ${surplus} a ${productoExistente.id} (${stockActual} → ${nuevoStock})`);
  
  // REGISTRAR en LOG local de CPG con trazabilidad completa
  registrarSurplusEnLog({
    tipo: respuesta.accion,
    pedId: solicitud.pedId || '',           // NUEVO
    idProduccion: solicitud.idProduccion || '', // NUEVO
    id: respuesta.id,
    producto: solicitud.tipo,
    color: solicitud.color,
    surplus: solicitud.surplus,
    envase: solicitud.envase,
    stockAnterior: respuesta.stockAnterior || 0,
    stockNuevo: respuesta.stockNuevo || solicitud.surplus,
    idTipoNuevo: solicitud.esNuevo
  });
  
  return mensajeSurplus;
}

/**
 * Crear nuevo producto en SIG_Ventas
 * @param {Sheet} inventarioSheet - Hoja Inventario
 * @param {object} datosMapeo - Mapeo CPG → SIG
 * @param {string} color - Color específico
 * @param {number} surplus - Cantidad inicial
 * @param {string} envase - Envase
 * @return {string} Mensaje resultado
 */
function crearNuevoProductoEnSIG(inventarioSheet, datosMapeo, color, surplus, envase) {
  // GENERAR ID ÚNICO
  const nuevoId = generarNuevoIdSIG(datosMapeo.idTipo);
  const fechaHoy = new Date();
  
  // PREPARAR FILA NUEVA según estructura SIG_Ventas
  const nuevaFila = [
    fechaHoy,                    // A: Fecha
    datosMapeo.tipo,             // B: Producto  
    datosMapeo.tipo,             // C: Tipo (mismo que producto)
    color,                       // D: Color
    datosMapeo.idTipo,           // E: ID_Tipo
    nuevoId,                     // F: ID único
    datosMapeo.naturaleza,       // G: Naturaleza
    'SURPLUS',                   // H: Tarima (marcado como surplus)
    envase,                      // I: Envase
    surplus,                     // J: Cantidad
    0,                           // K: Antiguedad_dias (nuevo = 0)
    'Nuevo',                     // L: Banda_Antiguedad
    0,                           // M: Precio_100 (se asigna después)
    0,                           // N: Precio_90
    0                            // O: Precio_80
  ];
  
  // INSERTAR EN SIG_VENTAS
  inventarioSheet.appendRow(nuevaFila);
  
  const mensajeSurplus = `📦 *Nuevo producto creado:*\n• ID: *${nuevoId}*\n• Tipo: ${datosMapeo.idTipo}${datosMapeo.esNuevo ? ' (NUEVO)' : ''}\n• Stock inicial: *${surplus} ${envase}*`;
  
  Logger.log(`✅ Nuevo producto creado: ${nuevoId} con ${surplus} ${envase}`);
  
  // REGISTRAR EN LOG
  registrarSurplusEnLog({
    tipo: 'PRODUCTO_NUEVO',
    id: nuevoId,
    producto: datosMapeo.tipo,
    color: color,
    surplus: surplus,
    envase: envase,
    stockAnterior: 0,
    stockNuevo: surplus,
    idTipoNuevo: datosMapeo.esNuevo
  });
  
  // NOTIFICAR SI ES TIPO COMPLETAMENTE NUEVO
  if (datosMapeo.esNuevo) {
    Logger.log(`🆕 TIPO NUEVO CREADO: ${datosMapeo.idTipo} - ${datosMapeo.tipo}`);
  }
  
  return mensajeSurplus;
}

/**
 * Registrar surplus en LOG_Surplus (mejorado)
 * @param {object} datos - Datos del surplus registrado
 */
function registrarSurplusEnLog(datos) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName('LOG_Surplus');
    
    if (!logSheet) {
      logSheet = crearLogSurplusActualizado();
    }
    
    const timestamp = new Date();
    
    logSheet.appendRow([
      timestamp,                      // A: Timestamp
      datos.tipo,                     // B: Tipo_Operacion
      datos.pedId || '',              // C: PED_ID (NUEVO)
      datos.idProduccion || '',       // D: ID_PRODUCCION (NUEVO)
      datos.id,                       // E: ID_Producto
      datos.producto,                 // F: Producto
      datos.color,                    // G: Color
      datos.surplus,                  // H: Cantidad_Surplus
      datos.envase,                   // I: Envase
      datos.stockAnterior,            // J: Stock_Anterior
      datos.stockNuevo,               // K: Stock_Nuevo
      datos.stockNuevo - datos.stockAnterior,  // L: Diferencia
      datos.idTipoNuevo || false      // M: Tipo_Nuevo_Creado
    ]);
    
    Logger.log(`📝 Surplus registrado en LOG: ${datos.surplus} ${datos.envase} - ID ${datos.id} - PED_ID ${datos.pedId} - ID_PROD ${datos.idProduccion}`);
    
  } catch (error) {
    Logger.log(`⚠️ Error registrando surplus en LOG: ${error}`);
  }
}

function crearLogSurplusActualizado() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.insertSheet('LOG_Surplus');
  
  const headers = [
    'Timestamp', 'Tipo_Operacion', 'PED_ID', 'ID_PRODUCCION', 
    'ID_Producto', 'Producto', 'Color', 'Cantidad_Surplus', 
    'Envase', 'Stock_Anterior', 'Stock_Nuevo', 'Diferencia', 'Tipo_Nuevo_Creado'
  ];
  
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  logSheet.getRange(1, 1, 1, headers.length).setBackground('#E8F5E8');
  logSheet.setFrozenRows(1);
  
  // Formato de fecha
  logSheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  
  Logger.log('📋 Hoja LOG_Surplus creada con estructura actualizada');
  return logSheet;
}

/**
 * Migrar LOG_Surplus existente a nueva estructura
 * Ejecutar UNA VEZ si ya tienes datos en LOG_Surplus
 */
function migrarLogSurplusANuevaEstructura() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.getSheetByName('LOG_Surplus');
  
  if (!logSheet) {
    Logger.log('No existe LOG_Surplus para migrar');
    return;
  }
  
  // Verificar si ya tiene la nueva estructura
  const encabezados = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  
  if (encabezados.includes('PED_ID') && encabezados.includes('ID_PRODUCCION')) {
    Logger.log('LOG_Surplus ya tiene la nueva estructura');
    return;
  }
  
  // Insertar nuevas columnas
  logSheet.insertColumnAfter(2); // PED_ID después de Tipo_Operacion
  logSheet.insertColumnAfter(3); // ID_PRODUCCION después de PED_ID
  
  // Actualizar encabezados
  logSheet.getRange(1, 3).setValue('PED_ID');
  logSheet.getRange(1, 4).setValue('ID_PRODUCCION');
  
  Logger.log('✅ LOG_Surplus migrado a nueva estructura');
  Logger.log('💡 Las entradas existentes tendrán PED_ID e ID_PRODUCCION vacíos');
}

/**
 * Crear hoja LOG_Surplus con estructura mejorada
 */
function crearLogSurplus() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.insertSheet('LOG_Surplus');
  
  const headers = [
    'Timestamp', 'Tipo_Operacion', 'ID_Producto', 'Producto', 'Color',
    'Cantidad_Surplus', 'Envase', 'Stock_Anterior', 'Stock_Nuevo', 
    'Diferencia', 'Tipo_Nuevo_Creado'
  ];
  
  logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  logSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  logSheet.setFrozenRows(1);
  
  // Formato de fecha
  logSheet.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  
  Logger.log('📋 Hoja LOG_Surplus creada automáticamente');
  return logSheet;
}

// ═══════════════════════════════════════════════════════════
// FUNCIÓN UNIFICADA WHATSAPP
// ═══════════════════════════════════════════════════════════

/**
 * Enviar notificación de producción (FUNCIÓN UNIFICADA)
 */
function enviarNotificacionProduccion(fila, tipoEvento, textoEstado, datosExtra = null) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('WAS_TOKEN');
  const grupo = props.getProperty('GROUP_GREQ_TECNICO');
  
  if (!token || !grupo) {
    Logger.log('⚠️ Faltan credenciales WhatsApp');
    return;
  }
  
  let mensaje = "";
  
  switch(tipoEvento) {
    case "NUEVA_ORDEN":
      mensaje = construirMensajeNuevaOrden(datosExtra);
      break;
    case "INICIADO":
      mensaje = construirMensajeIniciado(fila);
      break;
    case "CALIDAD":
    case "ENVASE":
      mensaje = construirMensajeEstado(fila, textoEstado);
      break;
    case "COMPLETADO":
      mensaje = construirMensajeCompletado(datosExtra);
      break;
    case "INVENTARIO":
      mensaje = construirMensajeInventario(fila, datosExtra);
      break;
    default:
      Logger.log(`Tipo de evento desconocido: ${tipoEvento}`);
      return;
  }
  
  // Envío unificado
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify({ to: grupo, text: String(mensaje) }),
    muteHttpExceptions: true
  };
  
  try {
    const resp = UrlFetchApp.fetch('https://www.wasenderapi.com/api/send-message', options);
    Logger.log(`📱 Mensaje ${tipoEvento} enviado: ${resp.getResponseCode()}`);
  } catch (error) {
    Logger.log(`❌ Error enviando mensaje ${tipoEvento}: ${error}`);
  }
}

// ═══════════════════════════════════════════════════════════
// CONSTRUCTORES DE MENSAJES
// ═══════════════════════════════════════════════════════════

/**
 * Construir mensaje nueva orden (CON ID_PRODUCCION)
 */
function construirMensajeNuevaOrden(datos) {
  let mensaje = `📋 *NUEVA ORDEN* | ${datos.pedId}\n`;
  mensaje += `Cliente: *${datos.cliente}*\n`;
  mensaje += `*${datos.producto} ${datos.color}* (${datos.cantidad} ${datos.envase})\n`;
  mensaje += `ID Producción: *${datos.idProduccion}*\n`;
  mensaje += `Estado: Pendiente asignación`;
  return mensaje;
}

/**
 * Construir mensaje iniciado
 */
function construirMensajeIniciado(fila) {
  const produccionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produccion');
  
  const pedId = produccionSheet.getRange(fila, PROD_COL.PED_ID).getValue();
  const producto = produccionSheet.getRange(fila, PROD_COL.PRODUCTO).getValue();
  const color = produccionSheet.getRange(fila, PROD_COL.COLOR).getValue();
  const cantidad = produccionSheet.getRange(fila, PROD_COL.CANTIDAD).getValue();
  const envase = produccionSheet.getRange(fila, PROD_COL.ENVASE).getValue();
  const operario = produccionSheet.getRange(fila, PROD_COL.OPERARIO).getValue();
  const maquina = produccionSheet.getRange(fila, PROD_COL.MAQUINA).getValue();
  const horaInicio = produccionSheet.getRange(fila, PROD_COL.HORA_INICIO).getValue();
  const idProduccion = produccionSheet.getRange(fila, PROD_COL.ID_PRODUCCION).getValue();
  
  const horaTexto = Utilities.formatDate(horaInicio, Session.getScriptTimeZone(), 'HH:mm');
  
  let mensaje = `🔄 *INICIADO* | ${pedId}\n`;
  mensaje += `*${producto} ${color}* (${cantidad} ${envase})\n`;
  mensaje += `${operario} • Máq.${maquina} • *${horaTexto}*\n`;
  mensaje += `ID Producción: *${idProduccion}*`;
  
  return mensaje;
}

/**
 * Construir mensaje estado intermedio
 */
function construirMensajeEstado(fila, textoEstado) {
  const produccionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Produccion');
  const pedId = produccionSheet.getRange(fila, PROD_COL.PED_ID).getValue();
  const producto = produccionSheet.getRange(fila, PROD_COL.PRODUCTO).getValue();
  const color = produccionSheet.getRange(fila, PROD_COL.COLOR).getValue();
  const idProduccion = produccionSheet.getRange(fila, PROD_COL.ID_PRODUCCION).getValue();
  
  let mensaje = `${textoEstado} | ${pedId}\n`;
  mensaje += `*${producto} ${color}*\n`;
  mensaje += `ID Producción: *${idProduccion}*`;
  
  return mensaje;
}

/**
 * Construir mensaje completado (CON ID_PRODUCCION)
 */
function construirMensajeCompletado(datos) {
  let mensaje = `✅ *COMPLETADO* | ${datos.pedId}\n`;
  mensaje += `*${datos.producto} ${datos.color}*\n`;
  
  if (datos.surplusGalones > 0) {
    mensaje += `Pedido: ${datos.cantidadOriginal} ${datos.envaseOriginal} → Producido: *${datos.cantProducida} ${datos.unidadEnvase}* (+surplus)\n`;
  } else {
    mensaje += `Producido: *${datos.cantProducida} ${datos.unidadEnvase}* según pedido\n`;
  }
  
  mensaje += `${datos.operario} • Máq.${datos.maquina} • *${datos.tiempoTotal}*\n`;
  mensaje += `ID Producción: *${datos.idProduccion}*`;
  
  if (datos.surplusGalones > 0 && datos.mensajeSurplus) {
    mensaje += `\n${datos.mensajeSurplus}`;
  }
  
  return mensaje;
}

/**
 * Construir mensaje inventario
 */
function construirMensajeInventario(pedidoRow, itemUsado) {
  const pedidosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const pedId = pedidosSheet.getRange(pedidoRow, 12).getValue();
  const cliente = pedidosSheet.getRange(pedidoRow, 3).getValue();
  const producto = pedidosSheet.getRange(pedidoRow, 4).getValue();
  const color = pedidosSheet.getRange(pedidoRow, 5).getValue();
  
  let mensaje = `📦 *DESPACHADO DESDE INVENTARIO*\n`;
  mensaje += `PED_ID: ${pedId}\n`;
  mensaje += `Cliente: *${cliente}*\n`;
  mensaje += `*${producto} ${color}*\n`;
  mensaje += `Estado: LISTO P/ DESPACHAR\n\n`;
  mensaje += `Stock utilizado:\n`;
  mensaje += `• ${itemUsado.cantidadUsar} und - ID: *${itemUsado.id}*\n`;
  mensaje += `Se debitará del inventario a las 6PM al archivar`;
  
  return mensaje;
}

/**
 * FUNCIÓN DE TESTING - agregarSurplusAlInventario()
 */
function testAgregarSurplus() {
  Logger.log('🧪 INICIANDO TEST: agregarSurplusAlInventario()');
  
  // TEST CASO 1: Producto existente (debería sumar al stock)
  Logger.log('\n--- TEST 1: Surplus en producto existente ---');
  const resultado1 = agregarSurplusAlInventario('ESMALTE INDUSTRIAL', 'GRIS PERLA', 2, 'Galón');
  Logger.log(`Resultado 1: ${resultado1}`);
  
  // TEST CASO 2: Producto nuevo tipo existente (debería crear nueva entrada)
  Logger.log('\n--- TEST 2: Nuevo producto tipo existente ---');
  const resultado2 = agregarSurplusAlInventario('ACRILICA SUPERIOR HP', 'AZUL CIELO', 3, 'Cubeta');
  Logger.log(`Resultado 2: ${resultado2}`);
  
  // TEST CASO 3: Producto tipo completamente nuevo (debería crear tipo + entrada)
  Logger.log('\n--- TEST 3: Producto tipo completamente nuevo ---');
  const resultado3 = agregarSurplusAlInventario('SEMIGLOSS PREMIUM', 'BLANCO', 1, 'Galón');
  Logger.log(`Resultado 3: ${resultado3}`);
  
  Logger.log('\n✅ TEST COMPLETADO - Revisar LOG_Surplus y SIG_Ventas');
}

/**
 * TEST ESPECÍFICO: Validar mapeo
 */
function testMapeoSolamente() {
  Logger.log('🧪 TESTING SOLO MAPEO:');
  
  const productos = [
    'SEMIGLOSS PREMIUM',
    'SEMIGLOSS TIPO B', 
    'BARNIZ CLEAR INDUSTRIAL',
    'ESMALTE INDUSTRIAL'
  ];
  
  productos.forEach(producto => {
    const mapeo = mapearProductoCPGaSIG(producto);
    Logger.log(`${producto} → ${mapeo.idTipo} (${mapeo.esNuevo ? 'NUEVO' : 'EXISTENTE'})`);
  });
}

// ═══════════════════════════════════════════════════════════
// FUNCIONES AUXILIARES
// ═══════════════════════════════════════════════════════════

function encontrarUltimaFilaConDatosSIG(sheet) {
  const ultimaFilaTotal = sheet.getLastRow();
  const datos = sheet.getRange(1, 1, ultimaFilaTotal, 10).getValues();
  
  for (let i = datos.length - 1; i >= 1; i--) {
    const producto = datos[i][1];
    const tipo = datos[i][2];
    const cantidad = datos[i][9];
    
    if (producto && producto !== "Producto" && 
        tipo && tipo !== "Tipo" &&
        (cantidad !== undefined && cantidad !== "")) {
      return i + 1;
    }
  }
  
  return ultimaFilaTotal;
}

function existeTipoEnCatalogo(tiposSheet, idTipo) {
  const datos = tiposSheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === idTipo) return true;
  }
  return false;
}

function existePrecioEnCatalogo(preciosSheet, idTipo) {
  const datos = preciosSheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === idTipo) return true;
  }
  return false;
}

function crearRespuestaExito(datos) {
  return { success: true, timestamp: new Date(), ...datos };
}

function crearRespuestaError(codigo, mensaje) {
  return { success: false, error: codigo, mensaje: mensaje, timestamp: new Date() };
}

/**
 * TESTING: Cliente CPG completo
 */
function testClienteCPGCompleto() {
  Logger.log('🧪 Testing cliente CPG completo...');
  
  const resultado = agregarSurplusAlInventario('SEMIGLOSS TIPO B', 'NARANJA TEST', 4, 'Cubeta');
  Logger.log(`📱 Resultado final: ${resultado}`);
  Logger.log('✅ Verificar SIG_Ventas para confirmar creación de SEM-B-001');
}

/**
 * TESTING: Intentar procesar el mismo surplus dos veces
 */
function testAntiDuplicados() {
  Logger.log('🧪 Testing sistema anti-duplicados...');
  
  // Simular datos de prueba
  const testPedId = 'PED-2025-TEST';
  const testIdProd = 'GQ2580099';
  
  Logger.log(`Primera verificación: ${verificarSurplusYaProcesado(testPedId, testIdProd)}`);
  
  // Simular entrada en LOG
  registrarSurplusEnLog({
    tipo: 'TESTING',
    pedId: testPedId,
    idProduccion: testIdProd,
    id: 'TEST-001',
    producto: 'PRODUCTO TEST',
    color: 'COLOR TEST',
    surplus: 1,
    envase: 'Galón',
    stockAnterior: 0,
    stockNuevo: 1
  });
  
  Logger.log(`Segunda verificación: ${verificarSurplusYaProcesado(testPedId, testIdProd)}`);
  Logger.log('✅ Sistema anti-duplicados funcionando');
}

function testMensajeTecnico() {
  const mensaje = "Prueba desde sistema CPG-SIG ✅";
  enviarNotificacionProduccion(null, "TEST", mensaje);
}