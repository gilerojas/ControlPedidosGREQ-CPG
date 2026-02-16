/**
 * ═══════════════════════════════════════════════════════════
 * MIGRACIÓN PED_ID: 3 DÍGITOS (2025) → 4 DÍGITOS (2026+)
 * ═══════════════════════════════════════════════════════════
 */

/**
 * PASO 1: Actualizar función getNextPedId() para usar 4 dígitos desde 2026
 * REEMPLAZAR la función existente en Code.gs
 */
function getNextPedId() {
  const props = PropertiesService.getDocumentProperties();
  const tz = Session.getScriptTimeZone();
  const year = Utilities.formatDate(new Date(), tz, 'yyyy');

  let lastYear = props.getProperty('PED_YEAR');
  let seq = Number(props.getProperty('NEXT_PED_SEQ')) || 1;

  // Reset anual
  if (lastYear !== year) {
    seq = 1;
    props.setProperty('PED_YEAR', year);
  }

  // Determinar padding según el año
  let padding;
  if (parseInt(year) <= 2025) {
    padding = 3; // 2025 y anteriores: PED-2025-001
  } else {
    padding = 4; // 2026+: PED-2026-0001
  }

  const id = `PED-${year}-${('0'.repeat(padding) + seq).slice(-padding)}`;
  props.setProperty('NEXT_PED_SEQ', seq + 1);
  props.setProperty('PED_YEAR', year);
  
  return id;
}

/**
 * PASO 2: Resincronizar solo pedidos de 2026
 * Reorganiza TODOS los pedidos de 2026 (Pedidos + Despachados) cronológicamente
 */
function resincronizarPedId2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const shPed = ss.getSheetByName("Pedidos");
  const shDes = ss.getSheetByName("Despachados");

  if (!shPed || !shDes) {
    ui.alert("❌ Falta la hoja 'Pedidos' o 'Despachados'");
    return;
  }

  // Confirmar acción
  const respuesta = ui.alert(
    '⚠️ RESINCRONIZAR PED_ID 2026',
    'Esto hará lo siguiente:\n\n' +
    '1. Buscar todos los pedidos de 2026 en Pedidos y Despachados\n' +
    '2. Renumerarlos cronológicamente: PED-2026-0001, PED-2026-0002...\n' +
    '3. Actualizar el contador global\n\n' +
    '⚠️ Los PED_ID de 2025 NO se tocarán\n\n' +
    '¿Desea continuar?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) {
    ui.alert('❌ Operación cancelada');
    return;
  }

  Logger.log('🚀 Iniciando resincronización de PED_ID 2026...\n');

  const dataPed = shPed.getDataRange().getValues();
  const dataDes = shDes.getDataRange().getValues();

  let registros2026 = [];

  // Recorrer Pedidos
  for (let i = 1; i < dataPed.length; i++) {
    const row = dataPed[i];
    const pedId = row[11]; // Columna L
    const fecha = row[12] || row[0]; // M = Último cambio o A = Fecha_pedido

    // Solo procesar pedidos de 2026
    if (fecha && new Date(fecha).getFullYear() === 2026) {
      registros2026.push({
        origen: "Pedidos",
        row: i + 1,
        fecha: fecha,
        pedIdViejo: pedId
      });
    }
  }

  // Recorrer Despachados
  for (let i = 1; i < dataDes.length; i++) {
    const row = dataDes[i];
    const pedId = row[11]; // Columna L
    const fecha = row[12] || row[15] || row[0]; // M = Último cambio, P = Fecha_archivo, A = Fecha_pedido

    // Solo procesar pedidos de 2026
    if (fecha && new Date(fecha).getFullYear() === 2026) {
      registros2026.push({
        origen: "Despachados",
        row: i + 1,
        fecha: fecha,
        pedIdViejo: pedId
      });
    }
  }

  if (registros2026.length === 0) {
    Logger.log('ℹ️ No se encontraron pedidos de 2026 para resincronizar');
    ui.alert('ℹ️ Sin Cambios', 'No se encontraron pedidos de 2026', ui.ButtonSet.OK);
    return;
  }

  Logger.log(`📊 Encontrados ${registros2026.length} pedidos de 2026\n`);

  // Filtrar registros sin fecha válida
  registros2026 = registros2026.filter(r => r.fecha && !isNaN(new Date(r.fecha).getTime()));

  // Ordenar por fecha
  registros2026.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));

  // Reasignar PED_ID con 4 dígitos
  let seq = 1;
  registros2026.forEach(reg => {
    reg.nuevoId = `PED-2026-${('0000' + seq).slice(-4)}`;
    seq++;
  });

  // Escribir nuevos PED_ID en las hojas
  registros2026.forEach(reg => {
    const sh = (reg.origen === "Pedidos") ? shPed : shDes;
    sh.getRange(reg.row, 12).setValue(reg.nuevoId); // Columna L = PED_ID
    Logger.log(`✅ ${reg.pedIdViejo} → ${reg.nuevoId} (${reg.origen})`);
  });

  // Actualizar contador global
  const props = PropertiesService.getDocumentProperties();
  props.setProperty("PED_YEAR", "2026");
  props.setProperty("NEXT_PED_SEQ", String(seq));

  const mensaje = `✅ Resincronización completada:\n\n` +
    `• Total pedidos 2026 procesados: ${registros2026.length}\n` +
    `• Último PED_ID asignado: ${registros2026[registros2026.length - 1].nuevoId}\n` +
    `• Próximo PED_ID disponible: PED-2026-${('0000' + seq).slice(-4)}\n\n` +
    `⚠️ Los PED_ID de 2025 permanecen sin cambios`;

  Logger.log('\n' + mensaje);
  ui.alert('✅ Resincronización Completada', mensaje, ui.ButtonSet.OK);
}

/**
 * PASO 3: Función de diagnóstico mejorada
 */
function diagnosticarPedId2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPed = ss.getSheetByName("Pedidos");
  const shDes = ss.getSheetByName("Despachados");
  const props = PropertiesService.getDocumentProperties();

  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('🔍 DIAGNÓSTICO PED_ID 2026');
  Logger.log('═══════════════════════════════════════════════════════\n');

  // Estado de Properties
  const year = props.getProperty('PED_YEAR');
  const seq = props.getProperty('NEXT_PED_SEQ');
  Logger.log(`📅 Año actual registrado: ${year}`);
  Logger.log(`🔢 Próximo secuencial: ${seq}`);
  
  if (parseInt(year) <= 2025) {
    Logger.log(`🆔 Próximo PED_ID: PED-${year}-${('000' + seq).slice(-3)} (3 dígitos)`);
  } else {
    Logger.log(`🆔 Próximo PED_ID: PED-${year}-${('0000' + seq).slice(-4)} (4 dígitos)`);
  }

  // Análisis de pedidos 2026
  const analizar = (sheet, nombre) => {
    const data = sheet.getDataRange().getValues();
    const pedidos2026 = [];
    
    for (let i = 1; i < data.length; i++) {
      const fecha = data[i][12] || data[i][15] || data[i][0];
      if (fecha && new Date(fecha).getFullYear() === 2026) {
        pedidos2026.push({
          pedId: data[i][11],
          fecha: fecha
        });
      }
    }
    
    return pedidos2026;
  };

  const pedidos2026 = analizar(shPed, "Pedidos");
  const despachados2026 = analizar(shDes, "Despachados");
  const total2026 = [...pedidos2026, ...despachados2026];

  Logger.log(`\n📋 PEDIDOS 2026:`);
  Logger.log(`   En Pedidos: ${pedidos2026.length}`);
  Logger.log(`   En Despachados: ${despachados2026.length}`);
  Logger.log(`   Total: ${total2026.length}`);

  // Detectar formato de dígitos
  const con3Digitos = total2026.filter(p => /^PED-2026-\d{3}$/.test(p.pedId)).length;
  const con4Digitos = total2026.filter(p => /^PED-2026-\d{4}$/.test(p.pedId)).length;
  const otrosFormatos = total2026.length - con3Digitos - con4Digitos;

  Logger.log(`\n🔢 FORMATOS DETECTADOS:`);
  Logger.log(`   3 dígitos (PED-2026-###): ${con3Digitos}`);
  Logger.log(`   4 dígitos (PED-2026-####): ${con4Digitos}`);
  if (otrosFormatos > 0) {
    Logger.log(`   ⚠️ Otros formatos: ${otrosFormatos}`);
  }

  // Detectar duplicados
  const pedIds = total2026.map(p => p.pedId);
  const duplicados = pedIds.filter((id, index) => pedIds.indexOf(id) !== index);
  
  if (duplicados.length > 0) {
    Logger.log(`\n⚠️ DUPLICADOS DETECTADOS:`);
    [...new Set(duplicados)].forEach(id => {
      const count = pedIds.filter(p => p === id).length;
      Logger.log(`   ${id} aparece ${count} veces`);
    });
  } else {
    Logger.log(`\n✅ Sin duplicados detectados`);
  }

  // Recomendaciones
  Logger.log(`\n💡 RECOMENDACIONES:`);
  if (con3Digitos > 0) {
    Logger.log(`   ⚠️ Ejecutar: resincronizarPedId2026()`);
    Logger.log(`   Razón: Hay ${con3Digitos} pedidos con formato de 3 dígitos`);
  }
  if (duplicados.length > 0) {
    Logger.log(`   ⚠️ Ejecutar: resincronizarPedId2026()`);
    Logger.log(`   Razón: Se detectaron ${duplicados.length} duplicados`);
  }
  if (con3Digitos === 0 && duplicados.length === 0) {
    Logger.log(`   ✅ Sistema funcionando correctamente`);
  }

  Logger.log('\n═══════════════════════════════════════════════════════');
}

/**
 * PASO 4: Función de verificación post-migración
 */
function verificarMigracion2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPed = ss.getSheetByName("Pedidos");
  const shDes = ss.getSheetByName("Despachados");

  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('🔍 VERIFICACIÓN POST-MIGRACIÓN 2026');
  Logger.log('═══════════════════════════════════════════════════════\n');

  const verificarHoja = (sheet, nombre) => {
    const data = sheet.getDataRange().getValues();
    const pedidos2026 = [];
    let errores = 0;

    for (let i = 1; i < data.length; i++) {
      const pedId = data[i][11];
      const fecha = data[i][12] || data[i][15] || data[i][0];

      if (fecha && new Date(fecha).getFullYear() === 2026) {
        // Verificar formato correcto (4 dígitos)
        if (!/^PED-2026-\d{4}$/.test(pedId)) {
          Logger.log(`   ❌ Fila ${i + 1}: ${pedId} (formato incorrecto)`);
          errores++;
        } else {
          pedidos2026.push(pedId);
        }
      }
    }

    Logger.log(`\n📋 ${nombre}:`);
    Logger.log(`   Total pedidos 2026: ${pedidos2026.length}`);
    Logger.log(`   Errores de formato: ${errores}`);

    return { total: pedidos2026.length, errores: errores };
  };

  const resultPed = verificarHoja(shPed, "PEDIDOS");
  const resultDes = verificarHoja(shDes, "DESPACHADOS");

  const totalErrores = resultPed.errores + resultDes.errores;

  Logger.log('\n═══════════════════════════════════════════════════════');
  if (totalErrores === 0) {
    Logger.log('✅ MIGRACIÓN EXITOSA - Todos los PED_ID 2026 tienen 4 dígitos');
  } else {
    Logger.log(`❌ MIGRACIÓN INCOMPLETA - ${totalErrores} error(es) detectado(s)`);
  }
  Logger.log('═══════════════════════════════════════════════════════');
}

/**
 * PASO 5: Proceso completo automatizado
 */
function migrarA4Digitos2026() {
  Logger.log('🚀 INICIANDO MIGRACIÓN COMPLETA A 4 DÍGITOS (2026)\n');
  
  // Diagnóstico inicial
  Logger.log('PASO 1: Diagnóstico inicial');
  diagnosticarPedId2026();
  
  Logger.log('\n' + '─'.repeat(50) + '\n');
  
  // Resincronización
  Logger.log('PASO 2: Resincronización de PED_ID 2026');
  resincronizarPedId2026();
  
  Logger.log('\n' + '─'.repeat(50) + '\n');
  
  // Verificación final
  Logger.log('PASO 3: Verificación post-migración');
  verificarMigracion2026();
  
  Logger.log('\n✅ PROCESO DE MIGRACIÓN COMPLETADO');
  Logger.log('💡 Recuerda: getNextPedId() ya está actualizada para usar 4 dígitos desde 2026');
}