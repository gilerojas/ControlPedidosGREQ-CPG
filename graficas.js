// ───────────────────────────────────────────────────────────
// graficasResumen.gs – Generación de gráficas visuales
// ───────────────────────────────────────────────────────────

/**
 * Función principal: Genera hoja con gráficas del mes de octubre
 */
function generarGraficasOctubre() {
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

  // === Filtrar datos de octubre 2025 ===
  const inicio = new Date(2025, 9, 1);  // Octubre
  const fin = new Date(2025, 9, 31);

  const enMes = data.filter(r => {
    const fecha = new Date(r[15]); // Fecha_archivo
    const pedId = r[11];           // PED_ID
    return fecha >= inicio && fecha <= fin && pedId && pedId.toString().trim() !== "";
  });

  if (enMes.length === 0) {
    Logger.log("No hay despachos en octubre 2025");
    return;
  }

  // === Crear o limpiar hoja de gráficas ===
  let chartSheet = ss.getSheetByName('Gráficas_Octubre');
  if (!chartSheet) {
    chartSheet = ss.insertSheet('Gráficas_Octubre');
  } else {
    chartSheet.clear();
    // Eliminar gráficas anteriores
    const charts = chartSheet.getCharts();
    charts.forEach(chart => chartSheet.removeChart(chart));
  }

  // === PROCESAR DATOS ===
  
  // 1. Top 5 productos por volumen (cubetas + galones)
  const productos = {}; // { Producto: {gal: xx, cub: xx} }
  
  enMes.forEach(r => {
    const producto = r[3];    // Producto
    const cantidad = Number(r[5]) || 0;
    const unidad = (r[6] || "").toString().toLowerCase();

    if (!productos[producto]) {
      productos[producto] = { gal: 0, cub: 0 };
    }

    if (unidad.includes("gal")) productos[producto].gal += cantidad;
    if (unidad.includes("cub")) productos[producto].cub += cantidad;
  });

  // Ordenar por volumen total (gal + cub) y tomar top 5
  const topProductos = Object.entries(productos)
    .map(([prod, vals]) => ({
      producto: prod,
      gal: vals.gal,
      cub: vals.cub,
      total: vals.gal + vals.cub
    }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 10);

  // 2. Evolución diaria de pedidos
  const pedidosPorDia = {}; // { "2025-10-01": 5, "2025-10-02": 8, ... }
  
  enMes.forEach(r => {
    const fecha = new Date(r[15]);
    const fechaStr = Utilities.formatDate(fecha, "America/Santo_Domingo", "yyyy-MM-dd");
    pedidosPorDia[fechaStr] = (pedidosPorDia[fechaStr] || 0) + 1;
  });

  // Ordenar por fecha
  const evolucionDiaria = Object.entries(pedidosPorDia)
    .sort((a, b) => new Date(a[0]) - new Date(b[0]))
    .map(([fecha, cantidad]) => ({
      dia: new Date(fecha).getDate(), // Solo el día (1-31)
      cantidad: cantidad
    }));

  // === ESCRIBIR DATOS EN LA HOJA ===

  // Título
  chartSheet.getRange('A1').setValue('📊 Gráficas - Resumen de Octubre 2025')
    .setFontSize(14)
    .setFontWeight('bold');

  // --- Tabla 1: Top 5 Productos ---
  chartSheet.getRange('A3').setValue('Top 5 Productos (Cubetas y Galones)')
    .setFontWeight('bold');
  
  chartSheet.getRange('A4:C4').setValues([['Producto', 'Cubetas', 'Galones']])
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  const dataProductos = topProductos.map(p => [p.producto, p.cub, p.gal]);
  chartSheet.getRange(5, 1, dataProductos.length, 3).setValues(dataProductos);

  // --- Tabla 2: Evolución Diaria ---
  chartSheet.getRange('E3').setValue('Evolución Diaria de Pedidos')
    .setFontWeight('bold');
  
  chartSheet.getRange('E4:F4').setValues([['Día', 'Pedidos']])
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF');

  const dataEvolucion = evolucionDiaria.map(e => [e.dia, e.cantidad]);
  chartSheet.getRange(5, 5, dataEvolucion.length, 2).setValues(dataEvolucion);

  // === CREAR GRÁFICA 1: BAR CHART DOBLE (Cubetas y Galones) ===
  
  const rangeProductos = chartSheet.getRange('A4:C' + (4 + dataProductos.length));
  
  const barChart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(rangeProductos)
    .setPosition(5, 8, 0, 0) // Fila 5, Columna H
    .setOption('title', 'Top 5 Productos Despachados - Octubre 2025')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('width', 700)
    .setOption('height', 400)
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', {
      title: 'Productos',
      textStyle: { fontSize: 11 }
    })
    .setOption('vAxis', {
      title: 'Cantidad',
      textStyle: { fontSize: 11 }
    })
    .setOption('colors', ['#FBBC04', '#4285F4']) // Cubetas amarillo, Galones azul
    .setOption('bar', { groupWidth: '75%' })
    .setOption('annotations', {
      alwaysOutside: true,
      textStyle: {
        fontSize: 11,
        bold: true,
        color: '#000'
      }
    })
    .setOption('isStacked', false) // Barras agrupadas, no apiladas
    .build();

  chartSheet.insertChart(barChart);

  // === CREAR GRÁFICA 2: LINE CHART (Evolución Diaria) ===
  
  const rangeEvolucion = chartSheet.getRange('E4:F' + (4 + dataEvolucion.length));
  
  const lineChart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeEvolucion)
    .setPosition(20, 8, 0, 0) // Fila 20, Columna H
    .setOption('title', 'Evolución de Pedidos por Día - Octubre 2025')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('width', 700)
    .setOption('height', 400)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', {
      title: 'Día del Mes',
      textStyle: { fontSize: 11 },
      gridlines: { count: 31 }
    })
    .setOption('vAxis', {
      title: 'Cantidad de Pedidos',
      textStyle: { fontSize: 11 },
      minValue: 0
    })
    .setOption('colors', ['#EA4335']) // Rojo Google
    .setOption('lineWidth', 3)
    .setOption('pointSize', 6)
    .setOption('curveType', 'function') // Línea suavizada
    .setOption('tooltip', { isHtml: true })
    .build();

  chartSheet.insertChart(lineChart);

  Logger.log("✅ Gráficas generadas en la hoja 'Gráficas_Octubre'");
  
  // Ajustar anchos de columnas para mejor visualización
  chartSheet.setColumnWidth(1, 200); // Columna A (Productos)
  chartSheet.setColumnWidth(2, 100); // Columna B (Cubetas)
  chartSheet.setColumnWidth(3, 100); // Columna C (Galones)
  chartSheet.setColumnWidth(5, 80);  // Columna E (Día)
  chartSheet.setColumnWidth(6, 100); // Columna F (Pedidos)
}

/**
 * Versión genérica para cualquier mes (opcional)
 * @param {number} year - Año (ej: 2025)
 * @param {number} month - Mes (1-12, donde 1=Enero)
 */
function generarGraficasMes(year, month) {
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

  // === Filtrar datos del mes especificado ===
  const inicio = new Date(year, month - 1, 1);
  const fin = new Date(year, month, 0); // Último día del mes

  const enMes = data.filter(r => {
    const fecha = new Date(r[15]);
    const pedId = r[11];
    return fecha >= inicio && fecha <= fin && pedId && pedId.toString().trim() !== "";
  });

  if (enMes.length === 0) {
    Logger.log(`No hay despachos en ${month}/${year}`);
    return;
  }

  const mesNombre = Utilities.formatDate(inicio, "America/Santo_Domingo", "MMMM yyyy");
  const sheetName = `Gráficas_${mesNombre.replace(' ', '_')}`;

  // === Crear o limpiar hoja de gráficas ===
  let chartSheet = ss.getSheetByName(sheetName);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(sheetName);
  } else {
    chartSheet.clear();
    const charts = chartSheet.getCharts();
    charts.forEach(chart => chartSheet.removeChart(chart));
  }

  // === PROCESAR DATOS (igual que generarGraficasOctubre) ===
  
  const productos = {};
  
  enMes.forEach(r => {
    const producto = r[3];
    const cantidad = Number(r[5]) || 0;
    const unidad = (r[6] || "").toString().toLowerCase();

    if (!productos[producto]) {
      productos[producto] = { gal: 0, cub: 0 };
    }

    if (unidad.includes("gal")) productos[producto].gal += cantidad;
    if (unidad.includes("cub")) productos[producto].cub += cantidad;
  });

  const topProductos = Object.entries(productos)
    .map(([prod, vals]) => ({
      producto: prod,
      gal: vals.gal,
      cub: vals.cub,
      total: vals.gal + vals.cub
    }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 5);

  const pedidosPorDia = {};
  
  enMes.forEach(r => {
    const fecha = new Date(r[15]);
    const fechaStr = Utilities.formatDate(fecha, "America/Santo_Domingo", "yyyy-MM-dd");
    pedidosPorDia[fechaStr] = (pedidosPorDia[fechaStr] || 0) + 1;
  });

  const evolucionDiaria = Object.entries(pedidosPorDia)
    .sort((a, b) => new Date(a[0]) - new Date(b[0]))
    .map(([fecha, cantidad]) => ({
      dia: new Date(fecha).getDate(),
      cantidad: cantidad
    }));

  // === ESCRIBIR DATOS ===

  chartSheet.getRange('A1').setValue(`📊 Gráficas - ${capitalizeFirst(mesNombre)}`)
    .setFontSize(14)
    .setFontWeight('bold');

  chartSheet.getRange('A3').setValue('Top 5 Productos (Cubetas y Galones)')
    .setFontWeight('bold');
  
  chartSheet.getRange('A4:C4').setValues([['Producto', 'Cubetas', 'Galones']])
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  const dataProductos = topProductos.map(p => [p.producto, p.cub, p.gal]);
  chartSheet.getRange(5, 1, dataProductos.length, 3).setValues(dataProductos);

  chartSheet.getRange('E3').setValue('Evolución Diaria de Pedidos')
    .setFontWeight('bold');
  
  chartSheet.getRange('E4:F4').setValues([['Día', 'Pedidos']])
    .setFontWeight('bold')
    .setBackground('#34A853')
    .setFontColor('#FFFFFF');

  const dataEvolucion = evolucionDiaria.map(e => [e.dia, e.cantidad]);
  chartSheet.getRange(5, 5, dataEvolucion.length, 2).setValues(dataEvolucion);

  // === GRÁFICAS ===
  
  const rangeProductos = chartSheet.getRange('A4:C' + (4 + dataProductos.length));
  
  const barChart = chartSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(rangeProductos)
    .setPosition(5, 8, 0, 0)
    .setOption('title', `Top 5 Productos Despachados - ${capitalizeFirst(mesNombre)}`)
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('width', 700)
    .setOption('height', 400)
    .setOption('legend', { position: 'top' })
    .setOption('hAxis', {
      title: 'Productos',
      textStyle: { fontSize: 11 }
    })
    .setOption('vAxis', {
      title: 'Cantidad',
      textStyle: { fontSize: 11 }
    })
    .setOption('colors', ['#FBBC04', '#4285F4'])
    .setOption('bar', { groupWidth: '75%' })
    .setOption('isStacked', false)
    .build();

  chartSheet.insertChart(barChart);

  const rangeEvolucion = chartSheet.getRange('E4:F' + (4 + dataEvolucion.length));
  
  const lineChart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(rangeEvolucion)
    .setPosition(20, 8, 0, 0)
    .setOption('title', `Evolución de Pedidos por Día - ${capitalizeFirst(mesNombre)}`)
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('width', 700)
    .setOption('height', 400)
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', {
      title: 'Día del Mes',
      textStyle: { fontSize: 11 }
    })
    .setOption('vAxis', {
      title: 'Cantidad de Pedidos',
      textStyle: { fontSize: 11 },
      minValue: 0
    })
    .setOption('colors', ['#EA4335'])
    .setOption('lineWidth', 3)
    .setOption('pointSize', 6)
    .setOption('curveType', 'function')
    .build();

  chartSheet.insertChart(lineChart);

  Logger.log(`✅ Gráficas generadas en la hoja '${sheetName}'`);
  
  chartSheet.setColumnWidth(1, 200);
  chartSheet.setColumnWidth(2, 100);
  chartSheet.setColumnWidth(3, 100);
  chartSheet.setColumnWidth(5, 80);
  chartSheet.setColumnWidth(6, 100);
}

/**
 * Helper: Capitalizar primera letra
 */
function capitalizeFirst(text) {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

/**
 * BONUS: Exportar la hoja de gráficas como PDF y obtener link
 * @returns {string} URL del PDF generado
 */
function exportarGraficasComoPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const chartSheet = ss.getSheetByName('Gráficas_Octubre');
  
  if (!chartSheet) {
    Logger.log("❌ Primero genera las gráficas con generarGraficasOctubre()");
    return null;
  }

  const sheetId = chartSheet.getSheetId();
  const ssId = ss.getId();
  
  // URL de exportación como PDF
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&gid=${sheetId}&portrait=false&fitw=true`;
  
  Logger.log("📄 Link de PDF: " + url);
  return url;
}