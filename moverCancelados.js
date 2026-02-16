function moverCancelados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPedidos = ss.getSheetByName('Pedidos');
  const hojaCancelados = ss.getSheetByName('Cancelados');

  const datos = hojaPedidos.getDataRange().getValues(); // incluye encabezado
  const encabezado = datos[0];
  const filas = datos.slice(1); // sin encabezado

  let filasParaMover = [];
  let indicesParaBorrar = [];

  filas.forEach((fila, i) => {
    const estado = fila[10]; // Columna 11 (índice 10) = Estado
    if (estado === 'CANCELADO') {
      filasParaMover.push(fila);
      indicesParaBorrar.push(i + 2); // +2 porque slice(1) y Spreadsheet es 1-based
    }
  });

  if (filasParaMover.length === 0) {
    Logger.log('No hay pedidos CANCELADOS para mover.');
    return;
  }

  // Asegura encabezado en hoja Cancelados
  if (hojaCancelados.getLastRow() === 0) {
    hojaCancelados.appendRow(encabezado);
  }

  // Agrega a Cancelados
  hojaCancelados.getRange(hojaCancelados.getLastRow() + 1, 1, filasParaMover.length, filasParaMover[0].length)
    .setValues(filasParaMover);

  // Borra desde abajo hacia arriba
  indicesParaBorrar.reverse().forEach(idx => hojaPedidos.deleteRow(idx));

  Logger.log(`Se movieron ${filasParaMover.length} pedidos CANCELADOS a la hoja Cancelados.`);
}
