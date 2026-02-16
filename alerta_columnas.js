function onChange(e) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pedidos');
  const colMax = 14; // Número máximo permitido de columnas (ajusta si es necesario)

  const lastCol = hoja.getLastColumn();
  if (lastCol > colMax) {
    const numExtras = lastCol - colMax;
    const ui = SpreadsheetApp.getUi();
    ui.alert(`⚠️ Atención:\n\nSe detectaron ${numExtras} columna(s) extra fuera del diseño establecido.\n\nPor favor, no agregues columnas nuevas. Si fue un error, puedes eliminarlas manualmente.`);
    Logger.log(`Se detectaron columnas extra: ${numExtras}`);
  }
}
