/**
 * Llenar los PED_ID vacíos en la hoja 'Pedidos', solo para filas donde Estado está lleno.
 */
function completarPedIdsPedidosSoloConEstado() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName('Pedidos');
  if (!sh) { SpreadsheetApp.getUi().alert('Hoja Pedidos no existe'); return; }

  const header  = sh.getDataRange().getValues()[0];
  const colPed  = header.indexOf('PED_ID');
  const colEstado = header.indexOf('Estado');
  const lastRow = sh.getLastRow();

  let seq = Number(PropertiesService.getDocumentProperties().getProperty('NEXT_PED_SEQ')) || 1;
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  let nuevos = 0;

  for (let r = 2; r <= lastRow; r++) {
    const estado = sh.getRange(r, colEstado + 1).getValue();
    const cellPed = sh.getRange(r, colPed + 1);
    const valPed  = String(cellPed.getValue()).trim();

    if (!valPed && estado) {
      // Solo asigna PED_ID si Estado no está vacío
      cellPed.setValue(`${year}-${('000' + seq).slice(-3)}`);
      seq++;
      nuevos++;
    }
  }

  PropertiesService.getDocumentProperties().setProperty('NEXT_PED_SEQ', seq);
  SpreadsheetApp.flush();
  Logger.log(`Completados ${nuevos} PED_ID en Pedidos (solo con Estado lleno). NEXT_PED_SEQ ahora es ${seq}`);
}
