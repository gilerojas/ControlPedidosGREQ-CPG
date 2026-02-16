/**
 * ───────────────────────────────────────────────────────────
 * Utilidades de normalización (quita tildes, recorta, mayúsculas)
 * ───────────────────────────────────────────────────────────
 */
function normalize(text) {
  return text
    .toString()
    .normalize('NFD')                // separa tildes
    .replace(/[\u0300-\u036f]/g, '') // elimina diacríticos
    .toUpperCase()
    .trim()
    .replace(/\s+/g, ' ');           // colapsa espacios múltiples
}

/**
 * Mapa crudo Producto → código (con tildes correctas).
 * Luego lo convertimos a claves normalizadas para búsquedas robustas.
 */
const rawTipoMap = {
  'ACRÍLICA SUPERIOR HP': '001',
  'ACRÍLICA SUPERIOR TIPO B': '002',
  'SEMIGLOSS PREMIUM': '003',
  'SEMIGLOSS TIPO B': '004',
  'SATINADA': '005',
  'PROYECTO O CONTRACTOR': '006',
  'PROYECTO P/ TECHOS': '007',
  'ECONÓMICA': '008',
  'PRIMER ACRÍLICO': '009',
  'SELLADOR TECHOS HP o PREMIUM': '010',
  'SELLADOR TECHOS NORMAL': '011',
  'ESMALTE SINTÉTICO': '012',
  'ESMALTE INDUSTRIAL': '013',
  'TEXTURIZADAS': '014',
  'EPÓXICA': '015',
  'BARNIZ CLEAR INDUSTRIAL': '016',
  'BARNIZ PORT EPOXI CLEAR': '017',
  'DRY WET': '018',
  'ESMALTE INDUSTRIAL ANTICORROSIVO': '019'
};

// Construimos el mapa normalizado
const tipoMap = {};
for (const [k, v] of Object.entries(rawTipoMap)) {
  tipoMap[normalize(k)] = v;
}

// Envases (también normalizados)
const rawEnvaseMap = { Galon: '1', Cubeta: '5', Cuartillo: '2' };
const envaseMap = {};
for (const [k, v] of Object.entries(rawEnvaseMap)) {
  envaseMap[normalize(k)] = v;
}

/**
 * ───────────────────────────────────────────────────────────
 * Asigna códigos de barra a los pedidos antiguos sin Cod_Barra
 * ───────────────────────────────────────────────────────────
 */
function asignarCodBarraPedidosAntiguos() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sh        = ss.getSheetByName('Pedidos');
  const catalogo  = ss.getSheetByName('Catalogo_Codigos');
  if (!sh || !catalogo) throw new Error('Hoja Pedidos o Catalogo_Codigos no existe');

  const data           = sh.getDataRange().getValues();
  const colCodBarraIdx = 14;            // Índice 0-based (columna O)
  let   pedidosActualizados = 0;

  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const producto  = row[3];
    const unidad    = row[6];
    const codBarra  = row[colCodBarraIdx];
    const pedId     = row[11];

    if (!producto || !unidad || codBarra) continue; // ya tiene código o faltan datos

    // Normaliza y busca códigos
    const tipoCodigo   = tipoMap[normalize(producto)] || '000';
    const envaseCodigo = envaseMap[normalize(unidad)] || '0';
    const prefijo      = `911${tipoCodigo}${envaseCodigo}`;

    // Secuencial correlativo
    const lastRow    = catalogo.getLastRow();
    const registros  = lastRow > 1
      ? catalogo.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      : [];
    const existentes = registros.filter(c => String(c).startsWith(prefijo));
    const secuencial = ('0' + (existentes.length + 1)).slice(-2);
    const nuevoCod   = `${prefijo}${secuencial}`;

    // Escribe en Pedidos (hoja usa índice 1-based)
    sh.getRange(i + 1, colCodBarraIdx + 1).setValue(nuevoCod);

    // Registra en Catálogo_Codigos
    catalogo.appendRow([
      nuevoCod,
      producto,
      unidad,
      Number(secuencial),
      pedId,
      new Date()
    ]);

    pedidosActualizados++;
  }

  SpreadsheetApp.flush();
  Logger.log(`✅ Asignados ${pedidosActualizados} códigos de barra a pedidos antiguos.`);
}
