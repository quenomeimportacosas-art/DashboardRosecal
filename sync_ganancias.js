/***********************
 * CALCULO DE GANANCIAS
 ***********************/
function addMenuGanancias_(ui) {
  ui.createMenu('Ganancias')
    .addItem('🔄 Pasar Importes a Calculo Ganancias', 'syncImportesACalculoGanancias')
    .addToUi();
}

function syncImportesACalculoGanancias() {
  const SHEET_CALC = 'Calculo Ganancias';
  const CALC_REQUIRED_HEADERS = [
    'Marca temporal',
    'Importe Facturado',
    'Status del pedido',
    'Importe original',
    'Costo Envio',
    'Costo Beato',
    'Telefono'
  ];
  
  const ss = SpreadsheetApp.getActive();
  const shResp = ss.getSheetByName(RESPONSES_SHEET);
  let shCalc = ss.getSheetByName(SHEET_CALC);
  if (!shResp) throw new Error('No existe la hoja: ' + RESPONSES_SHEET);
  if (!shCalc) {
    shCalc = ss.insertSheet(SHEET_CALC);
  }

  const respLastRow = shResp.getLastRow();
  const respLastCol = shResp.getLastColumn();
  if (respLastRow < 2) {
    SpreadsheetApp.getUi().alert('No hay datos en Respuestas de formulario 1');
    return;
  }
  const respValues = shResp.getRange(1, 1, respLastRow, respLastCol).getValues();
  const respHeaders = respValues[0].map(function(h) { return String(h).trim(); });

  // --- Detectar columnas en Respuestas (busca por contenido parcial, sin importar acentos)
  var norm = function(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  };
  var findCol = function(headers, keyword1, keyword2) {
    for (var i = 0; i < headers.length; i++) {
      var h = norm(headers[i]);
      if (h.indexOf(norm(keyword1)) >= 0) return i;
      if (keyword2 && h.indexOf(norm(keyword2)) >= 0) return i;
    }
    return -1;
  };

  var idxTsResp = findCol(respHeaders, 'marca temporal');
  var idxImpResp = findCol(respHeaders, 'importe');
  var idxStatusResp = findCol(respHeaders, 'status del pedido', 'status');
  var idxCostoEnvioResp = findCol(respHeaders, 'costo envio', 'costo envío');
  var idxCostoBeatoResp = findCol(respHeaders, 'costo beato');
  var idxTelefonoResp = findCol(respHeaders, 'numero de telefono', 'telefono');

  if (idxTsResp === -1) throw new Error('No encuentro "Marca temporal" en Respuestas');
  if (idxImpResp === -1) throw new Error('No encuentro "Importe" en Respuestas');
  if (idxStatusResp === -1) throw new Error('No encuentro "Status" en Respuestas');

  Logger.log('Respuestas → TS:' + idxTsResp + ' IMP:' + idxImpResp + ' STATUS:' + idxStatusResp + ' ENVIO:' + idxCostoEnvioResp + ' BEATO:' + idxCostoBeatoResp + ' TEL:' + idxTelefonoResp);

  // --- Asegurar headers en Calculo Ganancias
  ensureHeadersFixedCalc_(shCalc, CALC_REQUIRED_HEADERS);
  var numDataCols = CALC_REQUIRED_HEADERS.length;
  var calcLastRowInA = getLastRowInColumnCalc_(shCalc, 1);

  // --- Cargar mapa existente (Marca temporal → fila)
  var calcData = [];
  if (calcLastRowInA >= 2) {
    calcData = shCalc.getRange(2, 1, calcLastRowInA - 1, numDataCols).getValues();
  }
  var tsToRow = new Map();
  for (var r = 0; r < calcData.length; r++) {
    var ts = normalizeKeyCalc_(calcData[r][0]);
    if (ts) tsToRow.set(ts, r + 2);
  }

  // --- Recorrer TODAS las filas de Respuestas
  var updates = [];
  var appends = [];
  var telCount = 0;

  for (var i = 1; i < respValues.length; i++) {
    var row = respValues[i];
    var tsRaw = row[idxTsResp];
    var tsKey = normalizeKeyCalc_(tsRaw);
    if (!tsKey) continue; // Sin marca temporal = fila vacía, saltear

    // Leer campos — NUNCA saltear por importe 0 o vacío
    var impRaw = row[idxImpResp];
    var importeNum = parseImportePrimeroCalc_(impRaw);
    var importeFinal = (importeNum !== null) ? importeNum : 0;
    var statusFinal = String(row[idxStatusResp] || '').trim();
    var costoEnvioFinal = (idxCostoEnvioResp >= 0 && row[idxCostoEnvioResp] !== '') ? row[idxCostoEnvioResp] : 0;
    var costoBeatoFinal = (idxCostoBeatoResp >= 0 && row[idxCostoBeatoResp] !== '') ? row[idxCostoBeatoResp] : 0;
    var telefonoFinal = (idxTelefonoResp >= 0) ? String(row[idxTelefonoResp] || '').trim() : '';

    if (telefonoFinal) telCount++;

    var existingRow = tsToRow.get(tsKey);
    if (existingRow) {
      // SIEMPRE sobreescribir todo: status, importe, teléfono, costos
      updates.push({
        row: existingRow,
        data: [null, importeFinal, statusFinal, impRaw, costoEnvioFinal, costoBeatoFinal, telefonoFinal]
      });
    } else {
      appends.push([tsRaw, importeFinal, statusFinal, impRaw, costoEnvioFinal, costoBeatoFinal, telefonoFinal]);
    }
  }

  // --- Ejecutar updates: sobreescribir columnas B a G (2 a 7) de cada fila
  for (var u = 0; u < updates.length; u++) {
    var upd = updates[u];
    // Escribir columnas 2 a 7 (Importe Facturado, Status, Importe original, Costo Envio, Costo Beato, Telefono)
    shCalc.getRange(upd.row, 2, 1, 6).setValues([[upd.data[1], upd.data[2], upd.data[3], upd.data[4], upd.data[5], upd.data[6]]]);
  }

  // --- Ejecutar appends
  if (appends.length) {
    var insertRow = calcLastRowInA + 1;
    shCalc.getRange(insertRow, 1, appends.length, numDataCols).setValues(appends);
  }

  // --- Resumen
  SpreadsheetApp.getUi().alert(
    '✅ Sync completado\n\n' +
    'Sobreescritos: ' + updates.length + '\n' +
    'Nuevos: ' + appends.length + '\n' +
    'Con teléfono: ' + telCount + ' de ' + (updates.length + appends.length) + '\n' +
    'Col teléfono: ' + (idxTelefonoResp >= 0 ? 'col ' + (idxTelefonoResp + 1) + ' ("' + respHeaders[idxTelefonoResp] + '")' : 'NO encontrada')
  );
}

/***********************
 * HELPERS
 ***********************/
function ensureHeadersFixedCalc_(sh, requiredHeaders) {
  var numHeaders = requiredHeaders.length;
  var currentRange = sh.getRange(1, 1, 1, numHeaders);
  var currentValues = currentRange.getValues()[0];
  var needsUpdate = false;
  for (var i = 0; i < numHeaders; i++) {
    if (String(currentValues[i]).trim() !== requiredHeaders[i]) {
      needsUpdate = true;
      break;
    }
  }
  if (needsUpdate) {
    currentRange.setValues([requiredHeaders]);
  }
}

function getLastRowInColumnCalc_(sh, colNum) {
  var maxRows = sh.getMaxRows();
  if (maxRows === 0) return 1;
  var colData = sh.getRange(1, colNum, maxRows, 1).getValues();
  for (var i = colData.length - 1; i >= 0; i--) {
    if (colData[i][0] !== '' && colData[i][0] !== null && colData[i][0] !== undefined) {
      return i + 1;
    }
  }
  return 1;
}

function normalizeKeyCalc_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return v.toISOString();
  }
  return String(v).trim();
}

function parseImportePrimeroCalc_(input) {
  if (input === null || input === undefined) return null;
  var s0 = String(input).trim();
  if (!s0) return null;
  var match = s0.match(/[-+]?\s*[$€£]?\s*\d[\d.,\s]*/);
  if (!match) return null;
  var token = match[0].replace(/[$€£\s]/g, '').replace(/[^\d.,+-]/g, '');
  token = token.replace(/(?!^)[+-]/g, '');
  token = token.replace(/[.,]+$/g, '');
  var hasDot = token.includes('.');
  var hasComma = token.includes(',');
  if (hasDot && hasComma) {
    var lastDot = token.lastIndexOf('.');
    var lastComma = token.lastIndexOf(',');
    var decSep = lastDot > lastComma ? '.' : ',';
    var thouSep = decSep === '.' ? ',' : '.';
    token = token.split(thouSep).join('');
    if (decSep === ',') token = token.replace(',', '.');
  } else if (hasComma || hasDot) {
    var sep = hasComma ? ',' : '.';
    var parts = token.split(sep);
    if (parts.length === 2) {
      var after = parts[1];
      if (/^\d{3}$/.test(after)) {
        token = parts[0] + after;
      } else {
        token = parts[0] + '.' + after;
      }
    } else {
      token = token.split(sep).join('');
    }
  }
  var num = Number(token);
  return Number.isFinite(num) ? num : null;
}
