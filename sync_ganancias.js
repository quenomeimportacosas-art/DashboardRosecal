/***********************
 * CALCULO DE GANANCIAS
 * (No redeclara constantes que ya están en Código.gs)
 ***********************/
/***********************
 * MENÚ (agregar en onOpen de Código.gs)
 ***********************/
function addMenuGanancias_(ui) {
  ui.createMenu('Ganancias')
    .addItem('🔄 Pasar Importes a Calculo Ganancias', 'syncImportesACalculoGanancias')
    .addToUi();
}
/***********************
 * MAIN - SYNC IMPORTES
 ***********************/
function syncImportesACalculoGanancias() {
  const SHEET_CALC = 'Calculo Ganancias';
  const CALC_HEADER_TS = 'Marca temporal';
  const CALC_HEADER_IMPORTE_FACT = 'Importe Facturado';
  const CALC_HEADER_STATUS = 'Status del pedido';
  const CALC_HEADER_IMPORTE_ORIG = 'Importe original';
  const CALC_HEADER_COSTO_ENVIO = 'Costo Envio';
  const CALC_HEADER_COSTO_BEATO = 'Costo Beato';
  const CALC_HEADER_TELEFONO = 'Telefono';
  
  const CALC_REQUIRED_HEADERS = [
    CALC_HEADER_TS,
    CALC_HEADER_IMPORTE_FACT,
    CALC_HEADER_STATUS,
    CALC_HEADER_IMPORTE_ORIG,
    CALC_HEADER_COSTO_ENVIO,
    CALC_HEADER_COSTO_BEATO,
    CALC_HEADER_TELEFONO
  ];
  const ss = SpreadsheetApp.getActive();
  const shResp = ss.getSheetByName(RESPONSES_SHEET);
  let shCalc = ss.getSheetByName(SHEET_CALC);
  if (!shResp) throw new Error(`No existe la hoja: ${RESPONSES_SHEET}`);
  
  if (!shCalc) {
    shCalc = ss.insertSheet(SHEET_CALC);
    Logger.log(`Se creó la hoja: ${SHEET_CALC}`);
  }
  // --- Leer headers y data de Respuestas
  const respLastRow = shResp.getLastRow();
  const respLastCol = shResp.getLastColumn();
  if (respLastRow < 2) {
    SpreadsheetApp.getUi().alert('No hay datos en Respuestas de formulario 1');
    return;
  }
  const respValues = shResp.getRange(1, 1, respLastRow, respLastCol).getValues();
  const respHeaders = respValues[0].map(h => String(h).trim());
  const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
  const idxTsResp = respHeaders.indexOf(CALC_HEADER_TS);
  const idxImpResp = respHeaders.indexOf('Importe');
  const idxStatusResp = respHeaders.indexOf(CALC_HEADER_STATUS);
  const idxCostoEnvioResp = respHeaders.map(norm).findIndex(h => h === norm(CALC_HEADER_COSTO_ENVIO));
  const idxCostoBeatoResp = respHeaders.map(norm).findIndex(h => h === norm(CALC_HEADER_COSTO_BEATO));
  const idxTelefonoResp = respHeaders.map(norm).findIndex(h => h.indexOf('numero de telefono') >= 0 || h.indexOf('telefono') >= 0);
  
  if (idxTsResp === -1) throw new Error(`No encuentro la columna "${CALC_HEADER_TS}" en ${RESPONSES_SHEET}`);
  if (idxImpResp === -1) throw new Error(`No encuentro la columna "Importe" en ${RESPONSES_SHEET}`);
  if (idxStatusResp === -1) throw new Error(`No encuentro la columna "${CALC_HEADER_STATUS}" en ${RESPONSES_SHEET}`);
  
  // --- FORZAR las columnas en Calculo Ganancias
  ensureHeadersFixedCalc_(shCalc, CALC_REQUIRED_HEADERS);
  const numDataCols = CALC_REQUIRED_HEADERS.length;
  const calcLastRowInA = getLastRowInColumnCalc_(shCalc, 1);
  const calcHeaders = shCalc.getRange(1, 1, 1, numDataCols).getValues()[0].map(h => String(h).trim());
  const idxTsCalc = calcHeaders.indexOf(CALC_HEADER_TS);
  const idxImpFactCalc = calcHeaders.indexOf(CALC_HEADER_IMPORTE_FACT);
  const idxStatusCalc = calcHeaders.indexOf(CALC_HEADER_STATUS);
  const idxImpOrigCalc = calcHeaders.indexOf(CALC_HEADER_IMPORTE_ORIG);
  const idxCostoEnvioCalc = calcHeaders.indexOf(CALC_HEADER_COSTO_ENVIO);
  const idxCostoBeatoCalc = calcHeaders.indexOf(CALC_HEADER_COSTO_BEATO);
  const idxTelefonoCalc = calcHeaders.indexOf(CALC_HEADER_TELEFONO);
  
  if (idxTsCalc === -1 || idxImpFactCalc === -1 || idxStatusCalc === -1 || idxImpOrigCalc === -1) {
    throw new Error('Error al crear/encontrar las columnas en Calculo Ganancias');
  }
  
  // --- Cargar mapa (Marca temporal -> row)
  let calcData = [];
  if (calcLastRowInA >= 2) {
    calcData = shCalc.getRange(2, 1, calcLastRowInA - 1, numDataCols).getValues();
  }
  const tsToRow = new Map();
  for (let r = 0; r < calcData.length; r++) {
    const ts = normalizeKeyCalc_(calcData[r][idxTsCalc]);
    if (ts) tsToRow.set(ts, r + 2);
  }
  
  // --- Preparar escrituras
  const updates = [];
  const appends = [];
  for (let i = 1; i < respValues.length; i++) {
    const tsRaw = respValues[i][idxTsResp];
    const impRaw = respValues[i][idxImpResp];
    const statusRaw = respValues[i][idxStatusResp];
    const costoEnvioRaw = idxCostoEnvioResp !== -1 ? respValues[i][idxCostoEnvioResp] : '';
    const costoBeatoRaw = idxCostoBeatoResp !== -1 ? respValues[i][idxCostoBeatoResp] : '';
    const telefonoRaw = idxTelefonoResp !== -1 ? respValues[i][idxTelefonoResp] : '';
    
    const tsKey = normalizeKeyCalc_(tsRaw);
    if (!tsKey) continue;
    const importeNum = parseImportePrimeroCalc_(impRaw);
    if (importeNum === null) continue;
    const importeFinal = importeNum;
    const statusFinal = String(statusRaw).trim();
    const costoEnvioFinal = costoEnvioRaw !== '' ? costoEnvioRaw : 0;
    const costoBeatoFinal = costoBeatoRaw !== '' ? costoBeatoRaw : 0;
    const telefonoFinal = telefonoRaw !== '' ? String(telefonoRaw).trim() : '';
    
    const existingRow = tsToRow.get(tsKey);
    if (existingRow) {
      updates.push({
        row: existingRow,
        importeFact: importeFinal,
        status: statusFinal,
        orig: impRaw,
        costoEnvio: costoEnvioFinal,
        costoBeato: costoBeatoFinal,
        telefono: telefonoFinal
      });
    } else {
      const row = [tsRaw, importeFinal, statusFinal, impRaw, costoEnvioFinal, costoBeatoFinal, telefonoFinal];
      appends.push(row);
    }
  }
  
  // --- Ejecutar updates
  if (updates.length) {
    const impFactCol = idxImpFactCalc + 1;
    const statusCol = idxStatusCalc + 1;
    const origCol = idxImpOrigCalc + 1;
    const costoEnvioCol = idxCostoEnvioCalc + 1;
    const costoBeatoCol = idxCostoBeatoCalc + 1;
    const telefonoCol = idxTelefonoCalc + 1;
    
    updates.forEach((u) => {
      shCalc.getRange(u.row, impFactCol).setValue(u.importeFact);
      shCalc.getRange(u.row, statusCol).setValue(u.status);
      shCalc.getRange(u.row, origCol).setValue(u.orig);
      shCalc.getRange(u.row, costoEnvioCol).setValue(u.costoEnvio);
      shCalc.getRange(u.row, costoBeatoCol).setValue(u.costoBeato);
      shCalc.getRange(u.row, telefonoCol).setValue(u.telefono);
    });
  }
  
  // --- Ejecutar appends
  if (appends.length) {
    const insertRow = calcLastRowInA + 1;
    shCalc.getRange(insertRow, 1, appends.length, numDataCols).setValues(appends);
  }
}
/***********************
 * HELPERS (nombres únicos para evitar conflictos)
 ***********************/
function ensureHeadersFixedCalc_(sh, requiredHeaders) {
  const numHeaders = requiredHeaders.length;
  const currentRange = sh.getRange(1, 1, 1, numHeaders);
  const currentValues = currentRange.getValues()[0];
  
  let needsUpdate = false;
  
  for (let i = 0; i < numHeaders; i++) {
    const current = String(currentValues[i]).trim();
    const required = requiredHeaders[i];
    
    if (current !== required) {
      needsUpdate = true;
    }
  }
  
  if (needsUpdate) {
    currentRange.setValues([requiredHeaders]);
  }
}
function getLastRowInColumnCalc_(sh, colNum) {
  const maxRows = sh.getMaxRows();
  if (maxRows === 0) return 1;
  
  const colData = sh.getRange(1, colNum, maxRows, 1).getValues();
  
  for (let i = colData.length - 1; i >= 0; i--) {
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
  const s0 = String(input).trim();
  if (!s0) return null;
  const match = s0.match(/[-+]?\s*[$€£]?\s*\d[\d.,\s]*/);
  if (!match) return null;
  let token = match[0]
    .replace(/[$€£\s]/g, "")
    .replace(/[^\d.,+-]/g, "");
  token = token.replace(/(?!^)[+-]/g, "");
  token = token.replace(/[.,]+$/g, "");
  const hasDot = token.includes(".");
  const hasComma = token.includes(",");
  if (hasDot && hasComma) {
    const lastDot = token.lastIndexOf(".");
    const lastComma = token.lastIndexOf(",");
    const decSep = lastDot > lastComma ? "." : ",";
    const thouSep = decSep === "." ? "," : ".";
    token = token.split(thouSep).join("");
    if (decSep === ",") token = token.replace(",", ".");
  } else if (hasComma || hasDot) {
    const sep = hasComma ? "," : ".";
    const parts = token.split(sep);
    if (parts.length === 2) {
      const after = parts[1];
      if (/^\d{3}$/.test(after)) {
        token = parts[0] + after;
      } else {
        token = parts[0] + "." + after;
      }
    } else {
      token = token.split(sep).join("");
    }
  }
  const num = Number(token);
  return Number.isFinite(num) ? num : null;
}
