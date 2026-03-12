# Guía de Configuración — Dashboard Rosecal 2026

## Conexión Paso a Paso

### Paso 1: Abrir Apps Script

1. Abrí tu Google Sheets **"Pedidos Rosecal"**
2. Hacé clic en **Extensiones → Apps Script**
3. Borrá todo el código que aparece por defecto

### Paso 2: Pegar el código

Copiá y pegá **todo** este código:

```javascript
// ╔══════════════════════════════════════════════════════════════╗
// ║  DASHBOARD ROSECAL 2026 — Apps Script Backend v4           ║
// ║  Funciones:                                                ║
// ║  • configurarHojas()  → Ejecutar desde el editor (▶ Run)   ║
// ║  • doGet(e)           → API web (read / write)             ║
// ╚══════════════════════════════════════════════════════════════╝

// ══════════════════════════════════════════════════════════════
// FUNCIÓN 1: configurarHojas()
// ── Ejecutar desde el editor con el botón ▶ Run ──
// Borra las hojas viejas y crea "Datos Manuales"
// ══════════════════════════════════════════════════════════════
function configurarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Lista de hojas viejas a borrar
  var hojasViejas = [
    'Saldos Tesoreria',
    'Deudas y Compromisos',
    'Cuentas por Cobrar',
    'Ventas por Canal (Semanal)',
    'Marketing y Eficiencia',
    'Ranking Publicaciones'
  ];

  var borradas = 0;
  for (var i = 0; i < hojasViejas.length; i++) {
    var hoja = ss.getSheetByName(hojasViejas[i]);
    if (hoja) {
      ss.deleteSheet(hoja);
      borradas++;
    }
  }

  // Crear "Datos Manuales" si no existe
  var dm = ss.getSheetByName('Datos Manuales');
  if (!dm) {
    dm = ss.insertSheet('Datos Manuales');
    dm.appendRow([
      'Fecha',
      'Banco Cta Cte',
      'Saldo Mercado Pago',
      'Efectivo Caja',
      'Inversiones',
      'Deuda Proveedores',
      'Deuda Servicios',
      'Deuda Mercado Pago',
      'Inversion Publicidad',
      'Gasto Marketing',
      'Ventas Corporativas',
      'Gastos Operativos'
    ]);
    // Formato
    dm.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#f0f0f0');
    dm.setFrozenRows(1);
    dm.getRange('B:L').setNumberFormat('$#,##0');
    dm.getRange('A:A').setNumberFormat('dd/MM/yyyy');
    // Ancho de columnas
    dm.setColumnWidth(1, 120);
    for (var c = 2; c <= 12; c++) { dm.setColumnWidth(c, 160); }

    console.log('Fueron borradas ' + borradas + ' hojas viejas.');
    console.log('Creando hoja "Datos Manuales" con 12 columnas (incluyendo Gastos Operativos)...');

    SpreadsheetApp.getUi().alert(
      '✅ ¡Listo!\n\n' +
      '• Se borraron ' + borradas + ' hojas viejas.\n' +
      '• Se creó la hoja "Datos Manuales" con 12 columnas (incluyendo Gastos Operativos).\n\n' +
      'Ahora hacé una Nueva Implementación (Deploy) y pegá la URL en tu index.html.'
    );
    console.log('✅ ¡Listo! Ejecución finalizada con éxito.');
  } else {
    // Si la hoja ya existe, asegurarnos de que la columna Gastos Operativos exista
    var headers = dm.getRange(1, 1, 1, dm.getLastColumn()).getValues()[0];
    var tieneGastosOp = false;
    for (var h = 0; h < headers.length; h++) {
        if (String(headers[h]).toLowerCase().indexOf('gastos operativos') >= 0) tieneGastosOp = true;
    }
    if (!tieneGastosOp) {
        console.log('Falta la columna "Gastos Operativos", agregándola al final...');
        dm.getRange(1, 12).setValue('Gastos Operativos');
        dm.getRange(1, 12).setFontWeight('bold').setBackground('#f0f0f0');
        dm.setColumnWidth(12, 160);
        console.log('ℹ️ Se agregó la columna "Gastos Operativos" a Datos Manuales.');
        SpreadsheetApp.getUi().alert('ℹ️ Se agregó la columna "Gastos Operativos" a Datos Manuales.');
    } else {
        console.log('La hoja "Datos Manuales" ya existía y está correcta.');
        SpreadsheetApp.getUi().alert(
          'ℹ️ La hoja "Datos Manuales" ya existía.\n' +
          'Se borraron ' + borradas + ' hojas viejas.'
        );
    }
  }
}

// ══════════════════════════════════════════════════════════════
// FUNCIÓN 2: doGet(e)
// ── API Web — el Dashboard llama a esta función ──
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var p = (e && e.parameter) ? e.parameter : {};
    var action = p.action || 'read';

    // ── ESCRIBIR DATOS MANUALES ──
    if (action === 'write') {
      var sheet = ss.getSheetByName('Datos Manuales');
      if (!sheet) {
        return jsonOut({ error: 'La hoja "Datos Manuales" no existe. Ejecutá configurarHojas() primero.' });
      }
      sheet.appendRow([
        new Date(),
        toNum(p.bancoCuenta),
        toNum(p.saldoMP),
        toNum(p.efectivoCaja),
        toNum(p.inversiones),
        toNum(p.deudaProveedores),
        toNum(p.deudaServicios),
        toNum(p.deudaMP),
        toNum(p.inversionPublicidad),
        toNum(p.gastoMarketing),
        toNum(p.ventasCorporativas),
        toNum(p.gastosOperativos)
      ]);
      return jsonOut({ success: true, message: 'Datos guardados correctamente' });
    }

    // ── LEER TODO ──
    var targetDate = p.month ? new Date(p.month + '-01T12:00:00') : new Date();
    
    var orders = readOrders(ss, targetDate);
    var monthly = readMonthlySummary(ss);
    var rawGanancias = readRawGanancias(ss, targetDate);
    var cashFlow = readCashFlow(ss);
    var cheques = readCheques(ss);
    var manual = readManualData(ss, targetDate);
    var result = buildResponse(orders, monthly, rawGanancias, cashFlow, cheques, manual, targetDate);
    return jsonOut(result);

  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ══════ HELPER: JSON output ══════
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toNum(v) {
  return Number(v) || 0;
}

// ══════════════════════════════════════════════════════════════
// LECTORES DE HOJAS
// ══════════════════════════════════════════════════════════════

// ── LEER DATOS MANUALES ──
function readManualData(ss, targetDate) {
  var empty = {
    fecha: '', bancoCuenta: 0, saldoMP: 0, efectivoCaja: 0,
    inversiones: 0, deudaProveedores: 0, deudaServicios: 0,
    deudaMP: 0, inversionPublicidad: 0, gastoMarketing: 0,
    ventasCorporativas: 0, gastosOperativos: 0, 
    totalLiquidez: 0, totalPasivo: 0,
    prev: null, history: []
  };

  var sheet = ss.getSheetByName('Datos Manuales');
  if (!sheet) return empty;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return empty;

  var history = [];
  var targetYMonth = targetDate.getFullYear() + '-' + padZero(targetDate.getMonth() + 1);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var dStr = fmtDate(row[0]);
    if (dStr.indexOf(targetYMonth) !== 0) continue; // Solo del mes elegido

    history.push({
      fecha: dStr,
      bancoCuenta: toNum(row[1]),
      saldoMP: toNum(row[2]),
      efectivoCaja: toNum(row[3]),
      inversiones: toNum(row[4]),
      deudaProveedores: toNum(row[5]),
      deudaServicios: toNum(row[6]),
      deudaMP: toNum(row[7]),
      inversionPublicidad: toNum(row[8]),
      gastoMarketing: toNum(row[9]),
      ventasCorporativas: toNum(row[10]),
      gastosOperativos: toNum(row[11])
    });
  }

  if (history.length === 0) return empty;

  var last = history[history.length - 1];
  var totalLiquidez = last.bancoCuenta + last.saldoMP + last.efectivoCaja + last.inversiones;
  var totalPasivo = last.deudaProveedores + last.deudaServicios + last.deudaMP;
  var prev = history.length > 1 ? history[history.length - 2] : null;

  return {
    fecha: last.fecha,
    bancoCuenta: last.bancoCuenta,
    saldoMP: last.saldoMP,
    efectivoCaja: last.efectivoCaja,
    inversiones: last.inversiones,
    deudaProveedores: last.deudaProveedores,
    deudaServicios: last.deudaServicios,
    deudaMP: last.deudaMP,
    inversionPublicidad: last.inversionPublicidad,
    gastoMarketing: last.gastoMarketing,
    ventasCorporativas: last.ventasCorporativas,
    gastosOperativos: last.gastosOperativos,
    totalLiquidez: totalLiquidez,
    totalPasivo: totalPasivo,
    prev: prev,
    history: history
  };
}

// ── LEER GANANCIAS RAW (Ventas Telefónicas + Deuda Clientes) ──
function readRawGanancias(ss, targetDate) {
  var sheet = ss.getSheetByName('Calculo Ganancias');
  if (!sheet) return { ventasTelefonicas: 0, deudaClientes: 0 };
  var data = sheet.getDataRange().getValues();

  var ventasTelefonicas = 0;
  var deudaClientes = 0;

  var curMonth = targetDate.getMonth();
  var curYear = targetDate.getFullYear();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0];
    var amountVal = toNum(row[1]);
    var statusVal = String(row[2] || '').trim().toUpperCase();

    if (dateVal instanceof Date) {
      if (dateVal.getMonth() === curMonth && dateVal.getFullYear() === curYear) {
        // Ventas Telefónicas: status != ANULADO
        if (statusVal !== 'ANULADO') ventasTelefonicas += amountVal;
        // Deuda Clientes: total - cobrado
        deudaClientes += amountVal;
        if (statusVal === 'COBRADO') deudaClientes -= amountVal;
      }
    }
  }
  return { ventasTelefonicas: ventasTelefonicas, deudaClientes: deudaClientes };
}

// ── LEER PEDIDOS (Respuestas de formulario 1) ──
function readOrders(ss, targetDate) {
  var sheet = ss.getSheetByName('Respuestas de formulario 1');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = [];
  for (var j = 0; j < data[0].length; j++) {
    h.push(String(data[0][j]).trim().toLowerCase());
  }

  var ci = {
    fecha: findCol(h, 'marca temporal'),
    cliente: findCol(h, 'nombre del cliente') >= 0 ? findCol(h, 'nombre del cliente') : findCol(h, 'cliente'),
    provincia: findCol(h, 'provincia'),
    importe: findCol(h, 'importe'),
    condicion: findCol(h, 'condici'),
    origen: findCol(h, 'origen'),
    status: findCol(h, 'status')
  };

  var targetYMonth = targetDate.getFullYear() + '-' + padZero(targetDate.getMonth() + 1);
  var orders = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[ci.fecha >= 0 ? ci.fecha : 0]) continue;
    
    var fParsed = fmtDate(r[ci.fecha >= 0 ? ci.fecha : 0]);
    if (fParsed.indexOf(targetYMonth) !== 0) continue; // Filtrar por mes

    var imp = parseAmount(r[ci.importe >= 0 ? ci.importe : 0]);
    if (imp <= 0) continue;
    orders.push({
      date: fParsed,
      client: String(r[ci.cliente >= 0 ? ci.cliente : 1] || ''),
      province: String(r[ci.provincia >= 0 ? ci.provincia : 5] || ''),
      amount: imp,
      condition: String(r[ci.condicion >= 0 ? ci.condicion : 9] || ''),
      origin: String(r[ci.origen >= 0 ? ci.origen : 11] || ''),
      status: String(r[ci.status >= 0 ? ci.status : 14] || '')
    });
  }
  return orders;
}

// ── LEER RESUMEN MENSUAL (Calculo Ganancias - Costos incl.) ──
function readMonthlySummary(ss) {
  var sheet = ss.getSheetByName('Calculo Ganancias');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = [];
  for (var j = 0; j < data[0].length; j++) {
    h.push(String(data[0][j]).trim().toLowerCase());
  }

  var mi = findCol(h, 'mes');
  var vti = findCol(h, 'ventas total');
  var vci = findCol(h, 'ventas cobrad');
  var ipi = findCol(h, 'inversi');
  
  var cbi = findCol(h, 'costo beato');
  if (cbi < 0) cbi = findCol(h, 'beato');
  
  var cli = findCol(h, 'costo logistica');
  if (cli < 0) cli = findCol(h, 'logistica');

  if (mi < 0) return [];
  var monthly = [];
  for (var i = 1; i < data.length; i++) {
    var mes = String(data[i][mi] || '').trim();
    if (!mes) continue;
    monthly.push({
      mes: mes,
      ventasTotal: toNum(data[i][vti >= 0 ? vti : mi + 1]),
      ventasCobrado: toNum(data[i][vci >= 0 ? vci : mi + 2]),
      inversionPub: toNum(data[i][ipi >= 0 ? ipi : mi + 3]),
      costoBeato: cbi >= 0 ? toNum(data[i][cbi]) : 0,
      costoLogistica: cli >= 0 ? toNum(data[i][cli]) : 0
    });
  }
  return monthly;
}

// ── LEER FLUJO OPERATIVO (transpuesto) ──
function readCashFlow(ss) {
  var sheet = ss.getSheetByName('Flujo Operativo');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var months = [];
  for (var c = 1; c < data[0].length; c++) {
    var m = String(data[0][c]).trim().toLowerCase();
    if (m) months.push({ col: c, name: m });
  }

  var concepts = {};
  for (var i = 1; i < data.length; i++) {
    var label = String(data[i][0]).trim().toLowerCase();
    var vals = {};
    for (var k = 0; k < months.length; k++) {
      vals[months[k].name] = toNum(data[i][months[k].col]);
    }
    concepts[label] = vals;
  }

  var monthNames = [];
  for (var n = 0; n < months.length; n++) { monthNames.push(months[n].name); }
  return { months: monthNames, concepts: concepts };
}

// ── LEER CHEQUES ──
function readCheques(ss) {
  var sheet = ss.getSheetByName('Cheques');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = [];
  for (var j = 0; j < data[0].length; j++) {
    h.push(String(data[0][j]).trim().toLowerCase());
  }

  var ci = {
    numero: findCol(h, 'numero'),
    proveedor: findCol(h, 'proveedor'),
    vencimiento: findCol(h, 'vencimiento'),
    monto: findCol(h, 'monto'),
    estado: findCol(h, 'estado'),
    cierre: findCol(h, 'cierre')
  };
  if (ci.numero < 0) ci.numero = findCol(h, 'número');

  var cheques = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[ci.numero >= 0 ? ci.numero : 0]) continue;
    cheques.push({
      numero: String(r[ci.numero >= 0 ? ci.numero : 0]),
      proveedor: String(r[ci.proveedor >= 0 ? ci.proveedor : 1] || ''),
      vencimiento: fmtDate(r[ci.vencimiento >= 0 ? ci.vencimiento : 3]),
      monto: toNum(r[ci.monto >= 0 ? ci.monto : 4]),
      estado: String(r[ci.estado >= 0 ? ci.estado : 5] || ''),
      cierre: fmtDate(r[ci.cierre >= 0 ? ci.cierre : 6])
    });
  }
  return cheques;
}

// ══════════════════════════════════════════════════════════════
// CONSTRUIR RESPUESTA JSON
// ══════════════════════════════════════════════════════════════
function buildResponse(orders, monthly, rawGanancias, cashFlow, cheques, manual, targetDate) {
  var monthNames = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var currentMonthName = monthNames[targetDate.getMonth()];
  var curYearMonth = targetDate.getFullYear() + '-' + padZero(targetDate.getMonth() + 1);
  
  // --- Cash Flow KPIs ---
  var cfData = {
    currentMonth: currentMonthName, ingresosMercadoPago: 0, ingresosEfectivo: 0,
    ingresosDeudaDuek: 0, cajaFinal: 0, prevIngresosMercadoPago: 0,
    prevIngresosEfectivo: 0, prevIngresosDeudaDuek: 0, prevCajaFinal: 0,
    monthlyLabels: [], monthlyCajaFinal: []
  };

  if (cashFlow) {
    var ms = cashFlow.months;
    var c = cashFlow.concepts;
    var mpKey = findConceptKey(c, 'mercado pago', 'ingresos m');
    var efKey = findConceptKey(c, 'efectivo');
    var ddKey = findConceptKey(c, 'deuda duek', 'duek');
    var cajaKey = findConceptKey(c, 'caja final');

    var curIdx = ms.indexOf(currentMonthName);
    if (curIdx < 0) curIdx = ms.length - 1; // Fallback
    var curM = ms[curIdx] || '';
    var prevM = curIdx > 0 ? ms[curIdx - 1] : null;

    cfData.currentMonth = curM;
    cfData.ingresosMercadoPago = mpKey ? (c[mpKey][curM] || 0) : 0;
    cfData.ingresosEfectivo = efKey ? (c[efKey][curM] || 0) : 0;
    cfData.ingresosDeudaDuek = ddKey ? (c[ddKey][curM] || 0) : 0;
    cfData.cajaFinal = cajaKey ? (c[cajaKey][curM] || 0) : 0;

    if (prevM) {
      cfData.prevIngresosMercadoPago = mpKey ? (c[mpKey][prevM] || 0) : 0;
      cfData.prevIngresosEfectivo = efKey ? (c[efKey][prevM] || 0) : 0;
      cfData.prevIngresosDeudaDuek = ddKey ? (c[ddKey][prevM] || 0) : 0;
      cfData.prevCajaFinal = cajaKey ? (c[cajaKey][prevM] || 0) : 0;
    }

    if (cajaKey) {
      for (var m = 0; m < ms.length; m++) {
        if (c[cajaKey][ms[m]]) {
          cfData.monthlyLabels.push(ms[m]);
          cfData.monthlyCajaFinal.push(c[cajaKey][ms[m]]);
        }
      }
    }
  }

  // --- Sales KPIs ---
  var slData = {
    currentMonth: currentMonthName, ventasTotal: 0, ventasCobrado: 0,
    inversionPub: 0, costoBeato: 0, costoLogistica: 0,
    prevVentasTotal: 0, prevVentasCobrado: 0, prevInversionPub: 0,
    ventasTelefonicas: rawGanancias.ventasTelefonicas,
    deudaClientes: rawGanancias.deudaClientes,
    deudaDuek: 0, ventasMercadoLibre: 0, pedidosNuevosMonto: 0, chequesMes: 0
  };

  if (monthly.length > 0) {
    // Buscar el mes objetivo en el array mensual
    var mIdx = -1;
    for (var mCount = 0; mCount < monthly.length; mCount++) {
      if (monthly[mCount].mes === currentMonthName) { mIdx = mCount; break; }
    }
    if (mIdx < 0) mIdx = monthly.length - 1; // Fallback

    var last = monthly[mIdx];
    slData.currentMonth = last ? last.mes : currentMonthName;
    slData.ventasTotal = last ? last.ventasTotal : 0;
    slData.ventasCobrado = last ? last.ventasCobrado : 0;
    slData.inversionPub = last ? last.inversionPub : 0;
    slData.costoBeato = last ? last.costoBeato : 0;
    slData.costoLogistica = last ? last.costoLogistica : 0;
    if (mIdx > 0) {
      var prev = monthly[mIdx - 1];
      slData.prevVentasTotal = prev.ventasTotal;
      slData.prevVentasCobrado = prev.ventasCobrado;
      slData.prevInversionPub = prev.inversionPub;
    }
  }

  // --- Monthly Evolution ---
  var mEvolution = { labels: [], ventasTotal: [], ventasCobrado: [], inversionPub: [] };
  for (var me = 0; me < monthly.length; me++) {
    mEvolution.labels.push(monthly[me].mes);
    mEvolution.ventasTotal.push(monthly[me].ventasTotal);
    mEvolution.ventasCobrado.push(monthly[me].ventasCobrado);
    mEvolution.inversionPub.push(monthly[me].inversionPub);
  }

  // --- Orders by Status + Deuda Duek + ML ---
  var statusCount = {};
  var statusAmount = {};

  for (var o = 0; o < orders.length; o++) {
    var ord = orders[o];
    var s = String(ord.status || '');
    var sUp = s.toUpperCase();

    statusCount[s] = (statusCount[s] || 0) + 1;
    statusAmount[s] = (statusAmount[s] || 0) + ord.amount;
    
    // Contabilizamos todos los pedidos como monto nuevo bruto, excepto los anulados.
    if (sUp !== 'ANULADO') slData.pedidosNuevosMonto += ord.amount;

    // Deuda Duek
    if (sUp.indexOf('DUEK') >= 0) {
      slData.deudaDuek += ord.amount;
    }

    // Mercado Libre
    if (ord.date && ord.date.indexOf(curYearMonth) === 0) {
      var ori = String(ord.origin || '').toUpperCase();
      if (ori.indexOf('MERCADO LIBRE') >= 0 || ori.indexOf('ML') >= 0) {
        slData.ventasMercadoLibre += ord.amount;
      }
    }
  }

  // --- Cheques ---
  var chequesByMonth = {};
  var totalChequesActivos = 0;
  var cantidadActivos = 0;
  var chequesDelMesPagar = 0;

  for (var ch = 0; ch < cheques.length; ch++) {
    if ((cheques[ch].estado || '').toUpperCase() === 'ACTIVO') {
      var cDateStr = cheques[ch].cierre || cheques[ch].vencimiento || 'Sin fecha';
      var monthKey = cDateStr.substring(0, 7);
      
      chequesByMonth[monthKey] = (chequesByMonth[monthKey] || 0) + cheques[ch].monto;
      totalChequesActivos += cheques[ch].monto;
      cantidadActivos++;

      // Cheques a pagar en ESTE mes en particular (para el margen neto)
      if (monthKey === curYearMonth) {
          chequesDelMesPagar += cheques[ch].monto;
      }
    }
  }
  slData.chequesMes = chequesDelMesPagar;
  
  var sortedChequeKeys = Object.keys(chequesByMonth).sort();
  var chequesMontos = [];
  for (var ck = 0; ck < sortedChequeKeys.length; ck++) {
    chequesMontos.push(chequesByMonth[sortedChequeKeys[ck]]);
  }

  // --- Transactions (last 50 for the selected month) ---
  var txs = orders.slice(Math.max(0, orders.length - 50)).reverse();

  // Calcular totales (no enviar historicos de transacciones grandes para no saturar)
  return {
    lastUpdate: new Date().toISOString(),
    computedMonth: curYearMonth, // El mes que se computó (YYYY-MM)
    cashFlow: cfData,
    sales: slData,
    manual: manual,
    monthlyEvolution: mEvolution,
    ordersByStatus: statusCount,
    orderAmountByStatus: statusAmount,
    totalOrders: orders.length,
    cheques: {
      labels: sortedChequeKeys,
      montos: chequesMontos,
      totalActivos: totalChequesActivos,
      cantidad: cantidadActivos
    },
    transactions: txs
  };
}

// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════
function findCol(headers, keyword) {
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].indexOf(keyword) >= 0) return i;
  }
  return -1;
}

function findConceptKey(concepts, keyword1, keyword2) {
  var keys = Object.keys(concepts);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf(keyword1) >= 0) return keys[i];
    if (keyword2 && keys[i].indexOf(keyword2) >= 0) return keys[i];
  }
  return null;
}

function padZero(n) {
  return n < 10 ? '0' + n : '' + n;
}

function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  var parts = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (parts) {
    return parts[3] + '-' + padZero(parseInt(parts[2])) + '-' + padZero(parseInt(parts[1]));
  }
  return s;
}

function parseAmount(v) {
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^\d.,\-]/g, '');
  var clean = s.replace(/\./g, '').replace(',', '.');
  return parseFloat(clean) || 0;
}
```

### Paso 3: Ejecutar `configurarHojas` para actualizar a versión 4

1. En el editor de Apps Script, arriba seleccioná **`configurarHojas`** del menú desplegable.
2. Apretá **▶ Ejecutar (Run)**.
3. El script va a detectar que te falta la columna "Gastos Operativos" y la va a agregar automáticamente a tu hoja "Datos Manuales".

### Paso 4: Publicar (Deploy)

1. Hacé clic en **Implementar → Administrar implementaciones**
2. Lápiz (editar)
3. En **Versión**, cambiá a **Nueva versión**
4. Hacé clic en **Implementar**

> No hace falta que cambies la URL en tu HTML, la URL sigue siendo exactamente la misma.
