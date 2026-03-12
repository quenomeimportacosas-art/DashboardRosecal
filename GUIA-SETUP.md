# Guía de Configuración — Dashboard Rosecal 2026

## Conexión Paso a Paso

### Paso 1: Abrir Apps Script

1. Abrí tu Google Sheets **"Pedidos Rosecal"**
2. Hacé clic en **Extensiones → Apps Script**
3. Borrá todo el código que aparece por defecto

### Paso 2: Pegar el código

Copiá y pegá **todo** este código exacto:

```javascript
// ╔══════════════════════════════════════════════════════════════╗
// ║  DASHBOARD ROSECAL 2026 — Apps Script Backend v5           ║
// ║  Funciones:                                                ║
// ║  • configurarHojas()  → Ejecutar desde el editor (▶ Run)   ║
// ║  • doGet(e)           → API web (read / write)             ║
// ╚══════════════════════════════════════════════════════════════╝

function configurarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hojasViejas = [
    'Saldos Tesoreria', 'Deudas y Compromisos',
    'Cuentas por Cobrar', 'Ventas por Canal (Semanal)',
    'Marketing y Eficiencia', 'Ranking Publicaciones'
  ];

  var borradas = 0;
  for (var i = 0; i < hojasViejas.length; i++) {
    var hoja = ss.getSheetByName(hojasViejas[i]);
    if (hoja) {
      ss.deleteSheet(hoja);
      borradas++;
    }
  }

  var dm = ss.getSheetByName('Datos Manuales');
  if (!dm) {
    dm = ss.insertSheet('Datos Manuales');
    dm.appendRow([
      'Fecha', 'Banco Cta Cte', 'Saldo Mercado Pago',
      'Efectivo Caja', 'Inversiones', 'Deuda Proveedores',
      'Deuda Servicios', 'Deuda Mercado Pago', 'Inversion Publicidad',
      'Gasto Marketing', 'Ventas Corporativas', 'Gastos Operativos'
    ]);
    dm.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#f0f0f0');
    dm.setFrozenRows(1);
    dm.getRange('B:L').setNumberFormat('$#,##0');
    dm.getRange('A:A').setNumberFormat('dd/MM/yyyy');
    dm.setColumnWidth(1, 120);
    for (var c = 2; c <= 12; c++) { dm.setColumnWidth(c, 160); }

    console.log('Fueron borradas ' + borradas + ' hojas viejas.');
    console.log('Creando hoja "Datos Manuales"...');

    SpreadsheetApp.getUi().alert('✅ ¡Listo! Se borraron ' + borradas + ' hojas viejas y se creó Datos Manuales. Revisá el código.');
    console.log('✅ ¡Listo! Ejecución finalizada con éxito.');
  } else {
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
        SpreadsheetApp.getUi().alert('ℹ️ Se agregó "Gastos Operativos".');
    } else {
        console.log('La hoja "Datos Manuales" ya existía y está correcta.');
        SpreadsheetApp.getUi().alert('ℹ️ La hoja "Datos Manuales" ya existía.\nSe borraron ' + borradas + ' hojas viejas.');
    }
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var p = (e && e.parameter) ? e.parameter : {};
    var action = p.action || 'read';

    if (action === 'write') {
      var sheet = ss.getSheetByName('Datos Manuales');
      if (!sheet) return jsonOut({ error: 'La hoja no existe. Ejecutá configurarHojas() primero.' });
      sheet.appendRow([
        new Date(), toNum(p.bancoCuenta), toNum(p.saldoMP),
        toNum(p.efectivoCaja), toNum(p.inversiones), toNum(p.deudaProveedores),
        toNum(p.deudaServicios), toNum(p.deudaMP), toNum(p.inversionPublicidad),
        toNum(p.gastoMarketing), toNum(p.ventasCorporativas), toNum(p.gastosOperativos)
      ]);
      return jsonOut({ success: true });
    }

    var targetDate = p.month ? new Date(p.month + '-01T12:00:00') : new Date();
    
    var orders = readOrders(ss, targetDate);
    var monthly = readMonthlySummary(ss);
    var rawGanancias = readRawGanancias(ss, targetDate);
    var cheques = readCheques(ss);
    var manual = readManualData(ss, targetDate);
    
    var result = buildResponse(orders, monthly, rawGanancias, cheques, manual, targetDate);
    return jsonOut(result);

  } catch (err) {
    return jsonOut({ error: err.message, stack: err.stack });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toNum(v) { return Number(v) || 0; }

function readManualData(ss, targetDate) {
  var empty = { bancoCuenta: 0, saldoMP: 0, efectivoCaja: 0, inversiones: 0, deudaProveedores: 0, deudaServicios: 0, deudaMP: 0, inversionPublicidad: 0, gastoMarketing: 0, ventasCorporativas: 0, gastosOperativos: 0, prev: null };
  var sheet = ss.getSheetByName('Datos Manuales');
  if (!sheet) return empty;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return empty;

  var history = [], prevHistory = [];
  var targetYMonth = targetDate.getFullYear() + '-' + padZero(targetDate.getMonth() + 1);
  var prevDDate = new Date(targetDate.getFullYear(), targetDate.getMonth() - 1, 1);
  var prevYMonth = prevDDate.getFullYear() + '-' + padZero(prevDDate.getMonth() + 1);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var dStr = fmtDate(row[0]);
    var dObj = { fecha: dStr, bancoCuenta: toNum(row[1]), saldoMP: toNum(row[2]), efectivoCaja: toNum(row[3]), inversiones: toNum(row[4]), deudaProveedores: toNum(row[5]), deudaServicios: toNum(row[6]), deudaMP: toNum(row[7]), inversionPublicidad: toNum(row[8]), gastoMarketing: toNum(row[9]), ventasCorporativas: toNum(row[10]), gastosOperativos: toNum(row[11]) };

    if (dStr.indexOf(targetYMonth) === 0) history.push(dObj);
    if (dStr.indexOf(prevYMonth) === 0) prevHistory.push(dObj);
  }

  var last = history.length > 0 ? history[history.length - 1] : empty;
  last.prev = prevHistory.length > 0 ? prevHistory[prevHistory.length - 1] : empty;
  return last;
}

function readRawGanancias(ss, targetDate) {
  var sheet = ss.getSheetByName('Calculo Ganancias');
  if (!sheet) return { deudaClientes: 0, totalFacturado: 0 };
  var data = sheet.getDataRange().getValues();
  var deudaClientes = 0, totalFacturado = 0;
  var curMonth = targetDate.getMonth(), curYear = targetDate.getFullYear();

  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][0], amountVal = toNum(data[i][1]), statusVal = String(data[i][2] || '').trim().toUpperCase();
    if (dateVal instanceof Date && dateVal.getMonth() === curMonth && dateVal.getFullYear() === curYear) {
        if(statusVal !== 'ANULADO') totalFacturado += amountVal;
        deudaClientes += amountVal;
        if (statusVal === 'COBRADO' || statusVal === 'ANULADO') deudaClientes -= amountVal;
    }
  }
  return { deudaClientes: deudaClientes, totalFacturado: totalFacturado };
}

function readOrders(ss, targetDate) {
  var sheet = ss.getSheetByName('Respuestas de formulario 1');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = data[0].map(function(k){ return String(k).trim().toLowerCase() });
  
  var ci = {
    fecha: findCol(h, 'marca temporal'),
    cliente: findCol(h, 'nombre del cliente') >= 0 ? findCol(h, 'nombre del cliente') : findCol(h, 'cliente'),
    importe: findCol(h, 'importe'),
    condicion: findCol(h, 'condici'),
    status: findCol(h, 'status')
  };

  var targetYMonth = targetDate.getFullYear() + '-' + padZero(targetDate.getMonth() + 1);
  var orders = [];
  
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[ci.fecha >= 0 ? ci.fecha : 0]) continue;
    
    // Extraer TS full (e.g. "2026-03-01T15:30:00.000Z")
    var rawDate = r[ci.fecha >= 0 ? ci.fecha : 0];
    var tsFull = "";
    if (rawDate instanceof Date) { tsFull = rawDate.toISOString(); }
    else { tsFull = new Date(rawDate).toISOString(); }
    
    // Filtrar mes
    if (!tsFull || tsFull.indexOf(targetYMonth) !== 0) continue; 
    
    var imp = parseAmount(r[ci.importe >= 0 ? ci.importe : 0]);
    if (imp <= 0) continue;
    
    orders.push({
      date: tsFull,
      amount: imp,
      status: String(r[ci.status >= 0 ? ci.status : 14] || '').toUpperCase()
    });
  }
  return orders;
}

function readMonthlySummary(ss) {
  var sheet = ss.getSheetByName('Calculo Ganancias');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var h = data.length > 0 ? data[0].map(function(k){ return String(k).trim().toLowerCase() }) : [];

  var mi = findCol(h, 'mes');
  var cli = findCol(h, 'costo logistica', 'logistica');
  var cbi = findCol(h, 'costo beato', 'beato');
  
  if (mi < 0 || data.length < 2) return [];
  var monthly = [];
  for (var i = 1; i < data.length; i++) {
    var mes = String(data[i][mi] || '').trim();
    if (!mes) continue;
    monthly.push({
      mes: mes,
      costoBeato: cbi >= 0 ? toNum(data[i][cbi]) : 0,
      costoLogistica: cli >= 0 ? toNum(data[i][cli]) : 0
    });
  }
  return monthly;
}

function readCheques(ss) {
  var sheet = ss.getSheetByName('Cheques');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var h = data.length > 0 ? data[0].map(function(k){ return String(k).trim().toLowerCase() }) : [];

  var ci = { numero: findCol(h, 'numero', 'número'), proveedor: findCol(h, 'proveedor'), vencimiento: findCol(h, 'vencimiento'), monto: findCol(h, 'monto'), estado: findCol(h, 'estado') };
  var cheques = [];
  
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[ci.numero >= 0 ? ci.numero : 0]) continue;
    var est = String(r[ci.estado >= 0 ? ci.estado : 5] || '').toUpperCase();
    if (est === 'ACTIVO') {
        cheques.push({
          proveedor: String(r[ci.proveedor >= 0 ? ci.proveedor : 1] || ''),
          vencimiento: fmtDate(r[ci.vencimiento >= 0 ? ci.vencimiento : 3]),
          monto: toNum(r[ci.monto >= 0 ? ci.monto : 4])
        });
    }
  }
  return cheques;
}

function buildResponse(orders, monthly, rawGanancias, cheques, manual, targetDate) {
  var currentMonthName = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'][targetDate.getMonth()];
  var curYearMonth = targetDate.getFullYear() + '-' + padZero(targetDate.getMonth() + 1);
  
  var slData = { costoBeato: 0, costoLogistica: 0, deudaClientes: rawGanancias.deudaClientes, facturadoTotal: rawGanancias.totalFacturado };

  if (monthly.length > 0) {
    var last = monthly.slice().reverse().find(function(m){ return m.mes === currentMonthName; }) || monthly[monthly.length - 1];
    slData.costoBeato = last.costoBeato || 0;
    slData.costoLogistica = last.costoLogistica || 0;
  }

  // Aggregate Cheques Activos
  var totalChequesActivos = 0;
  for (var c = 0; c < cheques.length; c++) totalChequesActivos += cheques[c].monto;

  return {
    lastUpdate: new Date().toISOString(),
    computedMonth: curYearMonth, 
    manual: manual,
    sales: slData,
    chequesActivos: cheques,
    totalChequesActivos: totalChequesActivos,
    orders: orders // Export full month orders with FULL timestamp
  };
}

function findCol(headers, k1, k2) {
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].indexOf(k1) >= 0) return i;
    if (k2 && headers[i].indexOf(k2) >= 0) return i;
  }
  return -1;
}

function padZero(n) { return n < 10 ? '0' + n : '' + n; }

function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v).trim(), parts = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (parts) return parts[3] + '-' + padZero(parseInt(parts[2])) + '-' + padZero(parseInt(parts[1]));
  return s;
}

function parseAmount(v) {
  if (typeof v === 'number') return v;
  return parseFloat(String(v).replace(/[^\d.,\-]/g, '').replace(/\./g, '').replace(',', '.')) || 0;
}
```

### Paso 3: Publicar como Nueva Versión Oficial

⚠️ **ATENCIÓN: ES EL PASO MÁS IMPORTANTE PARA QUE FUNCIONE EL V5**

1. En el script abierto con el nuevo bloque, toca el botón azul de arriba a la derecha que dice **Implementar** y elegí **Administrar implementaciones**.
2. Te abre un recuadro. Del lado derecho está el ícono de un **lápiz (✏️)**. Hacé click ahí.
3. Te abre otro panel. Donde dice **"Versión"** (va a decir algo como V3 o V4), marcá la opción que dice **"Nueva versión"**.
4. Ahora, dale al botón azul **Implementar** un poco más abajo.

Con esto recarga la fuente y ya está listo tu sistema!
