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
// ║  DASHBOARD ROSECAL 2026 — Apps Script Backend v3           ║
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Lista de hojas viejas a borrar
  const hojasViejas = [
    'Saldos Tesoreria',
    'Deudas y Compromisos',
    'Cuentas por Cobrar',
    'Ventas por Canal (Semanal)',
    'Marketing y Eficiencia',
    'Ranking Publicaciones'
  ];

  let borradas = 0;
  hojasViejas.forEach(function(nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (hoja) {
      ss.deleteSheet(hoja);
      borradas++;
    }
  });

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
      'Ventas Corporativas'
    ]);
    // Formato
    dm.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f0f0f0');
    dm.setFrozenRows(1);
    dm.getRange('B:K').setNumberFormat('$#,##0');
    dm.getRange('A:A').setNumberFormat('dd/MM/yyyy');
    // Ancho de columnas
    dm.setColumnWidth(1, 120);
    for (var c = 2; c <= 11; c++) { dm.setColumnWidth(c, 160); }

    SpreadsheetApp.getUi().alert(
      '✅ ¡Listo!\n\n' +
      '• Se borraron ' + borradas + ' hojas viejas.\n' +
      '• Se creó la hoja "Datos Manuales" con 11 columnas.\n\n' +
      'Ahora hacé una Nueva Implementación (Deploy) y pegá la URL en tu index.html.'
    );
  } else {
    SpreadsheetApp.getUi().alert(
      'ℹ️ La hoja "Datos Manuales" ya existía.\n' +
      'Se borraron ' + borradas + ' hojas viejas.'
    );
  }
}

// ══════════════════════════════════════════════════════════════
// FUNCIÓN 2: doGet(e)
// ── API Web — el Dashboard llama a esta función ──
// Acciones:
//   (sin parámetro o action=read)  → devuelve JSON con todos los datos
//   action=write&bancoCuenta=X&... → graba fila en "Datos Manuales"
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
        toNum(p.ventasCorporativas)
      ]);
      return jsonOut({ success: true, message: 'Datos guardados correctamente' });
    }

    // ── LEER TODO ──
    var orders = readOrders(ss);
    var monthly = readMonthlySummary(ss);
    var rawGanancias = readRawGanancias(ss);
    var cashFlow = readCashFlow(ss);
    var cheques = readCheques(ss);
    var manual = readManualData(ss);
    var result = buildResponse(orders, monthly, rawGanancias, cashFlow, cheques, manual);
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
function readManualData(ss) {
  var empty = {
    fecha: '', bancoCuenta: 0, saldoMP: 0, efectivoCaja: 0,
    inversiones: 0, deudaProveedores: 0, deudaServicios: 0,
    deudaMP: 0, inversionPublicidad: 0, gastoMarketing: 0,
    ventasCorporativas: 0, totalLiquidez: 0, totalPasivo: 0,
    prev: null, history: []
  };

  var sheet = ss.getSheetByName('Datos Manuales');
  if (!sheet) return empty;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return empty;

  var history = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    history.push({
      fecha: fmtDate(row[0]),
      bancoCuenta: toNum(row[1]),
      saldoMP: toNum(row[2]),
      efectivoCaja: toNum(row[3]),
      inversiones: toNum(row[4]),
      deudaProveedores: toNum(row[5]),
      deudaServicios: toNum(row[6]),
      deudaMP: toNum(row[7]),
      inversionPublicidad: toNum(row[8]),
      gastoMarketing: toNum(row[9]),
      ventasCorporativas: toNum(row[10])
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
    totalLiquidez: totalLiquidez,
    totalPasivo: totalPasivo,
    prev: prev,
    history: history
  };
}

// ── LEER GANANCIAS RAW (Ventas Telefónicas + Deuda Clientes) ──
function readRawGanancias(ss) {
  var sheet = ss.getSheetByName('Calculo Ganancias');
  if (!sheet) return { ventasTelefonicas: 0, deudaClientes: 0 };
  var data = sheet.getDataRange().getValues();

  var ventasTelefonicas = 0;
  var deudaClientes = 0;

  var now = new Date();
  var curMonth = now.getMonth();
  var curYear = now.getFullYear();

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
function readOrders(ss) {
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

  var orders = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[ci.fecha >= 0 ? ci.fecha : 0]) continue;
    var imp = parseAmount(r[ci.importe >= 0 ? ci.importe : 0]);
    if (imp <= 0) continue;
    orders.push({
      date: fmtDate(r[ci.fecha >= 0 ? ci.fecha : 0]),
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

// ── LEER RESUMEN MENSUAL (Calculo Ganancias) ──
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

  if (mi < 0) return [];
  var monthly = [];
  for (var i = 1; i < data.length; i++) {
    var mes = String(data[i][mi] || '').trim();
    if (!mes) continue;
    monthly.push({
      mes: mes,
      ventasTotal: toNum(data[i][vti >= 0 ? vti : mi + 1]),
      ventasCobrado: toNum(data[i][vci >= 0 ? vci : mi + 2]),
      inversionPub: toNum(data[i][ipi >= 0 ? ipi : mi + 3])
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
function buildResponse(orders, monthly, rawGanancias, cashFlow, cheques, manual) {
  // --- Cash Flow KPIs ---
  var cfData = {
    currentMonth: '', ingresosMercadoPago: 0, ingresosEfectivo: 0,
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

    var curIdx = -1;
    if (cajaKey) {
      for (var i = ms.length - 1; i >= 0; i--) {
        if (c[cajaKey][ms[i]]) { curIdx = i; break; }
      }
    }
    if (curIdx < 0) curIdx = ms.length - 1;
    var curM = ms[curIdx];
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
    currentMonth: '', ventasTotal: 0, ventasCobrado: 0,
    inversionPub: 0, prevVentasTotal: 0, prevVentasCobrado: 0, prevInversionPub: 0,
    ventasTelefonicas: rawGanancias.ventasTelefonicas,
    deudaClientes: rawGanancias.deudaClientes,
    deudaDuek: 0, ventasMercadoLibre: 0
  };

  if (monthly.length > 0) {
    var last = monthly[monthly.length - 1];
    slData.currentMonth = last.mes;
    slData.ventasTotal = last.ventasTotal;
    slData.ventasCobrado = last.ventasCobrado;
    slData.inversionPub = last.inversionPub;
    if (monthly.length > 1) {
      var prev = monthly[monthly.length - 2];
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
  var provinceSet = {};
  var conditionSet = {};

  var now = new Date();
  var curYearMonth = now.getFullYear() + '-' + padZero(now.getMonth() + 1);

  for (var o = 0; o < orders.length; o++) {
    var ord = orders[o];
    var s = String(ord.status || '');
    var sUp = s.toUpperCase();

    statusCount[s] = (statusCount[s] || 0) + 1;
    statusAmount[s] = (statusAmount[s] || 0) + ord.amount;
    if (ord.province) provinceSet[ord.province] = 1;
    if (ord.condition) conditionSet[ord.condition] = 1;

    // Deuda Duek
    if (sUp.indexOf('DUEK') >= 0) {
      slData.deudaDuek += ord.amount;
    }

    // Mercado Libre (mes en curso)
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
  for (var ch = 0; ch < cheques.length; ch++) {
    if ((cheques[ch].estado || '').toUpperCase() === 'ACTIVO') {
      var monthKey = (cheques[ch].cierre || cheques[ch].vencimiento || 'Sin fecha').substring(0, 7);
      chequesByMonth[monthKey] = (chequesByMonth[monthKey] || 0) + cheques[ch].monto;
      totalChequesActivos += cheques[ch].monto;
      cantidadActivos++;
    }
  }
  var sortedChequeKeys = Object.keys(chequesByMonth).sort();
  var chequesMontos = [];
  for (var ck = 0; ck < sortedChequeKeys.length; ck++) {
    chequesMontos.push(chequesByMonth[sortedChequeKeys[ck]]);
  }

  // --- Comparison ---
  var comparison = [];
  for (var cm = monthly.length - 2; cm >= 0 && comparison.length < 3; cm--) {
    comparison.push({
      period: monthly[cm].mes.charAt(0).toUpperCase() + monthly[cm].mes.slice(1),
      value: monthly[cm].ventasTotal
    });
  }

  // --- Transactions (last 30) ---
  var txs = orders.slice(Math.max(0, orders.length - 30)).reverse();

  return {
    lastUpdate: new Date().toISOString(),
    cashFlow: cfData,
    sales: slData,
    manual: manual,
    monthlyEvolution: mEvolution,
    ordersByStatus: statusCount,
    orderAmountByStatus: statusAmount,
    cheques: {
      labels: sortedChequeKeys,
      montos: chequesMontos,
      totalActivos: totalChequesActivos,
      cantidad: cantidadActivos
    },
    comparison: comparison,
    transactions: txs,
    provinces: Object.keys(provinceSet).sort(),
    conditions: Object.keys(conditionSet).sort()
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

### Paso 3: Ejecutar `configurarHojas` (ESTO ES NUEVO)

Antes de publicar, vas a correr la función que borra las hojas viejas y crea "Datos Manuales":

1. En el editor de Apps Script, arriba hay un **menú desplegable** que dice `doGet` o `myFunction`
2. Hacé clic en ese menú y seleccioná **`configurarHojas`**
3. Hacé clic en el botón **▶ Ejecutar** (Run)
4. Si te pide permisos, **autorizá** (es tu cuenta)
5. Te va a aparecer un cartel diciendo cuántas hojas borró y que creó "Datos Manuales"

### Paso 4: Publicar (Deploy)

1. Hacé clic en **Implementar → Nueva implementación**
2. En **Tipo**, elegí **App web**
3. Configurá:
   - **Descripción:** Dashboard Rosecal API v3
   - **Ejecutar como:** Yo mismo
   - **Quién tiene acceso:** Cualquiera
4. Hacé clic en **Implementar**
5. **Copiá la URL** resultante y **pegámela en el chat**

### Paso 5: Conectar al Dashboard

1. Abrí `index.html` con un editor de texto
2. Buscá esta línea:
```javascript
APPS_SCRIPT_URL: '',
```
3. Pegá tu URL:
```javascript
APPS_SCRIPT_URL: 'https://script.google.com/macros/s/TU-NUEVA-URL/exec',
```
4. Guardá el archivo

---

## Hojas que lee el Dashboard

| Pestaña | Qué lee | Para qué |
|---------|---------|----------|
| `Respuestas de formulario 1` | Pedidos (fecha, cliente, importe, provincia, status) | Tabla, filtros, gráfico de estados |
| `Calculo Ganancias` | Columnas A (fecha), B (importe), C (status), H (mes) | Ventas Telefónicas, Deuda Clientes, evolución |
| `Flujo Operativo` | Conceptos transpuestos por mes | KPIs de caja, gráfico de flujo |
| `Cheques` | Número, monto, estado, vencimiento | Gráfico de cheques activos |
| `Datos Manuales` | Banco, Saldo MP, Efectivo, Inversiones, Deudas, Marketing | Tesorería, pasivos, publicidad |

> ⚠️ **Importante:** Los nombres de las pestañas deben coincidir exactamente.

---

## Hosting Gratuito

| Opción | Cómo |
|--------|------|
| **Local** | Doble clic en `index.html` |
| **Netlify** | Arrastrá la carpeta a [app.netlify.com](https://app.netlify.com) |
| **GitHub Pages** | Subí a un repo → Settings → Pages → Deploy |

---

## Mantenimiento

- **Datos manuales:** Se cargan desde el botón "📋 Cargar Datos" en el Dashboard
- **Datos automáticos:** Se leen solos de las hojas existentes
- **Actualización:** Cada 60 minutos (o botón "Sincronizar")
