# Guía de Configuración — Dashboard Rosecal 2026

## Conexión Paso a Paso

### Paso 1: Abrir Apps Script

1. Abrí tu Google Sheets **"Pedidos Rosecal"**
2. Hacé clic en **Extensiones → Apps Script**
3. Borrá todo el código que aparece por defecto

### Paso 2: Pegar el código

Copiá y pegá **todo** este código:

```javascript
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = readOrders(ss);
    const monthly = readMonthlySummary(ss);
    const cashFlow = readCashFlow(ss);
    const cheques = readCheques(ss);
    const result = buildResponse(orders, monthly, cashFlow, cheques);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ══════ LEER PEDIDOS (Respuestas de formulario 1) ══════
function readOrders(ss) {
  const sheet = ss.getSheetByName('Respuestas de formulario 1');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0].map(v => String(v).trim().toLowerCase());

  const ci = {
    fecha: h.findIndex(x => x.includes('marca temporal')),
    cliente: h.findIndex(x => x.includes('nombre del cliente') || x.includes('cliente')),
    provincia: h.findIndex(x => x.includes('provincia')),
    importe: h.findIndex(x => x.includes('importe')),
    condicion: h.findIndex(x => x.includes('condici')),
    origen: h.findIndex(x => x.includes('origen')),
    status: h.findIndex(x => x.includes('status'))
  };

  const orders = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[ci.fecha]) continue;
    const imp = parseAmount(r[ci.importe >= 0 ? ci.importe : 0]);
    if (imp <= 0) continue;
    orders.push({
      date: fmtDate(r[ci.fecha]),
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

// ══════ LEER RESUMEN MENSUAL (Calculo Ganancias) ══════
function readMonthlySummary(ss) {
  const sheet = ss.getSheetByName('Calculo Ganancias');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0].map(v => String(v).trim().toLowerCase());

  const mi = h.findIndex(x => x.includes('mes'));
  const vti = h.findIndex(x => x.includes('ventas total'));
  const vci = h.findIndex(x => x.includes('ventas cobrad'));
  const ipi = h.findIndex(x => x.includes('inversi'));

  if (mi < 0) return [];
  const monthly = [];
  for (let i = 1; i < data.length; i++) {
    const mes = String(data[i][mi] || '').trim();
    if (!mes) continue;
    monthly.push({
      mes: mes,
      ventasTotal: Number(data[i][vti >= 0 ? vti : mi+1]) || 0,
      ventasCobrado: Number(data[i][vci >= 0 ? vci : mi+2]) || 0,
      inversionPub: Number(data[i][ipi >= 0 ? ipi : mi+3]) || 0
    });
  }
  return monthly;
}

// ══════ LEER FLUJO OPERATIVO (formato transpuesto) ══════
function readCashFlow(ss) {
  const sheet = ss.getSheetByName('Flujo Operativo');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // Row 0 = headers: [Conceptos/Meses, enero, febrero, marzo, ...]
  const months = [];
  for (let c = 1; c < data[0].length; c++) {
    const m = String(data[0][c]).trim().toLowerCase();
    if (m) months.push({ col: c, name: m });
  }

  const concepts = {};
  for (let i = 1; i < data.length; i++) {
    const label = String(data[i][0]).trim().toLowerCase();
    const vals = {};
    months.forEach(m => { vals[m.name] = Number(data[i][m.col]) || 0; });
    concepts[label] = vals;
  }

  return { months: months.map(m => m.name), concepts: concepts };
}

// ══════ LEER CHEQUES ══════
function readCheques(ss) {
  const sheet = ss.getSheetByName('Cheques');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const h = data[0].map(v => String(v).trim().toLowerCase());

  const ci = {
    numero: h.findIndex(x => x.includes('numero') || x.includes('número')),
    proveedor: h.findIndex(x => x.includes('proveedor')),
    vencimiento: h.findIndex(x => x.includes('vencimiento')),
    monto: h.findIndex(x => x.includes('monto')),
    estado: h.findIndex(x => x.includes('estado')),
    cierre: h.findIndex(x => x.includes('cierre'))
  };

  const cheques = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[ci.numero >= 0 ? ci.numero : 0]) continue;
    cheques.push({
      numero: String(r[ci.numero >= 0 ? ci.numero : 0]),
      proveedor: String(r[ci.proveedor >= 0 ? ci.proveedor : 1] || ''),
      vencimiento: fmtDate(r[ci.vencimiento >= 0 ? ci.vencimiento : 3]),
      monto: Number(r[ci.monto >= 0 ? ci.monto : 4]) || 0,
      estado: String(r[ci.estado >= 0 ? ci.estado : 5] || ''),
      cierre: fmtDate(r[ci.cierre >= 0 ? ci.cierre : 6])
    });
  }
  return cheques;
}

// ══════ CONSTRUIR RESPUESTA JSON ══════
function buildResponse(orders, monthly, cashFlow, cheques) {
  // --- Cash Flow KPIs ---
  let cfData = { currentMonth:'', ingresosMercadoPago:0, ingresosEfectivo:0,
    ingresosDeudaDuek:0, cajaFinal:0, prevIngresosMercadoPago:0,
    prevIngresosEfectivo:0, prevIngresosDeudaDuek:0, prevCajaFinal:0,
    monthlyLabels:[], monthlyCajaFinal:[] };

  if (cashFlow) {
    const ms = cashFlow.months;
    const c = cashFlow.concepts;
    // Find concept rows by partial match
    const mpKey = Object.keys(c).find(k => k.includes('mercado pago') || k.includes('ingresos m'));
    const efKey = Object.keys(c).find(k => k.includes('efectivo'));
    const ddKey = Object.keys(c).find(k => k.includes('deuda duek') || k.includes('duek'));
    const cajaKey = Object.keys(c).find(k => k.includes('caja final'));

    // Find last month with caja final data
    let curIdx = -1;
    if (cajaKey) {
      for (let i = ms.length - 1; i >= 0; i--) {
        if (c[cajaKey][ms[i]]) { curIdx = i; break; }
      }
    }
    if (curIdx < 0) curIdx = ms.length - 1;
    const curM = ms[curIdx], prevM = curIdx > 0 ? ms[curIdx - 1] : null;

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

    // Monthly caja final for chart
    if (cajaKey) {
      ms.forEach(m => {
        if (c[cajaKey][m]) {
          cfData.monthlyLabels.push(m);
          cfData.monthlyCajaFinal.push(c[cajaKey][m]);
        }
      });
    }
  }

  // --- Sales KPIs (from monthly summary) ---
  let slData = { currentMonth:'', ventasTotal:0, ventasCobrado:0,
    inversionPub:0, prevVentasTotal:0, prevVentasCobrado:0, prevInversionPub:0 };
  if (monthly.length > 0) {
    const last = monthly[monthly.length - 1];
    slData.currentMonth = last.mes;
    slData.ventasTotal = last.ventasTotal;
    slData.ventasCobrado = last.ventasCobrado;
    slData.inversionPub = last.inversionPub;
    if (monthly.length > 1) {
      const prev = monthly[monthly.length - 2];
      slData.prevVentasTotal = prev.ventasTotal;
      slData.prevVentasCobrado = prev.ventasCobrado;
      slData.prevInversionPub = prev.inversionPub;
    }
  }

  // --- Monthly Evolution ---
  const mEvolution = {
    labels: monthly.map(m => m.mes),
    ventasTotal: monthly.map(m => m.ventasTotal),
    ventasCobrado: monthly.map(m => m.ventasCobrado),
    inversionPub: monthly.map(m => m.inversionPub)
  };

  // --- Orders by Status ---
  const statusCount = {}, statusAmount = {};
  const provinceSet = new Set(), conditionSet = new Set();
  orders.forEach(o => {
    const s = o.status || 'SIN ESTADO';
    statusCount[s] = (statusCount[s] || 0) + 1;
    statusAmount[s] = (statusAmount[s] || 0) + o.amount;
    if (o.province) provinceSet.add(o.province);
    if (o.condition) conditionSet.add(o.condition);
  });

  // --- Cheques grouping ---
  const activeCheques = cheques.filter(c => (c.estado || '').toUpperCase() === 'ACTIVO');
  const chequesByMonth = {};
  let totalChequesActivos = 0;
  activeCheques.forEach(c => {
    const month = c.cierre || c.vencimiento || 'Sin fecha';
    const key = month.substring(0, 7); // YYYY-MM
    chequesByMonth[key] = (chequesByMonth[key] || 0) + c.monto;
    totalChequesActivos += c.monto;
  });
  const sortedChequeKeys = Object.keys(chequesByMonth).sort();

  // --- Comparison ---
  const comparison = [];
  for (let i = monthly.length - 2; i >= 0 && comparison.length < 3; i--) {
    comparison.push({ period: monthly[i].mes.charAt(0).toUpperCase() + monthly[i].mes.slice(1), value: monthly[i].ventasTotal });
  }

  // --- Transactions (last 30) ---
  const txs = orders.slice(-30).reverse();

  return {
    lastUpdate: new Date().toISOString(),
    cashFlow: cfData,
    sales: slData,
    monthlyEvolution: mEvolution,
    ordersByStatus: statusCount,
    orderAmountByStatus: statusAmount,
    cheques: {
      labels: sortedChequeKeys,
      montos: sortedChequeKeys.map(k => chequesByMonth[k]),
      totalActivos: totalChequesActivos,
      cantidad: activeCheques.length
    },
    comparison: comparison,
    transactions: txs,
    provinces: [...provinceSet].sort(),
    conditions: [...conditionSet].sort()
  };
}

// ══════ HELPERS ══════
function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  // Handle dd/MM/yyyy format
  const parts = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (parts) return parts[3] + '-' + parts[2].padStart(2,'0') + '-' + parts[1].padStart(2,'0');
  return s;
}

function parseAmount(v) {
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[^\d.,\-]/g, '');
  const clean = s.replace(/\./g, '').replace(',', '.');
  return parseFloat(clean) || 0;
}
```

### Paso 3: Guardar y publicar

1. Hacé clic en **💾 Guardar** (o Ctrl+S)
2. Hacé clic en **Implementar → Nueva implementación**
3. En **Tipo**, elegí **App web**
4. Configurá:
   - **Descripción:** Dashboard Rosecal API
   - **Ejecutar como:** Yo mismo
   - **Quién tiene acceso:** Cualquiera
5. Hacé clic en **Implementar**
6. **Autorizá** los permisos (es tu cuenta)
7. **Copiá la URL** resultante (empieza con `https://script.google.com/macros/s/...`)

### Paso 4: Conectar al Dashboard

1. Abrí `index.html` con un editor de texto
2. Buscá esta línea:
```javascript
APPS_SCRIPT_URL: '',
```
3. Pegá tu URL:
```javascript
APPS_SCRIPT_URL: 'https://script.google.com/macros/s/TU-ID-AQUI/exec',
```
4. Guardá el archivo

---

## Hojas que lee el Dashboard

El script lee automáticamente de estas pestañas (no necesitás cambiar nada):

| Pestaña | Qué lee | Para qué |
|---------|---------|----------|
| `Respuestas de formulario 1` | Pedidos (fecha, cliente, importe, provincia, status) | Tabla, filtros, gráfico de estados |
| `Calculo Ganancias` | Columnas Mes, Ventas Total, Ventas Cobrado, Inversión publicitaria | KPIs de ventas, evolución mensual, ROI |
| `Flujo Operativo` | Ingresos MP, Efectivo, Deuda Duek, Caja Final | KPIs de caja, gráfico de flujo |
| `Cheques` | Número, monto, estado, vencimiento | Gráfico de cheques activos |

> ⚠️ **Importante:** Los nombres de las pestañas deben coincidir exactamente. Si tenés diferencias, editá los nombres en el código del Apps Script.

---

## Hosting Gratuito

| Opción | Cómo |
|--------|------|
| **Local** | Doble clic en `index.html` |
| **Netlify** | Arrastrá la carpeta a [app.netlify.com](https://app.netlify.com) |
| **GitHub Pages** | Subí a un repo → Settings → Pages → Deploy |

---

## Mantenimiento

- **Agregar datos:** Simplemente agregá filas en tu Sheets. El dashboard los lee automáticamente.
- **Actualización:** Cada 60 minutos (o botón "Actualizar")
- **Cambiar intervalo:** Editá `REFRESH_INTERVAL` en el HTML (ej: `5 * 60 * 1000` = 5 min)

## Soporte

Si algo no anda:
1. Abrí la consola del navegador (F12 → Console)
2. Verificá que la URL del Apps Script esté bien pegada
3. Verificá que los nombres de las pestañas coincidan exactamente
