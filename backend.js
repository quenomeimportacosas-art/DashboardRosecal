var gasBackend = {
    configurarHojas: function() {
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
    },

    doGet: function(e) {
        try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var p = (e && e.parameter) ? e.parameter : {};
            var action = p.action || 'read';

            if (action === 'write') {
                var sheet = ss.getSheetByName('Datos Manuales');
                if (!sheet) return this.jsonOut({ error: 'La hoja no existe. Ejecutá configurarHojas() primero.' });
                sheet.appendRow([
                    new Date(), this.toNum(p.bancoCuenta), this.toNum(p.saldoMP),
                    this.toNum(p.efectivoCaja), this.toNum(p.inversiones), this.toNum(p.deudaProveedores),
                    this.toNum(p.deudaServicios), this.toNum(p.deudaMP), this.toNum(p.inversionPublicidad),
                    this.toNum(p.gastoMarketing), this.toNum(p.ventasCorporativas), this.toNum(p.gastosOperativos)
                ]);
                return this.jsonOut({ success: true });
            }

            if (action === 'update_cheque') {
                var sheet = ss.getSheetByName('Cheques');
                if(!sheet) return this.jsonOut({error: 'No existe hoja Cheques'});
                var data = sheet.getDataRange().getValues();
                var numCol = -1, estCol = -1;
                var h = data[0].map(function(k){ return String(k).trim().toLowerCase() });
                for(var j=0; j<h.length; j++) {
                    if(h[j].indexOf('numero') >= 0 || h[j].indexOf('número') >= 0) numCol = j;
                    if(h[j].indexOf('estado') >= 0) estCol = j;
                }
                if(numCol < 0 || estCol < 0) return this.jsonOut({error: 'Columnas no encontradas en hoja Cheques'});
                
                var found = false;
                for (var i = 1; i < data.length; i++) {
                    if (String(data[i][numCol]) === String(p.numero)) {
                        sheet.getRange(i + 1, estCol + 1).setValue(p.nuevoEstado);
                        found = true;
                        break;
                    }
                }
                return this.jsonOut({ success: found });
            }

            var targetDate = p.month ? new Date(p.month + '-01T12:00:00') : new Date();
            
            var condicionesMap = this.readCondiciones(ss);
            var orders = this.readOrders(ss, targetDate, condicionesMap);
            var monthly = this.readMonthlySummary(ss);
            var rawGanancias = this.readRawGanancias(ss, targetDate);
            var cheques = this.readCheques(ss);
            var manual = this.readManualData(ss, targetDate);
            
            var result = this.buildResponse(orders, monthly, rawGanancias, cheques, manual, targetDate);
            return this.jsonOut(result);

        } catch (err) {
            return this.jsonOut({ error: err.message, stack: err.stack });
        }
    },

    jsonOut: function(obj) {
        return ContentService.createTextOutput(JSON.stringify(obj))
            .setMimeType(ContentService.MimeType.JSON);
    },

    toNum: function(v) { return Number(v) || 0; },

    readManualData: function(ss, targetDate) {
        var createEmpty = function() { return { bancoCuenta: 0, saldoMP: 0, efectivoCaja: 0, inversiones: 0, deudaProveedores: 0, deudaServicios: 0, deudaMP: 0, inversionPublicidad: 0, gastoMarketing: 0, ventasCorporativas: 0, gastosOperativos: 0 }; };
        var sheet = ss.getSheetByName('Datos Manuales');
        
        var currentSum = createEmpty();
        var prevSum = createEmpty();
        currentSum.prev = prevSum;

        if (!sheet) return currentSum;
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return currentSum;

        var targetYMonth = targetDate.getFullYear() + '-' + this.padZero(targetDate.getMonth() + 1);
        var prevDDate = new Date(targetDate.getFullYear(), targetDate.getMonth() - 1, 1);
        var prevYMonth = prevDDate.getFullYear() + '-' + this.padZero(prevDDate.getMonth() + 1);

        var keys = Object.keys(currentSum).filter(function(k) { return k !== 'prev'; });

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            if (!row[0]) continue;
            var dStr = this.fmtDate(row[0]);
            
            var isCurrent = dStr.indexOf(targetYMonth) === 0;
            var isPrev = dStr.indexOf(prevYMonth) === 0;

            if (isCurrent || isPrev) {
                var dObj = { 
                    bancoCuenta: this.toNum(row[1]), saldoMP: this.toNum(row[2]), 
                    efectivoCaja: this.toNum(row[3]), inversiones: this.toNum(row[4]), 
                    deudaProveedores: this.toNum(row[5]), deudaServicios: this.toNum(row[6]), 
                    deudaMP: this.toNum(row[7]), inversionPublicidad: this.toNum(row[8]), 
                    gastoMarketing: this.toNum(row[9]), ventasCorporativas: this.toNum(row[10]), 
                    gastosOperativos: this.toNum(row[11]) 
                };

                for (var k = 0; k < keys.length; k++) {
                    if (isCurrent) currentSum[keys[k]] += dObj[keys[k]];
                    if (isPrev) prevSum[keys[k]] += dObj[keys[k]];
                }
            }
        }
        return currentSum;
    },

    readRawGanancias: function(ss, targetDate) {
        var sheet = ss.getSheetByName('Calculo Ganancias');
        if (!sheet) return { deudaClientes: 0, totalFacturado: 0 };
        var data = sheet.getDataRange().getValues();
        var deudaClientes = 0, totalFacturado = 0;
        var curMonth = targetDate.getMonth(), curYear = targetDate.getFullYear();

        for (var i = 1; i < data.length; i++) {
            var dateVal = data[i][0], amountVal = this.toNum(data[i][1]), statusVal = String(data[i][2] || '').trim().toUpperCase();
            if (dateVal instanceof Date && dateVal.getMonth() === curMonth && dateVal.getFullYear() === curYear) {
                if(statusVal !== 'ANULADO') totalFacturado += amountVal;
                deudaClientes += amountVal;
                if (statusVal === 'COBRADO' || statusVal === 'ANULADO') deudaClientes -= amountVal;
            }
        }
        return { deudaClientes: deudaClientes, totalFacturado: totalFacturado };
    },

    // Lee condiciones de pedido desde "Respuestas de formulario 1"
    // Devuelve un mapa: "fecha|importe" → condición para cruzar con Calculo Ganancias
    readCondiciones: function(ss) {
        var sheet = ss.getSheetByName('Respuestas de formulario 1');
        if (!sheet) return {};
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return {};
        var h = data[0].map(function(k){ return String(k).trim().toLowerCase() });
        
        var ci = {
            fecha: this.findCol(h, 'marca temporal', 'fecha'),
            importe: this.findCol(h, 'importe factur', 'importe original', 'importe'),
            condicion: this.findCol(h, 'condicion', 'condición')
        };
        if (ci.fecha < 0) ci.fecha = 0;
        if (ci.importe < 0) ci.importe = 1;
        if (ci.condicion < 0) return {};
        
        var mapa = {};
        for (var i = 1; i < data.length; i++) {
            var r = data[i];
            if (!r[ci.fecha]) continue;
            var localDate = null;
            if (r[ci.fecha] instanceof Date) localDate = r[ci.fecha];
            else {
                var parsed = new Date(r[ci.fecha]);
                if (!isNaN(parsed.getTime())) localDate = parsed;
            }
            if (!localDate) continue;
            var dateStr = localDate.getFullYear() + '-' + this.padZero(localDate.getMonth() + 1) + '-' + this.padZero(localDate.getDate());
            var imp = this.parseAmount(r[ci.importe]);
            var cond = String(r[ci.condicion] || '').trim().toUpperCase();
            // Clave: fecha + importe para cruzar
            var key = dateStr + '|' + Math.round(imp);
            mapa[key] = cond;
        }
        return mapa;
    },

    readOrders: function(ss, targetDate, condicionesMap) {
        var sheet = ss.getSheetByName('Calculo Ganancias');
        if (!sheet) return { current: [], global: { totalBruto: 0, totalCobrado: 0, porMes: {}, totalDeuda: 0 } };
        var data = sheet.getDataRange().getValues();
        if (data.length < 2) return { current: [], global: { totalBruto: 0, totalCobrado: 0, porMes: {}, totalDeuda: 0 } };
        var h = data[0].map(function(k){ return String(k).trim().toLowerCase() });
        
        var ci = {
            fecha: this.findCol(h, 'marca temporal'),
            importe: this.findCol(h, 'importe factur', 'importe original', 'importe'),
            status: this.findCol(h, 'status del', 'status')
        };
        // Fallbacks directos si cambian los nombres en base a la foto compartida
        if (ci.fecha < 0) ci.fecha = 0;
        if (ci.importe < 0) ci.importe = 1;
        if (ci.status < 0) ci.status = 2;

        var targetYMonth = targetDate.getFullYear() + '-' + this.padZero(targetDate.getMonth() + 1);
        var orders = [];
        var pendingAll = [];
        var glob = { totalBruto: 0, totalCobrado: 0, porMes: {}, totalDeuda: 0 };
        
        for (var i = 1; i < data.length; i++) {
            var r = data[i];
            if (!r[ci.fecha]) continue;
            
            var rawDate = r[ci.fecha];
            var tsFull = "";
            var localDate = null;
            if (rawDate instanceof Date) { localDate = rawDate; }
            else { 
                var parsed = new Date(rawDate);
                if (!isNaN(parsed.getTime())) localDate = parsed; 
            }
            if(!localDate) continue;
            // Usar fecha local en vez de toISOString() para evitar desfase por zona horaria
            tsFull = localDate.getFullYear() + '-' + this.padZero(localDate.getMonth() + 1) + '-' + this.padZero(localDate.getDate()) + 'T' + this.padZero(localDate.getHours()) + ':' + this.padZero(localDate.getMinutes()) + ':' + this.padZero(localDate.getSeconds());
            
            var yMonth = tsFull.substring(0, 7);
            
            var imp = this.parseAmount(r[ci.importe]);
            if (imp <= 0) continue;

            var st = String(r[ci.status] || '').trim().toUpperCase();
            
            // Buscar condición del pedido cruzando con Respuestas de formulario
            var dateOnly = tsFull.substring(0, 10);
            var condKey = dateOnly + '|' + Math.round(imp);
            var condicion = (condicionesMap && condicionesMap[condKey]) ? condicionesMap[condKey] : '';
            
            if (st !== 'ANULADO') {
                glob.totalBruto += imp;
                if (!glob.porMes[yMonth]) glob.porMes[yMonth] = { bruto: 0, cobrado: 0 };
                glob.porMes[yMonth].bruto += imp;
                
                if (st === 'COBRADO') {
                    glob.totalCobrado += imp;
                    glob.porMes[yMonth].cobrado += imp;
                } else {
                    glob.totalDeuda += imp;
                    // Recolectar TODOS los pedidos pendientes de cualquier mes
                    pendingAll.push({ date: tsFull, amount: imp, status: st, condicion: condicion });
                }
            }
            
            if (yMonth === targetYMonth) {
                orders.push({
                    date: tsFull,
                    amount: imp,
                    status: st,
                    condicion: condicion
                });
            }
        }
        return { current: orders, global: glob, pendingAll: pendingAll };
    },

    readMonthlySummary: function(ss) {
        var sheet = ss.getSheetByName('Calculo Ganancias');
        if (!sheet) return [];
        var data = sheet.getDataRange().getValues();
        var h = data.length > 0 ? data[0].map(function(k){ return String(k).trim().toLowerCase() }) : [];

        var mi = this.findCol(h, 'mes');
        var cli = this.findCol(h, 'costo logistica', 'logistica');
        var cbi = this.findCol(h, 'costo beato', 'beato');
        
        if (mi < 0 || data.length < 2) return [];
        var monthly = [];
        for (var i = 1; i < data.length; i++) {
            var mes = String(data[i][mi] || '').trim();
            if (!mes) continue;
            monthly.push({
            mes: mes,
            costoBeato: cbi >= 0 ? this.toNum(data[i][cbi]) : 0,
            costoLogistica: cli >= 0 ? this.toNum(data[i][cli]) : 0
            });
        }
        return monthly;
    },

    readCheques: function(ss) {
        var sheet = ss.getSheetByName('Cheques');
        if (!sheet) return [];
        var data = sheet.getDataRange().getValues();
        var h = data.length > 0 ? data[0].map(function(k){ return String(k).trim().toLowerCase() }) : [];

        var ci = { numero: this.findCol(h, 'numero', 'número'), proveedor: this.findCol(h, 'proveedor'), vencimiento: this.findCol(h, 'vencimiento'), monto: this.findCol(h, 'monto'), estado: this.findCol(h, 'estado') };
        var cheques = [];
        
        for (var i = 1; i < data.length; i++) {
            var r = data[i];
            if (!r[ci.numero >= 0 ? ci.numero : 0]) continue;
            var est = String(r[ci.estado >= 0 ? ci.estado : 5] || '').trim().toUpperCase();
            cheques.push({
                numero: String(r[ci.numero >= 0 ? ci.numero : 0] || ''),
                proveedor: String(r[ci.proveedor >= 0 ? ci.proveedor : 1] || ''),
                vencimiento: this.fmtDate(r[ci.vencimiento >= 0 ? ci.vencimiento : 3]),
                monto: this.toNum(r[ci.monto >= 0 ? ci.monto : 4]),
                estado: est || 'PENDIENTE'
            });
        }
        return cheques;
    },

    buildResponse: function(orders, monthly, rawGanancias, cheques, manual, targetDate) {
        var currentMonthName = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'][targetDate.getMonth()];
        var curYearMonth = targetDate.getFullYear() + '-' + this.padZero(targetDate.getMonth() + 1);
        
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
            orders: orders.current, // Export full month orders with FULL timestamp
            pendingAll: orders.pendingAll, // Todos los pedidos pendientes de cualquier mes
            globalStats: orders.global // Agregado para histórico
        };
    },

    findCol: function(headers, k1, k2) {
        for (var i = 0; i < headers.length; i++) {
            if (headers[i].indexOf(k1) >= 0) return i;
            if (k2 && headers[i].indexOf(k2) >= 0) return i;
        }
        return -1;
    },

    padZero: function(n) { return n < 10 ? '0' + n : '' + n; },

    fmtDate: function(v) {
        if (!v) return '';
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var s = String(v).trim(), parts = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (parts) return parts[3] + '-' + this.padZero(parseInt(parts[2])) + '-' + this.padZero(parseInt(parts[1]));
        return s;
    },

    parseAmount: function(v) {
        if (typeof v === 'number') return v;
        return parseFloat(String(v).replace(/[^\d.,\-]/g, '').replace(/\./g, '').replace(',', '.')) || 0;
    }
};
