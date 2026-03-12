# Estructura Google Sheets — Dashboard Rosecal 2026

## Hojas que YA tenés (no tocar)

El Dashboard lee directamente de estas hojas. No necesitan cambios.

| Hoja | Datos que lee el script | Para qué |
|------|------------------------|----------|
| `Respuestas de formulario 1` | Fecha, Cliente, Importe, Provincia, Condición, Origen, Status | Tabla de pedidos, gráfico por estado, Deuda Duek, Ventas ML |
| `Calculo Ganancias` | Col A (fecha), B (importe), C (status), H (mes) | Ventas Telefónicas, Deuda Clientes, Evolución mensual |
| `Flujo Operativo` | Conceptos transpuestos por mes | KPIs de caja, Mercado Pago, Efectivo, Caja Final |
| `Cheques` | Número, Proveedor, Monto, Estado, Vencimiento | Gráfico de cheques activos, Deuda Cheques |

---

## Hoja NUEVA: `Datos Manuales`

Esta hoja se crea automáticamente con el botón de setup del Apps Script, o al guardar datos desde el Dashboard.

Se llena **desde el botón "📋 Cargar Datos" del Dashboard** (no hace falta abrir Google Sheets).

| Columna | Ejemplo | Fuente |
|---------|---------|--------|
| Fecha | 11/03/2026 | Automático (fecha del día) |
| Banco Cta Cte | $ 1.500.000 | Manual (Home Banking) |
| Saldo Mercado Pago | $ 850.000 | Manual (App MP) |
| Efectivo Caja | $ 120.000 | Manual (conteo físico) |
| Inversiones | $ 5.000.000 | Manual (FCI, Plazos fijos) |
| Deuda Proveedores | $ 2.300.000 | Manual (facturas pendientes) |
| Deuda Servicios | $ 150.000 | Manual (luz, internet, etc.) |
| Deuda Mercado Pago | $ 0 | Manual (créditos MP) |
| Inversion Publicidad | $ 300.000 | Manual (gasto Ads semanal) |
| Gasto Marketing | $ 100.000 | Manual (agencia, diseño) |
| Ventas Corporativas | $ 3.000.000 | Manual (mayoristas) |

---

## Cálculos que hace el Dashboard (automáticos)

Estos valores NO se guardan en ninguna hoja. El Dashboard los calcula al vuelo:

| Métrica | Fórmula |
|---------|---------|
| **Total Liquidez** | Banco + Saldo MP + Efectivo + Inversiones |
| **Total Pasivo** | Deuda Proveedores + Deuda Servicios + Deuda MP |
| **Posición Neta** | Total Liquidez − Total Pasivo |
| **Deuda Clientes** | Ventas Total − Ventas Cobrado (del mes, hoja Calculo Ganancias) |
| **Deuda Duek** | Suma de pedidos con status "DEUDA DUEK" |
| **Ventas Telefónicas** | Importe facturado del mes con status ≠ ANULADO |
| **ROI Publicitario** | Ventas Cobrado / Inversión Publicidad |
