# 05c — UI Telegram: alertas y auto-disparos

## Alertas por nivel

| Nivel | Evento | Formato resumido |
|---|---|---|
| 🔴 | Saldo proyectado negativo | Mes crítico, causas, acciones sugeridas |
| 🔴 | Saldo bajo umbral ($36M) | Saldo actual, diferencia, qué viene |
| 🟡 | Categoría >90% del mes | Gastado vs proyectado, última factura grande |
| 🟡 | Factura grande vence en 3 días | Proveedor, monto, fecha vencimiento |
| 🟢 | Resumen semanal (lunes 8am) | Saldo, gastos semana, top cat, vencimientos |
| 🟢 | Cierre mes (día 1, 8am) | Ingresos, egresos, saldo, desvío + PDF |

## Auto-disparos (sin comando del usuario)

| Trigger | Mensaje |
|---|---|
| Guía cerezas 7d sin nueva Y >=8-dic | "X kg despachados. ¿Cerramos cosecha?" |
| Guía nueces 7d sin nueva Y >=23-may | "Y kg despachados. ¿Cerramos cosecha?" |
| Depósito sin match cosecha | "Llegó $X de [origen]. ¿Ingreso extraordinario?" |
| Categorización confianza <0.85 | Inline keyboard revisión inmediato |
| Match banco↔factura ambiguo | Inline keyboard selección |
| 1 día antes vencimiento factura | "Vence mañana: factura X $Y" |
| Día 1 mes + scraper + projector | Resumen mensual + PDF directorio |

## Deduplicación
Una alerta se envía 1 sola vez por combinación tipo+mes+categoría.
Se resetea al inicio de cada mes.
