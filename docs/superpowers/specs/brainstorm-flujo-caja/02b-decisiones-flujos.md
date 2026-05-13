# 02b — Decisiones: flujos, dashboard, alertas

## D11 — Matching banco↔factura: **B + Telegram para ambiguos**
- Match auto por monto + proveedor + nº factura
- Ambiguo → bot pregunta por Telegram con inline keyboard

## D12 — Dashboard táctico (B) + extras
- Saldo actual + proyección mensual hasta abril 2027
- Facturas pendientes desglosadas
- Gasto del mes vs proyectado por categoría
- Top 5 proveedores del mes
- Comparación gasto 2025 ajustado por hc
- Simulador replante interactivo
- "¿Cuánto puedo gastar?" widget

## D13 — Saldo mínimo de seguridad: **10% del gasto anual proyectado**
~$36M CLP con base 2025. Recalculado automáticamente cada año.

## D14 — Alertas estándar (Opción B) + reporte mensual formal
- 🔴 Saldo proyectado negativo
- 🔴 Saldo bajo umbral
- 🟡 Categoría >90% del mes
- 🟡 Factura por vencer (3 días antes)
- 🟢 Resumen semanal (lunes 8am)
- 🟢 Cierre de mes (día 1 + PDF directorio)

## D15 — Simulador replante: **simple affordability check**
No es ROI calculator. Solo: ¿alcanza la plata para replantar X hc?
Costo por hc calculado desde facturas categorizadas como Inversión/Replante del año anterior.

## D16 — Wizard post-cosecha por cultivo
Bot pregunta por Telegram: kg, exportadora, precio, cuotas, fechas, montos.
Usuario confirma/ajusta. Misma lógica para cerezas. Avellanos placeholder hasta 2028.

## D17 — Refresh banco: **1×/día a las 18:00**
Cron diario. Si falla → reintenta 19:00 y 20:00.

## D18 — Dashboard: **combinado Telegram + browser local**
Flask puerto 5000. Visión futura: app móvil (Fase 7).

## D19 — Arquitectura: **Enfoque 1 — extensión modular**
Módulo `cash_flow/` paralelo a tareas, inventario, vacaciones. Excel-only (sin SQLite).

## D20 — Año base proyección detallado
2025 como base. 2024 mostrado como columna comparativa. Ajustes manuales del usuario en hoja propia.
