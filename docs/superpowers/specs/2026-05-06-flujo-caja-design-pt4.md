# Spec — Fase 1: Flujo de Caja — Parte 4/4

## 8. Dashboard Web

Flask puerto 5000. 4 pestañas: Flujo de Caja | Cosechas | Replante | Reportes

### Pestaña Flujo de Caja
- KPIs: saldo actual, pendiente pago, disponible mes, mínimo seguridad
- Gráfico barras+línea: egresos proyectados vs reales, ingresos, saldo mes a mes
- Toggle "mostrar 2025 ajustado por hc" (barras semi-transparentes)
- Click en barra → drill-down: desglose por categoría, semáforo, top facturas desvío
- Tabla proyección mensual (meses × categorías)
- Alertas activas + facturas por vencer

### Pestaña Cosechas
- Card por temporada/cultivo: kg, exportadoras, ingresos esperados/recibidos, barra progreso
- Gastos asociados al cultivo (filtrado por Cultivo=X)
- Prorrateo gastos GENERAL proporcional a hectáreas
- Retorno neto, $/kg, $/hc
- Tabla comparativa histórica (2024/2025/2026)

### Pestaña Replante
- Input: cultivo + hectáreas → affordability check
- Resultado: alcanza con caja / requiere financiamiento / déficit
- Simulador crédito: tasa, plazo, cuota mensual, impacto en flujo
- Plan multi-año editable: transición nogales→avellanos año por año

### Pestaña Reportes
- Historial de PDFs mensuales generados, descargables

### Tech: Chart.js, Bootstrap/Tailwind, Jinja2. Sin framework JS pesado.

## 9. Ingresos USD (hallazgo B1)

Exportadoras USD depositan en ScotiaUSD (no ScotiaBCO).
Bot debe leer ScotiaUSD para detectar ingresos por exportadora.
Vitakai paga CLP directo en ScotiaBCO.
"Venta de dólares" = transferencia USD→CLP interna, no tiene nombre exportadora.

## 10. Errores y recuperación

- Excel locked → retry 5×4seg (existente)
- Scraper falla → retry 19h, 20h, alerta si 3 fallos
- Claude API down → cola pendiente, procesa al volver
- Match incorrecto → usuario corrige vía /revisar
- Master corrupto → restaurar de Backups/Master/snapshots/

## 11. Testing

- Unit: categorizer (20 gold set), matcher (50 pares), projector (datos fijos)
- Integration: factura end-to-end, scraper diario, wizard cosecha completo
- Validación: categorización vs PRESUPUESTO 2024-2025, saldo vs banco real

## 12. Dependencias bloqueantes

- Datos hectáreas 2024/2025/2026 (Daniel)
- Scraper Scotiabank funcional (existente parcial, completar)
- Limpieza de código del bot (en paralelo, otro chat)

## 13. Entregables Fase 1

1. Módulo cash_flow/ completo (7 archivos)
2. Router documentos + guías despacho
3. Historical importer (onboarding)
4. Dashboard /cash-flow (4 pestañas)
5. Backup automático Dropbox
6. MANUAL_TELEGRAM.md auto-generado
7. Tests unitarios + integración
