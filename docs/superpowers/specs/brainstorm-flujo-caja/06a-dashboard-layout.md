# 06a — Dashboard Web: layout general

## Estructura de pestañas

```
[ Flujo de Caja ] [ Cosechas ] [ Replante ] [ Reportes ]
```

Flask existente puerto 5000. Cada pestaña es una ruta:
- `/cash-flow` — vista principal flujo de caja (default)
- `/cash-flow/cosechas` — detalle por cosecha/cultivo
- `/cash-flow/replante` — simulador replante con deuda
- `/cash-flow/reportes` — historial de reportes PDF

## Pestaña: Flujo de Caja (principal)

Layout vertical, secciones apiladas:

### Sección 1 — KPIs arriba (cards)
```
┌─────────────┐ ┌─────────────┐ ┌─────────────┐ ┌─────────────┐
│ Saldo actual │ │ Pendiente   │ │ Disponible  │ │ Mínimo seg. │
│ $130.6M     │ │ pago: $50M  │ │ mes: $44M   │ │ $36M (10%)  │
│ 🟢 OK      │ │ 🟡 ojo     │ │             │ │             │
└─────────────┘ └─────────────┘ └─────────────┘ └─────────────┘
```

### Sección 2 — Gráfico flujo mes a mes (ver 06b para detalle)
Barras apiladas con línea de saldo proyectado + real.

### Sección 3 — Tabla proyección mensual
Tabla con meses como columnas, filas: ingresos, egresos por categoría, saldo.
Celdas clickeables → drill-down al "por qué".

### Sección 4 — Alertas activas
Lista de alertas vigentes con nivel de severidad.

### Sección 5 — Facturas por vencer
Top 10 facturas pendientes ordenadas por fecha vencimiento.

## Tech stack dashboard
- Flask + Jinja2 templates (existente)
- Chart.js para gráficos
- Bootstrap o Tailwind para layout responsive
- Sin framework JS pesado (no React/Vue)
