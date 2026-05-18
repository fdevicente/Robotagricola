# Plan 6: Dashboard Web — Cash Flow

**Goal:** Pestaña `/cash-flow` con KPIs + tabla + grafico + simulador replante.

**Tasks:**
1. API `/api/cash-flow` devuelve proyeccion 12 meses (JSON)
2. API `/api/cash-flow/replante` POST cultivo+hc → JSON afford check
3. Template `cash_flow.html` con Chart.js + link en nav

Reusa: modules/cash_flow/projector.py + replante.py.
