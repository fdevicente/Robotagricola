# Handoff — Bot Agrícola Santa Elisa

**Fecha:** 2026-05-22
**Estado:** ✅ Bot funcional + refactor completo + cash flow operativo

---

## Lo que hicimos en esta conversación

### Plan 1-6 (Cash Flow Fase 1) — ✅ COMPLETO

| Plan | Descripción | Tests |
|---|---|---|
| 1 Fundación | Master + columnas + backups | 15 ✅ |
| 2 Categorización | Claude + cache + onboarding | 26 ✅ |
| 3 Banco + Matching | Scraper + matcher + linking | 21 ✅ |
| 4 Motor proyección | Projector mes×cat×cultivo | 22 ✅ |
| 5 Telegram | /proyeccion /categoria /cosecha + alertas | 22+ ✅ |
| 6 Dashboard Web | /cash-flow Flask + Chart.js + replante | 2 ✅ |

**Resultados reales:**
- 1324 facturas categorizadas (Claude + 302 vía cruce Camarico)
- Master.Cosechas 2026 cargado: Valbifrut $223.6M CLP recibidos + Pacific Nuts 4 cuotas
- 7 auto-matches banco↔factura aplicados
- Proyección 12 meses: ingresos $383.6M, egresos $488.4M, saldo final $25.8M

### Refactor main.py — ✅ COMPLETO

main.py: **1689 → 270 líneas** (-84%)

Estructura final:
- `utils/` — formatting, keyboards, parsing
- `handlers/` — facturas, finanzas, tareas, personal, inventario_h, chat, cash_flow_cmds, cash_flow_jobs
- `modules/cash_flow/` — projector, categorizer, matcher, alerts, replante, cosecha_wizard, historical_importer
- `infrastructure/` — backups
- `main.py` — arranque + registro

### Graphify instalado

`graphifyy 0.8.15` registrado en Claude Code. Comando: `/graphify .` desde Robot/ para generar grafo del proyecto.

---

## Scripts disponibles

| Script | Uso |
|---|---|
| `main.py` | Bot Telegram (polling) |
| `setup_cash_flow.py` | Setup inicial (1 vez) |
| `onboarding_cash_flow.py` | Categoriza histórico Claude |
| `daily_banco_18h.py` | Cron 18:00 banco+match |
| `recalc_flujo_caja.py` | Regenera hoja Flujo Caja |
| `src/dashboard.py` | Dashboard Flask :5000 |

Scheduler Windows: `AgricolaSantaElisa-DailyBanco-18h` registrado.

---

## Comandos Telegram nuevos

- `/proyeccion [N]` — saldo proyectado N meses
- `/categoria <nombre>` — gasto mensual por categoría
- `/cosecha <CULTIVO>` — wizard FSM cierre cosecha
- Job lunes 8am: resumen semanal automático

---

## Próximos pasos sugeridos

1. **Probar `/graphify .`** en próxima sesión
2. **Revisar 286 facturas con `Categoria=REVISAR`** (baja confianza Claude)
3. **Correr `daily_banco_18h.py`** cuando haya datos banco 2026 nuevos
4. **Verificar dashboard** http://localhost:5000/cash-flow

---

## Issues conocidos / pendientes

- Master.Cuenta Banco solo tiene hasta ene-2025 — scraper Scotiabank traerá 2026 cuando se corra
- 4220 cargos banco sin clasificar (esperando categorizer Claude)
- 776 movimientos ambiguos en matcher (resolución manual via Telegram en Plan futuro)
