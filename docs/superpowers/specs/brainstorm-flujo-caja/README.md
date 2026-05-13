# Brainstorm — Fase 1: Flujo de Caja

**Inicio:** 2026-05-06
**Última actualización:** 2026-05-12
**Estado:** Diseño completo — pendiente spec final + user review

## Índice de archivos

| Archivo | Contenido |
|---|---|
| [01-contexto.md](01-contexto.md) | Objetivo, urgencia, hectáreas, archivos fuente |
| [02a-decisiones-categorias.md](02a-decisiones-categorias.md) | D1-D10: categorías, proyección, históricos |
| [02b-decisiones-flujos.md](02b-decisiones-flujos.md) | D11-D20: matching, dashboard, alertas |
| [02c-decisiones-extras.md](02c-decisiones-extras.md) | D21-D31: guías, backup, NLP, replante deuda |
| [03a-arquitectura-modulos.md](03a-arquitectura-modulos.md) | Estructura módulos + triggers |
| [03b-arquitectura-componentes.md](03b-arquitectura-componentes.md) | Detalle cada componente |
| [03c-arquitectura-flujos.md](03c-arquitectura-flujos.md) | 7 data flows paso a paso |
| [04a-excel-columnas-nuevas.md](04a-excel-columnas-nuevas.md) | Columnas nuevas Facturas + Cuenta Banco |
| [04b-excel-hojas-cosechas-guias.md](04b-excel-hojas-cosechas-guias.md) | Hojas Cosechas + Guias Despacho |
| [04c-excel-hojas-flujo-config.md](04c-excel-hojas-flujo-config.md) | Hojas Flujo Caja, Ajustes, Config, Hectareas |
| [05a-telegram-comandos.md](05a-telegram-comandos.md) | 14 comandos / nuevos |
| [05b-telegram-wizards.md](05b-telegram-wizards.md) | Wizard cosecha + inline keyboards |
| [05c-telegram-alertas.md](05c-telegram-alertas.md) | Alertas + auto-disparos |
| [B1-validacion-venta-dolares.md](B1-validacion-venta-dolares.md) | Hallazgo: ingresos USD en ScotiaUSD |
| [06a-dashboard-layout.md](06a-dashboard-layout.md) | Layout pestañas + sección principal |
| [06b-dashboard-grafico.md](06b-dashboard-grafico.md) | Gráfico presupuesto vs real + drill-down |
| [06c-dashboard-cosechas.md](06c-dashboard-cosechas.md) | Pestaña cosechas: gastos + retornos por cultivo |
| [06d-dashboard-replante.md](06d-dashboard-replante.md) | Simulador replante con escenario deuda |
| [07-errores-testing.md](07-errores-testing.md) | Errores, recuperación, unit/integration tests |

## Estado

- [x] Sección 1-3 — Arquitectura + componentes + data flows
- [x] Sección 4 — Cambios al Master Excel
- [x] Sección 5 — UI Telegram
- [x] B1 — Validación ingresos USD (ScotiaUSD)
- [x] Sección 6 — Dashboard Web (4 archivos)
- [x] Sección 7 — Errores + Testing
- [ ] **E1 — Spec final consolidado** ← siguiente
- [ ] E2 — Spec self-review
- [ ] E3 — User review + invocar writing-plans

## Pendientes del usuario

- Datos exactos hectáreas 2024/2025/2026 (Daniel revisando)

## Fases del proyecto

```
FASE 1 — Flujo de caja (URGENTE) ← diseño completo
FASE 2 — Bitácora Inteligente NLP
FASE 3 — Multi-Excel output
FASE 4 — Inventario automático
FASE 5 — Maquinaria con mantenciones
FASE 6 — Reportes + dashboard avanzado
FASE 7 — App móvil
```
