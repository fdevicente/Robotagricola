# Spec — Fase 1: Flujo de Caja — Parte 2/4

## 4. Arquitectura

### Módulos nuevos

```
modules/cash_flow/
  categorizer.py    — Claude clasifica factura/cargo → (categoría, cultivo, confianza)
  matcher.py        — Match banco↔factura, escribe Fecha Pago
  projector.py      — Proyección mes × categoría × cultivo
  income_wizard.py  — FSM Telegram cierre cosecha
  replante.py       — Affordability check con escenario deuda
  reporter.py       — PDF mensual formato directorio
  alerts.py         — Genera/dispara notificaciones Telegram

modules/documentos/router.py  — Clasifica tipo documento entrante
modules/guias_despacho/        — Procesa guías, acumula kg por temporada
modules/historical_importer.py — Onboarding: categoriza + importa histórico

infrastructure/backups.py      — Backup Master+código a Dropbox
infrastructure/manual_gen.py   — Auto-genera MANUAL_TELEGRAM.md
```

### Principios
- Excel = fuente única (Master.xlsx), sin SQLite
- Cache en RAM (pandas dataframe), invalidable
- Reutiliza infraestructura existente (Telegram, retry Excel, Claude)
- Compatible con limpieza de código en paralelo

### Triggers

| Evento | Componente |
|---|---|
| Factura nueva (Telegram) | categorizer al guardar |
| 18:00 diario | scraper → matcher → categorizer → projector |
| Lunes 8am | alerts.weekly_summary() |
| Día 1 mes 8am | reporter.monthly_close() |
| /cosecha | income_wizard.start() |
| /replante | replante.afford_check() |
| Categoría >90% | alerts.budget_warning() |
| Guía despacho nueva | guias_despacho + accumulator |

## 5. Data flows (7 flujos)

1. **Onboarding** (1×): backup → batch categorize 1351 facturas + banco → import cosechas desde Dropbox → import ScotiaUSD → revisar dudosos
2. **Diario 18:00**: scraper → matcher → categorizer no-factura → projector → alerts → backup
3. **Factura nueva**: router → facturas existente → categorizer → projector → alerts → backup
4. **Wizard cosecha**: FSM Telegram → hoja Cosechas → projector recalcula
5. **Replante check**: saldo proyectado - mínimo - pendientes = disponible ÷ costo/hc
6. **Reporte mensual**: reporter genera PDF → Telegram con adjunto
7. **Guía despacho**: router → extrae campos → Master.Guias Despacho → accumulator → auto-cierre cosecha si fecha límite
