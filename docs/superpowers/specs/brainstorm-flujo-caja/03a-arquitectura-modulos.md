# 03a — Arquitectura: estructura de módulos

## Vista general

```
robot/
├── modules/
│   ├── facturas/           (existente, agregar categorización al guardar)
│   ├── boletas/            (existente)
│   ├── tareas/             (existente)
│   ├── inventario/         (existente, profundiza en Fase 4)
│   ├── vacaciones/         (existente, mejoras en Fase 6)
│   ├── caja_chica/         (existente)
│   ├── documentos/         ← NUEVO (router de tipos de documento)
│   │   └── router.py
│   ├── guias_despacho/     ← NUEVO
│   │   └── process_guia.py
│   ├── cash_flow/          ← NUEVO (corazón de Fase 1)
│   │   ├── categorizer.py
│   │   ├── matcher.py
│   │   ├── projector.py
│   │   ├── income_wizard.py
│   │   ├── replante.py
│   │   ├── reporter.py
│   │   └── alerts.py
│   └── historical_importer.py  ← NUEVO (onboarding 1×)
├── infrastructure/         ← NUEVO bloque cross-cutting
│   ├── backups.py
│   ├── manual_gen.py
│   └── (bitacora_nlp.py)   ← futuro Fase 2
├── scraper/                (existente parcial, completar a 18:00)
└── dashboard/              (existente Flask, +ruta /cash-flow)
```

## Principios de diseño

- **Excel = fuente única de verdad** (Master.xlsx)
- **Sin SQLite** (cache en RAM con invalidación)
- **Reutiliza infraestructura existente** (Telegram bot, retry logic Excel, Claude integration)
- **Cada módulo tiene una responsabilidad clara**
- **Compatible con la limpieza de código que se está haciendo en paralelo**

## Triggers (cuándo corre cada cosa)

| Evento | Disparador | Componente |
|---|---|---|
| Factura nueva ingresa | Telegram | `categorizer` al guardar |
| 18:00 cada día | Cron diario | `scraper → matcher → categorizer → projector` |
| Lunes 8am | Cron semanal | `alerts.weekly_summary()` |
| Día 1 mes, 8am | Cron mensual | `reporter.monthly_close()` + alertas |
| Cosecha terminada | `/cosecha` o auto | `income_wizard.start()` |
| Replante check | Dashboard o `/replante` | `replante.afford_check()` |
| Categoría >90% | Después de update | `alerts.budget_warning()` |
| Guía despacho nueva | Telegram | `guias_despacho` + accumulator |
