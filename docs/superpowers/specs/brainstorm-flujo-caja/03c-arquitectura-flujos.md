# 03c — Arquitectura: data flows

## Flujo 1 — Onboarding inicial (1 sola vez)

```
1. backups.backup_master("pre-onboarding")
2. categorizer.batch_categorize_history()
   ├─ Lee 1351 facturas + 4710 cargos banco
   ├─ Claude clasifica con cache
   └─ Marca confianza <0.85 como REVISAR
3. historical_importer.import_cosechas()
   ├─ Lee DATOS COSECHA.xlsx
   ├─ Lee FXP.Ingresos + "2023 y 2024 cosecha"
   ├─ Detecta ventas de dólares en ScotiaBCO
   └─ Reconstruye Cosechas + Guias Despacho histórico
4. Bot Telegram: "X OK, Y a revisar. /revisar para empezar"
5. Usuario corrige dudosos (5-10 min)
6. backups.backup_master("post-onboarding")
```

## Flujo 2 — Diario 18:00

```
scraper.fetch_scotiabank()
   ▼
matcher.match_new_bank_movements()
   ├─ Match único → escribe Fecha Pago
   ├─ Ambiguo → cola Telegram
   └─ No-factura → categorizer
   ▼
categorizer.categorize_bank_movement() (no-factura)
   ▼
projector.invalidate_cache() + recompute()
   ▼
alerts.fire_pending()
   ▼
backups.backup_master("post-scraper-18h")
```

Tiempo: 2-3 min. Retry 19h/20h si falla.

## Flujo 3 — Factura nueva por Telegram

```
Usuario → foto factura
   ▼
documentos/router (detecta tipo)
   ▼ tipo=factura
facturas (existente): OCR → Claude → preview → guarda
   ▼
categorizer.categorize_invoice()
   ▼
projector.invalidate_cache()
   ▼
alerts.check_new_thresholds()
   ▼
backups.backup_master("post-factura")
```

## Flujo 4 — Wizard post-cosecha

```
/cosecha nogales
   ▼
FSM Telegram:
  kg total → exportadoras → precio × exportadora
  → cuotas → fechas → montos → liquidación
   ▼
Escribe en hoja Cosechas
   ▼
projector.invalidate_cache()
   ▼
Bot resume confirmación
```

## Flujo 5 — Replante check

```
/replante avellanos 4
   ▼
replante.afford_check("avellanos", 4)
   ├─ saldo_proyectado_fin_año
   ├─ - saldo_minimo (10%)
   ├─ - pendientes + proyectados
   ├─ costo/hc = histórico Inversión/Replante
   └─ ¿alcanza? + cuántas hc + monto deuda si no
   ▼
Dashboard widget + respuesta Telegram
```

## Flujo 6 — Reporte mensual (día 1, 8am)

```
reporter.monthly_close(mes_anterior)
   ├─ Resumen ejecutivo
   ├─ Tablas ingresos/egresos
   ├─ Top proveedores
   ├─ Comparativo año anterior
   └─ Proyección 3 meses
   ▼
Reportes/Cierre_YYYY-MM.pdf
   ▼
Telegram con PDF adjunto + 3 líneas resumen
```

## Flujo 7 — Guía de despacho nueva

```
Usuario → PDF guía por Telegram
   ▼
documentos/router → guía_despacho
   ▼
guias_despacho.process_guia()
   ├─ Claude extrae campos
   ├─ Preview Telegram
   ├─ Guarda en Master.Guias Despacho
   └─ Copia PDF a Dropbox/Documentos/Guias Despacho
   ▼
accumulator.update(cultivo, kg_temporada)
   ├─ Compara vs wizard si existe
   └─ Si difiere → alerta amarilla
   ▼
Auto-cierre cosecha:
  ├─ Cerezas ≥15-dic OR 7d sin guías → sugiere /cosecha
  └─ Nueces  ≥30-may OR 7d sin guías → sugiere /cosecha
   ▼
projector.invalidate_cache()
   ▼
Dashboard widget "Cosecha en curso"
```
