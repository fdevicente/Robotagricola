# 07 — Manejo de errores y testing

## Errores y recuperación

| Escenario | Acción |
|---|---|
| Excel locked (otro proceso) | Retry 5 intentos, 4 seg espera (existente) |
| Scraper Scotiabank falla 18:00 | Retry 19:00, 20:00. Si 3 fallos → alerta Telegram |
| Claude API down | Categorización queda en cola "pendiente". Procesa al volver |
| Claude confianza <0.85 | Marca REVISAR, no asume categoría |
| Match banco↔factura incorrecto | Usuario corrige via /revisar. Bot aprende proveedor |
| Backup Dropbox falla | Retry 3x. Si falla → alerta "backup falló, revisar Dropbox" |
| Master corrupto | Restaurar último snapshot de Backups/Master/snapshots/ |
| Wizard cosecha interrumpido | FSM guarda estado parcial. Al reabrir continúa donde quedó |
| Guía despacho ilegible | Bot avisa "no pude leer", guarda PDF igual, pide ingreso manual |

## Principio: nunca perder datos silenciosamente

- Toda escritura a Master va con backup previo
- Todo error de categorización queda visible en /revisar
- Todo match ambiguo requiere confirmación humana
- Nunca sobrescribir Fecha Pago si ya tiene valor

## Testing

### Unit tests (pytest)

| Módulo | Qué testear |
|---|---|
| categorizer | Prompt → JSON correcto. 20 facturas etiquetadas manualmente como gold set |
| matcher | Dataset de 50 pares banco↔factura conocidos. Probar match, ambiguo, no-match |
| projector | Proyección con datos fijos → resultado esperado. Verificar escalamiento hc |
| replante | Cálculo con saldo fijo → afford/no-afford correcto |
| alerts | Simular saldo bajo umbral → alerta generada |

### Integration tests

| Test | Qué valida |
|---|---|
| Factura nueva end-to-end | Telegram → OCR → Claude → categoría → Master → backup |
| Scraper diario | Fetch → match → categorize → project → alert |
| Wizard cosecha | FSM completo → hoja Cosechas llena → projector actualizado |

### Validación datos históricos

- Comparar categorización Claude vs categorías en PRESUPUESTO 2024-2025 (CARGO)
- Comparar saldo proyectado vs saldo real banco (meses pasados = error 0)
- Comparar kg guías despacho importados vs totales en DATOS COSECHA
