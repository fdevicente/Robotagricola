# 04a — Excel: columnas nuevas en hojas existentes

## `Facturas` — agregar 4 columnas al final

| Col | Nombre | Tipo | Llenado por |
|---|---|---|---|
| Q | Categoria | str | categorizer (Claude) |
| R | Cultivo | str | categorizer (Claude) |
| S | Confianza | float 0-1 | categorizer |
| T | Categorizado_por | str | "claude"/"manual"/"heuristic" |

- Categoria: una de las 11 (ver 02a)
- Cultivo: NOGALES / CEREZOS / AVELLANOS / GENERAL
- Confianza <0.85 → marca REVISAR + Telegram

## `Cuenta Banco` — agregar 4 columnas al final

| Col | Nombre | Tipo | Notas |
|---|---|---|---|
| G | Tipo | str | "factura"/"sueldo"/"honorario"/"transferencia"/"venta_dolares"/"otro" |
| H | Categoria | str | Para no-factura, llenado por categorizer |
| I | Cultivo | str | Idem |
| J | Factura_linkeada | str | Si Tipo=factura: ref al N° factura |

- Tipo="venta_dolares" se usa para ingresos en USD que se ven como venta de divisas
- Linking con factura permite trazabilidad bidireccional

## Backup pre-cambio
Antes de la 1ra ejecución del bot v1.0:
`backups.backup_master("pre-fase-1")` deja la versión actual en Dropbox intacta.
