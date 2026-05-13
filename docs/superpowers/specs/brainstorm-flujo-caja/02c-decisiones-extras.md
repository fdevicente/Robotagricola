# 02c — Decisiones: guías, backups, NLP, extras

## D21 — Backup automático en Dropbox
- `Dropbox/Agricola Santa Elisa/Backups/Master/current.xlsx` (último)
- `Backups/Master/snapshots/YYYY-MM-DD_HH-MM.xlsx` (rotación 30 días)
- `Backups/Robot/current/` (código completo, semanal)

## D22 — Documento manual Telegram auto-generado
`docs/MANUAL_TELEGRAM.md` con comandos, alertas, palabras clave, ejemplos.
Se regenera automáticamente cuando se agrega un comando o alerta nueva.

## D23 — Bitácora Inteligente = **Fase 2 nueva**
Capa NLP que entiende lenguaje natural del cuidador y rutea al módulo correcto.
Habilita el resto de las fases con UX natural.

## D24 — Re-secuencia de fases
```
FASE 1 — Flujo de caja (URGENTE) ← actual
FASE 2 — Bitácora Inteligente NLP
FASE 3 — Multi-Excel output
FASE 4 — Inventario automático
FASE 5 — Maquinaria con mantenciones
FASE 6 — Reportes + dashboard avanzado + vacaciones mejoradas
FASE 7 — App móvil
```

## D25 — Guías de Despacho como tipo de documento nuevo
Router en la entrada del bot clasifica:
factura | boleta | guía_despacho | otro
Cada tipo se procesa con su pipeline.

## D26 — Carpeta Dropbox documentos
```
Documentos/
├── Guias Despacho/
├── Certificados SAG/
├── Contratos/
└── Otros/
```

## D27 — Fechas límite de cosecha (hardcoded)
- Cerezas: máximo 15 diciembre
- Nueces: máximo 30 mayo
Bot auto-sugiere cerrar cosecha al acercarse o tras 7 días sin guías.

## D28 — Cosechas históricas: importar automáticamente
Fuente: `DATOS COSECHA.xlsx` (Dropbox CAMARICO) + `FXP.Ingresos` + FXP `2023 y 2024 cosecha`.
Reconstruye Cosechas y Guias Despacho con flag "histórico".

## D29 — Flujo Caja contrastado con banco
Cada mes muestra Saldo Inicial proyectado + Saldo Inicial REAL (de banco) + Diferencia.
Alerta amarilla si diferencia >5%.

## D30 — Ingresos extraordinarios
Línea dinámica en Flujo Caja. Bot detecta depósitos sin match con cosechas
y pregunta tipo: devolución IVA / venta equipo / indemnización / otro.

## D31 — Replante con deuda (NUEVO)
Replante de avellanos se hace igual aunque haya que endeudarse.
Simulador muestra: ¿alcanza con caja? Si no → ¿cuánto endeudamiento se necesita?
