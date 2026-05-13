# 04c — Excel: hojas Flujo Caja, Ajustes, Config, Hectareas

## Hoja nueva: `Flujo Caja` (solo lectura, regenerada por projector)

Estructura: meses como columnas, categorías como filas.

```
                          May-26  Jun-26  Jul-26 ... Abr-27
SALDO INICIAL
  Proyectado
  Real (banco)
  Diferencia

INGRESOS
  Valbifrut adelanto
  Pacific Nuts adelanto
  Pacific Nuts liquidación
  Cerezas
  Ingresos extraordinarios
  TOTAL INGRESOS

EGRESOS (11 categorías)
  Mano obra planta
  Mano obra temporal
  Fertilizantes
  ... (las 11)
  TOTAL EGRESOS

SALDO CIERRE MES
  Proyectado
  Real (cuando el mes cierra)
  Comparación 2025 ajustado por hc
  Desviación %
```

No se edita a mano. Se regenera cada vez que projector recalcula.

## Hoja nueva: `Ajustes Manuales`

| Col | Nombre | Ejemplo |
|---|---|---|
| A | Fecha agregado | 2026-05-06 |
| B | Mes proyectado | 2026-07 |
| C | Categoria | Riego |
| D | Cultivo | GENERAL |
| E | Monto | +5000000 |
| F | Razón | "Bomba nueva" |
| G | Activo | TRUE/FALSE |

Se suman al gasto base proyectado. Desactivables sin borrar.

## Hoja nueva: `Config`

| Parámetro | Valor |
|---|---|
| saldo_minimo_pct | 0.10 |
| año_base_proyeccion | 2025 |
| fecha_limite_cerezas | 12-15 |
| fecha_limite_nueces | 05-30 |
| dias_sin_guia_cierre | 7 |
| umbral_alerta_cat_pct | 0.90 |
| umbral_confianza | 0.85 |
| ventana_match_dias | 15 |
| dropbox_backup_path | (ruta) |
| usd_clp_estimado | 1000 |

## Hoja nueva: `Hectareas`

| Año | Nogales | Cerezos | Avellanos | Total | Notas |
|---|---|---|---|---|---|
| 2024 | 65 hc | 1.8 hc | 0 hc | 66.8 hc | Sin avellanos |
| 2025 | 54 hc | 3.8 hc | 11.5 hc | 69.3 hc | Inicio replante avellanos |
| 2026 | 43 hc | 3.8 hc | 26.5 hc | 73.3 hc | +15 hc avellanos, -11 nogales |

Factores de escalamiento 2025→2026:
- Nogales: 43/54 = 0.796 (baja 20%)
- Cerezos: 3.8/3.8 = 1.0 (sin cambio)
- Avellanos: 26.5/11.5 = 2.30 (sube 130%, pero sin producción aún)

## Resumen Master post-Fase 1

```
Master Agricola Santa Elisa.xlsx
├─ Facturas          (+4 cols)
├─ Proveedores       (sin cambios)
├─ Cuenta Banco      (+4 cols)
├─ Tareas            (sin cambios)
├─ Bitácora          (sin cambios, Fase 2)
├─ Inventario        (sin cambios, Fase 4)
├─ Aplicaciones      (sin cambios)
├─ Personal          (sin cambios)
├─ Vacaciones        (sin cambios, Fase 6)
├─ Cosechas          ← NUEVA
├─ Guias Despacho    ← NUEVA
├─ Flujo Caja        ← NUEVA
├─ Ajustes Manuales  ← NUEVA
├─ Config            ← NUEVA
└─ Hectareas         ← NUEVA
```
