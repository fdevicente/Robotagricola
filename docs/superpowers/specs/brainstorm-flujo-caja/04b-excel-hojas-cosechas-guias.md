# 04b — Excel: hojas nuevas Cosechas y Guias Despacho

## Hoja nueva: `Cosechas` (output del income_wizard)

| Col | Nombre | Ejemplo |
|---|---|---|
| A | Año | 2026 |
| B | Cultivo | NOGALES |
| C | Kg total | 240000 |
| D | Exportadora | Valbifrut |
| E | Kg asignados | 140000 |
| F | Precio USD/kg | 1.80 |
| G | N° cuotas | 2 |
| H | Cuota # | 1 |
| I | Fecha estimada | 2026-06-15 |
| J | Monto USD estimado | 126000 |
| K | Tipo cuota | "adelanto" / "liquidación final" |
| L | Estado | "esperado" / "recibido" / "ajustado" |
| M | Fecha real recibido | (cuando llega) |
| N | Monto real recibido | (cuando llega) |
| O | Moneda recibida | "USD" / "CLP" |
| P | Notas | |

- Cada exportadora ocupa N filas (una por cuota)
- Proyección de ingresos suma col J (o N si realizado)
- Columna O distingue Vitakai (CLP) vs Valbifrut/Pacific (USD vía "venta de dólares")

## Hoja nueva: `Guias Despacho`

| Col | Nombre |
|---|---|
| A | Fecha |
| B | N° Guía |
| C | Cultivo |
| D | Kg |
| E | Exportadora destino |
| F | Camión / Conductor |
| G | Sector / Equipo |
| H | Año cosecha |
| I | Origen | (manual / importado_historico) |
| J | PDF_path | link a Dropbox |
| K | Notas |

- Importadas históricamente desde DATOS COSECHA.xlsx con flag origen="importado_historico"
- Nuevas vienen por Telegram → guias_despacho/process_guia.py
- Sirven para validar wizard de cosecha (kg reales vs declarados)
