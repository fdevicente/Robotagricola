# 02a — Decisiones: categorías y proyección

## D1 — Nivel de proyección: **C (categoría + mes)**
Estacional, captura que ciertos meses son más caros (cosecha, fumigación).

## D2 — Esquema de categorías: **Híbrido (Opción 3)**
Categorías nuevas orientadas a flujo de caja + campo CARGO libre para retrocompatibilidad con planilla vieja.

## D3 — Lista de 11 categorías + cruce con cultivo

| # | Categoría |
|---|---|
| 1 | Mano de obra planta |
| 2 | Mano de obra temporal |
| 3 | Fertilizantes |
| 4 | Fitosanitarios |
| 5 | Combustible |
| 6 | Maquinaria — mantención |
| 7 | Riego |
| 8 | Servicios profesionales |
| 9 | Arriendos / Patentes / Seguros |
| 10 | Inversión / Replante |
| 11 | Caja chica / Imprevistos |

**Cruce cultivo:** NOGALES / CEREZOS / AVELLANOS / GENERAL

## D4 — Categorización del histórico: **Claude AI**
~1300 facturas, costo único ~USD $5. Confianza < 0.85 → revisar.

## D5 — Año base proyección: **2025 + 2024 referencia + ajustes manuales**
- Base: 2025 escalado por hectáreas 2026
- 2024 ajustado mostrado en paralelo para comparación visual
- Ajustes manuales del usuario sumados encima

## D6 — Master = fuente única
Master fue actualizado por el usuario al día. NO hay migración masiva desde FXP.

## D7 — Faltantes en Master.Facturas
- Liquidaciones de personal (vienen del banco)
- Boletas de honorarios de Francisco (se cargan al Master normal)

## D8 — Mano de obra: **Híbrido**
- Liquidaciones masivas → categorizadas desde Master.Cuenta Banco
- Honorarios Francisco → en Master.Facturas como una factura más

## D9 — Ingresos: vienen del banco + wizard post-cosecha
- Ingresos pasados: detectables del banco como "venta de dólares" + depósitos directos
- Ingresos futuros: wizard cosecha pregunta exportadora, precio, cuotas, fechas

## D10 — Moneda
- Españoles (Valbifrut, Pacific Nuts) pagan en USD → aparece como "venta de dólares" en el banco
- Vitakai paga en pesos directo
- Bot debe detectar ambos patrones
