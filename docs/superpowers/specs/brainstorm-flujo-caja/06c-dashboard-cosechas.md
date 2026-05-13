# 06c — Pestaña Cosechas

## Vista: una card por temporada/cultivo

```
┌─ NOGALES 2026 ──────────────────────────────┐
│ Producción: 240.000 kg                      │
│ Exportadoras: Valbifrut (140k), Pacific (100k)│
│ Ingresos esperados: $452k USD ≈ $452M CLP   │
│ Ingresos recibidos: $223M CLP (49%)         │
│ [████████░░░░░░] 49%                        │
│                                              │
│ Gastos asociados al cultivo:                 │
│   Fertilizantes NOGALES: $18.2M             │
│   Fitosanitarios NOGALES: $12.1M            │
│   Mano obra temporal NOGALES: $34.5M        │
│   ... (solo categorías con cultivo=NOGALES)  │
│   TOTAL gastos NOGALES: $89.3M              │
│                                              │
│ Retorno neto: $452M - $89.3M = $362.7M      │
│ Retorno por kg: $1.511 CLP/kg               │
│ Retorno por hc: $9.5M CLP/hc (38 hc)       │
└──────────────────────────────────────────────┘
```

## Datos que alimentan la card

- Ingresos: hoja Cosechas (wizard) + ScotiaUSD (real)
- Gastos: Master.Facturas filtradas por Cultivo=NOGALES
- Gastos GENERAL: se prorratean por hc entre cultivos
- Kg: hoja Guias Despacho acumulados por cultivo+año

## Comparación histórica

Debajo de cada card: tabla mini

| Temporada | Kg | Ingresos | Gastos | Retorno neto | $/kg | $/hc |
|---|---|---|---|---|---|---|
| 2024 | (importado) | (importado) | (categorizado) | calc | calc | calc |
| 2025 | 290.000 | (del banco) | (categorizado) | calc | calc | calc |
| 2026 | 240.000 | $452M est. | $89.3M | $362.7M | $1.511 | $9.5M |

## Prorrateo gastos GENERAL

Gastos con Cultivo=GENERAL se distribuyen proporcional a hectáreas:
- Nogales 38hc / Cerezos 8hc / Avellanos Xhc
- Ej: gasto GENERAL $100M → Nogales recibe 38/(38+8+X) = ~80%
