# 06d — Pestaña Replante (con escenario deuda)

## Layout

### Input (arriba)
```
Cultivo: [Avellanos ▼]  Hectáreas: [___]  [Calcular]
```

### Resultado (abajo)

```
┌─ RESULTADO REPLANTE ────────────────────────┐
│ Costo estimado: 4 hc × $8M/hc = $32M       │
│ (base: promedio facturas Inversión/Replante 2025) │
│                                              │
│ Disponible fin de año:     $87M             │
│ - Saldo mínimo seguridad: -$36M            │
│ - Gastos pendientes:      -$14M            │
│ = Disponible para replante: $37M            │
│                                              │
│ ✅ ALCANZA CON CAJA PROPIA                  │
│ Margen sobrante: $5M                        │
│                                              │
│ Si NO alcanzara:                            │
│ ❌ REQUIERE FINANCIAMIENTO                  │
│ Déficit: $XX M                              │
│ Opciones:                                   │
│   - Reducir hc replante a Y (máx con caja) │
│   - Crédito por $XX M                      │
│   - Postergar a próxima temporada           │
└──────────────────────────────────────────────┘
```

## Escenario deuda (D31)

Dado que el replante se hace igual aunque haya que endeudarse:

### Sub-sección "Simulador con crédito"
```
Si necesitás crédito:
  Monto necesario: $XXM
  Con tasa 8% anual a 3 años:
    Cuota mensual: $X.XM
    Total intereses: $X.XM

  Impacto en flujo de caja:
    Mes actual → Abr-27: cuota mensual se suma a egresos
    Saldo proyectado CON crédito: [mini gráfico línea]
    ¿Algún mes queda negativo? Sí/No → alerta
```

El usuario puede ajustar: tasa, plazo, meses de gracia.
El projector recalcula incluyendo las cuotas como egreso fijo.

## Tabla plan multi-año

```
Plan de transición Nogales → Avellanos

| Año | Nogales | Sacar | Avellanos | Plantar | Costo est. | Fuente |
|-----|---------|-------|-----------|---------|------------|--------|
| 2025| 52 hc   | 14    | 0→12      | 12      | $96M       | caja   |
| 2026| 38 hc   | (TBD) | 12→(TBD)  | (TBD)   | (TBD)      | (TBD)  |
| 2027| (TBD)   |       |           |         |            |        |
```

Editable: el usuario ajusta cuántas sacar/plantar cada año.
El sistema recalcula impacto en flujo de caja multi-año.
