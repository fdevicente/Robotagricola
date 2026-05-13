# B1 — Validación patrón ingresos USD

## Hallazgo

Los ingresos por exportadoras NO aparecen directamente en ScotiaBCO (CLP).
El flujo real es:

```
Exportadora deposita USD → ScotiaUSD (cuenta dólares)
     ▼
Usuario vende dólares → CLP llega a ScotiaBCO
     ▼
En ScotiaBCO aparece como "Agricola Santa Elisa Bco Chile" (transferencia interna)
```

## Ejemplo real encontrado (ScotiaUSD)

- R6: 2023-01-06 | "Venta nueces Valbifrut F 177" | +72.184 USD | Nota: "Liquidacion final temporada 2022"
- R7: 2023-02-03 | "Venta dólares misma empresa y banco" | -15.000 USD (sale a CLP)

## Implicaciones para el diseño

1. **Para detectar ingresos de exportadoras el bot debe leer ScotiaUSD**, no solo ScotiaBCO
2. En ScotiaUSD las descripciones SÍ nombran a la exportadora ("Valbifrut", etc.)
3. Las "ventas de dólares" en ScotiaBCO son transferencias internas — no tienen info de exportadora
4. Vitakai paga directo en CLP → aparece en ScotiaBCO con su nombre

## Acción de diseño

- `historical_importer` debe leer ScotiaUSD para reconstruir ingresos históricos por exportadora
- `matcher` para ingresos: linkear depósitos ScotiaUSD con cuotas en hoja Cosechas
- El campo Estado en Cosechas se actualiza cuando matchea con un depósito USD real
- Dashboard muestra ingresos consolidados (CLP + USD convertidos)

## Fuentes de ingresos por tipo

| Exportadora | Moneda | Dónde aparece | Patrón descripción |
|---|---|---|---|
| Valbifrut | USD | ScotiaUSD | "Venta nueces Valbifrut" |
| Pacific Nuts | USD | ScotiaUSD | (verificar, probablemente similar) |
| Vitakai | CLP | ScotiaBCO | "Vitakai" directo |
