# 05b — UI Telegram: wizards e inline keyboards

## Wizard cosecha (FSM)

```
/cosecha nogales
→ ¿Kg totales? → 240000
→ ¿Exportadoras? → "Valbifrut 140000, Pacific Nuts 100000"
→ Confirmación inline: [Sí] [Editar]
→ Valbifrut: ¿precio USD/kg? → 1.80
→ ¿Cuotas? → [1] [2] [3] [4] [Otro]
→ Cuota 1: fecha? monto?
→ Cuota 2: fecha? monto?
→ Pacific Nuts: idem
→ ¿Liquidación final? → [Sí, en diciembre] [No] [Otra fecha]
→ ¿Monto estimado USD? → 40000
→ Resumen final: [Guardar] [Editar] [Cancelar]
```

## Inline keyboard: match ambiguo banco↔factura

```
🏦 Cargo 06/05: $1.291.894 - "Cals F2569088"
¿A qué factura corresponde?
[1) Cals N°2569088 $1.291.894]
[2) Cals N°2569090 $1.300.000]
[3) Otra (ver lista)]
[4) No es pago de factura]
```

## Inline keyboard: categorización dudosa

```
📋 Confianza baja (0.72)
Proveedor: COPEVAL
Glosa: "ACEITE Y FILTRO TRACTOR"
Categoría: [Maquinaria-mantención] [Caja chica] [Otro...]
Cultivo: [NOGALES] [CEREZOS] [AVELLANOS] [GENERAL]
```

## Inline keyboard: ingreso extraordinario

```
💰 Depósito 12/07: $8.500.000 - "TES.GRAL REPUBLICA"
No matchea cosecha.
[Devolución IVA] [Venta equipo]
[Indemnización] [Otro] [No contar]
```
