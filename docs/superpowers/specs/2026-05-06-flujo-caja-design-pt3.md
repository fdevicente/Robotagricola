# Spec — Fase 1: Flujo de Caja — Parte 3/4

## 6. Cambios al Master Excel

### Columnas nuevas en Facturas (Q-T)
- Q: Categoria (str, una de 11)
- R: Cultivo (NOGALES/CEREZOS/AVELLANOS/GENERAL)
- S: Confianza (float 0-1)
- T: Categorizado_por (claude/manual/heuristic)

### Columnas nuevas en Cuenta Banco (G-J)
- G: Tipo (factura/sueldo/honorario/venta_dolares/otro)
- H: Categoria
- I: Cultivo
- J: Factura_linkeada (ref N° factura si aplica)

### 6 hojas nuevas
- **Cosechas**: año, cultivo, kg, exportadora, precio, cuotas, fechas, estado, moneda
- **Guias Despacho**: fecha, n°guía, cultivo, kg, exportadora, camión, sector, origen
- **Flujo Caja**: proyección mes×categoría (regenerada por projector, no editar)
- **Ajustes Manuales**: fecha, mes, categoría, cultivo, monto, razón, activo
- **Config**: parámetros del sistema (saldo_min, umbrales, fechas límite, etc.)
- **Hectareas**: año × cultivo (pendiente datos Daniel)

## 7. UI Telegram

### 14 comandos nuevos
/saldo, /proyeccion, /proyeccion_completa, /categoria,
/cosecha, /cosecha_actual, /replante, /reporte,
/revisar, /ajuste, /refresh, /manual, /guias, /cuota

### Wizard cosecha (FSM)
kg → exportadoras → precio × cada una → cuotas → fechas → liquidación → resumen → confirmar

### Inline keyboards para:
- Match ambiguo banco↔factura (opciones + "no es factura")
- Categorización dudosa (categoría + cultivo)
- Ingreso extraordinario (tipo + clasificación)

### Alertas
- 🔴 Saldo negativo proyectado / bajo umbral
- 🟡 Categoría >90% mes / factura por vencer 3 días
- 🟢 Resumen semanal lunes 8am / cierre mes día 1 + PDF

### Auto-disparos
- Cierre cosecha sugerido (fecha límite o 7d sin guías)
- Depósito sin match → ¿ingreso extraordinario?
- Categorización baja confianza → inline keyboard
- Match ambiguo → inline keyboard
