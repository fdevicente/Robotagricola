# Conciliador bancario estilo Chipax — Diseño

Fecha: 2026-07-29
Referencia: capturas de Chipax (Cartolas · Sugerencias de conciliación)

## Objetivo
Llevar la página `/conciliacion` actual (169 líneas, tabla plana) a una interfaz de
trabajo diaria como la de Chipax: ver todos los movimientos del banco con su estado,
conciliarlos contra documentos, y aprobar sugerencias en lote.

## Qué tenemos hoy
- `modules/conciliador.py` — motor de matching (nº factura en glosa, monto+proveedor, IA para dudosos).
- `handlers/conciliacion.py` — `/conciliar` por Telegram con botones Aplicar/Cancelar.
- `src/templates/conciliacion.html` — página básica: analizar, checkboxes, aplicar.
- `Cuenta Banco`: 4.929 movimientos, 10 columnas. Estado de conciliación = col J `Factura_linkeada` (texto libre, solo 7 usadas).

## Limitación de fondo
La col J es un texto: no permite conciliación **parcial** ni **N:M**. Chipax muestra
"saldo por asignar" porque un movimiento puede pagar varias facturas y una factura
puede pagarse en varias transferencias. Sin eso no se pueden resolver los casos que
ya tenemos pendientes (Copeval "ND506808 al ND506813" en un cargo; S-Invest en cuotas).

---

## FASE 1 — Modelo de datos (fundación)

Hoja nueva **`Conciliaciones`**, una fila por vínculo:

| Campo | Detalle |
|---|---|
| ID | correlativo |
| Fecha conciliación | cuándo se hizo |
| Fila banco | fila en Cuenta Banco |
| Fecha mov / Descripción / Monto mov | copia para lectura |
| Tipo documento | Factura / Boleta / Ajuste / Traspaso / Terceros |
| Fila documento / N° doc / Proveedor | el respaldo |
| Monto asignado | permite parcial |
| Criterio | auto-nro / auto-monto / IA / manual |
| Usuario | quién lo concilió |
| Nota | comentario libre |

Funciones en `modules/conciliacion_store.py`:
- `conciliar(fila_banco, doc, monto, criterio, usuario, nota)`
- `desconciliar(id)`
- `saldo_por_asignar(fila_banco)` → monto − Σ asignados
- `estado(fila_banco)` → `por conciliar` · `parcial` · `conciliado`
- Migrar los 7 links de la col J; col J pasa a ser resumen legible.

## FASE 2 — Página principal (`/conciliacion`)

**Header KPIs** (como Chipax):
- **Por Pagar** — facturas sin fecha de pago (regla NN: solo saldo numérico)
- **Por Cobrar** — cosechas/ingresos esperados de la hoja Cosechas
- **Saldo Cta. Corriente** — último saldo de Cuenta Banco

**Filtros**: fecha desde/hasta · monto · descripción · N° documento · categoría · estado.

**Pestañas** con contador: `Todos` · `Sugerencias ⚡(n)` · `Abonos (n)` · `Cargos (n)` + checkbox **"Por conciliar"**.

**Tabla**: Fecha · Descripción (badge Transf./Cargo/Abono) · Categoría · Cargo · Abono · **Estado**
(monto por asignar + botón **Conciliar** + menú).

**Menú del botón Conciliar** (adaptado a la agrícola):
- *Vincular a factura/boleta* → modal de búsqueda (por nº, proveedor o monto)
- *Traspasar a balance* → gasto sin documento (comisiones, impuestos)
- *Cuenta a terceros* → préstamos, Gestora E, traspasos entre cuentas
- *Marcar como no conciliable* → con motivo

## FASE 3 — Vista de sugerencias (`/conciliacion/sugerencias`)

Tarjetas pareadas **Movimiento bancario ⇄ Documento de respaldo**, con:
- El motivo visible: *Recomendado por monto · por RUT · por descripción · por IA (92%)*
- Botones **Conciliar** / **Rechazar** por par
- **Conciliar todas (N)** arriba
- Filtros: fecha del movimiento, fecha del documento, tipo, proveedor

Reusa `conciliador.analizar()` + `resolver_dudosos_ia()`; se agrega registro de
**rechazos** para no volver a sugerir lo mismo.

## FASE 4 — Conciliación parcial / avanzada
- Un movimiento ↔ **varias facturas** (Copeval ND agrupadas): seleccionar varios documentos hasta cubrir el monto.
- Una factura ↔ **varios movimientos** (pagos en cuotas: S-Invest, Misael/plantas avellano).
- Mostrar siempre el **saldo por asignar** restante.

## FASE 5 — Extras
- Comentario por movimiento (ícono 💬 de Chipax)
- Exportar a Excel lo conciliado / lo pendiente
- Resumen en el reporte mensual: % conciliado del mes

---

## Fuera de alcance (de Chipax, no aplica acá)
Transbank · Líneas de negocio · Automatizaciones/reglas · Multi-cuenta corriente
(por ahora solo Scotiabank).

## Orden sugerido
Fase 1 → Fase 2 → Fase 3 son el núcleo utilizable.
Fase 4 resuelve los casos pendientes concretos. Fase 5 es cosmética.
