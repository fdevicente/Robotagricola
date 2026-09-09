# Pantalla de labores y costo de mano de obra

**Fecha:** 2026-09-09
**Estado:** diseño aprobado por el dueño

## El problema

El front no tiene ninguna pantalla que acumule el trabajo hecho. La hoja
`Bitácora` solo la lee `get_reporte_mensual`, y ahí sale como lista del mes: no
suma días de poda, no suma jornadas de herbicida, no cruza nada contra plata.

El dueño preguntó, textual: *"¿hay alguna pantalla que muestre las tareas que se
han hecho y los tiempos? por ejemplo que acumule los días de poda o de
herbicida, y que saque data de eso?"* Y después: *"saber cuánto me cuesta cada
trabajador (sin considerar el mío) pero según los sueldos y las jornadas
hombre"*.

## Lo que hay hoy, medido

284 filas en `Bitácora`, del 8-jun-2026 al 1-sep-2026 (más una fila suelta de
marzo). **54 etiquetas de labor distintas, 382 jornadas-hombre.** Las columnas
que sirven: `Fecha`, `Actividad`, `Jornadas Hombre`, `Trabajadores`, `Cultivo`,
`Sector`, `Horas Día`.

Las siete labores más grandes:

| labor | jornadas | días | personas | período |
|---|---:|---:|---:|---|
| Sacar restos poda nogales | 98 | 16 | 20 | 4-ago → 1-sep |
| Poda nogales | 46 | 25 | 3 | 16-jun → 10-ago |
| Mantención maquinaria | 33 | 16 | 5 | 10-jun → 28-ago |
| Poda avellanos | 28 | 25 | 2 | 16-jun → 10-ago |
| Pintar poda nogales | 22 | 23 | 3 | 17-jun → 10-ago |
| Desaguar | 19 | 6 | 7 | 24-jul → 28-ago |
| Aplicación herbicida | 14 | 10 | 5 | 11-jun → 1-sep |

Por mes: junio 95, julio 123, agosto 145, septiembre 19.

⚠️ **La historia es corta.** Arranca en junio de 2026. "Acumulado" son cuatro
meses, no años. La pantalla no puede sugerir lo contrario.

## De dónde sale el costo

**No hay ningún valor de jornada en el Master.** `Personal` tiene nombre, RUT,
cargo, fecha de ingreso y vacaciones — **ningún sueldo**. Se descartó pedirle al
dueño que cargue una tarifa en `Config`: envejece en silencio y la pantalla
miente sin avisar.

Lo que sí hay es plata que salió: **672 movimientos categorizados `MANO DE OBRA
PLANTA` y 715 `MANO DE OBRA TEMPORAL`** en `Cuenta Banco`. Y las transferencias
a la gente de planta **llevan el RUT en la glosa**:

```
TEF 13373052-4 Juan Parada Cas
TEF  9850887-2 FELICITO AMIGO
TEF 12318508-0 AGUSTIN MORA HE
```

Esos RUT están en `Personal` (los 6). El cruce es determinista, por RUT, no por
nombre.

### Las reglas

1. **Planta.** Costo por jornada de una persona en un mes = lo que se le
   transfirió ese mes ÷ las jornadas que hizo ese mes. Cada labor donde aparece
   se lleva `sus jornadas × ese costo`.
2. **Previred.** `PAGO COTIZ.PREVIRED` ($6.195.572 entre junio y hoy) es costo
   laboral real pero se paga en un monto único. Se reparte entre la gente de
   planta **a prorrata de sus jornadas del mes**.
3. **Temporada.** La cuadrilla —los 20 que sacaron restos de poda— no está en
   `Personal` ni cobra por transferencia nominativa. Costo por jornada del mes =
   total `MANO DE OBRA TEMPORAL` del mes ÷ jornadas de temporada del mes. **Es
   un prorrateo y la pantalla tiene que decirlo**, no presentarlo como exacto.
4. **Fuera del cálculo:** `CRAVE SPA` (RUT 77912665-K) y las remuneraciones de
   Felix De Vicente. Es lo del dueño y él pidió excluirlo.
5. ⚠️ **Un mes sin pagos cargados da costo `None`, que se muestra "sin datos".
   Nunca cero.** Un cero se lee como "salió gratis".

⚠️ Una persona puede recibir **dos transferencias el mismo mes** (Juan en
septiembre: $1.356.322 + $1.080.000). Se suman, no se toma la primera.

## Agrupación de las 54 etiquetas

Agrupación **por palabra clave**, elegida por el dueño sabiendo que va a juntar
cosas mal la primera vez. Un diccionario chico en el módulo: poda, herbicida,
mantención, riego, aseo, rastra, replante, desaguar.

⚠️ **La pantalla muestra qué etiquetas cayeron en cada grupo.** Sin eso la
agrupación es una caja negra y nadie puede corregirla. El dueño mira, dice qué
quedó mal, y se ajusta el diccionario.

Ojo con los falsos hermanos: `Poda nogales`, `Pintar poda nogales` y `Sacar
restos poda nogales` **son tres trabajos distintos** aunque compartan la
palabra. El diccionario tiene que distinguirlos, no colapsarlos.

## Arquitectura

**`modules/labores.py` — puro.** Lee el Master, no escribe nada, no importa
Flask. Se prueba entero en memoria contra un libro de prueba. El día que el
mismo dato se quiera por Telegram, ya está disponible.

Cuatro funciones, una por vista:

| función | devuelve |
|---|---|
| `resumen_labores(desde, hasta, path=None)` | por labor agrupada: jornadas, días, personas, costo, etiquetas que la componen |
| `costo_por_trabajador(desde, hasta, path=None)` | por persona: pagado, jornadas, costo por jornada, labores en las que trabajó |
| `por_cultivo_sector(desde, hasta, path=None)` | jornadas y costo abiertos por cultivo y sector |
| `evolucion_mensual(desde, hasta, path=None)` | jornadas y costo por mes y por grupo de labor |

Todas aceptan `path` para poder probarse contra un libro de prueba en vez de
contra el Master de producción. Es el mismo patrón que se acaba de aplicar a
`get_facturas_summary`.

**En el front:** ruta `/labores` + `src/templates/labores.html` + `/api/labores`,
siguiendo el patrón de `/vacaciones`, que ya existe y es la referencia más
parecida. Cuatro bloques, uno por función.

## Errores

- Mes sin pagos → costo `None` → "sin datos". Nunca cero.
- Persona con pagos pero **cero jornadas** ese mes → no se divide; su costo por
  jornada queda `None` y su plata no se imputa a ninguna labor.
- Fila sin `Jornadas Hombre` → cuenta como día trabajado pero no suma jornadas.
- Hoja `Bitácora` ausente o ilegible → las cuatro funciones devuelven vacío y
  logean, como hace `opciones_capataz`. La pantalla no puede reventar por eso.

## Tests

Libro de prueba con los casos que rompen, no con los que funcionan:

- Un mes **sin pagos** cargados: costo `None`, no 0.
- Una persona con sueldo y **cero jornadas** ese mes: no divide por cero.
- **Felix / CRAVE SPA no aparecen** en ningún resultado.
- **Dos transferencias el mismo mes** a la misma persona: se suman.
- Tres etiquetas distintas de la misma labor **quedan en un grupo**, y
  `Poda nogales` / `Pintar poda nogales` / `Sacar restos poda nogales`
  **quedan en grupos separados**.
- Previred repartido a prorrata: la suma de lo imputado **es igual** a lo pagado.
- Una labor con gente de planta y de temporada mezclada suma los dos costos.

Y, como en todo lo de este proyecto: **medir contra el Master real antes de dar
nada por bueno.** Los tests verdes no han cazado ninguno de los defectos reales
de las últimas semanas.

## Fuera de alcance

- Editar la bitácora desde el front. Es solo lectura.
- Proyectar costos a futuro. Eso ya vive en el cash flow.
- Cambiar cómo Juan reporta las labores por Telegram.
