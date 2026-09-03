# Diseño: la IA lee lo que escribe Juan y lo anota donde corresponda

**Fecha:** 2026-09-02
**Estado:** aprobado, pendiente de plan de implementación

## El problema

Juan no va a aprender un formato ni se va a acordar de uno. Hoy el bot supone que sí.

El 2-sep-2026, al recuperar 12 días de partes perdidos (ver `project-flujo-trabado-bitacora`), quedó a la vista que el problema de fondo no era el flujo trabado. Tres cosas medidas, no supuestas:

1. **`parsear_asistencia` exige `Nombre : actividad` y descarta EN SILENCIO las líneas sin dos puntos.** Juan dejó de ponerlos. En el parte del 26-ago (7 trabajadores, 2 líneas con dos puntos) devuelve **2 jornadas-hombre de 7**. No falla: falla a medias, que es peor.
2. **El extractor con IA devuelve UN objeto — una fila.** Por construcción no puede representar el parte del 31-ago: 19 personas en 3 labores distintas. Se probó: la IA colapsa todo a `Sacar restos de poda / NOGALES / 17 JH` y **se come a los 4 que aplicaron herbicida**, atribuyéndoselos a la poda. Confianza 0,6.
3. **El ruteo se decide ANTES de llamar a la IA**, con adivinanzas deterministas (`parece_maquinaria`, `parsear_asistencia`, `es_mensaje_sin_contenido`). La IA entra al final y ya encasillada en un destino.

Además hay dos vocabularios distintos: el prompt de `bitacora_extractor` conoce **9 trabajadores** y máquinas genéricas (`TRACTOR`, `CAMION`), mientras la hoja `Maquinaria` tiene **16 máquinas con modelo** y la cuadrilla de temporada trae unos 12 nombres nuevos. Por eso la IA nunca puede decir `TRACTOR MASSEY FERGUSON 4292`.

## Objetivo

Que **cualquier cosa que escriba Juan**, en el formato que sea, la lea la IA y quede anotada donde corresponde, sin que él tenga que recordar nada ni apretar nada.

## Decisiones tomadas

| Decisión | Elegido | Por qué |
|---|---|---|
| Quién revisa | **El dueño, solo lo dudoso** | Se guarda siempre y no se frena a Juan; el dueño ve el texto original y lo entendido, con botón para corregir |
| Trabajadores nuevos | **Preguntar la primera vez** | Mismo patrón que ya usa el bot con los proveedores nuevos de una factura. El alta automática ensuciaría `Personal` con las erratas de Juan |
| Alcance | **Bitácora + máquinas** | Asistencia, labores, eventos, horómetros, mantenciones y fichas. Es lo que Juan escribe hoy y lo que hoy se rutea mal |
| Enfoque | **La IA lee primero y devuelve una LISTA** | Es lo único que puede representar un mensaje con varias labores, varias máquinas o varios días |
| Los parsers | **Segunda opinión, no escriben** | La IA corre siempre; los parsers corren en paralelo, gratis, y su desacuerdo es la señal de duda |

### Por qué los parsers pasan a controlar en vez de a leer

Se descartó dejarlos como **vía rápida antes** de la IA: reintroduce dos comportamientos según cómo escriba Juan —la divergencia que causó el bug de los 12 días— y no caza un parseo *completo pero equivocado* (un aviso de mantención con forma de asistencia).

Como control valen mucho y no cuestan nada. El caso concreto: **`extraer_odometro` es un regex sobre el texto y la IA no.** Ya sabemos que la IA se equivoca en dígitos: el 2-sep leyó `3.169.778` donde decía `3.159.778`. Si el número de la IA no calza con el del regex sobre el propio texto, esa lectura no se guarda.

## Arquitectura

Tres piezas nuevas. **Ninguna escribe Excel nueva**: la escritura sigue pasando por lo que ya existe.

### `modules/parte_extractor.py` — el lector

```
leer_parte(texto: str, fecha_recepcion: date, contexto: Contexto) -> Parte
```

Una sola llamada a la IA por mensaje. Devuelve:

```
Parte = {
    "anotaciones": [Anotacion, ...],   # 0..N, no una
    "lineas_sin_anotar": [str, ...],   # las que no convirtió en nada
    "confianza": float,                # 0.0 a 1.0
}
Anotacion = {
    "destino": "BITACORA" | "HOROMETRO" | "MANTENCION" | "FICHA",
    "fecha": "YYYY-MM-DD",             # la que dice el texto, no la de recepción
    ...campos según destino
}
```

`Contexto` es lo que el bot ya sabe y la IA necesita para normalizar. Se arma una vez por mensaje y se le pasa **igual al lector y al juez**, para que los dos midan contra lo mismo.

⚠️ **Los trabajadores NO salen solo de la hoja `Personal`.** Medido el 2-sep: `Personal` tiene **6 filas** y con el nombre legal completo (`Felicito Amigo Soto`, `Luis Ramiro Amigo Soto`), mientras la columna `Trabajadores` de la bitácora usa **8 nombres canónicos** distintos (`Felicito Amigo`, `Ramiro Amigo`) y **Richard Padilla y Richard Padilla Crespo no están en `Personal`**. Armar el contexto solo con esa hoja dejaría a la IA peor informada que hoy.

La lista es la **unión de tres fuentes**, y manda la primera:

1. la columna `Trabajadores` de la hoja `Bitácora` — el vocabulario que el bot ya usa;
2. `TRABAJADORES_CONOCIDOS` y `ALIAS` de `bitacora_extractor` — trae los apodos (`pato` → `Patricio Mora`) y la regla de que `richard` a secas es el padre;
3. la hoja `Personal` — para que un trabajador recién dado de alta aparezca aunque todavía no tenga ninguna fila en la bitácora.

Las máquinas sí salen de `maquinas_conocidas()`, que ya une la hoja `Maquinaria` con lo visto en la bitácora y trae la última lectura y la unidad (h o km).

Dos diferencias con `bitacora_extractor` de hoy:

- **Devuelve una lista.** Un mensaje mixto —asistencia arriba, horómetro abajo— produce las dos cosas, que hoy es imposible.
- **El `contexto` se lee de las hojas `Personal` y `Maquinaria`**, no se escribe a mano en el prompt. Así deja de haber dos vocabularios y la IA puede nombrar la máquina con su modelo.

**`lineas_sin_anotar` es la pieza central.** Sin ella, "la IA lo lee todo" es un acto de fe, y lo que hoy falla en silencio seguiría fallando en silencio, solo que con más confianza.

### `modules/parte_control.py` — el juez

```
revisar(texto: str, parte: Parte, contexto: Contexto) -> list[Duda]
```

Corre `parsear_asistencia_multi`, `detectar_maquina` y `extraer_odometro` sobre el mismo texto y compara contra lo que dijo la IA. Puro: sin red y sin Excel, se puede probar entero con datos en memoria.

### `handlers/partes.py` — el que orquesta

Leer → controlar → escribir → responder a Juan → avisar al dueño si hay dudas.

### Flujo

```
texto de Juan (modo capataz)
  → leer_parte()      [IA, siempre]
  → revisar()         [parsers, gratis]
  → escribir cada anotación
        BITACORA / HOROMETRO  → registrar_bitacora_estructurada()   (ya existe)
        MANTENCION / FICHA    → modules/maquinaria                  (ya existe)
  → responder a Juan lo que quedó anotado
  → si hay dudas → aviso al dueño con el texto original
```

**Sale del camino:** en `handlers/chat.py`, para los usuarios en `AUTO_SAVE_USERS` desaparece el portón `parece_maquinaria` — la IA decide qué es. Para el dueño el camino de hoy no cambia: usa comandos.

**No se toca:** `/bitacora` y `/maquinaria` siguen igual, las facturas por foto siguen igual, y la caducidad de flujos de `modules/flujos.py` queda como está.

## Reglas de duda

Cualquiera manda aviso al dueño. La anotación se guarda igual **salvo las dos marcadas**.

| | Disparador | ¿Se guarda? |
|---|---|---|
| 1 | La IA declara líneas que no anotó | Sí, lo demás |
| 2 | El parser vio más personas que la IA | Sí |
| 3 | **El odómetro de la IA ≠ el del regex sobre el texto** | **No** |
| 4 | **La IA nombró una máquina que no está en la hoja `Maquinaria`** | **No** |
| 5 | Aparecen trabajadores que no están en `Personal` | Sí, con los nombres tal cual |
| 6 | La IA declara confianza menor a 0,6 | Sí |

Las dos que retienen son las que **contaminan hacia adelante**: un odómetro malo descuadra el cálculo de horas de todas las lecturas siguientes de esa máquina, y una máquina inventada crea una ficha fantasma. Lo demás es una fila que se puede borrar.

Se mantiene tal cual lo que ya existe: `validar_odometro` rechaza el salto imposible, no guarda y se lo dice a Juan.

### El aviso al dueño

Un mensaje con el texto de Juan **tal cual**, lo que quedó anotado, el motivo de la duda, y botones:

`✅ Está bien` · `🗑️ Borrar lo anotado` · `➕ Dar de alta a los nuevos` (cuando aplique)

## Manejo de errores

- **La IA se cae o devuelve basura:** no se pierde nada, el mensaje ya está en `files/telegram/*.jsonl` desde antes de que ningún handler lo toque. Se guarda una fila `OTRO` con el texto completo (como hoy) y se avisa al dueño.
- **Excel bloqueado:** sigue el respaldo a `files/logs/bitacora_fallback.txt` que ya existe.
- **Toda decisión de descarte deja línea en el log.** La falta de rastro fue lo que escondió el bug de los 12 días; un rechazo educado que no se loguea es indistinguible de que no haya pasado nada.

## Pruebas

Los casos son **los partes reales del respaldo crudo**, no ejemplos inventados. Están en `files/telegram/2026-08.jsonl` y `2026-09.jsonl`:

| Parte | `message_id` | Resultado esperado |
|---|---|---|
| 24-ago, con dos puntos | 2933 | 2 filas, 7 JH |
| 26-ago, mixto | 2992 | **7 personas, no 2** — hoy falla callado |
| 25, 27 y 28-ago, sin dos puntos | 2934, 3000, 3003 | 7 personas cada uno |
| 31-ago, 19 personas | 3106 | 19 JH, cero líneas sin anotar, y **los 4 del herbicida NO quedan en la poda** |
| 1-sep, 19 personas | 3122 | 19 JH, cero líneas sin anotar, cada labor con su gente |
| 7 horómetros | 3130, 3143, 3146, 3149, 3152, 3155, 3158 | máquina **con modelo** y odómetro exacto |

Los tests de asistencia afirman **jornadas-hombre, líneas sin anotar y a quién se le atribuye cada labor** — no el número de filas. Cuántas filas salen depende de cómo se normalice la actividad (`aplicación herbicida` y `aplicación herbicida nogales` pueden ser una o dos), y esa decisión se toma al implementar. Un test que fije el número de filas se rompería con un cambio de normalización que no tiene nada de malo.

La IA no es determinista, así que la suite corre contra **respuestas grabadas** del extractor. Aparte, un puñado de casos marcados (`@pytest.mark.ia`) que se corren contra la IA real cuando uno quiera comprobar el prompt. Sin eso la suite dependería de la red y del humor del modelo.

`parte_control.revisar()` se prueba entero sin red: se le pasa un `Parte` a mano y se comprueba cada regla de duda.

## Migración

Cuando esto ande, `scripts/carga/recuperar_bitacora_perdida.py` pasa a usar el lector nuevo y reingresa los **6 partes** que quedaron esperando (mid 2934, 2992, 3000, 3003, 3106, 3122), con `--simular` primero y contrastando contra lo que Juan escribió.

## Fuera de alcance

- No le corrige el formato a Juan ni le pide que escriba mejor.
- No aprende solo de las correcciones del dueño.
- No toca facturas ni fotos.
- No toca tareas, inventario ni vencimientos: cada destino nuevo es una forma más de escribir en la hoja que no era.
