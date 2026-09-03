# La IA lee lo que escribe Juan — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que cualquier cosa que escriba Juan por Telegram, en el formato que sea, la lea la IA y quede anotada donde corresponde, sin que él tenga que recordar nada ni apretar nada.

**Architecture:** Se da vuelta el orden actual. Hoy el ruteo se decide con adivinanzas deterministas *antes* de llamar a la IA, y la IA devuelve un solo objeto (una fila). Ahora la IA lee primero y devuelve una **lista** de anotaciones con su destino, más las líneas que no supo anotar. Los parsers deterministas dejan de escribir y pasan a **controlar** a la IA: su desacuerdo es la señal de duda que le llega al dueño.

**Tech Stack:** Python 3.11, python-telegram-bot 20+, openpyxl, requests contra la API de Claude, pytest.

**Diseño:** `docs/superpowers/specs/2026-09-02-partes-juan-lectura-ia-design.md`

> 🔄 **Este plan pasó a ser el SEGUNDO, y con menos alcance.** El 3-sep-2026, midiendo qué manda Juan de verdad, se vio que hace tres cosas y que sus **15 intentos de comando** dicen que está buscando un menú. Primero va `2026-09-03-teclado-y-horometro-guiado.md`, que se entrega solo.
>
> Qué cambia acá cuando llegue el turno:
> - **`MANTENCION` y `FICHA` salen de los destinos.** Medido: Juan mandó **cero**. El dueño las carga por `/maquinaria`, que no se toca. Quedan `BITACORA` y `HOROMETRO`.
> - **El horómetro deja de ser el camino principal** y queda como red: Juan va a seguir escribiendo partes a mano por costumbre y esos hay que leerlos igual. Sus casos de prueba siguen valiendo.
> - **La asistencia es la razón de ser de este plan**: 19 personas en 3 labores no se ingresan a botonazos.
> - La Task 1 (el contexto) ya está hecha y la usan los dos planes.

---

## Antes de empezar — cosas de esta máquina

**No hay `python` en el PATH.** El intérprete del proyecto es:

```
C:\Users\Windows\AppData\Local\Python\bin\python3.11.exe
```

Todos los comandos de este plan se corren desde `Robot/` con ese intérprete. En Git Bash:

```bash
alias py='/c/Users/Windows/AppData/Local/Python/bin/python3.11.exe'
```

**El bot está corriendo** y lo levanta la tarea programada `AgricolaBotWatchdog`. Ninguna tarea de este plan necesita reiniciarlo salvo la 8. Cuando toque reiniciar, seguir [feedback-reiniciar-bot-watchdog]: `Stop-ScheduledTask` **no mata el python**, hay que matarlo a mano y comprobar que queden 0 procesos.

**Mensajes de commit en español**, terminados en `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`.

### Si se trabaja en un worktree

Se ejecutó en `~/.config/superpowers/worktrees/Robot/partes-ia`, rama `feature/partes-lectura-ia`, para que el código a medio escribir no quede en la carpeta desde la que corre el bot. Un worktree recién creado **no arranca con la suite en verde**; hacen falta tres cosas que el `.gitignore` deja fuera:

1. **`files/` está ignorado**, así que el respaldo crudo de Telegram no viaja. Copiar `files/telegram/*.jsonl`: sin eso, las tareas 3 y 10 no tienen los partes de Juan.
2. **`.env` está ignorado.** Copiarlo (la tarea 3 necesita `ANTHROPIC_API_KEY`) y **apuntar `EXCEL_PATH` a una copia del Master**, para que nada del worktree pueda tocar el de producción.
3. **39 tests se saltan `EXCEL_PATH`** y arman a mano `tests/../../MASTER Agricola Santa Elisa.xlsx`. Hay que dejar una copia del Master **un nivel por encima del worktree** o esos 39 dan `FileNotFoundError`. (Que esos tests ignoren la configuración es un problema aparte; no es de este plan.)

Con las tres cosas, el baseline da **710 passed**, igual que `main`.

---

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `modules/parte_contexto.py` (nuevo) | Arma el `Contexto`: trabajadores y máquinas que el bot ya conoce. Lee Excel, no decide nada. |
| `modules/parte_extractor.py` (nuevo) | El lector. Una llamada a la IA, devuelve `Parte` (lista de anotaciones + líneas sin anotar). |
| `modules/parte_control.py` (nuevo) | El juez. Puro: sin red y sin Excel. Compara la IA contra los parsers y devuelve `Duda`s. |
| `handlers/partes.py` (nuevo) | Orquesta: leer → controlar → escribir → responder → avisar. Y los callbacks del aviso. |
| `handlers/chat.py` (modificar) | El modo capataz llama a `handlers.partes` en vez de a `auto_guardar_bitacora`. |
| `main.py` (modificar) | Registrar los tres callbacks nuevos. |
| `scripts/debug/grabar_fixtures_partes.py` (nuevo) | Graba respuestas reales de la IA como fixtures. Se corre a mano. |
| `tests/fixtures/partes/*.json` (nuevo) | Respuestas grabadas, para que la suite no dependa de la red. |
| `scripts/carga/recuperar_bitacora_perdida.py` (modificar) | Pasa a usar el lector nuevo y reingresa los 6 partes pendientes. |

**Lo que NO se toca:** `bitacora_manager.registrar_bitacora_estructurada`, `modules/maquinaria`, `modules/bitacora_asistencia`, las facturas, `modules/flujos.py`, y los comandos `/bitacora` y `/maquinaria`.

---

### Tipos que se usan en todo el plan

```python
Contexto = {
    "trabajadores": ["Felicito Amigo", ...],        # nombres canónicos
    "alias": {"pato": "Patricio Mora", ...},        # minúscula sin tildes → canónico
    "maquinas": [{"maquina": "TRACTOR MASSEY FERGUSON 4292",
                  "ultimo_odometro": 5231.0, "fecha": date(2026, 8, 11),
                  "unidad": "h"}, ...],
}

Anotacion = {
    "destino": "BITACORA" | "HOROMETRO" | "MANTENCION" | "FICHA",
    "fecha": "2026-08-31",          # la que dice el texto, no la de recepción
    "actividad": str, "cultivo": str, "sector": str,
    "jornadas_hombre": int | None, "trabajadores": [str, ...],
    "insumo": str, "cantidad": float | None, "unidad": str,
    "maquina": str, "odometro": float | None, "superficie_ha": float | None,
    "detalle": str,                 # para MANTENCION y FICHA
}

Parte = {
    "anotaciones": [Anotacion, ...],
    "lineas_sin_anotar": [str, ...],
    "confianza": float,             # 0.0 a 1.0
    "error": str,                   # "" si salió bien
}

Duda = {
    "regla": str,        # lineas_sin_anotar | parser_vio_mas | odometro_no_calza
                         # | maquina_desconocida | trabajadores_nuevos | confianza_baja
    "detalle": str,      # frase que se le muestra al dueño
    "retiene": bool,     # True = la anotación NO se guarda
    "indice": int | None,  # cuál anotación; None si la duda es del mensaje entero
    "datos": dict,       # extra, p.ej. {"nombres": [...]}
}
```

---

## Task 1: El contexto

**Files:**
- Create: `modules/parte_contexto.py`
- Test: `tests/test_parte_contexto.py`

⚠️ **Ojo con esto**, que es contraintuitivo: los trabajadores **no** salen solo de la hoja `Personal`. Medido el 2-sep-2026: `Personal` tiene 6 filas con el nombre legal completo (`Felicito Amigo Soto`), la bitácora usa 8 nombres canónicos (`Felicito Amigo`), y **Richard Padilla y Richard Padilla Crespo no están en `Personal`**. Manda el vocabulario de la bitácora.

> 🔴 **EJECUTADA — y el código de abajo quedó corto.** La revisión de calidad, midiendo contra el Master real, encontró que unir las dos listas tal cual da **15 nombres para 10 personas**: de las 6 filas de `Personal`, 5 son el nombre legal de alguien que la bitácora ya conoce y **ninguna es gente nueva**. Eso le da a la IA dos nombres válidos para la misma persona, y el que elija se escribe en la hoja y desaparece de las jornadas-hombre, porque `bitacora_asistencia` solo cuenta canónicos.
>
> El estado bueno está en los commits de la rama, no en este bloque. Los tres cambios: `Personal` pasa por `bitacora_asistencia._canonico` y solo entra quien no se conozca; los nombres se comparan sin tildes, mayúsculas ni espacios de más (la columna ahora la escribe la IA y el ciclo es cerrado); y se arregló en `modules/maquinaria.py` un `next()` sin proteger que, con una hoja `Bitácora` vacía, se llevaba la hoja `Maquinaria` **entera** en silencio.
>
> Ver la spec, sección del `Contexto`, que ya está corregida.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_parte_contexto.py
# -*- coding: utf-8 -*-
"""El contexto es lo que el bot ya sabe: a quien conoce y que maquinas tiene.

OJO: los trabajadores NO salen solo de la hoja Personal. Medido el 2-sep-2026,
Personal tiene 6 filas con el nombre legal completo ("Felicito Amigo Soto") y no
incluye a Richard Padilla ni a su hijo, mientras la columna Trabajadores de la
bitacora usa los 8 nombres canonicos que el bot viene usando hace meses.
Armar el contexto solo con Personal dejaria a la IA peor informada que hoy.
"""
from openpyxl import Workbook

from modules.parte_contexto import construir


def _excel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(["Fecha", "Hora", "Tipo", "Actividad", "Cultivo", "Sector",
               "Jornadas Hombre", "Trabajadores", "Insumo", "Cantidad",
               "Unidad", "Registro", "Registrado por", "Máquina", "Odómetro",
               "Horas Día", "Superficie ha", "Días Cubiertos"])
    ws.append(["2026-08-20", "14:09", "LABOR", "Poda", "NOGALES", "", 2,
               "Richard Padilla, Richard Padilla Crespo", "", None, "",
               "texto", "Juan Parada", "", None, None, None, None])
    per = wb.create_sheet("Personal")
    per.append(["Nombre", "RUT", "Cargo", "Fecha Ingreso"])
    per.append(["Felicito Amigo Soto", "9.850.887-2", None, None])
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_trae_los_nombres_de_la_bitacora(tmp_path):
    ctx = construir(_excel(tmp_path))
    assert "Richard Padilla" in ctx["trabajadores"]
    assert "Richard Padilla Crespo" in ctx["trabajadores"]


def test_tambien_trae_los_de_personal(tmp_path):
    ctx = construir(_excel(tmp_path))
    assert "Felicito Amigo Soto" in ctx["trabajadores"]


def test_trae_los_canonicos_de_siempre_aunque_no_esten_en_el_excel(tmp_path):
    """Sin esto se perderian los apodos y la regla del padre/hijo."""
    ctx = construir(_excel(tmp_path))
    assert "Patricio Mora" in ctx["trabajadores"]
    assert ctx["alias"]["pato"] == "Patricio Mora"
    assert ctx["alias"]["richard"] == "Richard Padilla"


def test_no_repite_nombres(tmp_path):
    ctx = construir(_excel(tmp_path))
    assert len(ctx["trabajadores"]) == len(set(ctx["trabajadores"]))


def test_los_canonicos_van_antes_que_los_nombres_legales_de_personal(tmp_path):
    """El orden es el que ve el modelo: primero los nombres que el bot escribe."""
    ctx = construir(_excel(tmp_path))
    nombres = ctx["trabajadores"]
    assert nombres.index("Patricio Mora") < nombres.index("Felicito Amigo Soto")


def test_las_maquinas_traen_unidad_y_ultima_lectura(tmp_path):
    ctx = construir(_excel(tmp_path))
    assert isinstance(ctx["maquinas"], list)
    for m in ctx["maquinas"]:
        assert set(m) >= {"maquina", "ultimo_odometro", "fecha", "unidad"}


def test_un_excel_sin_hojas_no_revienta(tmp_path):
    wb = Workbook()
    ruta = tmp_path / "vacio.xlsx"
    wb.save(ruta)
    ctx = construir(str(ruta))
    assert ctx["trabajadores"]          # quedan los canónicos de siempre
    assert ctx["maquinas"] == []
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_parte_contexto.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'modules.parte_contexto'`

- [ ] **Step 3: Escribir la implementación**

```python
# modules/parte_contexto.py
# -*- coding: utf-8 -*-
"""Lo que el bot ya sabe y la IA necesita para normalizar.

Se arma una vez por mensaje y se le pasa IGUAL al lector y al juez: si midieran
contra vocabularios distintos, el juez marcaria como duda lo que la IA hizo bien.

OJO CON LOS TRABAJADORES: no salen solo de la hoja Personal. Medido el
2-sep-2026, Personal tiene 6 filas con el nombre legal completo ("Felicito
Amigo Soto") y no incluye a Richard Padilla ni a su hijo, mientras la columna
Trabajadores de la bitacora usa los 8 nombres canonicos que el bot viene usando.
Manda el vocabulario de la bitacora; Personal solo agrega a los recien dados de
alta que todavia no tienen ninguna fila.
"""
import logging

logger = logging.getLogger(__name__)

BITACORA_SHEET = "Bitácora"
PERSONAL_SHEET = "Personal"


def construir(excel_path: str | None = None) -> dict:
    """Devuelve {"trabajadores": [...], "alias": {...}, "maquinas": [...]}."""
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    from modules.bitacora_extractor import ALIAS, TRABAJADORES_CONOCIDOS
    from modules.maquinaria import maquinas_conocidas

    ruta = excel_path or EXCEL_PATH
    nombres, orden = set(), []

    def _sumar(n):
        n = str(n or "").strip()
        if n and n not in nombres:
            nombres.add(n)
            orden.append(n)

    # El ORDEN importa: es el orden en que la lista se le muestra al modelo.
    # Primero los nombres que el bot realmente escribe en la hoja, después los
    # nombres legales largos de Personal.
    def _leer(hoja, saca):
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            try:
                if hoja in wb.sheetnames:
                    saca(wb[hoja])
            finally:
                wb.close()
        except Exception as e:                # un Excel raro no puede voltear esto
            logger.warning("parte_contexto: no pude leer %s de %s: %s",
                           hoja, ruta, e)

    def _de_bitacora(ws):
        enc = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if "Trabajadores" not in enc:
            return
        i = enc.index("Trabajadores")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > i and row[i]:
                for n in str(row[i]).split(","):
                    _sumar(n)

    def _de_personal(ws):
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row:
                _sumar(row[0])

    _leer(BITACORA_SHEET, _de_bitacora)       # 1 — el vocabulario que ya se usa
    for n in TRABAJADORES_CONOCIDOS:          # 2 — los de siempre, con sus apodos
        _sumar(n)
    _leer(PERSONAL_SHEET, _de_personal)       # 3 — los recién dados de alta

    try:
        maquinas = maquinas_conocidas(ruta)
    except Exception as e:
        logger.warning("parte_contexto: no pude leer las máquinas: %s", e)
        maquinas = []

    return {"trabajadores": orden, "alias": dict(ALIAS), "maquinas": maquinas}
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_parte_contexto.py -q`
Expected: PASS, 7 passed

- [ ] **Step 5: Comprobarlo contra el Master de verdad**

Run:
```bash
py -c "import sys; sys.path.insert(0,'.'); from modules.parte_contexto import construir; c=construir(); print(len(c['trabajadores']),'trabajadores'); print(len(c['maquinas']),'máquinas')"
```
Expected: al menos 10 trabajadores (los 8 de la bitácora + los de Personal que faltaban) y 16 máquinas.

- [ ] **Step 6: Commit**

```bash
git add modules/parte_contexto.py tests/test_parte_contexto.py
git commit -m "El contexto que la IA necesita para normalizar

Une tres fuentes y manda la primera: la columna Trabajadores de la bitacora
(el vocabulario que el bot ya usa), los canonicos con sus apodos, y la hoja
Personal. Personal sola no alcanza: tiene 6 filas con el nombre legal completo
y no incluye a Richard Padilla ni a su hijo.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 2: El lector — el prompt y la forma de la respuesta

**Files:**
- Create: `modules/parte_extractor.py`
- Test: `tests/test_parte_extractor.py`

En esta tarea **no se llama a la IA**: se prueba `_normalizar`, que es lo que convierte lo que devuelva el modelo en un `Parte` con forma garantizada. La llamada real va en la Task 3.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_parte_extractor.py
# -*- coding: utf-8 -*-
"""La forma del Parte esta garantizada aunque la IA devuelva cualquier cosa.

El extractor viejo devolvia UN objeto, o sea una fila, y por eso no podia
representar el parte del 31-ago-2026: 19 personas en 3 labores distintas.
Este devuelve una LISTA, y ademas declara que lineas NO anoto.
"""
import pytest

from modules.parte_extractor import _normalizar

CTX = {"trabajadores": ["Felicito Amigo", "Patricio Mora"],
       "alias": {"pato": "Patricio Mora"},
       "maquinas": [{"maquina": "TRACTOR MASSEY FERGUSON 4292",
                     "ultimo_odometro": 5231.0, "fecha": None, "unidad": "h"}]}


def test_devuelve_una_lista_de_anotaciones():
    p = _normalizar({"anotaciones": [
        {"destino": "BITACORA", "actividad": "Poda", "jornadas_hombre": 2},
        {"destino": "HOROMETRO", "maquina": "MF 4292", "odometro": 5237},
    ]}, "texto", CTX)
    assert len(p["anotaciones"]) == 2
    assert p["anotaciones"][0]["destino"] == "BITACORA"


def test_un_destino_desconocido_se_cae_a_bitacora():
    p = _normalizar({"anotaciones": [{"destino": "INVENTADO", "actividad": "x"}]},
                    "texto", CTX)
    assert p["anotaciones"][0]["destino"] == "BITACORA"


def test_rellena_todos_los_campos_aunque_la_ia_los_omita():
    p = _normalizar({"anotaciones": [{"destino": "BITACORA"}]}, "texto", CTX)
    a = p["anotaciones"][0]
    for campo in ("fecha", "actividad", "cultivo", "sector", "jornadas_hombre",
                  "trabajadores", "insumo", "cantidad", "unidad", "maquina",
                  "odometro", "superficie_ha", "detalle"):
        assert campo in a


def test_normaliza_los_apodos_a_nombre_canonico():
    p = _normalizar({"anotaciones": [
        {"destino": "BITACORA", "trabajadores": ["pato", "Felicito Amigo"]}]},
        "texto", CTX)
    assert p["anotaciones"][0]["trabajadores"] == ["Patricio Mora", "Felicito Amigo"]


def test_un_nombre_desconocido_se_conserva_tal_cual():
    """Se guarda igual: el dueño decide despues si lo da de alta."""
    p = _normalizar({"anotaciones": [
        {"destino": "BITACORA", "trabajadores": ["Josefina quiroga"]}]},
        "texto", CTX)
    assert p["anotaciones"][0]["trabajadores"] == ["Josefina quiroga"]


def test_jornadas_hombre_sale_de_la_lista_si_la_ia_no_la_dio():
    p = _normalizar({"anotaciones": [
        {"destino": "BITACORA", "trabajadores": ["a", "b", "c"]}]},
        "texto", CTX)
    assert p["anotaciones"][0]["jornadas_hombre"] == 3


def test_lineas_sin_anotar_siempre_es_lista():
    assert _normalizar({}, "texto", CTX)["lineas_sin_anotar"] == []
    p = _normalizar({"lineas_sin_anotar": "una sola"}, "texto", CTX)
    assert p["lineas_sin_anotar"] == ["una sola"]


@pytest.mark.parametrize("valor,esperado", [(None, 0.0), ("0.8", 0.8),
                                            (1.5, 1.0), (-1, 0.0), ("x", 0.0)])
def test_la_confianza_siempre_queda_entre_0_y_1(valor, esperado):
    assert _normalizar({"confianza": valor}, "t", CTX)["confianza"] == esperado


def test_una_respuesta_que_no_es_dict_no_revienta():
    p = _normalizar(["no", "es", "un", "dict"], "texto", CTX)
    assert p["anotaciones"] == []
    assert p["error"]
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_parte_extractor.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'modules.parte_extractor'`

- [ ] **Step 3: Escribir el prompt y `_normalizar`**

```python
# modules/parte_extractor.py
# -*- coding: utf-8 -*-
"""Lee lo que escribe Juan y devuelve TODAS las anotaciones que trae.

POR QUE EXISTE
bitacora_extractor devuelve UN objeto, o sea una fila. Con eso no se puede
representar el parte del 31-ago-2026: 19 personas en 3 labores distintas. Se
probo pasarselo igual y la IA colapsa todo a "Sacar restos de poda / 17 JH",
comiendose a los 4 que aplicaron herbicida. Este devuelve una LISTA.

Y devuelve `lineas_sin_anotar`. Sin eso, "la IA lo lee todo" es un acto de fe:
lo que hoy falla en silencio seguiria fallando en silencio, con mas confianza.
"""
import json
import logging
import re

import requests

from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)

CLAUDE_URL = "https://api.anthropic.com/v1/messages"
CLAUDE_MODELS = ["claude-sonnet-4-6", "claude-haiku-4-5-20251001"]
MAX_TOKENS = 4000          # un parte de 19 personas no cabe en 600
TIMEOUT_SEC = 60

DESTINOS = ("BITACORA", "HOROMETRO", "MANTENCION", "FICHA")
CULTIVOS = ("NOGALES", "CEREZOS", "AVELLANOS", "GENERAL")

_CAMPOS = {
    "fecha": "", "actividad": "", "cultivo": "GENERAL", "sector": "",
    "jornadas_hombre": None, "trabajadores": [], "insumo": "",
    "cantidad": None, "unidad": "", "maquina": "", "odometro": None,
    "superficie_ha": None, "detalle": "",
}


def _sin_tildes(s: str) -> str:
    for a, b in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"),
                 ("ñ", "n")):
        s = s.replace(a, b)
    return s


def _num(v):
    if v in (None, "", []):
        return None
    try:
        return float(str(v).replace(".", "").replace(",", ".")) if isinstance(v, str) else float(v)
    except (TypeError, ValueError):
        return None


def _canonico(nombre: str, ctx: dict) -> str:
    """Apodo o grafia suelta -> nombre canonico. Si no se reconoce, se conserva."""
    n = str(nombre or "").strip()
    if not n:
        return ""
    clave = _sin_tildes(n.lower())
    if clave in ctx.get("alias", {}):
        return ctx["alias"][clave]
    for conocido in ctx.get("trabajadores", []):
        if _sin_tildes(conocido.lower()) == clave:
            return conocido
    return n                                  # desconocido: se conserva tal cual


def _normalizar(data, texto: str, ctx: dict) -> dict:
    """Convierte lo que sea que devolvio la IA en un Parte con forma garantizada."""
    if not isinstance(data, dict):
        return {"anotaciones": [], "lineas_sin_anotar": [], "confianza": 0.0,
                "error": "la IA no devolvió un objeto", "texto_original": texto}

    anotaciones = []
    for cruda in (data.get("anotaciones") or []):
        if not isinstance(cruda, dict):
            continue
        a = dict(_CAMPOS)
        a.update({k: v for k, v in cruda.items() if k in _CAMPOS})
        destino = str(cruda.get("destino") or "").strip().upper()
        a["destino"] = destino if destino in DESTINOS else "BITACORA"
        a["trabajadores"] = [_canonico(n, ctx) for n in (a["trabajadores"] or [])
                             if str(n).strip()]
        if a["jornadas_hombre"] in (None, "") and a["trabajadores"]:
            a["jornadas_hombre"] = len(a["trabajadores"])
        else:
            jh = _num(a["jornadas_hombre"])
            a["jornadas_hombre"] = int(jh) if jh is not None else None
        for campo in ("cantidad", "odometro", "superficie_ha"):
            a[campo] = _num(a[campo])
        cultivo = str(a["cultivo"] or "").strip().upper()
        a["cultivo"] = cultivo if cultivo in CULTIVOS else "GENERAL"
        a["fecha"] = str(a["fecha"] or "")[:10]
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", a["fecha"]):
            a["fecha"] = ""
        anotaciones.append(a)

    sin_anotar = data.get("lineas_sin_anotar") or []
    if isinstance(sin_anotar, str):
        sin_anotar = [sin_anotar]
    sin_anotar = [str(l) for l in sin_anotar if str(l).strip()]

    conf = _num(data.get("confianza"))
    conf = 0.0 if conf is None else max(0.0, min(1.0, conf))

    return {"anotaciones": anotaciones, "lineas_sin_anotar": sin_anotar,
            "confianza": conf, "error": "", "texto_original": texto}
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_parte_extractor.py -q`
Expected: PASS, 13 passed

- [ ] **Step 5: Commit**

```bash
git add modules/parte_extractor.py tests/test_parte_extractor.py
git commit -m "El Parte tiene forma garantizada aunque la IA devuelva cualquier cosa

Devuelve una LISTA de anotaciones, no un objeto: un parte con 19 personas en 3
labores no cabe en una fila. Y declara las lineas que NO anoto, que es lo que
hace comprobable el 'la IA lo lee todo'.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 3: El lector — la llamada a la IA y los fixtures grabados

**Files:**
- Modify: `modules/parte_extractor.py` (agregar `leer_parte` y `_build_prompt`)
- Create: `scripts/debug/grabar_fixtures_partes.py`
- Create: `tests/fixtures/partes/` (se llena al correr el script)

- [ ] **Step 1: Escribir el prompt y `leer_parte`**

Agregar al final de `modules/parte_extractor.py`:

```python
def _build_prompt(texto: str, fecha_recepcion: str, ctx: dict) -> str:
    trabajadores = "\n".join(f"- {n}" for n in ctx.get("trabajadores", []))
    maquinas = "\n".join(
        f"- {m['maquina']} (última lectura: {m['ultimo_odometro']} {m['unidad']})"
        for m in ctx.get("maquinas", []))
    return f"""Eres el que anota los partes de una agrícola en Chile.

Juan Parada, el jefe de campo, escribe como le sale: a veces con dos puntos
("Felicito amigo : poda"), a veces sin nada ("Felicito amigo poda"), a veces
varios días en un mismo mensaje, a veces asistencia y horómetros juntos.
NO le corrijas el formato. Solo entiende qué dice y devuélvelo estructurado.

Fecha en que llegó el mensaje: {fecha_recepcion}

TRABAJADORES CONOCIDOS (normaliza a estos nombres exactos):
{trabajadores}

Si aparece alguien que NO está en la lista, ponlo TAL CUAL lo escribió Juan.
NO lo omitas y NO lo cambies por alguien parecido: en temporada entra gente
nueva y perderla significa perder jornadas de trabajo.

MÁQUINAS CONOCIDAS (normaliza al nombre completo con modelo):
{maquinas}

Si nombra una máquina que no está en la lista, ponla tal cual.

DESTINOS:
- BITACORA: trabajo de personas (poda, aplicación, riego, aseo, cosecha),
  vacaciones, clima o incidentes.
- HOROMETRO: una lectura de horómetro u odómetro de una máquina.
- MANTENCION: mantención, reparación, cambio de aceite o filtros, taller.
- FICHA: datos de la máquina (patente, año, marca, modelo, n° de serie,
  si es propia o arrendada).

REGLAS QUE IMPORTAN

1. UNA ANOTACIÓN POR LABOR. Si en un mismo día unos podan y otros aplican
   herbicida, son DOS anotaciones, cada una con SU gente. Nunca juntes gente
   que hizo cosas distintas.
2. VARIOS DÍAS = VARIAS ANOTACIONES, cada una con la fecha de su encabezado.
3. LA FECHA ES LA DEL TEXTO, no la de recepción: Juan reporta días después.
   Formato "YYYY-MM-DD". Si el texto no dice ninguna fecha, devuelve "".
4. HORÓMETRO: si dice inicio y término, el odómetro es el de TÉRMINO.
   Copia el número EXACTO, dígito por dígito. No lo redondees ni lo arregles.
5. jornadas_hombre = cuántas personas hicieron ESA labor.
6. NO INVENTES. Si algo no está en el texto, déjalo vacío o null.
7. lineas_sin_anotar: copia TAL CUAL las líneas del mensaje que no
   convertiste en ninguna anotación. Si anotaste todo, devuelve [].
   Esta lista es lo que revisa una persona después: si te la saltas, el
   trabajo que no anotaste se pierde sin que nadie se entere.

Devuelve SOLO este JSON, sin texto alrededor y sin markdown:
{{
  "anotaciones": [
    {{
      "destino": "BITACORA|HOROMETRO|MANTENCION|FICHA",
      "fecha": "YYYY-MM-DD o vacío",
      "actividad": "descripción corta",
      "cultivo": "NOGALES|CEREZOS|AVELLANOS|GENERAL",
      "sector": "",
      "jornadas_hombre": null,
      "trabajadores": [],
      "insumo": "", "cantidad": null, "unidad": "",
      "maquina": "", "odometro": null, "superficie_ha": null,
      "detalle": "solo para MANTENCION y FICHA"
    }}
  ],
  "lineas_sin_anotar": [],
  "confianza": 0.0
}}

Mensaje de Juan:
\"\"\"{texto}\"\"\""""


def _pedir_a_la_ia(prompt: str) -> str | None:
    """Devuelve el texto crudo de la respuesta, o None si fallaron todos."""
    headers = {"x-api-key": ANTHROPIC_API_KEY,
               "anthropic-version": "2023-06-01",
               "content-type": "application/json"}
    base = {"max_tokens": MAX_TOKENS,
            "messages": [{"role": "user", "content": prompt}]}
    for modelo in CLAUDE_MODELS:
        try:
            r = requests.post(CLAUDE_URL, headers=headers,
                              json={**base, "model": modelo}, timeout=TIMEOUT_SEC)
        except requests.RequestException as e:
            logger.warning("Claude %s excepción leyendo parte: %s", modelo, e)
            continue
        if r.status_code != 200:
            logger.warning("Claude %s HTTP %s: %s", modelo, r.status_code,
                           r.text[:150])
            continue
        try:
            return r.json()["content"][0]["text"].strip()
        except (KeyError, IndexError, ValueError) as e:
            logger.warning("Claude %s respuesta rara: %s", modelo, e)
    return None


def _json_de(crudo: str):
    """Saca el JSON aunque venga envuelto en fences."""
    txt = crudo.strip()
    if txt.startswith("```"):
        txt = txt.split("```")[1]
        if txt.startswith("json"):
            txt = txt[4:]
    return json.loads(txt)


def leer_parte(texto: str, fecha_recepcion: str, ctx: dict) -> dict:
    """Lee el mensaje entero y devuelve un Parte. Nunca lanza."""
    crudo = _pedir_a_la_ia(_build_prompt(texto, fecha_recepcion, ctx))
    if crudo is None:
        logger.error("Todos los modelos fallaron leyendo el parte")
        return {"anotaciones": [], "lineas_sin_anotar": [], "confianza": 0.0,
                "error": "IA no disponible", "texto_original": texto}
    try:
        return _normalizar(_json_de(crudo), texto, ctx)
    except (ValueError, json.JSONDecodeError) as e:
        logger.error("La IA devolvió algo no parseable: %s", e)
        return {"anotaciones": [], "lineas_sin_anotar": [], "confianza": 0.0,
                "error": "respuesta no parseable", "texto_original": texto}
```

- [ ] **Step 2: Escribir el grabador de fixtures**

```python
# scripts/debug/grabar_fixtures_partes.py
# -*- coding: utf-8 -*-
"""Graba las respuestas REALES de la IA para los partes de Juan del respaldo.

POR QUE
La IA no es determinista. Si la suite la llamara de verdad, dependeria de la red
y del humor del modelo, y un test rojo no diria si se rompio el codigo o si el
modelo contesto distinto. Se graba una vez y la suite corre contra lo grabado.

Se vuelve a correr a mano cuando se cambia el prompt, y se MIRA el diff: ahi se
ve si el prompt nuevo mejora o empeora.

USO
    python scripts/debug/grabar_fixtures_partes.py
"""
import json
import os
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, RAIZ)

DESTINO = os.path.join(RAIZ, "tests", "fixtures", "partes")
MIDS = [2933, 2934, 2992, 3000, 3003, 3106, 3122,
        3130, 3143, 3146, 3149, 3152, 3155, 3158]


def main():
    from modules.parte_contexto import construir
    from modules.parte_extractor import _build_prompt, _pedir_a_la_ia
    from modules.telegram_backup import leer_mes

    os.makedirs(DESTINO, exist_ok=True)
    por_id = {f["message_id"]: f for mes in ("2026-08", "2026-09")
              for f in leer_mes(mes) if f.get("text")}
    ctx = construir()

    for mid in MIDS:
        fila = por_id.get(mid)
        if fila is None:
            print("  falta el mensaje %s en el respaldo" % mid)
            continue
        recibido = fila["recibido_utc"][:10]
        crudo = _pedir_a_la_ia(_build_prompt(fila["text"], recibido, ctx))
        if crudo is None:
            print("  %s: la IA no respondio" % mid)
            continue
        ruta = os.path.join(DESTINO, "%s.json" % mid)
        with open(ruta, "w", encoding="utf-8") as fh:
            json.dump({"message_id": mid, "texto": fila["text"],
                       "fecha_recepcion": recibido, "respuesta_cruda": crudo},
                      fh, ensure_ascii=False, indent=2)
        print("  %s grabado" % mid)


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Grabar los fixtures**

Run: `py scripts/debug/grabar_fixtures_partes.py`
Expected: 14 líneas `<mid> grabado`, y 14 archivos en `tests/fixtures/partes/`.

- [ ] **Step 4: Mirar a ojo dos fixtures antes de fijarlos como verdad**

Run:
```bash
py -c "import json;d=json.load(open('tests/fixtures/partes/3106.json',encoding='utf-8'));print(d['respuesta_cruda'][:1500])"
```
Expected: un JSON con **varias** anotaciones, las de herbicida separadas de las de poda, y `lineas_sin_anotar` vacío o casi.

Si sale una sola anotación con las 19 personas juntas, **el prompt no sirve todavía**: ajustar la regla 1 del prompt, volver a grabar y volver a mirar. No seguir hasta que este fixture salga bien — es el caso que motivó todo el diseño.

- [ ] **Step 5: Commit**

```bash
git add modules/parte_extractor.py scripts/debug/grabar_fixtures_partes.py tests/fixtures/partes/
git commit -m "La llamada a la IA y los partes reales grabados como fixtures

La IA no es determinista: si la suite la llamara de verdad, un test rojo no
diria si se rompio el codigo o si el modelo contesto distinto. Se graban las
respuestas de los 14 partes reales de Juan y la suite corre contra eso.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 4: Los partes reales, contra los fixtures

**Files:**
- Create: `tests/test_partes_reales.py`

Esta es la tarea que dice si el diseño sirve. Los casos son los partes de verdad, no ejemplos inventados.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_partes_reales.py
# -*- coding: utf-8 -*-
"""Los partes REALES de Juan, contra las respuestas grabadas de la IA.

Estos son los casos que hoy fallan:
- 2992 (26-ago): 7 lineas, 2 con dos puntos. parsear_asistencia lee 2 JH de 7.
- 3106 (31-ago): 19 personas en 3 labores. El extractor viejo devuelve UNA fila
  y se come a los 4 que aplicaron herbicida.
- los 7 horometros: la maquina tiene que salir CON MODELO, no "TRACTOR" a secas.

Se afirman jornadas-hombre, lineas sin anotar y a quien se le atribuye cada
labor. NO se afirma el numero de filas: si "aplicacion herbicida" y "aplicacion
herbicida nogales" son una labor o dos es una decision de normalizacion, y un
test que la fije se rompe con un cambio que no tiene nada de malo.
"""
import json
import os

import pytest

from modules.parte_extractor import _json_de, _normalizar

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures", "partes")
CTX = {"trabajadores": ["Felicito Amigo", "Patricio Mora", "Ramiro Amigo",
                        "Agustin Mora", "Javier Gonzalez", "Richard Padilla",
                        "Richard Padilla Crespo", "Juan Parada"],
       "alias": {"pato": "Patricio Mora", "richard": "Richard Padilla"},
       "maquinas": []}


def parte(mid):
    with open(os.path.join(FIXTURES, "%s.json" % mid), encoding="utf-8") as fh:
        d = json.load(fh)
    return _normalizar(_json_de(d["respuesta_cruda"]), d["texto"], CTX)


def personas(p):
    """Total de personas anotadas, contando a cada una una sola vez por labor."""
    return sum(len(a["trabajadores"]) for a in p["anotaciones"]
               if a["destino"] == "BITACORA")


def gente_de(p, palabra):
    return {n for a in p["anotaciones"] if palabra in a["actividad"].lower()
            for n in a["trabajadores"]}


@pytest.mark.parametrize("mid", [2933, 2934, 2992, 3000, 3003])
def test_los_partes_de_7_personas_anotan_a_las_7(mid):
    p = parte(mid)
    assert personas(p) == 7, p["anotaciones"]
    assert p["lineas_sin_anotar"] == []


@pytest.mark.parametrize("mid", [3106, 3122])
def test_los_partes_de_temporada_anotan_a_las_19(mid):
    p = parte(mid)
    assert personas(p) == 19, p["anotaciones"]
    assert p["lineas_sin_anotar"] == []


def test_el_31_de_agosto_no_mezcla_herbicida_con_poda():
    """El caso exacto que el extractor viejo se comia."""
    p = parte(3106)
    poda = gente_de(p, "poda")
    herbicida = gente_de(p, "herbicida")
    assert herbicida, "no anoto a nadie en herbicida"
    assert not (poda & herbicida), "hay gente en las dos labores"
    assert "Patricio Mora" in herbicida


@pytest.mark.parametrize("mid,maquina,odo", [
    (3130, "TRACTOR MASSEY FERGUSON 6711", 2055),
    (3143, "TRACTOR MASSEY FERGUSON 4292", 5237),
    (3146, "TRACTOR MASSEY FERGUSON 4275", 3456),
    (3149, "TRACTOR MASSEY FERGUSON 6711", 2057),
    (3152, "TRACTOR MASSEY FERGUSON 4292", 5239),
    (3155, "TRACTOR MASSEY FERGUSON 4275", 3460),
    (3158, "TRACTOR JOHN DEERE 5085", 3265),
])
def test_los_horometros_traen_maquina_con_modelo_y_el_numero_de_termino(mid, maquina, odo):
    p = parte(mid)
    lecturas = [a for a in p["anotaciones"] if a["destino"] == "HOROMETRO"]
    assert len(lecturas) == 1, p["anotaciones"]
    assert lecturas[0]["maquina"].upper() == maquina
    assert lecturas[0]["odometro"] == odo


@pytest.mark.parametrize("mid,fecha", [
    (2933, "2026-08-24"), (3106, "2026-08-31"), (3130, "2026-08-31"),
    (3158, "2026-09-01"),
])
def test_la_fecha_es_la_del_texto_no_la_de_recepcion(mid, fecha):
    """Juan reporta dias despues: los 7 horometros llegaron todos el 2-sep."""
    p = parte(mid)
    assert all(a["fecha"] == fecha for a in p["anotaciones"]), p["anotaciones"]
```

- [ ] **Step 2: Correr el test**

Run: `py -m pytest tests/test_partes_reales.py -q`
Expected: PASS, 20 passed.

Si alguno falla, **el que está mal es el prompt, no el test**: ajustar `_build_prompt`, volver a correr `scripts/debug/grabar_fixtures_partes.py`, y volver a correr. Los números de este test salen de lo que Juan escribió; no se tocan para que pase.

- [ ] **Step 3: Registrar el marcador `ia` para que no salgan avisos**

El proyecto no tiene `pytest.ini`. Crearlo:

```ini
# pytest.ini
[pytest]
markers =
    ia: llama a la IA de verdad. No corre en la suite normal.
addopts = -m "not ia"
```

⚠️ **El `addopts` no es opcional.** Un marcador solo etiqueta: sin él, `pytest tests/` correría igual los casos que llaman a la IA, y la suite pasaría a depender de la red. Con esto, para correrlos hay que pedirlo: `-m ia`.

⚠️ Crear un `pytest.ini` fija el `rootdir` y puede cambiar cómo se descubren los tests que ya existen. **Correr la suite entera justo después** y comprobar que el total no bajó:

Run: `py -m pytest tests/ -q`
Expected: el mismo número de tests que antes de crear el archivo (más los nuevos).

- [ ] **Step 4: Escribir los casos que sí llaman a la IA de verdad**

Los fixtures congelan lo que el modelo contestó *ese día*. Estos casos son para comprobar el prompt contra el modelo de hoy, y se corren a mano.

```python
# tests/test_partes_ia_real.py
# -*- coding: utf-8 -*-
"""Casos que llaman a la IA DE VERDAD. No corren en la suite normal.

    py -m pytest tests/test_partes_ia_real.py -m ia -q

Los fixtures congelan lo que el modelo contesto el dia que se grabaron. Esto
comprueba que el prompt sigue funcionando con el modelo de hoy. Si esto falla y
la suite normal pasa, no se rompio el codigo: cambio el modelo.
"""
import json
import os

import pytest

from modules.parte_extractor import leer_parte

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures", "partes")
CTX = {"trabajadores": ["Felicito Amigo", "Patricio Mora", "Ramiro Amigo",
                        "Agustin Mora", "Javier Gonzalez", "Richard Padilla",
                        "Richard Padilla Crespo"],
       "alias": {},
       "maquinas": [{"maquina": "TRACTOR MASSEY FERGUSON 4292",
                     "ultimo_odometro": 5231.0, "fecha": None, "unidad": "h"}]}


def _texto(mid):
    with open(os.path.join(FIXTURES, "%s.json" % mid), encoding="utf-8") as fh:
        return json.load(fh)


@pytest.mark.ia
def test_el_parte_de_19_personas_sigue_saliendo_completo():
    d = _texto(3106)
    p = leer_parte(d["texto"], d["fecha_recepcion"], CTX)
    personas = sum(len(a["trabajadores"]) for a in p["anotaciones"])
    assert personas == 19, p["anotaciones"]
    assert p["lineas_sin_anotar"] == []


@pytest.mark.ia
def test_el_horometro_sigue_saliendo_con_el_numero_exacto():
    d = _texto(3143)
    p = leer_parte(d["texto"], d["fecha_recepcion"], CTX)
    lecturas = [a for a in p["anotaciones"] if a["destino"] == "HOROMETRO"]
    assert len(lecturas) == 1
    assert lecturas[0]["odometro"] == 5237
```

- [ ] **Step 5: Correr los casos contra la IA una vez**

Run: `py -m pytest tests/test_partes_ia_real.py -m ia -q`
Expected: PASS, 2 passed (tarda ~10 s: está llamando a la IA de verdad).

Comprobar que **no** corren en la suite normal:
Run: `py -m pytest tests/test_partes_ia_real.py -q`
Expected: `2 deselected`. Si dice `2 passed`, el `addopts` no quedó bien y la suite depende de la red.

- [ ] **Step 6: Commit**

```bash
git add tests/test_partes_reales.py tests/test_partes_ia_real.py pytest.ini
git commit -m "Los partes reales de Juan como casos de prueba

El del 26-ago, donde el parser lee 2 jornadas de 7. El del 31-ago, con 19
personas en 3 labores, donde el extractor viejo se comia a los del herbicida.
Y los 7 horometros, que tienen que salir con modelo y el numero exacto.

Se afirman jornadas, lineas sin anotar y a quien se le atribuye cada labor,
no el numero de filas: eso depende de como se normalice la actividad.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 5: El juez

**Files:**
- Create: `modules/parte_control.py`
- Test: `tests/test_parte_control.py`

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_parte_control.py
# -*- coding: utf-8 -*-
"""Los parsers dejan de escribir y pasan a controlar a la IA.

La regla que mas vale: extraer_odometro es un regex sobre el texto y la IA no.
El 2-sep-2026 la IA leyo 3.169.778 donde el texto decia 3.159.778. Un odometro
malo no es una fila mala: descuadra el calculo de horas de TODAS las lecturas
siguientes de esa maquina. Por eso esa no se guarda.
"""
from modules.parte_control import revisar

CTX = {"trabajadores": ["Felicito Amigo", "Patricio Mora"],
       "alias": {},
       "maquinas": [{"maquina": "TRACTOR MASSEY FERGUSON 4292",
                     "ultimo_odometro": 5231.0, "fecha": None, "unidad": "h"}]}


def _anot(**kw):
    base = {"destino": "BITACORA", "fecha": "2026-08-31", "actividad": "Poda",
            "cultivo": "NOGALES", "sector": "", "jornadas_hombre": None,
            "trabajadores": [], "insumo": "", "cantidad": None, "unidad": "",
            "maquina": "", "odometro": None, "superficie_ha": None, "detalle": ""}
    base.update(kw)
    return base


def _parte(anotaciones, sin_anotar=(), confianza=0.9):
    return {"anotaciones": list(anotaciones), "lineas_sin_anotar": list(sin_anotar),
            "confianza": confianza, "error": "", "texto_original": ""}


def reglas(dudas):
    return {d["regla"] for d in dudas}


def test_un_parte_limpio_no_levanta_dudas():
    texto = "Lunes 31 de agosto 2026\nFelicito amigo : poda"
    p = _parte([_anot(trabajadores=["Felicito Amigo"], jornadas_hombre=1)])
    assert revisar(texto, p, CTX) == []


def test_lineas_sin_anotar():
    p = _parte([_anot()], sin_anotar=["Richard padilla desaguar"])
    d = revisar("texto", p, CTX)
    assert "lineas_sin_anotar" in reglas(d)
    assert all(not x["retiene"] for x in d)


def test_el_parser_vio_mas_personas_que_la_ia():
    texto = ("Lunes 31 de agosto 2026\nFelicito amigo : poda\n"
             "Patricio Mora : poda")
    p = _parte([_anot(trabajadores=["Felicito Amigo"], jornadas_hombre=1)])
    d = revisar(texto, p, CTX)
    assert "parser_vio_mas" in reglas(d)


def test_el_parser_no_se_queja_si_la_ia_vio_mas_que_el():
    """Es lo esperado: el parser no lee las lineas sin dos puntos."""
    texto = "Lunes 31 de agosto 2026\nFelicito amigo poda\nPatricio Mora poda"
    p = _parte([_anot(trabajadores=["Felicito Amigo", "Patricio Mora"],
                      jornadas_hombre=2)])
    assert "parser_vio_mas" not in reglas(revisar(texto, p, CTX))


def test_odometro_que_no_calza_con_el_texto_retiene():
    texto = "Tractor massey ferguson 4292\nHorometro termino 5237"
    p = _parte([_anot(destino="HOROMETRO", maquina="TRACTOR MASSEY FERGUSON 4292",
                      odometro=5273)])
    d = revisar(texto, p, CTX)
    assert "odometro_no_calza" in reglas(d)
    assert [x for x in d if x["regla"] == "odometro_no_calza"][0]["retiene"] is True
    assert [x for x in d if x["regla"] == "odometro_no_calza"][0]["indice"] == 0


def test_odometro_que_calza_no_molesta():
    texto = "Tractor massey ferguson 4292\nHorometro inicio 5234\nHorometro termino 5237"
    p = _parte([_anot(destino="HOROMETRO", maquina="TRACTOR MASSEY FERGUSON 4292",
                      odometro=5237)])
    assert revisar(texto, p, CTX) == []


def test_maquina_que_no_existe_retiene():
    texto = "El tractor fantasma horometro 100"
    p = _parte([_anot(destino="HOROMETRO", maquina="TRACTOR FANTASMA 999",
                      odometro=100)])
    d = revisar(texto, p, CTX)
    assert "maquina_desconocida" in reglas(d)
    assert [x for x in d if x["regla"] == "maquina_desconocida"][0]["retiene"] is True


def test_trabajadores_nuevos_avisan_pero_no_retienen():
    p = _parte([_anot(trabajadores=["Felicito Amigo", "Josefina Quiroga"],
                      jornadas_hombre=2)])
    d = revisar("texto", p, CTX)
    duda = [x for x in d if x["regla"] == "trabajadores_nuevos"][0]
    assert duda["retiene"] is False
    assert duda["datos"]["nombres"] == ["Josefina Quiroga"]


def test_confianza_baja():
    p = _parte([_anot()], confianza=0.4)
    d = revisar("texto", p, CTX)
    assert "confianza_baja" in reglas(d)
    assert all(not x["retiene"] for x in d)


def test_confianza_justo_en_el_limite_no_se_queja():
    assert "confianza_baja" not in reglas(revisar("t", _parte([_anot()], confianza=0.6), CTX))


def test_la_ia_caida_es_una_duda_del_mensaje_entero():
    p = _parte([], confianza=0.0)
    p["error"] = "IA no disponible"
    d = revisar("texto", p, CTX)
    assert d and all(x["indice"] is None for x in d)
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_parte_control.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'modules.parte_control'`

- [ ] **Step 3: Escribir la implementación**

```python
# modules/parte_control.py
# -*- coding: utf-8 -*-
"""Los parsers deterministas controlan a la IA en vez de escribir.

POR QUE ASI
Se penso dejarlos como via rapida ANTES de la IA. Se descarto: reintroduce dos
comportamientos segun como escriba Juan --la divergencia que causo el bug de los
12 dias-- y no caza un parseo completo pero equivocado.

Como control valen y no cuestan nada. El cruce que mas vale: extraer_odometro es
un regex sobre el texto y la IA no. Ya sabemos que la IA se equivoca en digitos.

Este modulo es PURO: no toca red ni Excel, asi se prueba entero en memoria.
"""
import logging

logger = logging.getLogger(__name__)

CONFIANZA_MINIMA = 0.6


def _duda(regla, detalle, retiene=False, indice=None, datos=None):
    return {"regla": regla, "detalle": detalle, "retiene": retiene,
            "indice": indice, "datos": datos or {}}


def _norm(s):
    return str(s or "").strip().upper()


def revisar(texto: str, parte: dict, ctx: dict) -> list:
    """Compara lo que dijo la IA contra los parsers. Devuelve las dudas."""
    from modules.bitacora_asistencia import parsear_asistencia_multi
    from modules.maquinaria import extraer_odometro

    dudas = []

    if parte.get("error"):
        return [_duda("ia_caida", parte["error"])]

    # 1 — lineas que la IA declaro no haber anotado
    if parte.get("lineas_sin_anotar"):
        dudas.append(_duda(
            "lineas_sin_anotar",
            "no anotó %d línea(s): %s" % (len(parte["lineas_sin_anotar"]),
                                          " · ".join(parte["lineas_sin_anotar"][:5])),
            datos={"lineas": parte["lineas_sin_anotar"]}))

    anotaciones = parte.get("anotaciones") or []

    # 2 — el parser determinista vio mas gente que la IA
    try:
        dias = parsear_asistencia_multi(texto) or []
    except Exception as e:                     # el parser no puede voltear el juez
        logger.warning("parte_control: el parser de asistencia falló: %s", e)
        dias = []
    del_parser = sum(len(g["trabajadores"]) for d in dias for g in d["grupos"])
    de_la_ia = sum(len(a.get("trabajadores") or []) for a in anotaciones
                   if a.get("destino") == "BITACORA")
    if del_parser > de_la_ia:
        dudas.append(_duda(
            "parser_vio_mas",
            "el parser contó %d personas y la IA anotó %d" % (del_parser, de_la_ia),
            datos={"parser": del_parser, "ia": de_la_ia}))

    # 3 y 4 — por anotacion
    conocidas = {_norm(m["maquina"]) for m in (ctx.get("maquinas") or [])}
    del_texto = extraer_odometro(texto)
    for i, a in enumerate(anotaciones):
        maquina = _norm(a.get("maquina"))
        if maquina and conocidas and maquina not in conocidas:
            dudas.append(_duda(
                "maquina_desconocida",
                "«%s» no está en la hoja Maquinaria" % a.get("maquina"),
                retiene=True, indice=i, datos={"maquina": a.get("maquina")}))
            continue                            # ya se retiene, no revisar el odómetro
        if a.get("destino") == "HOROMETRO" and a.get("odometro") is not None:
            if del_texto is not None and float(a["odometro"]) != float(del_texto):
                dudas.append(_duda(
                    "odometro_no_calza",
                    "la IA leyó %g y en el texto dice %g" % (a["odometro"], del_texto),
                    retiene=True, indice=i,
                    datos={"ia": a["odometro"], "texto": del_texto}))

    # 5 — gente que el bot no conoce
    conocidos = {str(n).strip().lower() for n in (ctx.get("trabajadores") or [])}
    nuevos, vistos = [], set()
    for a in anotaciones:
        for n in (a.get("trabajadores") or []):
            clave = str(n).strip().lower()
            if clave and clave not in conocidos and clave not in vistos:
                vistos.add(clave)
                nuevos.append(n)
    if nuevos:
        dudas.append(_duda(
            "trabajadores_nuevos",
            "%d nombre(s) que no conozco: %s" % (len(nuevos), ", ".join(nuevos)),
            datos={"nombres": nuevos}))

    # 6 — la propia IA declara poca confianza
    if parte.get("confianza", 0.0) < CONFIANZA_MINIMA:
        dudas.append(_duda("confianza_baja",
                           "la IA quedó con %.0f%% de confianza"
                           % (parte.get("confianza", 0.0) * 100)))

    return dudas


def indices_retenidos(dudas: list) -> set:
    """Anotaciones que NO se guardan."""
    return {d["indice"] for d in dudas if d["retiene"] and d["indice"] is not None}
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_parte_control.py -q`
Expected: PASS, 11 passed

- [ ] **Step 5: Commit**

```bash
git add modules/parte_control.py tests/test_parte_control.py
git commit -m "Los parsers pasan a controlar a la IA en vez de escribir

El cruce que mas vale: extraer_odometro es un regex sobre el texto y la IA no.
El 2-sep la IA leyo 3.169.778 donde decia 3.159.778. Un odometro malo no es una
fila mala: descuadra las horas de TODAS las lecturas siguientes de esa maquina,
asi que esa anotacion no se guarda. Lo mismo con una maquina que no existe.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 6: El que escribe

**Files:**
- Create: `handlers/partes.py`
- Test: `tests/test_partes_escritura.py`

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_partes_escritura.py
# -*- coding: utf-8 -*-
"""Cada anotacion va a su destino, y lo retenido NO se escribe."""
from handlers.partes import escribir_anotaciones


class _Espia:
    def __init__(self):
        self.bitacora, self.mantenciones, self.fichas = [], [], []

    def bita(self, campos, quien):
        self.bitacora.append((campos, quien))
        return {"horas_dia": None, "odo_previo": None, "es_baseline": False}

    def mant(self, m, quien):
        self.mantenciones.append(m)
        return len(self.mantenciones)

    def ficha(self, d):
        self.fichas.append(d)
        return d.get("maquina", "")


def _anot(**kw):
    base = {"destino": "BITACORA", "fecha": "2026-08-31", "actividad": "Poda",
            "cultivo": "NOGALES", "sector": "", "jornadas_hombre": 1,
            "trabajadores": ["Felicito Amigo"], "insumo": "", "cantidad": None,
            "unidad": "", "maquina": "", "odometro": None,
            "superficie_ha": None, "detalle": ""}
    base.update(kw)
    return base


def test_una_anotacion_de_bitacora_va_a_la_bitacora():
    e = _Espia()
    escritas = escribir_anotaciones([_anot()], "texto", "Juan Parada", set(),
                                    _bita=e.bita, _mant=e.mant, _ficha=e.ficha)
    assert len(e.bitacora) == 1
    assert e.bitacora[0][0]["actividad"] == "Poda"
    assert e.bitacora[0][1] == "Juan Parada"
    assert len(escritas) == 1


def test_un_horometro_va_a_la_bitacora_como_maquinaria():
    e = _Espia()
    escribir_anotaciones([_anot(destino="HOROMETRO", maquina="TRACTOR X",
                                odometro=100, trabajadores=[])],
                         "texto", "Juan", set(),
                         _bita=e.bita, _mant=e.mant, _ficha=e.ficha)
    assert e.bitacora[0][0]["tipo"] == "MAQUINARIA"
    assert e.bitacora[0][0]["odometro"] == 100


def test_una_mantencion_va_a_mantenciones():
    e = _Espia()
    escribir_anotaciones([_anot(destino="MANTENCION", maquina="TRACTOR X",
                                detalle="cambio de aceite")],
                         "texto", "Juan", set(),
                         _bita=e.bita, _mant=e.mant, _ficha=e.ficha)
    assert len(e.mantenciones) == 1
    assert e.bitacora == []


def test_una_ficha_va_a_fichas():
    e = _Espia()
    escribir_anotaciones([_anot(destino="FICHA", maquina="TRACTOR X",
                                detalle="patente ABCD12")],
                         "texto", "Juan", set(),
                         _bita=e.bita, _mant=e.mant, _ficha=e.ficha)
    assert len(e.fichas) == 1
    assert e.bitacora == []


def test_lo_retenido_no_se_escribe():
    e = _Espia()
    escritas = escribir_anotaciones([_anot(), _anot(actividad="Riego")],
                                    "texto", "Juan", {1},
                                    _bita=e.bita, _mant=e.mant, _ficha=e.ficha)
    assert len(e.bitacora) == 1
    assert e.bitacora[0][0]["actividad"] == "Poda"
    assert len(escritas) == 1


def test_si_una_anotacion_revienta_las_demas_igual_se_escriben(tmp_path):
    """Nunca perder el resto del parte por una fila mala."""
    e = _Espia()

    def bita_que_falla(campos, quien):
        if campos["actividad"] == "Poda":
            raise RuntimeError("Excel bloqueado")
        return e.bita(campos, quien)

    escritas = escribir_anotaciones([_anot(), _anot(actividad="Riego")],
                                    "texto", "Juan", set(),
                                    _bita=bita_que_falla, _mant=e.mant,
                                    _ficha=e.ficha,
                                    _carpeta_respaldo=str(tmp_path))
    assert len(escritas) == 1
    assert escritas[0]["actividad"] == "Riego"


def test_lo_que_no_se_pudo_escribir_queda_respaldado(tmp_path):
    """Excel bloqueado no puede costar el parte."""
    def siempre_falla(campos, quien):
        raise RuntimeError("Excel bloqueado")

    escribir_anotaciones([_anot()], "el parte de Juan", "Juan Parada", set(),
                         _bita=siempre_falla, _mant=None, _ficha=None,
                         _carpeta_respaldo=str(tmp_path))
    respaldo = tmp_path / "bitacora_fallback.txt"
    assert respaldo.exists()
    assert "el parte de Juan" in respaldo.read_text(encoding="utf-8")


def test_los_tests_no_escriben_en_el_respaldo_de_produccion(tmp_path):
    """El 26-ago-2026 los tests ensuciaron la cola de Drive por esto mismo."""
    import handlers.partes as mod
    fuente = __import__("inspect").getsource(mod.escribir_anotaciones)
    assert "_carpeta_respaldo" in fuente


def test_el_texto_original_queda_en_cada_fila():
    e = _Espia()
    escribir_anotaciones([_anot()], "el parte completo", "Juan", set(),
                         _bita=e.bita, _mant=e.mant, _ficha=e.ficha)
    assert e.bitacora[0][0]["texto_original"] == "el parte completo"
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_partes_escritura.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'handlers.partes'`

- [ ] **Step 3: Escribir la implementación**

```python
# handlers/partes.py
# -*- coding: utf-8 -*-
"""Lee lo que escribe Juan, lo controla, lo anota y avisa si hay dudas.

Reemplaza a auto_guardar_bitacora para los usuarios en AUTO_SAVE_USERS. Juan no
va a aprender un formato: lo que escriba, la IA lo lee y lo anota donde va.
"""
import logging

logger = logging.getLogger(__name__)


def escribir_anotaciones(anotaciones, texto, quien, retenidos,
                         _bita=None, _mant=None, _ficha=None,
                         _carpeta_respaldo=None) -> list:
    """Escribe cada anotacion en su destino. Devuelve las que quedaron escritas.

    Los `_bita`/`_mant`/`_ficha` son para poder probar esto sin tocar el Excel, y
    `_carpeta_respaldo` para que los tests no escriban en el archivo de
    produccion. El 26-ago-2026 pasó justo eso con la cola de Drive: los tests
    pasaban rutas propias pero el encolado leia la ruta de config, y se
    acumularon 13 entradas apuntando a carpetas de pytest ya borradas.

    Una anotacion que revienta NO se lleva a las demas: perder el parte entero
    por una fila mala seria peor que perder la fila.
    """
    if _bita is None:
        from bitacora_manager import registrar_bitacora_estructurada as _bita
    if _mant is None:
        from modules.maquinaria import registrar_mantencion as _mant
    if _ficha is None:
        from modules.maquinaria import guardar_ficha as _ficha

    escritas = []
    for i, a in enumerate(anotaciones):
        if i in retenidos:
            logger.info("Parte: anotación %d retenida por el control", i)
            continue
        try:
            destino = a.get("destino")
            if destino == "MANTENCION":
                _mant({"maquina": a.get("maquina"), "fecha": a.get("fecha"),
                       "detalle": a.get("detalle") or a.get("actividad"),
                       "odometro": a.get("odometro")}, quien)
            elif destino == "FICHA":
                _ficha({"maquina": a.get("maquina"),
                        "notas": a.get("detalle") or a.get("actividad")})
            else:
                campos = {
                    "fecha": a.get("fecha") or "",
                    "tipo": "MAQUINARIA" if destino == "HOROMETRO" else _tipo_de(a),
                    "actividad": (a.get("actividad")
                                  or ("Lectura de horómetro"
                                      if destino == "HOROMETRO" else "")),
                    "cultivo": a.get("cultivo") or "GENERAL",
                    "sector": a.get("sector") or "",
                    "jornadas_hombre": a.get("jornadas_hombre"),
                    "trabajadores": a.get("trabajadores") or [],
                    "insumo": a.get("insumo") or "",
                    "cantidad": a.get("cantidad"),
                    "unidad": a.get("unidad") or "",
                    "maquina": a.get("maquina") or "",
                    "odometro": a.get("odometro"),
                    "superficie_ha": a.get("superficie_ha"),
                    "texto_original": texto,
                }
                res = _bita(campos, quien)
                if isinstance(res, dict) and res.get("error_odometro"):
                    logger.warning("Parte: odómetro rechazado — %s",
                                   res["error_odometro"])
                    continue
                # Guardar en qué fila quedó: es lo que necesita el botón 🗑️
                if isinstance(res, dict) and res.get("fila"):
                    a = {**a, "_fila": res["fila"]}
            escritas.append(a)
        except Exception as e:
            logger.error("Parte: no pude escribir la anotación %d (%s): %s",
                         i, a.get("destino"), e)
            _a_respaldo(texto, quien, e, _carpeta_respaldo)
    return escritas


def _a_respaldo(texto, quien, error, carpeta=None):
    """Excel bloqueado u otro fallo: la entrada NUNCA se pierde."""
    import os
    from datetime import datetime
    try:
        carpeta = carpeta or os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "files", "logs")
        os.makedirs(carpeta, exist_ok=True)
        with open(os.path.join(carpeta, "bitacora_fallback.txt"), "a",
                  encoding="utf-8") as fh:
            fh.write("%s | %s | %s | %s\n"
                     % (datetime.now().isoformat(), quien, error, texto))
    except Exception as e:
        logger.error("Parte: falló también el respaldo a archivo: %s", e)


def _tipo_de(a) -> str:
    from modules.bitacora_asistencia import tipo_de
    return tipo_de(a.get("actividad") or "")
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_partes_escritura.py -q`
Expected: PASS, 9 passed

- [ ] **Step 5: Commit**

```bash
git add handlers/partes.py tests/test_partes_escritura.py
git commit -m "Cada anotacion va a su destino y lo retenido no se escribe

Una anotacion que revienta no se lleva a las demas: perder el parte entero por
una fila mala seria peor que perder la fila.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 7: Poder borrar lo que se anotó

**Files:**
- Modify: `bitacora_manager.py` (agregar `borrar_filas`)
- Test: `tests/test_bitacora_borrar.py`

Es el escape del "se guarda siempre": sin esto, una fila mal leída solo se arregla a mano en el Excel. No existe hoy — `excel_manager` tiene `delete_last_rows` pero es de la hoja `Facturas`.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_bitacora_borrar.py
# -*- coding: utf-8 -*-
"""Borrar filas de la bitacora por numero de fila.

Es el escape del "se guarda siempre" del modo capataz: el dueno tiene que poder
deshacer lo que la IA leyo mal sin abrir el Excel a mano.

OJO: se borran de mayor a menor. Borrando de menor a mayor, la primera baja a
todas las de abajo y las siguientes se llevan la fila equivocada.
"""
import pytest
from openpyxl import Workbook, load_workbook

from bitacora_manager import borrar_filas


@pytest.fixture
def excel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(["Fecha", "Hora", "Tipo", "Actividad"])
    for i in range(1, 6):
        ws.append(["2026-08-%02d" % i, "10:00", "LABOR", "Labor %d" % i])
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def actividades(ruta):
    wb = load_workbook(ruta, read_only=True, data_only=True)
    out = [r[3] for r in wb["Bitácora"].iter_rows(min_row=2, values_only=True)]
    wb.close()
    return out


def test_borra_una_fila(excel):
    assert borrar_filas([3], excel) == 1
    assert actividades(excel) == ["Labor 1", "Labor 3", "Labor 4", "Labor 5"]


def test_borra_varias_sin_correrse(excel):
    """El bug clasico: borrar de menor a mayor se lleva la fila equivocada."""
    assert borrar_filas([2, 4, 6], excel) == 3
    assert actividades(excel) == ["Labor 2", "Labor 4"]


def test_el_orden_en_que_se_piden_da_igual(excel):
    assert borrar_filas([6, 2, 4], excel) == 3
    assert actividades(excel) == ["Labor 2", "Labor 4"]


def test_no_borra_el_encabezado(excel):
    assert borrar_filas([1], excel) == 0
    assert len(actividades(excel)) == 5


def test_una_fila_que_no_existe_se_ignora(excel):
    assert borrar_filas([99], excel) == 0
    assert len(actividades(excel)) == 5


def test_sin_filas_no_hace_nada(excel):
    assert borrar_filas([], excel) == 0
    assert len(actividades(excel)) == 5
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_bitacora_borrar.py -q`
Expected: FAIL con `ImportError: cannot import name 'borrar_filas'`

- [ ] **Step 3: Escribir la implementación**

Agregar a `bitacora_manager.py`:

```python
def borrar_filas(filas: list, excel_path: str | None = None) -> int:
    """Borra filas de la bitacora por numero de fila. Devuelve cuantas borro.

    Es el escape del "se guarda siempre" del modo capataz.

    OJO: se borran de MAYOR A MENOR. Borrando al reves, la primera baja de
    posicion a todas las de abajo y las siguientes se llevan la fila equivocada.
    """
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    ruta = excel_path or EXCEL_PATH
    wb = load_workbook(ruta)
    try:
        ws = wb[BITACORA_SHEET]
        objetivo = sorted({int(f) for f in filas
                           if isinstance(f, (int, float)) and 2 <= int(f) <= ws.max_row},
                          reverse=True)
        for fila in objetivo:
            ws.delete_rows(fila)
        if objetivo:
            _save_wb(wb)
            logger.info("Bitácora: borradas %d fila(s): %s",
                        len(objetivo), objetivo)
        return len(objetivo)
    finally:
        wb.close()
```

⚠️ `registrar_bitacora_estructurada` hace `ws.append()`, así que la fila que acaba de escribir es `ws.max_row`. Para que el botón sepa qué borrar, hay que devolver ese número: agregar `"fila": ws.max_row` al dict que devuelve, justo antes del `return` final.

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_bitacora_borrar.py -q`
Expected: PASS, 6 passed

- [ ] **Step 5: Commit**

```bash
git add bitacora_manager.py tests/test_bitacora_borrar.py
git commit -m "Poder borrar filas de la bitacora

Es el escape del 'se guarda siempre': sin esto una fila que la IA leyo mal solo
se arregla a mano en el Excel. Se borran de mayor a menor, porque al reves la
primera baja a las de abajo y las siguientes se llevan la fila equivocada.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 8: El aviso al dueño y los botones

**Files:**
- Modify: `handlers/partes.py` (agregar `procesar_parte`, `_texto_aviso` y los callbacks)
- Test: `tests/test_partes_aviso.py`

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_partes_aviso.py
# -*- coding: utf-8 -*-
"""El aviso lleva el texto de Juan TAL CUAL, lo anotado y el motivo."""
from handlers.partes import _texto_aviso


def _duda(regla, detalle, retiene=False, datos=None):
    return {"regla": regla, "detalle": detalle, "retiene": retiene,
            "indice": None, "datos": datos or {}}


ESCRITA = {"destino": "BITACORA", "fecha": "2026-08-31", "actividad": "Poda",
           "jornadas_hombre": 6, "trabajadores": ["Felicito Amigo"],
           "maquina": "", "odometro": None}


def test_lleva_el_texto_original_completo():
    txt = _texto_aviso("Lunes 31\nFelicito amigo poda", [ESCRITA],
                       [_duda("confianza_baja", "quedó con 40%")], "Juan Parada")
    assert "Felicito amigo poda" in txt


def test_dice_el_motivo():
    txt = _texto_aviso("t", [ESCRITA], [_duda("confianza_baja", "quedó con 40%")],
                       "Juan Parada")
    assert "quedó con 40%" in txt


def test_muestra_lo_que_quedo_anotado():
    txt = _texto_aviso("t", [ESCRITA], [_duda("confianza_baja", "x")], "Juan")
    assert "Poda" in txt
    assert "2026-08-31" in txt


def test_avisa_cuando_algo_no_se_guardo():
    txt = _texto_aviso("t", [], [_duda("odometro_no_calza", "leyó 5273 y dice 5237",
                                       retiene=True)], "Juan")
    assert "no guardé" in txt.lower()


def test_no_se_pasa_del_limite_de_telegram():
    txt = _texto_aviso("x" * 6000, [ESCRITA] * 50,
                       [_duda("confianza_baja", "y" * 500)], "Juan")
    assert len(txt) <= 4000
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_partes_aviso.py -q`
Expected: FAIL con `ImportError: cannot import name '_texto_aviso'`

- [ ] **Step 3: Escribir la implementación**

Agregar al final de `handlers/partes.py`:

```python
from telegram import InlineKeyboardButton, InlineKeyboardMarkup

MAX_TELEGRAM = 4000


def _linea_de(a) -> str:
    if a.get("destino") == "HOROMETRO":
        return "• %s — %s %s" % (a.get("fecha") or "?", a.get("maquina"),
                                 a.get("odometro"))
    gente = ", ".join(n.split()[0] for n in (a.get("trabajadores") or []))
    jh = a.get("jornadas_hombre")
    return "• %s — %s%s%s" % (a.get("fecha") or "?", a.get("actividad") or "?",
                              " (%s JH)" % jh if jh else "",
                              ": %s" % gente if gente else "")


def _texto_aviso(texto_original, escritas, dudas, quien) -> str:
    partes = ["🔎 *Revisa este parte de %s*" % quien, ""]
    for d in dudas:
        partes.append("⚠️ %s" % d["detalle"])
    retenidas = [d for d in dudas if d["retiene"]]
    if retenidas:
        partes.append("")
        partes.append("🚫 *No guardé* %d anotación(es) por eso." % len(retenidas))
    partes.append("")
    if escritas:
        partes.append("*Quedó anotado:*")
        partes.extend(_linea_de(a) for a in escritas[:15])
        if len(escritas) > 15:
            partes.append("… y %d más" % (len(escritas) - 15))
    else:
        partes.append("*No quedó nada anotado.*")
    partes.append("")
    partes.append("*Lo que escribió:*")
    partes.append(texto_original)
    txt = "\n".join(partes)
    if len(txt) > MAX_TELEGRAM:
        txt = txt[:MAX_TELEGRAM - 20].rstrip() + "\n…(recortado)"
    return txt


async def procesar_parte(update, context):
    """Punto de entrada del modo capataz: lee, controla, anota y avisa."""
    import asyncio
    from datetime import date

    from modules.bitacora_extractor import es_mensaje_sin_contenido
    from modules.parte_contexto import construir
    from modules.parte_control import indices_retenidos, revisar
    from modules.parte_extractor import leer_parte

    texto = (update.message.text or "").strip()
    if len(texto) < 3 or es_mensaje_sin_contenido(texto):
        await update.message.reply_text(
            "📓 Te leo, pero ahí no viene nada que anotar.\n\n"
            "Mándame el parte completo y yo lo ordeno.")
        return

    quien = update.effective_user.full_name if update.effective_user else ""
    estado = await update.message.reply_text("📓 Leyendo el parte…")

    ctx = await asyncio.to_thread(construir)
    parte = await asyncio.to_thread(leer_parte, texto,
                                    date.today().strftime("%Y-%m-%d"), ctx)
    dudas = revisar(texto, parte, ctx)

    escritas = await asyncio.to_thread(
        escribir_anotaciones, parte["anotaciones"], texto, quien,
        indices_retenidos(dudas))

    if not escritas and parte.get("error"):
        # Nunca perder la entrada: se guarda cruda como OTRO
        await asyncio.to_thread(_guardar_crudo, texto, quien)
        await estado.edit_text(
            "⚠️ No pude leerlo ahora, pero lo guardé completo. No se perdió.")
    else:
        await estado.edit_text(_respuesta_a_juan(escritas))

    if dudas:
        await _avisar_al_dueno(context, texto, escritas, dudas, quien)


def _respuesta_a_juan(escritas) -> str:
    if not escritas:
        return "📓 Lo recibí, pero no pude anotar nada. Ya le avisé al patrón."
    lineas = ["✅ Anotado:"] + [_linea_de(a) for a in escritas[:15]]
    if len(escritas) > 15:
        lineas.append("… y %d más" % (len(escritas) - 15))
    return "\n".join(lineas)[:MAX_TELEGRAM]


def _guardar_crudo(texto, quien):
    from bitacora_manager import registrar_bitacora_estructurada
    registrar_bitacora_estructurada(
        {"fecha": "", "tipo": "OTRO", "actividad": texto[:40],
         "cultivo": "GENERAL", "sector": "", "jornadas_hombre": None,
         "trabajadores": [], "insumo": "", "cantidad": None, "unidad": "",
         "maquina": "", "odometro": None, "superficie_ha": None,
         "texto_original": texto}, quien)


async def _avisar_al_dueno(context, texto, escritas, dudas, quien):
    owner = context.bot_data.get("owner_chat_id")
    if not owner:
        logger.warning("Parte con dudas y no hay owner_chat_id: %s",
                       [d["regla"] for d in dudas])
        return
    nuevos = [n for d in dudas if d["regla"] == "trabajadores_nuevos"
              for n in d["datos"].get("nombres", [])]
    filas = [a["_fila"] for a in escritas if a.get("_fila")]
    context.bot_data.setdefault("partes_pendientes", {})
    clave = str(update_id_de(context))
    context.bot_data["partes_pendientes"][clave] = {"nuevos": nuevos,
                                                    "filas": filas}
    fila = [InlineKeyboardButton("✅ Está bien", callback_data="parte_ok:%s" % clave)]
    if filas:
        fila.append(InlineKeyboardButton("🗑️ Borrar lo anotado",
                                         callback_data="parte_borrar:%s" % clave))
    if nuevos:
        fila.append(InlineKeyboardButton("➕ Dar de alta a los nuevos",
                                         callback_data="parte_alta:%s" % clave))
    try:
        await context.bot.send_message(
            owner, _texto_aviso(texto, escritas, dudas, quien),
            parse_mode="Markdown", reply_markup=InlineKeyboardMarkup([fila]))
    except Exception:
        await context.bot.send_message(
            owner, _texto_aviso(texto, escritas, dudas, quien),
            reply_markup=InlineKeyboardMarkup([fila]))


def update_id_de(context) -> int:
    """Un identificador corto y unico para el aviso."""
    n = context.bot_data.get("_parte_seq", 0) + 1
    context.bot_data["_parte_seq"] = n
    return n


async def cb_parte_ok(update, context):
    query = update.callback_query
    await query.answer("Listo")
    await query.edit_message_reply_markup(reply_markup=None)


async def cb_parte_borrar(update, context):
    query = update.callback_query
    await query.answer()
    import asyncio

    from bitacora_manager import borrar_filas
    clave = query.data.split(":", 1)[1]
    datos = (context.bot_data.get("partes_pendientes") or {}).get(clave) or {}
    filas = datos.get("filas") or []
    if not filas:
        await query.edit_message_reply_markup(reply_markup=None)
        await query.message.reply_text("⚠️ Ya no tengo esas filas registradas.")
        return
    borradas = await asyncio.to_thread(borrar_filas, filas)
    datos["filas"] = []                     # que no se pueda borrar dos veces
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text("🗑️ Borré %d fila(s) de la bitácora." % borradas)


async def cb_parte_alta(update, context):
    query = update.callback_query
    await query.answer()
    import asyncio

    from vacaciones_manager import agregar_trabajador
    clave = query.data.split(":", 1)[1]
    datos = (context.bot_data.get("partes_pendientes") or {}).get(clave) or {}
    nuevos = datos.get("nuevos") or []
    altas = 0
    for nombre in nuevos:
        try:
            if await asyncio.to_thread(agregar_trabajador, nombre, "", "Temporero", ""):
                altas += 1
        except Exception as e:
            logger.error("No pude dar de alta a %s: %s", nombre, e)
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text("➕ Di de alta a %d trabajador(es)." % altas)
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_partes_aviso.py -q`
Expected: PASS, 5 passed

- [ ] **Step 5: Commit**

```bash
git add handlers/partes.py tests/test_partes_aviso.py
git commit -m "El aviso al dueno con el texto de Juan tal cual y el motivo

Se guarda siempre para no frenar a Juan; al dueno le llega solo lo dudoso, con
lo que quedo anotado, lo que no se guardo y por que, y boton para dar de alta a
la gente nueva.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 9: Cablear y probar en vivo

**Files:**
- Modify: `handlers/chat.py:44-58`
- Modify: `main.py` (imports y registro de callbacks)
- Test: `tests/test_chat_modo_capataz.py`

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_chat_modo_capataz.py
# -*- coding: utf-8 -*-
"""En modo capataz manda la IA, no el porton parece_maquinaria.

El porton decidia el destino ANTES de leer: un parte mixto (asistencia arriba,
horometro abajo) se iba entero por el camino de maquinaria y la asistencia se
perdia. Ahora la IA lee el mensaje completo y decide.
"""
import inspect

import handlers.chat as chat


def test_el_modo_capataz_llama_a_procesar_parte():
    fuente = inspect.getsource(chat.handle_text)
    assert "procesar_parte" in fuente


def test_el_porton_de_maquinaria_ya_no_precede_al_modo_capataz():
    fuente = inspect.getsource(chat.handle_text)
    i_capataz = fuente.index("AUTO_SAVE_USERS")
    i_maquinaria = fuente.index("parece_maquinaria")
    assert i_capataz < i_maquinaria, (
        "el portón de maquinaria vuelve a decidir antes que la IA")


def test_auto_guardar_bitacora_ya_no_se_usa_en_el_dispatcher():
    assert "auto_guardar_bitacora" not in inspect.getsource(chat.handle_text)
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_chat_modo_capataz.py -q`
Expected: FAIL, 3 failed

- [ ] **Step 3: Cambiar `handlers/chat.py`**

Reemplazar el bloque que hoy va desde `# ── Maquinaria: horómetros...` hasta el `return` del modo capataz por:

```python
    # ── Modo capataz: la IA lee el mensaje entero y decide qué es ──
    # Va ANTES del portón de maquinaria a propósito. El portón decidía el
    # destino sin leer, así que un parte mixto —asistencia arriba, horómetro
    # abajo— se iba entero por el camino de maquinaria y la asistencia se perdía.
    from config import AUTO_SAVE_USERS
    if update.effective_user and update.effective_user.id in AUTO_SAVE_USERS:
        from handlers.partes import procesar_parte
        await procesar_parte(update, context)
        return

    # ── Maquinaria: horómetros, mantenciones y fichas (para el dueño) ──
    from handlers.maquinaria import (modo_activo, parece_maquinaria,
                                      procesar_texto_maquinaria)
    if modo_activo(context) or parece_maquinaria(update.message.text or ""):
        if await procesar_texto_maquinaria(update, context):
            return
```

- [ ] **Step 4: Registrar los callbacks en `main.py`**

Junto a los otros `from handlers...` de bitácora, agregar:

```python
from handlers.partes import cb_parte_ok, cb_parte_borrar, cb_parte_alta
```

Y junto a los `add_handler` de `cb_bita_*`:

```python
    app.add_handler(CallbackQueryHandler(cb_parte_ok,     pattern="^parte_ok:"))
    app.add_handler(CallbackQueryHandler(cb_parte_borrar, pattern="^parte_borrar:"))
    app.add_handler(CallbackQueryHandler(cb_parte_alta,   pattern="^parte_alta:"))
```

- [ ] **Step 5: Correr toda la suite**

Run: `py -m pytest tests/ -q`
Expected: PASS, cero fallos. Antes de este plan la suite eran **710**; cada tarea suma los suyos, así que el total tiene que ser bastante mayor y **ninguno de los 710 viejos puede haberse caído**. Si el total bajó, lo más probable es que el `pytest.ini` de la Task 4 haya cambiado el descubrimiento de tests.

- [ ] **Step 6: Commit**

```bash
git add handlers/chat.py main.py tests/test_chat_modo_capataz.py
git commit -m "En modo capataz manda la IA, no el porton parece_maquinaria

El porton decidia el destino ANTES de leer, asi que un parte mixto se iba entero
por el camino de maquinaria y la asistencia se perdia. Para el dueno, que usa
comandos, el camino de hoy no cambia.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

- [ ] **Step 7: Reiniciar el bot y probar en vivo**

⚠️ **Antes de matar el bot, mirar el log** (`tail -40 bot.log`): si hay una lectura en curso, esperar.

```powershell
Stop-ScheduledTask -TaskName "AgricolaBotWatchdog"
Get-CimInstance Win32_Process -Filter "Name = 'python3.11.exe'" | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
@(Get-CimInstance Win32_Process -Filter "Name = 'python3.11.exe'").Count   # tiene que dar 0
Start-ScheduledTask -TaskName "AgricolaBotWatchdog"
```

Después mandar por Telegram, desde la cuenta de Juan, el parte del 31-ago tal cual está en el respaldo. Comprobar:
1. Juan recibe `✅ Anotado:` con **varias** líneas.
2. En la hoja `Bitácora` quedan las labores separadas y suman **19 JH**.
3. Al dueño le llega el aviso de trabajadores nuevos con el botón.

---

## Task 10: Reingresar los 6 partes que quedaron esperando

**Files:**
- Modify: `scripts/carga/recuperar_bitacora_perdida.py`

⚠️ **El script de hoy trabaja con dicts de `registrar_bitacora_estructurada`, que llevan `tipo`. Las `Anotacion` del lector llevan `destino` y NO llevan `tipo`.** Si se le pasan directo, tanto el `print` de `main()` como la escritura revientan con `KeyError: 'tipo'`. Por eso la escritura pasa a hacerla `escribir_anotaciones`, que ya sabe traducir y está probada.

- [ ] **Step 1: Cambiar `PERDIDOS` para que queden solo los 6**

Los otros 8 (`2933, 3130, 3143, 3146, 3149, 3152, 3155, 3158`) ya se reingresaron el 2-sep-2026. Dejarlos duplicaría las filas. La lista queda:

```python
PERDIDOS = [
    (2934, "Asistencia Martes 25 de agosto"),
    (2992, "Asistencia miércoles 26 agosto"),
    (3000, "Asistencia jueves 27 de agosto"),
    (3003, "Asistencia viernes 28 de agosto"),
    (3106, "Lunes 31 de agosto 2026"),
    (3122, "Martes 1 de septiembre 2026"),
]
```

- [ ] **Step 2: Cambiar `_filas_de` para que use el lector nuevo**

Reemplazar el cuerpo entero de `_filas_de` por:

```python
def _filas_de(fila):
    """Lee el mensaje con la IA. Devuelve (anotaciones, motivo)."""
    from modules.parte_contexto import construir
    from modules.parte_control import indices_retenidos, revisar
    from modules.parte_extractor import leer_parte

    texto = fila["text"]
    ctx = construir()
    parte = leer_parte(texto, fila["recibido_utc"][:10], ctx)
    if parte.get("error"):
        return [], parte["error"]
    dudas = revisar(texto, parte, ctx)
    retenidos = indices_retenidos(dudas)
    filas = [a for i, a in enumerate(parte["anotaciones"]) if i not in retenidos]
    motivos = [d["detalle"] for d in dudas if d["regla"] in
               ("lineas_sin_anotar", "parser_vio_mas", "odometro_no_calza",
                "maquina_desconocida")]
    return filas, ("; ".join(motivos) if motivos else None)
```

- [ ] **Step 3: Cambiar `main()` para que imprima y escriba `Anotacion`**

Reemplazar el `print` de cada fila y el bloque de escritura de `main()` por:

```python
    filas.sort(key=lambda f: (f["fecha"], f["destino"] != "HOROMETRO"))

    print("%d anotaciones a reingresar" % len(filas))
    print()
    for f in filas:
        extra = ("  %s odo=%g" % (f["maquina"], f["odometro"])
                 if f["destino"] == "HOROMETRO"
                 else "  %s JH  %s" % (f["jornadas_hombre"] or 0,
                                       ", ".join(f["trabajadores"])))
        print("  %s  %-10s %-38s%s" % (f["fecha"], f["destino"],
                                       f["actividad"][:38], extra))
```

y el bloque que escribe por:

```python
    from handlers.partes import escribir_anotaciones
    escritas = escribir_anotaciones(filas, "reingreso desde el respaldo crudo",
                                    QUIEN, set())
    print()
    print("%d anotaciones escritas" % len(escritas))
```

- [ ] **Step 4: Simular y mirar**

Run: `py scripts/carga/recuperar_bitacora_perdida.py --simular`

Expected: los 6 partes que antes salían en PENDIENTES ahora aparecen con sus anotaciones. Comprobar a ojo contra el texto de Juan que el del 31-ago dé **19 personas** y que los del herbicida no estén en la poda.

- [ ] **Step 5: Respaldar el Master antes de escribir**

```bash
cp "../MASTER Agricola Santa Elisa.xlsx" "../MASTER Agricola Santa Elisa_bak_antes_reingreso_partes.xlsx"
```

- [ ] **Step 6: Reingresar de verdad**

Run: `py scripts/carga/recuperar_bitacora_perdida.py`
Expected: `N anotaciones escritas`, sin líneas de error.

- [ ] **Step 7: Comprobar contra la hoja**

```bash
py -c "import sys;sys.path.insert(0,'.');from openpyxl import load_workbook;from config import EXCEL_PATH;wb=load_workbook(EXCEL_PATH,read_only=True,data_only=True);ws=wb['Bitácora'];rows=[r for r in ws.iter_rows(min_row=2,values_only=True) if str(r[0]).startswith(('2026-08-25','2026-08-26','2026-08-27','2026-08-28','2026-08-31','2026-09-01'))];print(len(rows),'filas');[print(r[0],r[3],r[6],r[7]) for r in rows]"
```
Expected: los días 25, 26, 27 y 28-ago con **7 JH** cada uno, y el 31-ago y 1-sep con **19 JH** repartidos por labor.

- [ ] **Step 8: Commit**

```bash
git add scripts/carga/recuperar_bitacora_perdida.py
git commit -m "Reingresados los 6 partes que el parser no sabia leer

Los que no traian dos puntos y el del 26-ago, donde parsear_asistencia leia 2
jornadas de 7. Ahora los lee la IA y quedan con la gente de cada labor separada.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Al terminar

- [ ] Correr la suite entera: `py -m pytest tests/ -q`
- [ ] Actualizar `project-asistencia-sin-dos-puntos` en memoria: pasa de PENDIENTE a resuelto
- [ ] Actualizar la cabecera de `project-pendientes-roadmap`
- [ ] Pushear
