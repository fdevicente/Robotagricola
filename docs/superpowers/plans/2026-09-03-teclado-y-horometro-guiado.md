# Teclado fijo y horómetro guiado — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que Juan tenga siempre tres botones a la vista para lo único que hace —asistencia, horómetro y factura— y que la lectura de horómetro se ingrese guiada, validando el número contra la última lectura mientras él todavía está frente a la máquina.

**Architecture:** Un teclado persistente de Telegram (`ReplyKeyboardMarkup`) con tres botones. Dos de ellos solo instruyen y no dejan ningún estado abierto. El tercero, el horómetro, es un flujo corto de cuatro pasos con botones, que escribe por `registrar_bitacora_estructurada` como todo lo demás.

**Tech Stack:** Python 3.11, python-telegram-bot 20+, openpyxl, pytest.

**Diseño:** `docs/superpowers/specs/2026-09-02-partes-juan-lectura-ia-design.md`, sección «Revisión del 3-sep-2026».

---

## Por qué este plan va antes que el de la IA

Se midió qué manda Juan: **15 intentos de comando**, 8 fotos de factura, 8 partes de horómetro y 7 de asistencia. Muchos de esos comandos venían rotos (`/ cancelar`, `/ bitacora`, `/ Asistencia`, `/` a secas). **Juan busca un menú y no lo encuentra.**

Para el horómetro los botones son mejores que la IA, no un parche: son tres datos, la máquina sale de una lista cerrada, y el número se contrasta contra la última lectura **en el momento**. Eso mata los dos riesgos que el diseño con IA obligaba a vigilar — la máquina inventada y el error de dígitos.

Para la asistencia no: el parte del 31-ago trae 19 personas en 3 labores, y preguntarle uno por uno serían ~38 idas y vueltas. Ese sigue siendo trabajo de la IA, en el otro plan.

---

## Antes de empezar

**No hay `python` en el PATH.** El intérprete es `C:\Users\Windows\AppData\Local\Python\bin\python3.11.exe`. En Git Bash: `alias py='/c/Users/Windows/AppData/Local/Python/bin/python3.11.exe'`.

Si se trabaja en un worktree, ver las tres copias que hacen falta para que la suite arranque en verde, anotadas en `2026-09-02-partes-juan-lectura-ia.md`.

**Mensajes de commit en español**, terminados en `Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>`.

## 🔴 La regla que no se puede olvidar

**Todo flujo conversacional nuevo tiene que registrar su clave de estado en `modules/flujos.py`.** Un `/deposito` sin cerrar se comió 12 días de partes de Juan en silencio (ver `project-flujo-trabado-bitacora`). La Task 3 agrega un flujo; si su `horo_state` no queda en `CLAVES_ESTADO`, no caduca nunca y el bug vuelve. La Task 3 incluye un test que falla si alguien vuelve a agregar un flujo sin registrarlo.

---

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `modules/opciones_capataz.py` (nuevo) | Qué botones ofrecer: máquinas por lectura más reciente y labores más frecuentes, sacadas del Master. |
| `handlers/teclado.py` (nuevo) | El teclado persistente y el ruteo de sus tres botones. |
| `handlers/horometro.py` (nuevo) | El flujo guiado de cuatro pasos. |
| `modules/flujos.py` (modificar) | Registrar `horo_state` y `horo_data`. |
| `handlers/chat.py` (modificar) | Los botones se atienden antes que nada. |
| `main.py` (modificar) | Teclado en `/start` y registro de los callbacks. |

---

## Task 1: Qué botones ofrecer

**Files:**
- Create: `modules/opciones_capataz.py`
- Test: `tests/test_opciones_capataz.py`

Los botones **no van escritos a mano**. Las labores cambian con la temporada —hoy es poda, en marzo es cosecha— y una lista fija envejece sin que nadie la note. Salen del propio Master.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_opciones_capataz.py
# -*- coding: utf-8 -*-
"""Los botones salen del Master, no van escritos a mano.

Las labores cambian con la temporada: hoy poda, en marzo cosecha. Una lista fija
envejece sin que nadie lo note y Juan termina apretando "Otra..." siempre.
"""
from openpyxl import Workbook

from modules.opciones_capataz import labores_frecuentes, maquinas_recientes

CTX = {"maquinas": [
    {"maquina": "TRACTOR MASSEY FERGUSON 6711", "ultimo_odometro": 2057,
     "fecha": "2026-09-01", "unidad": "h"},
    {"maquina": "TRACTOR MASSEY FERGUSON 4292", "ultimo_odometro": 5239,
     "fecha": "2026-09-01", "unidad": "h"},
    {"maquina": "EXCAVADORA", "ultimo_odometro": 7240, "fecha": "2026-06-12",
     "unidad": "h"},
    {"maquina": "CAMIONETA RAM MODELO 1500", "ultimo_odometro": None,
     "fecha": None, "unidad": "km"},
]}


def test_las_maquinas_van_de_mas_reciente_a_mas_vieja():
    m = maquinas_recientes(CTX)
    assert m[0] in ("TRACTOR MASSEY FERGUSON 6711", "TRACTOR MASSEY FERGUSON 4292")
    assert m.index("EXCAVADORA") > m.index("TRACTOR MASSEY FERGUSON 6711")


def test_las_maquinas_sin_ninguna_lectura_van_al_final():
    """Nunca se usaron; no pueden ocupar los primeros botones."""
    m = maquinas_recientes(CTX)
    assert m[-1] == "CAMIONETA RAM MODELO 1500"


def test_no_devuelve_una_pared_de_botones():
    ctx = {"maquinas": [{"maquina": "M%d" % i, "ultimo_odometro": i,
                         "fecha": "2026-01-%02d" % (i + 1), "unidad": "h"}
                        for i in range(20)]}
    assert len(maquinas_recientes(ctx)) <= 6


def test_sin_maquinas_devuelve_lista_vacia():
    assert maquinas_recientes({"maquinas": []}) == []
    assert maquinas_recientes({}) == []


def _excel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(["Fecha", "Hora", "Tipo", "Actividad"])
    for _ in range(9):
        ws.append(["2026-08-20", "10:00", "LABOR", "Poda nogales"])
    for _ in range(4):
        ws.append(["2026-08-20", "10:00", "LABOR", "Aplicación herbicida"])
    for _ in range(7):
        ws.append(["2026-08-20", "10:00", "MAQUINARIA", "Lectura de horómetro"])
    ws.append(["2026-08-20", "10:00", "LABOR", "poda nogales"])   # misma, otra grafía
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_las_labores_van_de_mas_a_menos_usada(tmp_path):
    l = labores_frecuentes(_excel(tmp_path))
    assert l[0] == "Poda nogales"
    assert "Aplicación herbicida" in l


def test_lectura_de_horometro_no_es_una_labor(tmp_path):
    """La escribe el propio bot; ofrecersela a Juan como labor no tiene sentido."""
    assert "Lectura de horómetro" not in labores_frecuentes(_excel(tmp_path))


def test_la_misma_labor_con_otra_grafia_no_se_cuenta_dos_veces(tmp_path):
    l = labores_frecuentes(_excel(tmp_path))
    assert sum(1 for x in l if x.lower() == "poda nogales") == 1


def test_un_excel_ilegible_no_revienta(tmp_path):
    assert labores_frecuentes(str(tmp_path / "no_existe.xlsx")) == []
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_opciones_capataz.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'modules.opciones_capataz'`

- [ ] **Step 3: Escribir la implementación**

```python
# modules/opciones_capataz.py
# -*- coding: utf-8 -*-
"""Que botones ofrecerle a Juan. Salen del Master, no van escritos a mano.

Las labores cambian con la temporada --hoy poda, en marzo cosecha-- y una lista
fija envejece sin que nadie lo note: Juan terminaria apretando "Otra..." siempre
y el boton dejaria de servir.
"""
import logging

logger = logging.getLogger(__name__)

MAX_MAQUINAS = 6
MAX_LABORES = 6
BITACORA_SHEET = "Bitácora"

# La escribe el propio bot al guardar una lectura: no es una labor que Juan haga.
_NO_ES_LABOR = {"lectura de horómetro", "lectura de horometro"}


def maquinas_recientes(ctx: dict, tope: int = MAX_MAQUINAS) -> list:
    """Maquinas ordenadas por lectura mas reciente. Las que nunca se leyeron, al final."""
    maquinas = ctx.get("maquinas") or []
    def _orden(m):
        # Sin lectura -> al fondo. str() porque la fecha viene como date o texto.
        return (m.get("fecha") is None, "" if m.get("fecha") is None
                else str(m["fecha"]))
    ordenadas = sorted(maquinas, key=_orden, reverse=False)
    con_fecha = [m for m in ordenadas if m.get("fecha") is not None]
    sin_fecha = [m for m in ordenadas if m.get("fecha") is None]
    con_fecha.reverse()                        # la más reciente primero
    return [m["maquina"] for m in (con_fecha + sin_fecha)][:tope]


def labores_frecuentes(excel_path: str | None = None,
                       tope: int = MAX_LABORES) -> list:
    """Las labores mas usadas de la bitacora, de mas a menos."""
    from collections import Counter

    from openpyxl import load_workbook

    from config import EXCEL_PATH
    ruta = excel_path or EXCEL_PATH
    cuenta, grafia = Counter(), {}
    try:
        wb = load_workbook(ruta, read_only=True, data_only=True)
        try:
            if BITACORA_SHEET not in wb.sheetnames:
                return []
            ws = wb[BITACORA_SHEET]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 4 or not row[3]:
                    continue
                act = str(row[3]).strip()
                clave = act.lower()
                if clave in _NO_ES_LABOR:
                    continue
                cuenta[clave] += 1
                grafia.setdefault(clave, act)   # la primera grafía que se vio
        finally:
            wb.close()
    except Exception as e:
        logger.warning("opciones_capataz: no pude leer las labores: %r", e)
        return []
    return [grafia[c] for c, _ in cuenta.most_common(tope)]
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_opciones_capataz.py -q`
Expected: PASS, 8 passed

- [ ] **Step 5: Mirar qué botones salen con el Master real**

Run:
```bash
py -c "import sys;sys.path.insert(0,'.');from modules.parte_contexto import construir;from modules.opciones_capataz import maquinas_recientes,labores_frecuentes;print('MAQUINAS:',maquinas_recientes(construir()));print('LABORES:',labores_frecuentes())"
```
Expected: las máquinas empiezan por los tractores que Juan reportó el 1-sep (MF 6711, MF 4292, MF 4275, JD 5085) y las labores incluyen `Poda nogales`, `Poda avellanos` y `Sacar restos poda nogales`. **Si sale `Lectura de horómetro` entre las labores, el filtro no está funcionando.**

- [ ] **Step 6: Commit**

```bash
git add modules/opciones_capataz.py tests/test_opciones_capataz.py
git commit -m "Los botones salen del Master, no van escritos a mano

Las labores cambian con la temporada: hoy poda, en marzo cosecha. Una lista fija
envejece sin que nadie lo note y Juan terminaria apretando Otra siempre. Las
maquinas van por lectura mas reciente y las que nunca se leyeron quedan al final.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 2: El teclado persistente

**Files:**
- Create: `handlers/teclado.py`
- Test: `tests/test_teclado.py`

Dos de los tres botones **no abren ningún flujo**: solo le dicen a Juan qué mandar. Es a propósito — cada estado conversacional nuevo es una forma más de que se trabe, y estos dos no lo necesitan.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_teclado.py
# -*- coding: utf-8 -*-
"""Los tres botones fijos de Juan.

Se midio que manda: 15 intentos de comando (muchos rotos, "/ cancelar",
"/ Asistencia"), 8 fotos de factura, 8 partes de horometro, 7 de asistencia.
Busca un menu y no lo encuentra.

Asistencia y Factura NO abren flujo: solo instruyen. Cada estado conversacional
nuevo es una forma mas de que se trabe, y esos dos no lo necesitan.
"""
from handlers.teclado import (BOTON_ASISTENCIA, BOTON_FACTURA, BOTON_HOROMETRO,
                              es_boton, teclado_capataz, texto_de_ayuda)


def test_el_teclado_trae_los_tres_botones():
    filas = teclado_capataz().keyboard
    textos = [b.text for fila in filas for b in fila]
    assert set(textos) == {BOTON_ASISTENCIA, BOTON_HOROMETRO, BOTON_FACTURA}


def test_el_teclado_es_persistente_y_no_se_esconde():
    kb = teclado_capataz()
    assert kb.resize_keyboard is True
    assert getattr(kb, "one_time_keyboard", False) is not True


def test_reconoce_los_botones():
    assert es_boton(BOTON_ASISTENCIA)
    assert es_boton(BOTON_HOROMETRO)
    assert es_boton(BOTON_FACTURA)


def test_no_confunde_un_parte_con_un_boton():
    """El parte de Juan empieza con la fecha, no puede parecerse a un boton."""
    assert not es_boton("Lunes 31 de agosto 2026\nFelicito amigo poda")
    assert not es_boton("")
    assert not es_boton(None)


def test_reconoce_el_boton_aunque_venga_con_espacios():
    assert es_boton("  " + BOTON_ASISTENCIA + " ")


def test_asistencia_y_factura_solo_instruyen():
    """No abren flujo: el texto tiene que decir QUE mandar."""
    assert "parte" in texto_de_ayuda(BOTON_ASISTENCIA).lower()
    assert "foto" in texto_de_ayuda(BOTON_FACTURA).lower()


def test_el_horometro_no_tiene_texto_de_ayuda():
    """Ese si abre flujo, lo maneja handlers/horometro.py."""
    assert texto_de_ayuda(BOTON_HOROMETRO) is None
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_teclado.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'handlers.teclado'`

- [ ] **Step 3: Escribir la implementación**

```python
# handlers/teclado.py
# -*- coding: utf-8 -*-
"""Los tres botones fijos que Juan tiene siempre sobre el teclado.

POR QUE
Se midio que manda: 15 intentos de comando --muchos rotos, "/ cancelar",
"/ bitacora", "/ Asistencia", "/" a secas--, 8 fotos de factura, 8 partes de
horometro y 7 de asistencia. Busca un menu y no lo encuentra.

Se usa ReplyKeyboardMarkup y no botones inline a proposito: el teclado queda
fijo sobre el de Telegram, siempre visible, sin gastar un mensaje ni obligar a
preguntarle "quieres ingresar datos?" antes de cada cosa.

ASISTENCIA Y FACTURA NO ABREN FLUJO. Solo le dicen que mandar. Cada estado
conversacional nuevo es una forma mas de que se trabe --un /deposito sin cerrar
se comio 12 dias de partes-- y esos dos no lo necesitan.
"""
from telegram import KeyboardButton, ReplyKeyboardMarkup

BOTON_ASISTENCIA = "📋 Asistencia"
BOTON_HOROMETRO = "🚜 Horómetro"
BOTON_FACTURA = "🧾 Factura"

BOTONES = (BOTON_ASISTENCIA, BOTON_HOROMETRO, BOTON_FACTURA)

_AYUDA = {
    BOTON_ASISTENCIA: (
        "📋 Dale, mándame el *parte* del día.\n\n"
        "Escríbelo como te salga, una línea por persona:\n"
        "_Lunes 1 de septiembre 2026_\n"
        "_Felicito amigo sacar restos poda nogales_\n"
        "_Patricio Mora aplicación herbicida_"),
    BOTON_FACTURA: (
        "🧾 Mándame la *foto* de la factura.\n\n"
        "Si quieres, escríbele al lado de qué es "
        "(_«compra petróleo campo»_) y lo anoto."),
}


def teclado_capataz() -> ReplyKeyboardMarkup:
    """El teclado fijo. Queda puesto hasta que alguien lo reemplace."""
    return ReplyKeyboardMarkup(
        [[KeyboardButton(BOTON_ASISTENCIA)],
         [KeyboardButton(BOTON_HOROMETRO), KeyboardButton(BOTON_FACTURA)]],
        resize_keyboard=True, is_persistent=True)


def es_boton(texto) -> bool:
    return str(texto or "").strip() in BOTONES


def texto_de_ayuda(boton) -> str | None:
    """Que responderle. None para los botones que abren un flujo propio."""
    return _AYUDA.get(str(boton or "").strip())
```

- [ ] **Step 4: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_teclado.py -q`
Expected: PASS, 7 passed

⚠️ Si falla por `is_persistent`, la versión de python-telegram-bot es anterior a 20.1. Comprobar con `py -c "import telegram; print(telegram.__version__)"` y reportarlo en vez de sacar el parámetro: sin él el teclado se esconde y el punto del diseño se pierde.

- [ ] **Step 5: Commit**

```bash
git add handlers/teclado.py tests/test_teclado.py
git commit -m "Los tres botones fijos que Juan tiene siempre a la vista

Manda 15 intentos de comando, muchos rotos, contra 23 partes utiles: busca un
menu y no lo encuentra. Teclado persistente y no botones inline, para no gastar
un mensaje ni preguntarle antes de cada cosa.

Asistencia y Factura no abren flujo, solo instruyen: cada estado conversacional
nuevo es otra forma de trabarse.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 3: El flujo guiado de horómetro

**Files:**
- Create: `handlers/horometro.py`
- Modify: `modules/flujos.py`
- Test: `tests/test_horometro_guiado.py`

Cuatro pasos: máquina → inicio → término → labor. El valor está en el paso del inicio: se contrasta contra la última lectura guardada **mientras Juan está frente a la máquina**.

🔴 **Este flujo abre estado. Su clave TIENE que quedar registrada en `modules/flujos.py`** o no caduca nunca y vuelve el bug que costó 12 días.

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_horometro_guiado.py
# -*- coding: utf-8 -*-
"""El flujo guiado de horometro, paso a paso.

Lo que compra frente a leerlo con IA: la maquina sale de una lista cerrada (no
se puede inventar) y el numero de inicio se contrasta contra la ultima lectura
EN EL MOMENTO, con Juan parado frente al tractor. Hoy un error de tipeo entra
callado y descuadra las horas de todas las lecturas siguientes de esa maquina.
"""
import pytest

from handlers.horometro import (PASOS, avanzar, iniciar, revisar_inicio)

CTX = {"maquinas": [
    {"maquina": "TRACTOR MASSEY FERGUSON 4292", "ultimo_odometro": 5239.0,
     "fecha": "2026-09-01", "unidad": "h"},
    {"maquina": "TRACTOR JOHN DEERE 5085", "ultimo_odometro": 3265.0,
     "fecha": "2026-09-01", "unidad": "h"},
]}


def test_iniciar_deja_el_flujo_esperando_la_maquina():
    ud = {}
    iniciar(ud)
    assert ud["horo_state"] == PASOS.MAQUINA


def test_elegir_maquina_pasa_a_pedir_el_inicio():
    ud = {}
    iniciar(ud)
    r = avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    assert ud["horo_state"] == PASOS.INICIO
    assert ud["horo_data"]["maquina"] == "TRACTOR MASSEY FERGUSON 4292"
    assert r["ok"] is True


def test_una_maquina_que_no_existe_no_avanza():
    ud = {}
    iniciar(ud)
    r = avanzar(ud, "TRACTOR FANTASMA", CTX)
    assert ud["horo_state"] == PASOS.MAQUINA
    assert r["ok"] is False


def test_el_inicio_que_calza_con_la_ultima_lectura_pasa_derecho():
    assert revisar_inicio(5239, 5239.0) is None


def test_el_inicio_que_no_calza_avisa_con_los_dos_numeros():
    aviso = revisar_inicio(5137, 5239.0)
    assert aviso is not None
    assert "5.239" in aviso or "5239" in aviso


def test_una_diferencia_chica_tambien_avisa():
    """5237 vs 5239 es justo el error de tipeo que hay que cazar."""
    assert revisar_inicio(5237, 5239.0) is not None


def test_sin_lectura_previa_no_hay_con_que_contrastar():
    assert revisar_inicio(5239, None) is None


def test_un_termino_menor_que_el_inicio_no_avanza():
    ud = {}
    iniciar(ud)
    avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    avanzar(ud, "5239", CTX)
    r = avanzar(ud, "5230", CTX)
    assert r["ok"] is False
    assert ud["horo_state"] == PASOS.TERMINO


def test_un_numero_con_letras_no_avanza():
    ud = {}
    iniciar(ud)
    avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    r = avanzar(ud, "como cinco mil", CTX)
    assert r["ok"] is False
    assert ud["horo_state"] == PASOS.INICIO


def test_el_flujo_completo_deja_los_campos_listos_para_guardar():
    ud = {}
    iniciar(ud)
    avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    avanzar(ud, "5239", CTX)
    avanzar(ud, "5242", CTX)
    r = avanzar(ud, "Sacar restos poda nogales", CTX)
    assert r["ok"] is True
    campos = r["campos"]
    assert campos["tipo"] == "MAQUINARIA"
    assert campos["maquina"] == "TRACTOR MASSEY FERGUSON 4292"
    assert campos["odometro"] == 5242
    assert campos["actividad"] == "Sacar restos poda nogales"
    assert ud.get("horo_state") in (None, "")


@pytest.mark.parametrize("clave", ["horo_state"])
def test_el_flujo_esta_registrado_para_caducar(clave):
    """Un flujo sin registrar no caduca nunca: es el bug que costo 12 dias."""
    from modules.flujos import CLAVES_ESTADO
    assert clave in CLAVES_ESTADO


def test_ningun_flujo_del_proyecto_queda_sin_registrar():
    """Guard contra volver a agregar un flujo y olvidar registrarlo."""
    import os
    import re

    from modules.flujos import CLAVES_ESTADO
    raiz = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "handlers")
    usadas = set()
    for nombre in os.listdir(raiz):
        if not nombre.endswith(".py"):
            continue
        with open(os.path.join(raiz, nombre), encoding="utf-8") as fh:
            usadas |= set(re.findall(r'user_data\.get\("(\w+_state)"\)', fh.read()))
    faltan = usadas - set(CLAVES_ESTADO)
    assert not faltan, "flujos sin registrar en modules/flujos.py: %s" % faltan
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_horometro_guiado.py -q`
Expected: FAIL con `ModuleNotFoundError: No module named 'handlers.horometro'`

- [ ] **Step 3: Registrar el flujo en `modules/flujos.py`**

En `CLAVES_ESTADO`, agregar después de `"venc_state"`:

```python
    "horo_state",          # flujo guiado de horómetro
```

Y en `CLAVES_DATOS`, después de `"venc_pendientes", "venc_idx",`:

```python
    "horo_data",
```

- [ ] **Step 4: Escribir la implementación**

```python
# handlers/horometro.py
# -*- coding: utf-8 -*-
"""Flujo guiado para anotar una lectura de horometro.

POR QUE GUIADO Y NO LEIDO POR IA
Son tres datos y la maquina sale de una lista cerrada, asi que no se puede
inventar. Y el numero de inicio se contrasta contra la ultima lectura guardada
EN EL MOMENTO, con Juan todavia frente a la maquina. Hoy un error de tipeo entra
callado y descuadra las horas de todas las lecturas siguientes de esa maquina.

Este modulo es PURO: no toca Telegram ni Excel. Recibe el user_data y el texto,
y devuelve que responder. Asi se prueba entero en memoria.
"""
import logging
from datetime import date

logger = logging.getLogger(__name__)


class PASOS:
    MAQUINA = "esperando_maquina"
    INICIO = "esperando_inicio"
    TERMINO = "esperando_termino"
    LABOR = "esperando_labor"


def iniciar(user_data) -> None:
    user_data["horo_state"] = PASOS.MAQUINA
    user_data["horo_data"] = {}


def _cerrar(user_data) -> None:
    user_data["horo_state"] = None
    user_data["horo_data"] = None


def _numero(texto):
    """El numero que escribio Juan, o None. Acepta '5.239' y '5239'."""
    limpio = str(texto or "").strip().replace(".", "").replace(",", ".")
    try:
        return float(limpio)
    except ValueError:
        return None


def _ultima_de(maquina, ctx):
    for m in (ctx.get("maquinas") or []):
        if str(m.get("maquina", "")).upper() == str(maquina).upper():
            return m.get("ultimo_odometro")
    return None


def revisar_inicio(inicio, ultima) -> str | None:
    """Devuelve el aviso si el inicio no calza con la ultima lectura, o None.

    Esta es la razon de ser del flujo: cazar el tipeo mientras Juan mira la
    maquina, no tres dias despues cuando ya nadie se acuerda.
    """
    if ultima is None:
        return None                            # nunca se leyo: nada con que comparar
    if float(inicio) == float(ultima):
        return None
    return ("⚠️ La última que tengo es *%s*, y me pusiste *%s*.\n"
            "¿Está bien?" % (f"{ultima:,.0f}".replace(",", "."),
                             f"{float(inicio):,.0f}".replace(",", ".")))


def avanzar(user_data, texto, ctx) -> dict:
    """Procesa un paso. Devuelve {"ok", "mensaje", "aviso", "campos"}.

    `campos` solo viene en el ultimo paso, listo para
    registrar_bitacora_estructurada.
    """
    paso = user_data.get("horo_state")
    datos = user_data.setdefault("horo_data", {})
    texto = str(texto or "").strip()

    if paso == PASOS.MAQUINA:
        conocidas = {str(m["maquina"]).upper(): m["maquina"]
                     for m in (ctx.get("maquinas") or [])}
        if texto.upper() not in conocidas:
            return {"ok": False, "mensaje": "No conozco esa máquina. Elige una "
                                            "de la lista.", "campos": None}
        datos["maquina"] = conocidas[texto.upper()]
        user_data["horo_state"] = PASOS.INICIO
        ultima = _ultima_de(datos["maquina"], ctx)
        datos["ultima"] = ultima
        pista = ("La última que tengo es *%s*.\n"
                 % f"{ultima:,.0f}".replace(",", ".")) if ultima is not None else ""
        return {"ok": True, "mensaje": pista + "¿En cuánto *partió* hoy?",
                "campos": None}

    if paso == PASOS.INICIO:
        n = _numero(texto)
        if n is None:
            return {"ok": False, "mensaje": "Necesito el número del horómetro.",
                    "campos": None}
        datos["inicio"] = n
        user_data["horo_state"] = PASOS.TERMINO
        return {"ok": True, "mensaje": "¿Y en cuánto *terminó*?",
                "aviso": revisar_inicio(n, datos.get("ultima")), "campos": None}

    if paso == PASOS.TERMINO:
        n = _numero(texto)
        if n is None:
            return {"ok": False, "mensaje": "Necesito el número del horómetro.",
                    "campos": None}
        if n < datos.get("inicio", 0):
            return {"ok": False,
                    "mensaje": "El término no puede ser menor que el inicio "
                               "(%g). ¿Cuál es?" % datos["inicio"],
                    "campos": None}
        datos["termino"] = n
        user_data["horo_state"] = PASOS.LABOR
        return {"ok": True, "mensaje": "¿Qué *labor*?", "campos": None}

    if paso == PASOS.LABOR:
        campos = {
            "fecha": date.today().strftime("%Y-%m-%d"),
            "tipo": "MAQUINARIA", "actividad": texto or "Lectura de horómetro",
            "cultivo": "GENERAL", "sector": "", "jornadas_hombre": None,
            "trabajadores": [], "insumo": "", "cantidad": None, "unidad": "",
            "maquina": datos.get("maquina", ""), "odometro": datos.get("termino"),
            "superficie_ha": None,
            "texto_original": "Horómetro guiado: %s %g → %g · %s"
                              % (datos.get("maquina", ""), datos.get("inicio", 0),
                                 datos.get("termino", 0), texto),
        }
        _cerrar(user_data)
        return {"ok": True, "mensaje": "", "campos": campos}

    return {"ok": False, "mensaje": "", "campos": None}
```

- [ ] **Step 5: Correr el test y comprobar que pasa**

Run: `py -m pytest tests/test_horometro_guiado.py -q`
Expected: PASS, 12 passed

- [ ] **Step 6: Correr la suite entera**

Run: `py -m pytest tests/ -q`
Expected: sin fallos. Ojo con `tests/test_flujos_vencen.py`, que recorre `CLAVES_ESTADO` con `parametrize`: va a sumar un caso por la clave nueva, y eso está bien.

- [ ] **Step 7: Commit**

```bash
git add handlers/horometro.py modules/flujos.py tests/test_horometro_guiado.py
git commit -m "Flujo guiado de horometro, con el numero contrastado en el momento

Cuatro pasos: maquina de una lista cerrada, inicio, termino y labor. Lo que
compra frente a leerlo con IA es el paso del inicio: se contrasta contra la
ultima lectura guardada mientras Juan esta parado frente a la maquina. Hoy un
error de tipeo entra callado y descuadra las horas de todas las lecturas
siguientes de esa maquina.

horo_state queda registrado en modules/flujos.py para que caduque, y hay un test
que falla si alguien vuelve a agregar un flujo sin registrarlo.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

## Task 4: Cablear y probar en vivo

**Files:**
- Modify: `handlers/chat.py`
- Modify: `main.py`
- Test: `tests/test_teclado_ruteo.py`

- [ ] **Step 1: Escribir el test que falla**

```python
# tests/test_teclado_ruteo.py
# -*- coding: utf-8 -*-
"""Los botones se atienden ANTES que cualquier otra cosa.

Si no, "Asistencia" a secas cae en es_mensaje_sin_contenido, que la trata como
basura --y con razon, porque Juan escribia la palabra sola antes del parte-- y
el boton no haria nada.
"""
import inspect

import handlers.chat as chat


def test_los_botones_se_atienden_antes_que_los_flujos():
    fuente = inspect.getsource(chat.handle_text)
    assert "es_boton" in fuente
    assert fuente.index("es_boton") < fuente.index("handle_text_deposito")


def test_el_flujo_de_horometro_se_atiende_en_el_dispatcher():
    assert "horo_state" in inspect.getsource(chat.handle_text)
```

- [ ] **Step 2: Correr el test y comprobar que falla**

Run: `py -m pytest tests/test_teclado_ruteo.py -q`
Expected: FAIL, 2 failed

- [ ] **Step 3: Cablear en `handlers/chat.py`**

Al principio de `handle_text`, **antes** de `revisar_flujos` y de todos los `handle_text_*`:

```python
    # ── Los botones fijos se atienden antes que nada ──
    # "Asistencia" a secas cae en es_mensaje_sin_contenido y se trataria como
    # basura --con razon: Juan escribia la palabra sola antes del parte-- asi
    # que el boton no haria nada si no se mira primero.
    from handlers.teclado import es_boton
    if es_boton(update.message.text):
        from handlers.horometro_h import atender_boton
        await atender_boton(update, context)
        return

    # ── Flujo guiado de horómetro en curso ──
    if context.user_data.get("horo_state"):
        from handlers.horometro_h import atender_paso
        if await atender_paso(update, context):
            return
```

- [ ] **Step 4: Escribir la capa de Telegram del flujo**

`handlers/horometro.py` es puro a propósito. La parte que habla con Telegram va aparte, en `handlers/horometro_h.py`:

```python
# -*- coding: utf-8 -*-
"""La capa de Telegram del flujo de horometro. La logica esta en horometro.py."""
import asyncio
import logging

from telegram import ReplyKeyboardMarkup

from handlers.horometro import PASOS, avanzar, iniciar
from handlers.teclado import (BOTON_HOROMETRO, teclado_capataz, texto_de_ayuda)

logger = logging.getLogger(__name__)


def _menu(opciones):
    """Teclado de una columna con las opciones, mas Otra..."""
    filas = [[o] for o in opciones] + [["Otra…"]]
    return ReplyKeyboardMarkup(filas, resize_keyboard=True, is_persistent=True)


async def atender_boton(update, context):
    texto = (update.message.text or "").strip()
    ayuda = texto_de_ayuda(texto)
    if ayuda:                                   # Asistencia y Factura solo instruyen
        await update.message.reply_text(ayuda, parse_mode="Markdown",
                                        reply_markup=teclado_capataz())
        return
    if texto != BOTON_HOROMETRO:
        return
    from modules.opciones_capataz import maquinas_recientes
    from modules.parte_contexto import construir
    ctx = await asyncio.to_thread(construir)
    context.user_data["horo_ctx"] = ctx
    iniciar(context.user_data)
    await update.message.reply_text(
        "🚜 ¿Qué *máquina*?", parse_mode="Markdown",
        reply_markup=_menu(maquinas_recientes(ctx)))


async def atender_paso(update, context) -> bool:
    ctx = context.user_data.get("horo_ctx") or {}
    r = avanzar(context.user_data, update.message.text, ctx)

    if not r["ok"]:
        await update.message.reply_text(r["mensaje"])
        return True

    if r.get("aviso"):
        await update.message.reply_text(r["aviso"], parse_mode="Markdown")

    if r["campos"] is None:
        teclado = teclado_capataz()
        if context.user_data.get("horo_state") == PASOS.LABOR:
            from modules.opciones_capataz import labores_frecuentes
            teclado = _menu(await asyncio.to_thread(labores_frecuentes))
        await update.message.reply_text(r["mensaje"], parse_mode="Markdown",
                                        reply_markup=teclado)
        return True

    from bitacora_manager import registrar_bitacora_estructurada
    quien = update.effective_user.full_name if update.effective_user else ""
    try:
        res = await asyncio.to_thread(registrar_bitacora_estructurada,
                                      r["campos"], quien)
    except Exception as e:
        logger.error("Horómetro guiado: no pude guardar: %s", e)
        await update.message.reply_text(
            "❌ No pude guardarla: %s" % str(e)[:120],
            reply_markup=teclado_capataz())
        return True

    if isinstance(res, dict) and res.get("error_odometro"):
        await update.message.reply_text(
            "🤔 No la guardé: %s" % res["error_odometro"],
            reply_markup=teclado_capataz())
        return True

    horas = res.get("horas_dia") if isinstance(res, dict) else None
    msg = "✅ *%s* — %g" % (r["campos"]["maquina"], r["campos"]["odometro"])
    if horas is not None:
        msg += "\n🕐 *%g h* desde la lectura anterior" % horas
    await update.message.reply_text(msg, parse_mode="Markdown",
                                    reply_markup=teclado_capataz())
    return True
```

- [ ] **Step 5: Poner el teclado en `/start`**

`cmd_start` está en `main.py:174` y hoy hace un solo `await update.message.reply_text("👋 ¡Hola! Soy el bot de *Agrícola Santa Elisa*.\n\n" ... )` con un texto largo de varias líneas concatenadas.

Justo **antes** de esa llamada, insertar:

```python
    # El teclado fijo solo para los capataces: el dueño usa comandos y no le
    # sirve de nada ocuparle el teclado con tres botones que no va a usar.
    from config import AUTO_SAVE_USERS
    from handlers.teclado import teclado_capataz
    _kb = (teclado_capataz()
           if update.effective_user and update.effective_user.id in AUTO_SAVE_USERS
           else None)
```

y agregarle a esa misma llamada, junto al `parse_mode` que ya tiene, el argumento:

```python
        reply_markup=_kb,
```

No hay que tocar el texto del mensaje.

- [ ] **Step 6: Correr todo**

Run: `py -m pytest tests/ -q`
Expected: sin fallos.

- [ ] **Step 7: Commit**

```bash
git add handlers/chat.py handlers/horometro_h.py main.py tests/test_teclado_ruteo.py
git commit -m "Cableado del teclado y del flujo de horometro

Los botones se miran antes que nada: Asistencia a secas cae en
es_mensaje_sin_contenido y se trataria como basura. La logica del flujo queda en
horometro.py, que es puro y se prueba en memoria; horometro_h.py es la capa que
habla con Telegram.

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

- [ ] **Step 8: Probar en vivo**

⚠️ **Antes de matar el bot, mirar el log** (`tail -40 bot.log`). Y ojo: `Stop-ScheduledTask` **no mata el python**, hay que matarlo a mano y comprobar que queden 0 (ver `feedback-reiniciar-bot-watchdog`).

Después, desde el teléfono de Juan: `/start` → tienen que aparecer los tres botones → tocar `🚜 Horómetro` → elegir un tractor → poner un inicio **equivocado a propósito** y comprobar que avisa con la última lectura → completar → comprobar la fila nueva en la hoja `Bitácora`.

---

## Al terminar

- [ ] Actualizar la spec y la cabecera de `project-pendientes-roadmap`
- [ ] Pushear
- [ ] Recién ahí seguir con el plan de la IA para la asistencia
