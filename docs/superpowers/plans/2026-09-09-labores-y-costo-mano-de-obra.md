# Pantalla de labores y costo de mano de obra — plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Una pantalla `/labores` que acumule el trabajo hecho —jornadas, días, personas— y lo cruce con la plata que realmente se pagó, para saber cuánto costó cada labor y cada trabajador.

**Architecture:** Un módulo puro `modules/labores.py` que lee el Master y devuelve agregados, sin tocar Flask ni escribir nada. Cuatro funciones, una por vista. El front es una ruta `/labores` + un template, siguiendo el patrón de `/vacaciones`.

**Tech Stack:** Python 3.11, openpyxl, Flask, pytest. Sin dependencias nuevas.

**Spec:** `docs/superpowers/specs/2026-09-09-labores-y-costo-mano-de-obra-design.md`

---

## Estructura de archivos

| archivo | responsabilidad |
|---|---|
| `modules/labores.py` (nuevo) | Todo el cálculo. Puro: lee el Master, no escribe, no importa Flask. |
| `tests/test_labores.py` (nuevo) | Los casos que rompen, contra libros de prueba en memoria. |
| `src/dashboard.py` (modificar) | Dos rutas: `/labores` y `/api/labores`. |
| `src/templates/labores.html` (nuevo) | Los cuatro bloques. Mismo estilo que `vacaciones.html`. |

`modules/labores.py` se divide en cuatro capas dentro del archivo: lectura de la
bitácora, lectura de los pagos, cálculo del costo por jornada, y las cuatro
funciones públicas. Si crece más de ~350 líneas, separar la lectura de pagos a
`modules/labores_pagos.py`.

**Correr los tests:** desde `Robot/`, con
`PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -v`

---

### Task 1: Leer la bitácora y agrupar las labores

**Files:**
- Create: `modules/labores.py`
- Test: `tests/test_labores.py`

- [ ] **Step 1: Write the failing test**

```python
# -*- coding: utf-8 -*-
"""El acumulado de labores sale de la Bitácora del Master.

OJO CON LOS FALSOS HERMANOS: "Poda nogales", "Pintar poda nogales" y "Sacar
restos poda nogales" comparten la palabra pero son TRES trabajos distintos.
Agrupar por "poda" a secas los colapsa y borra la información que importa.
"""
from openpyxl import Workbook

from modules.labores import resumen_labores

ENC = ["Fecha", "Hora", "Tipo", "Actividad", "Cultivo", "Sector",
       "Jornadas Hombre", "Trabajadores", "Insumo", "Cantidad", "Unidad",
       "Registro", "Registrado por", "Máquina", "Odómetro", "Horas Día",
       "Superficie ha", "Días Cubiertos"]


def _fila(fecha, actividad, jh, trabajadores, cultivo="NOGALES", sector=""):
    r = [None] * len(ENC)
    r[0], r[2], r[3] = fecha, "LABOR", actividad
    r[4], r[5], r[6], r[7] = cultivo, sector, jh, trabajadores
    return r


def _libro(tmp_path, filas):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(ENC)
    for f in filas:
        ws.append(f)
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_la_misma_labor_escrita_distinto_queda_en_un_grupo(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "Felicito Amigo, Ramiro Amigo"),
        _fila("2026-06-11", "poda  nogales", 1, "Felicito Amigo"),
        _fila("2026-06-12", "Poda de nogales", 1, "Felicito Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert len(r) == 1
    assert r[0]["jornadas"] == 4
    assert r[0]["dias"] == 3
    assert set(r[0]["etiquetas"]) == {"Poda nogales", "poda  nogales", "Poda de nogales"}


def test_poda_pintar_y_sacar_restos_son_tres_labores(tmp_path):
    """Comparten la palabra 'poda' y son tres trabajos distintos."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-06-10", "Pintar poda nogales", 1, "B"),
        _fila("2026-06-10", "Sacar restos poda nogales", 3, "C"),
    ])
    nombres = {x["labor"] for x in resumen_labores(path=ruta)}
    assert len(nombres) == 3


def test_cuenta_personas_distintas_no_apariciones(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Desaguar", 2, "Felicito Amigo, Ramiro Amigo"),
        _fila("2026-06-11", "Desaguar", 1, "Felicito Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert r[0]["personas"] == 2
    assert r[0]["jornadas"] == 3


def test_una_fila_sin_jornadas_suma_dia_pero_no_jornadas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Aseo general", None, "Felicito Amigo"),
        _fila("2026-06-11", "Aseo general", 2, "Felicito Amigo, Ramiro Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert r[0]["jornadas"] == 2
    assert r[0]["dias"] == 2


def test_la_lectura_de_horometro_no_es_una_labor(tmp_path):
    """La escribe el propio bot al guardar una lectura."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Lectura de horómetro", None, ""),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda nogales"}


def test_sin_hoja_bitacora_devuelve_vacio(tmp_path):
    wb = Workbook()
    ruta = tmp_path / "vacio.xlsx"
    wb.save(ruta)
    assert resumen_labores(path=str(ruta)) == []


def test_un_excel_ilegible_no_revienta(tmp_path):
    assert resumen_labores(path=str(tmp_path / "no_existe.xlsx")) == []


def test_filtra_por_rango_de_fechas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-08-10", "Poda nogales", 3, "A"),
    ])
    r = resumen_labores(desde="2026-08-01", path=ruta)
    assert r[0]["jornadas"] == 3
```

- [ ] **Step 2: Run test to verify it fails**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -v`
Expected: FAIL con `ModuleNotFoundError: No module named 'modules.labores'`

- [ ] **Step 3: Write minimal implementation**

```python
# -*- coding: utf-8 -*-
"""Cuanto trabajo se hizo y cuanto costo.

Modulo PURO: lee el Master, no escribe nada y no sabe de Flask. Asi se prueba
entero en memoria y el mismo dato sirve despues por Telegram.
"""
import datetime as dt
import logging
import re
import unicodedata

logger = logging.getLogger(__name__)

BITACORA_SHEET = "Bitácora"

# La escribe el propio bot al guardar una lectura: no es trabajo de nadie.
_NO_ES_LABOR = {"lectura de horometro"}

# Agrupacion por palabra clave. El ORDEN IMPORTA: gana la primera que calza, y
# las mas especificas van primero. "Sacar restos poda nogales" tiene que caer en
# su grupo ANTES de que "poda" se lo lleve.
GRUPOS = [
    ("sacar restos", "Sacar restos de poda"),
    ("pintar", "Pintar poda"),
    ("bajar ramas", "Bajar ramas"),
    ("poda", "Poda"),
    ("herbicida", "Aplicación herbicida"),
    ("fungicida", "Aplicación fungicida"),
    ("fertiliza", "Fertilización"),
    ("desagu", "Desaguar"),
    ("riego", "Mantención riego"),
    ("maquinaria", "Mantención maquinaria"),
    ("planta", "Mantención planta"),
    ("mantencion", "Mantención general"),
    ("aseo", "Aseo"),
    ("rastra", "Pasar rastra"),
    ("replante", "Replante"),
    ("cosecha", "Cosecha"),
]


def _clave(t) -> str:
    """Minusculas, sin tildes y con los espacios colapsados."""
    t = "".join(c for c in unicodedata.normalize("NFD", str(t or "").lower())
                if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", t).split())


def grupo_de(actividad) -> str:
    """A que labor pertenece una etiqueta. La etiqueta cruda si no calza."""
    k = _clave(actividad)
    for palabra, nombre in GRUPOS:
        if palabra in k:
            return nombre
    return str(actividad or "").strip()


def _fecha(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    for molde in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(str(v or "")[:10], molde).date()
        except ValueError:
            continue
    return None


def _leer_bitacora(path=None):
    """Filas utiles de la bitacora: (fecha, actividad, jh, personas, cultivo, sector)."""
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    filas = []
    try:
        wb = load_workbook(path or EXCEL_PATH, read_only=True, data_only=True)
        try:
            if BITACORA_SHEET not in wb.sheetnames:
                return []
            ws = wb[BITACORA_SHEET]
            cab = next(ws.iter_rows(min_row=1, max_row=1), None)
            enc = [c.value for c in cab] if cab else []
            idx = {n: i for i, n in enumerate(enc)}

            def col(r, n):
                return r[idx[n]] if n in idx and len(r) > idx[n] else None

            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or not any(r):
                    continue
                act = str(col(r, "Actividad") or "").strip()
                if not act or _clave(act) in _NO_ES_LABOR:
                    continue
                try:
                    jh = float(col(r, "Jornadas Hombre") or 0)
                except (TypeError, ValueError):
                    jh = 0.0
                personas = [p.strip() for p in
                            str(col(r, "Trabajadores") or "").split(",") if p.strip()]
                filas.append({
                    "fecha": _fecha(col(r, "Fecha")),
                    "actividad": act,
                    "jornadas": jh,
                    "personas": personas,
                    "cultivo": str(col(r, "Cultivo") or "").strip(),
                    "sector": str(col(r, "Sector") or "").strip(),
                })
        finally:
            wb.close()
    except Exception as e:
        logger.warning("labores: no pude leer la bitácora: %r", e)
        return []
    return filas


def _en_rango(f, desde, hasta) -> bool:
    if f is None:
        return False
    if desde and f < _fecha(desde):
        return False
    if hasta and f > _fecha(hasta):
        return False
    return True


def resumen_labores(desde=None, hasta=None, path=None) -> list:
    """Por labor agrupada: jornadas, dias, personas y las etiquetas que la componen."""
    acc = {}
    for f in _leer_bitacora(path):
        if not _en_rango(f["fecha"], desde, hasta):
            continue
        g = grupo_de(f["actividad"])
        d = acc.setdefault(g, {"labor": g, "jornadas": 0.0, "_dias": set(),
                               "_pers": set(), "_etq": set()})
        d["jornadas"] += f["jornadas"]
        d["_dias"].add(f["fecha"])
        d["_pers"].update(_clave(p) for p in f["personas"])
        d["_etq"].add(f["actividad"])
    salida = []
    for d in acc.values():
        salida.append({
            "labor": d["labor"],
            "jornadas": d["jornadas"],
            "dias": len(d["_dias"]),
            "personas": len(d["_pers"]),
            "desde": str(min(d["_dias"])) if d["_dias"] else "",
            "hasta": str(max(d["_dias"])) if d["_dias"] else "",
            "etiquetas": sorted(d["_etq"]),
        })
    return sorted(salida, key=lambda x: -x["jornadas"])
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -v`
Expected: PASS, 8 tests

- [ ] **Step 5: Medir contra el Master real**

Run:
```bash
PYTHONPATH="$PWD" python3.11 -c "
from modules.labores import resumen_labores
r = resumen_labores()
print('grupos:', len(r), '· jornadas:', sum(x['jornadas'] for x in r))
for x in r[:12]: print('  %-26s %6.0f jh  %3d días  %2d pers  %s' % (x['labor'], x['jornadas'], x['dias'], x['personas'], x['etiquetas'][:3]))
"
```
Expected: ~382 jornadas en total (el mismo número que las 54 etiquetas sin
agrupar) y menos de 54 grupos. **Mirar qué etiquetas cayeron juntas y anotar las
que estén mal** — se corrigen en `GRUPOS` en el paso siguiente, no ahora.

- [ ] **Step 6: Commit**

```bash
git add modules/labores.py tests/test_labores.py
git commit -m "Acumulado de labores desde la bitácora, agrupando etiquetas"
```

---

### Task 2: Leer lo que se pagó, cruzando por RUT

**Files:**
- Modify: `modules/labores.py`
- Test: `tests/test_labores.py`

- [ ] **Step 1: Write the failing test**

```python
from modules.labores import pagos_por_mes

PERSONAL = ["Nombre", "RUT", "Cargo", "Fecha Ingreso", "Días Pendientes",
            "Días Tomados Total", "Última Vacación", "Notas"]
BANCO = ["Fecha", "Descripcion", "Referencia", "Cargo", "Abono", "Saldo",
         "Tipo", "Categoria", "Cultivo", "Factura_linkeada"]


def _libro_pagos(tmp_path, personal, movs, filas_bit=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(ENC)
    for f in filas_bit:
        ws.append(f)
    wp = wb.create_sheet("Personal")
    wp.append(PERSONAL)
    for p in personal:
        wp.append(list(p) + [None] * (len(PERSONAL) - len(p)))
    wbco = wb.create_sheet("Cuenta Banco")
    wbco.append(BANCO)
    for m in movs:
        fila = [None] * len(BANCO)
        fila[0], fila[1], fila[3], fila[7] = m[0], m[1], m[2], m[3]
        wbco.append(fila)
    ruta = tmp_path / "pagos.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_cruza_el_pago_por_RUT_no_por_nombre(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-01", "TEF  9850887-2 FELICITO AMIGO", 885110, "MANO DE OBRA PLANTA")])
    p = pagos_por_mes(path=ruta)
    assert p["planta"]["9850887-2"]["2026-06"] == 885110


def test_dos_transferencias_el_mismo_mes_se_suman(tmp_path):
    """Juan cobró dos veces en septiembre: 1.356.322 + 1.080.000."""
    ruta = _libro_pagos(
        tmp_path,
        [("Juan Parada Castillo", "13.373.052-4")],
        [("2026-09-01", "TEF 13373052-4 Juan Parada Cas", 1356322, "MANO DE OBRA PLANTA"),
         ("2026-09-01", "TEF 13373052-4 Juan Parada Cas", 1080000, "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["planta"]["13373052-4"]["2026-09"] == 2436322


def test_lo_del_dueno_queda_fuera(tmp_path):
    """CRAVE SPA y las remuneraciones de Felix De Vicente no son costo de labor."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-01", "TEF 77912665-K CRAVE SPA", 2029455, "MANO DE OBRA PLANTA"),
         ("2026-06-01", "Remuneracion Mayo Felix De Vicente", 2024358, "MANO DE OBRA PLANTA"),
         ("2026-06-01", "TEF  9359341-3 FELIX DE VICENT", 905117, "MANO DE OBRA PLANTA")])
    p = pagos_por_mes(path=ruta)
    assert p["planta"] == {}
    assert p["previred"] == {}
    assert p["temporal"] == {}


def test_previred_se_guarda_aparte(tmp_path):
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-06-01", "PAGO COTIZ.PREVIRED", 2042431, "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["previred"]["2026-06"] == 2042431


def test_la_mano_de_obra_temporal_va_a_su_bolsa(tmp_path):
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-08-01", "PAGO CUADRILLA", 1500000, "MANO DE OBRA TEMPORAL")])
    assert pagos_por_mes(path=ruta)["temporal"]["2026-08"] == 1500000


def test_un_abono_no_es_un_pago(tmp_path):
    """Los abonos entran; solo los cargos son plata que salió."""
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-08-01", "DEVOLUCION", 0, "MANO DE OBRA TEMPORAL")])
    assert pagos_por_mes(path=ruta)["temporal"] == {}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -k pagos -v`
Expected: FAIL con `ImportError: cannot import name 'pagos_por_mes'`

- [ ] **Step 3: Write minimal implementation**

Agregar a `modules/labores.py`:

```python
PERSONAL_SHEET = "Personal"
BANCO_SHEET = "Cuenta Banco"
CAT_PLANTA = "MANO DE OBRA PLANTA"
CAT_TEMPORAL = "MANO DE OBRA TEMPORAL"

# Lo del dueño: su remuneración y su sociedad. Lo pidió fuera explícitamente.
FUERA = ("CRAVE SPA", "77912665", "FELIX DE VICENT", "9359341", "17407271")


def rut_key(v) -> str:
    """RUT normalizado '9850887-2', o '' si no hay uno."""
    s = re.sub(r"[^0-9kK]", "", str(v or "")).upper()
    if len(s) < 2:
        return ""
    return s[:-1].lstrip("0") + "-" + s[-1]


def _rut_en(texto) -> str:
    """El RUT que aparece en la glosa del banco, o ''."""
    m = re.search(r"(\d{7,8})\s*-\s*([0-9kK])", str(texto or ""))
    return rut_key(m.group(1) + m.group(2)) if m else ""


def pagos_por_mes(path=None) -> dict:
    """Lo que se pagó, por mes.

    {"planta": {rut: {"2026-06": monto}}, "previred": {mes: monto},
     "temporal": {mes: monto}}
    """
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    salida = {"planta": {}, "previred": {}, "temporal": {}}
    try:
        wb = load_workbook(path or EXCEL_PATH, read_only=True, data_only=True)
        try:
            ruts = set()
            if PERSONAL_SHEET in wb.sheetnames:
                for r in wb[PERSONAL_SHEET].iter_rows(min_row=2, values_only=True):
                    if r and r[0] and rut_key(r[1]):
                        ruts.add(rut_key(r[1]))
            if BANCO_SHEET not in wb.sheetnames:
                return salida
            for r in wb[BANCO_SHEET].iter_rows(min_row=2, values_only=True):
                if not r or not r[0]:
                    continue
                f = _fecha(r[0])
                if not f:
                    continue
                try:
                    cargo = float(r[3] or 0)
                except (TypeError, ValueError):
                    continue
                if cargo <= 0:                      # un abono no es un pago
                    continue
                desc = str(r[1] or "")
                cat = str(r[7] or "")
                if any(x.upper() in desc.upper() for x in FUERA):
                    continue
                mes = "%04d-%02d" % (f.year, f.month)
                if "PREVIRED" in desc.upper():
                    salida["previred"][mes] = salida["previred"].get(mes, 0) + cargo
                    continue
                rk = _rut_en(desc)
                if rk and rk in ruts:
                    salida["planta"].setdefault(rk, {})
                    salida["planta"][rk][mes] = salida["planta"][rk].get(mes, 0) + cargo
                elif cat == CAT_TEMPORAL:
                    salida["temporal"][mes] = salida["temporal"].get(mes, 0) + cargo
        finally:
            wb.close()
    except Exception as e:
        logger.warning("labores: no pude leer los pagos: %r", e)
    return salida
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -v`
Expected: PASS, 14 tests

- [ ] **Step 5: Medir contra el Master real**

Run:
```bash
PYTHONPATH="$PWD" python3.11 -c "
from modules.labores import pagos_por_mes
p = pagos_por_mes()
print('personas con pago:', len(p['planta']))
for rut, meses in p['planta'].items(): print(' ', rut, {m: int(v) for m,v in sorted(meses.items())})
print('previred:', {m: int(v) for m,v in sorted(p['previred'].items())})
print('temporal:', {m: int(v) for m,v in sorted(p['temporal'].items())})
"
```
Expected: los 6 de `Personal` con pagos mensuales, previred con montos por mes,
y **cero rastro de CRAVE SPA o Felix**. Si aparece alguien de más, revisar
`FUERA` antes de seguir.

- [ ] **Step 6: Commit**

```bash
git add modules/labores.py tests/test_labores.py
git commit -m "Leer lo que se pagó de mano de obra, cruzando por RUT"
```

---

### Task 3: Costo por jornada y costo por trabajador

**Files:**
- Modify: `modules/labores.py`
- Test: `tests/test_labores.py`

- [ ] **Step 1: Write the failing test**

```python
from modules.labores import costo_por_trabajador


def test_el_costo_por_jornada_sale_del_sueldo_dividido_las_jornadas(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-01", "TEF  9850887-2 FELICITO AMIGO", 900000, "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 1, "Felicito Amigo"),
         _fila("2026-06-11", "Poda nogales", 1, "Felicito Amigo"),
         _fila("2026-06-12", "Poda nogales", 1, "Felicito Amigo")])
    r = {x["persona"]: x for x in costo_por_trabajador(path=ruta)}
    f = r["Felicito Amigo Soto"]
    assert f["jornadas"] == 3
    assert f["pagado"] == 900000
    assert f["costo_jornada"] == 300000


def test_un_mes_sin_pagos_da_sin_datos_no_cero(tmp_path):
    """Un cero se lee como 'salió gratis' y es mentira."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [],
        [_fila("2026-06-10", "Poda nogales", 1, "Felicito Amigo")])
    f = costo_por_trabajador(path=ruta)[0]
    assert f["costo_jornada"] is None
    assert f["costo"] is None


def test_una_persona_con_sueldo_y_cero_jornadas_no_divide_por_cero(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-01", "TEF  9850887-2 FELICITO AMIGO", 900000, "MANO DE OBRA PLANTA")],
        [])
    f = costo_por_trabajador(path=ruta)[0]
    assert f["jornadas"] == 0
    assert f["costo_jornada"] is None


def test_previred_se_reparte_a_prorrata_y_suma_lo_pagado(tmp_path):
    """Lo imputado tiene que ser IGUAL a lo que salió del banco."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2"), ("Ramiro Amigo Soto", "11.768.374-5")],
        [("2026-06-01", "TEF  9850887-2 FELICITO AMIGO", 600000, "MANO DE OBRA PLANTA"),
         ("2026-06-01", "TEF 11768374-5 RAMIRO AMIGO", 400000, "MANO DE OBRA PLANTA"),
         ("2026-06-05", "PAGO COTIZ.PREVIRED", 300000, "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 3, "Felicito Amigo, Felicito Amigo, Felicito Amigo"),
         _fila("2026-06-11", "Poda nogales", 1, "Ramiro Amigo")])
    r = {x["persona"]: x for x in costo_por_trabajador(path=ruta)}
    # 4 jornadas en el mes: Felicito 3, Ramiro 1 → previred 225.000 / 75.000
    assert round(r["Felicito Amigo Soto"]["previred"]) == 225000
    assert round(r["Ramiro Amigo Soto"]["previred"]) == 75000
    assert round(sum(x["previred"] for x in r.values())) == 300000
    assert round(r["Felicito Amigo Soto"]["costo"]) == 825000


def test_el_dueno_no_aparece_entre_los_trabajadores(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-01", "TEF 77912665-K CRAVE SPA", 2029455, "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 1, "Felicito Amigo")])
    nombres = {x["persona"] for x in costo_por_trabajador(path=ruta)}
    assert not any("FELIX" in n.upper() or "CRAVE" in n.upper() for n in nombres)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -k trabajador -v`
Expected: FAIL con `ImportError: cannot import name 'costo_por_trabajador'`

- [ ] **Step 3: Write minimal implementation**

Agregar a `modules/labores.py`:

```python
def _personas_del_master(path=None) -> dict:
    """{clave del nombre: (rut, nombre legal)} desde la hoja Personal."""
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    salida = {}
    try:
        wb = load_workbook(path or EXCEL_PATH, read_only=True, data_only=True)
        try:
            if PERSONAL_SHEET not in wb.sheetnames:
                return {}
            for r in wb[PERSONAL_SHEET].iter_rows(min_row=2, values_only=True):
                if not r or not r[0]:
                    continue
                salida[_clave(r[0])] = (rut_key(r[1]), str(r[0]))
        finally:
            wb.close()
    except Exception as e:
        logger.warning("labores: no pude leer Personal: %r", e)
    return salida


def _persona_de(nombre, personas) -> tuple:
    """(rut, nombre legal) de un nombre de la bitácora, o (None, None).

    La bitácora usa el nombre canónico ("Felicito Amigo") y Personal el legal
    ("Felicito Amigo Soto"): se calza por prefijo de palabras, no por igualdad.
    """
    k = _clave(nombre)
    if k in personas:
        return personas[k]
    for kp, v in personas.items():
        if kp.startswith(k) or k.startswith(kp):
            return v
    return (None, None)


def costo_por_trabajador(desde=None, hasta=None, path=None) -> list:
    """Por persona: jornadas, pagado, previred imputado, costo y costo/jornada."""
    personas = _personas_del_master(path)
    pagos = pagos_por_mes(path)
    filas = [f for f in _leer_bitacora(path) if _en_rango(f["fecha"], desde, hasta)]

    # jornadas por (rut, mes) y por (rut, labor)
    jh_mes, jh_total, labores = {}, {}, {}
    for f in filas:
        if not f["personas"]:
            continue
        parte = f["jornadas"] / len(f["personas"]) if f["personas"] else 0
        mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
        for p in f["personas"]:
            rut, _ = _persona_de(p, personas)
            if not rut:
                continue
            jh_mes[(rut, mes)] = jh_mes.get((rut, mes), 0) + parte
            jh_total[rut] = jh_total.get(rut, 0) + parte
            labores.setdefault(rut, set()).add(grupo_de(f["actividad"]))

    # previred del mes, a prorrata de las jornadas de planta de ese mes
    jh_por_mes = {}
    for (rut, mes), v in jh_mes.items():
        jh_por_mes[mes] = jh_por_mes.get(mes, 0) + v

    salida = []
    for k, (rut, nombre) in personas.items():
        if not rut:
            continue
        meses = pagos["planta"].get(rut, {})
        pagado = sum(meses.values())
        previred = 0.0
        for mes, total_mes in pagos["previred"].items():
            mias, todas = jh_mes.get((rut, mes), 0), jh_por_mes.get(mes, 0)
            if todas:
                previred += total_mes * mias / todas
        jornadas = jh_total.get(rut, 0)
        costo = (pagado + previred) if meses else None
        salida.append({
            "persona": nombre,
            "rut": rut,
            "jornadas": jornadas,
            "pagado": pagado if meses else None,
            "previred": previred,
            "costo": costo,
            "costo_jornada": (costo / jornadas) if costo and jornadas else None,
            "labores": sorted(labores.get(rut, [])),
        })
    return sorted(salida, key=lambda x: -(x["costo"] or 0))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -v`
Expected: PASS, 19 tests

- [ ] **Step 5: Medir contra el Master real**

Run:
```bash
PYTHONPATH="$PWD" python3.11 -c "
from modules.labores import costo_por_trabajador
for x in costo_por_trabajador():
    print('%-30s %6.1f jh  pagado %-12s previred %-10s costo/jornada %s' % (
        x['persona'][:30], x['jornadas'],
        int(x['pagado']) if x['pagado'] else 'sin datos',
        int(x['previred']), int(x['costo_jornada']) if x['costo_jornada'] else 'sin datos'))
"
```
Expected: los 6 de `Personal`, ninguno con costo 0 disfrazado, y ningún Felix.
**Si un costo por jornada sale absurdo (bajo $5.000 o sobre $200.000), parar y
entender por qué antes de seguir.**

- [ ] **Step 6: Commit**

```bash
git add modules/labores.py tests/test_labores.py
git commit -m "Costo por trabajador: sueldo y previred repartidos por jornada"
```

---

### Task 4: Costo por labor, por cultivo y evolución mensual

**Files:**
- Modify: `modules/labores.py`
- Test: `tests/test_labores.py`

- [ ] **Step 1: Write the failing test**

```python
from modules.labores import evolucion_mensual, por_cultivo_sector


def test_el_costo_de_una_labor_suma_planta_y_temporada(tmp_path):
    """Una labor con gente de planta y cuadrilla suma los dos costos."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-08-01", "TEF  9850887-2 FELICITO AMIGO", 900000, "MANO DE OBRA PLANTA"),
         ("2026-08-02", "PAGO CUADRILLA", 600000, "MANO DE OBRA TEMPORAL")],
        [_fila("2026-08-10", "Sacar restos poda nogales", 1, "Felicito Amigo"),
         _fila("2026-08-10", "Sacar restos poda nogales", 3, "Pedro, Ana, Luis")])
    r = {x["labor"]: x for x in resumen_labores(path=ruta)}["Sacar restos de poda"]
    assert r["jornadas"] == 4
    # planta: 900.000 / 1 jornada = 900.000 · temporada: 600.000 / 3 = 200.000 c/u
    assert round(r["costo"]) == 1500000


def test_por_cultivo_abre_las_jornadas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A", cultivo="NOGALES", sector="1"),
        _fila("2026-06-10", "Poda avellanos", 3, "B", cultivo="AVELLANOS", sector="2"),
    ])
    r = {(x["cultivo"], x["sector"]): x for x in por_cultivo_sector(path=ruta)}
    assert r[("NOGALES", "1")]["jornadas"] == 2
    assert r[("AVELLANOS", "2")]["jornadas"] == 3


def test_la_evolucion_va_por_mes_y_labor(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-07-10", "Poda nogales", 5, "A"),
    ])
    r = evolucion_mensual(path=ruta)
    meses = {x["mes"]: x for x in r}
    assert meses["2026-06"]["jornadas"] == 2
    assert meses["2026-07"]["jornadas"] == 5


def test_sin_pagos_el_costo_de_la_labor_es_sin_datos(tmp_path):
    ruta = _libro(tmp_path, [_fila("2026-06-10", "Poda nogales", 2, "A")])
    assert resumen_labores(path=ruta)[0]["costo"] is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -k "cultivo or evolucion or costo_de_una_labor or sin_pagos" -v`
Expected: FAIL — `ImportError` y `KeyError: 'costo'` en `resumen_labores`

- [ ] **Step 3: Write minimal implementation**

Agregar a `modules/labores.py` un `_costo_de_fila` y usarlo en `resumen_labores`,
`por_cultivo_sector` y `evolucion_mensual`:

```python
def _tabla_costos(path=None, filas=None) -> dict:
    """Costo por jornada de cada persona y mes, y de la cuadrilla por mes.

    Devuelve {"planta": {(rut, mes): $/jornada}, "temporal": {mes: $/jornada}}.
    Un mes sin pagos NO entra: quien lo consulte recibe None y muestra
    "sin datos", nunca cero.
    """
    personas = _personas_del_master(path)
    pagos = pagos_por_mes(path)
    filas = _leer_bitacora(path) if filas is None else filas

    jh_planta, jh_temporal, jh_mes = {}, {}, {}
    for f in filas:
        if not f["fecha"] or not f["personas"]:
            continue
        mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
        parte = f["jornadas"] / len(f["personas"])
        for p in f["personas"]:
            rut, _ = _persona_de(p, personas)
            if rut:
                jh_planta[(rut, mes)] = jh_planta.get((rut, mes), 0) + parte
                jh_mes[mes] = jh_mes.get(mes, 0) + parte
            else:
                jh_temporal[mes] = jh_temporal.get(mes, 0) + parte

    planta = {}
    for (rut, mes), jh in jh_planta.items():
        base = pagos["planta"].get(rut, {}).get(mes)
        if not base or not jh:
            continue
        prev = pagos["previred"].get(mes, 0)
        cuota = prev * jh / jh_mes[mes] if jh_mes.get(mes) else 0
        planta[(rut, mes)] = (base + cuota) / jh

    temporal = {}
    for mes, jh in jh_temporal.items():
        total = pagos["temporal"].get(mes)
        if total and jh:
            temporal[mes] = total / jh
    return {"planta": planta, "temporal": temporal}


def _costo_fila(f, personas, tabla):
    """Costo de una fila de bitácora, o None si falta el dato de algún mes."""
    if not f["fecha"] or not f["personas"]:
        return None
    mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
    parte = f["jornadas"] / len(f["personas"])
    total, visto = 0.0, False
    for p in f["personas"]:
        rut, _ = _persona_de(p, personas)
        cj = tabla["planta"].get((rut, mes)) if rut else tabla["temporal"].get(mes)
        if cj is None:
            continue
        total += cj * parte
        visto = True
    return total if visto else None
```

En `resumen_labores`, `por_cultivo_sector` y `evolucion_mensual`, acumular el
costo con `_costo_fila` y dejarlo en `None` si **ninguna** fila del grupo aportó
costo. Firmas:

```python
def por_cultivo_sector(desde=None, hasta=None, path=None) -> list:
    """[{cultivo, sector, jornadas, dias, costo}], de más a menos jornadas."""


def evolucion_mensual(desde=None, hasta=None, path=None) -> list:
    """[{mes, jornadas, costo, por_labor: {labor: jornadas}}], por mes ascendente."""
```

- [ ] **Step 4: Run test to verify it passes**

Run: `PYTHONPATH="$PWD" python3.11 -m pytest tests/test_labores.py -v`
Expected: PASS, 23 tests

- [ ] **Step 5: Medir contra el Master real**

Run:
```bash
PYTHONPATH="$PWD" python3.11 -c "
from modules.labores import resumen_labores, evolucion_mensual
r = resumen_labores()
print('COSTO POR LABOR')
for x in r[:10]: print('  %-26s %6.0f jh  %s' % (x['labor'], x['jornadas'], ('\$%s' % format(int(x['costo']),',d').replace(',','.')) if x['costo'] else 'sin datos'))
print('TOTAL imputado: ', sum(x['costo'] or 0 for x in r))
print('POR MES'); [print('  ', m['mes'], int(m['jornadas']), int(m['costo'] or 0)) for m in evolucion_mensual()]
"
```
Expected: el total imputado tiene que ser **cercano** a la suma de mano de obra
del período. Si difiere mucho, hay meses sin pagos o gente sin calzar — mirarlo
antes de seguir.

- [ ] **Step 6: Commit**

```bash
git add modules/labores.py tests/test_labores.py
git commit -m "Costo por labor, por cultivo y evolución mensual"
```

---

### Task 5: La pantalla

**Files:**
- Modify: `src/dashboard.py` (después de `api_vacaciones`, ~línea 440)
- Create: `src/templates/labores.html`

- [ ] **Step 1: Agregar las rutas**

En `src/dashboard.py`, después de `api_vacaciones`:

```python
@app.route("/labores")
def labores_page():
    return render_template("labores.html")


@app.route("/api/labores")
def api_labores():
    from modules.labores import (costo_por_trabajador, evolucion_mensual,
                                 por_cultivo_sector, resumen_labores)
    desde = request.args.get("desde") or None
    hasta = request.args.get("hasta") or None
    return jsonify({
        "labores": resumen_labores(desde, hasta),
        "trabajadores": costo_por_trabajador(desde, hasta),
        "cultivos": por_cultivo_sector(desde, hasta),
        "meses": evolucion_mensual(desde, hasta),
    })
```

- [ ] **Step 2: Escribir el template**

`src/templates/labores.html`, copiando la estructura y el CSS de
`vacaciones.html` (cabecera oscura, `.container`, `.card`, tabla con `th`
mayúsculas y `td.num` alineado a la derecha). Cuatro tarjetas:

1. **Acumulado por labor** — labor, jornadas, días, personas, período, costo.
   Cada fila despliega las **etiquetas** que la componen: es lo que permite
   corregir la agrupación.
2. **Costo por trabajador** — persona, jornadas, pagado, previred, costo,
   costo por jornada, labores en las que trabajó.
3. **Por cultivo y sector** — cultivo, sector, jornadas, costo.
4. **Evolución mensual** — mes, jornadas, costo, con barras simples en CSS.

Reglas de pintado, que no son cosméticas:
- `costo === null` se pinta **"sin datos"** en gris, nunca `$0`.
- Debajo de la tarjeta de labores, una nota fija: *"El costo de la cuadrilla de
  temporada es un prorrateo del total del mes, no un valor por persona."*
- Debajo de todo, la nota del alcance: *"La bitácora arranca en junio de 2026."*

- [ ] **Step 3: Levantar y mirar**

Run: `preview_start` con la configuración `dashboard`, entrar a `/labores`.
Expected: las cuatro tarjetas con datos, ningún `$0` donde debería decir
"sin datos", ningún `NaN`, ningún `undefined`.

- [ ] **Step 4: Commit**

```bash
git add src/dashboard.py src/templates/labores.html
git commit -m "Pantalla /labores: trabajo hecho y lo que costó"
```

---

### Task 6: Enlace desde el dashboard y cierre

**Files:**
- Modify: `src/templates/dashboard.html` (barra de navegación de la cabecera)

- [ ] **Step 1: Agregar el enlace**

Buscar en `dashboard.html` la cabecera donde están los enlaces a
`/cash-flow`, `/vacaciones` y `/conciliacion`, y agregar `/labores` con la misma
forma.

- [ ] **Step 2: Correr la suite completa**

Run: `python3.11 -m pytest -q`
Expected: 817 + 23 = **840 passed**

- [ ] **Step 3: Commit y push**

```bash
git add -A
git commit -m "Enlace a /labores desde el dashboard"
git push origin main
```

- [ ] **Step 4: Mostrarle los números al dueño**

Presentar el acumulado por labor y el costo por trabajador, **y la lista de
etiquetas que cayeron en cada grupo**, para que corrija la agrupación. Es el
punto que él pidió explícitamente y el que más probable esté mal en la primera
pasada.
