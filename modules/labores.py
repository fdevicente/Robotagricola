# -*- coding: utf-8 -*-
"""Cuanto trabajo se hizo y cuanto costo.

Modulo PURO: lee el Master, no escribe nada y no sabe de Flask. Asi se prueba
entero en memoria y el mismo dato sirve despues por Telegram.

DE DONDE SALE EL COSTO
No hay ningun valor de jornada en el Master --la hoja Personal no tiene
sueldos-- y pedirle al dueño que cargue una tarifa envejece en silencio. Lo que
si hay es plata que salio: las transferencias a la gente de planta llevan el RUT
en la glosa ("TEF 9850887-2 FELICITO AMIGO") y esos RUT estan en Personal. El
cruce es por RUT, no por nombre.
"""
import datetime as dt
import logging
import re
import unicodedata

logger = logging.getLogger(__name__)

BITACORA_SHEET = "Bitácora"

# La escribe el propio bot al guardar una lectura: no es trabajo de nadie.
_NO_ES_LABOR = {"lectura de horometro"}

# NO ES TRABAJO, aunque ocupe una fila de la bitacora. Medido contra el Master:
# por Tipo hay LABOR 313 jornadas, RIEGO 41, APLICACION 24, MAQUINARIA 4 y
# OTRO 0. Sin este filtro, "Vacaciones" y "Ausente" salian en la lista de
# labores como si alguien hubiera trabajado.
_TIPO_NO_TRABAJO = {"OTRO"}
_AUSENCIAS = {"vacaciones", "ausente", "aucente", "sin uso", "licencia"}

# Agrupacion por palabra clave. EL ORDEN IMPORTA: gana la primera que calza y
# las mas especificas van primero. "Sacar restos poda nogales" tiene que caer en
# su grupo ANTES de que "poda" se lo lleve; son trabajos distintos y medido
# contra el Master, sacar los restos costo mas del doble de jornadas que podar.
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


def _leer_bitacora(path=None) -> list:
    """Filas utiles de la bitacora, ya normalizadas."""
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
                if _clave(act) in _AUSENCIAS:
                    continue
                tipo = str(col(r, "Tipo") or "").strip().upper()
                if tipo in _TIPO_NO_TRABAJO:
                    continue
                try:
                    jh = float(col(r, "Jornadas Hombre") or 0)
                except (TypeError, ValueError):
                    jh = 0.0
                # Una fila de MAQUINARIA sin jornadas es el bot anotando la
                # maquina --combustible, horometro roto--, no trabajo de nadie.
                # Con jornadas si lo es: "Destroncar nogales" entra por ahi.
                if tipo == "MAQUINARIA" and jh <= 0:
                    continue
                personas = [p.strip() for p in
                            str(col(r, "Trabajadores") or "").split(",")
                            if p.strip()]
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
    """Por labor agrupada: jornadas, dias, personas y las etiquetas que la componen.

    Las etiquetas van en la salida a proposito: sin verlas, la agrupacion es una
    caja negra y nadie puede corregirla.
    """
    filas, personas, tabla = _con_costo(desde, hasta, path)
    acc = {}
    for f in filas:
        g = grupo_de(f["actividad"])
        d = acc.setdefault(g, {"labor": g, "jornadas": 0.0, "costo": None,
                               "_dias": set(), "_pers": set(), "_etq": set()})
        d["jornadas"] += f["jornadas"]
        d["_dias"].add(f["fecha"])
        d["_pers"].update(_clave(p) for p in f["personas"])
        d["_etq"].add(f["actividad"])
        c = _costo_fila(f, personas, tabla)
        if c is not None:
            d["costo"] = (d["costo"] or 0) + c

    salida = []
    for d in acc.values():
        salida.append({
            "labor": d["labor"],
            "jornadas": d["jornadas"],
            "costo": d["costo"],
            "dias": len(d["_dias"]),
            "personas": len(d["_pers"]),
            "desde": str(min(d["_dias"])) if d["_dias"] else "",
            "hasta": str(max(d["_dias"])) if d["_dias"] else "",
            "etiquetas": sorted(d["_etq"]),
        })
    return sorted(salida, key=lambda x: -x["jornadas"])


# ── Lo que se pagó ─────────────────────────────────────────────────────────

PERSONAL_SHEET = "Personal"
BANCO_SHEET = "Cuenta Banco"
CAT_TEMPORAL = "MANO DE OBRA TEMPORAL"

# Lo del dueño: su remuneracion y su sociedad. Lo pidio fuera explicitamente el
# 9-sep-2026: "saber cuanto me cuesta cada trabajador (sin considerar el mio)".
FUERA = ("CRAVE SPA", "77912665", "FELIX DE VICENT", "9359341", "17407271")

# Bajo esto, dividir el sueldo por las jornadas no significa nada. Juan es jefe
# de campo: reporta los partes y casi no se anota a si mismo. Medido, tiene 1
# jornada y $6.388.889 en tres meses, o sea $6.412.067 la jornada. Su costo se
# muestra igual --es plata real--, pero el costo POR JORNADA queda sin datos.
MIN_JORNADAS = 5


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


# Un pago hecho en los primeros dias del mes es el sueldo del mes ANTERIOR.
_DIA_CORTE = 5


def _mes_devengado(f) -> str:
    """El mes que se TRABAJO, no el dia que se pago.

    Medido contra el Master: el 1-jul se pago el sueldo de junio, el 31-jul el
    de julio y el 1-sep el de agosto. Atribuyendo por dia de pago, agosto
    quedaba sin sueldos y julio con el doble, y el costo por jornada salia
    disparatado en los dos meses.
    """
    if f.day <= _DIA_CORTE:
        anterior = f.replace(day=1) - dt.timedelta(days=1)
        return "%04d-%02d" % (anterior.year, anterior.month)
    return "%04d-%02d" % (f.year, f.month)


def pagos_por_mes(path=None) -> dict:
    """Lo que se pago de mano de obra, por mes.

    {"planta": {rut: {"2026-06": monto}}, "previred": {mes: monto},
     "temporal": {mes: monto}}

    Previred va aparte porque se paga en un monto unico por todos: se reparte
    despues, a prorrata de las jornadas del mes.
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
                if cargo <= 0:                  # un abono no es un pago
                    continue
                desc = str(r[1] or "")
                cat = str(r[7] or "")
                if any(x.upper() in desc.upper() for x in FUERA):
                    continue
                mes = _mes_devengado(f)
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


# ── Costo por trabajador ───────────────────────────────────────────────────


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
    """(rut, nombre legal) de un nombre de la bitacora, o (None, None).

    La bitacora usa el nombre canonico ("Ramiro Amigo") y Personal el legal
    ("Luis Ramiro Amigo Soto"). Se calza por SUBCONJUNTO DE PALABRAS, no por
    prefijo: medido contra el Master, por prefijo "Ramiro Amigo" no calzaba con
    "Luis Ramiro Amigo Soto" ni "Patricio Mora" con "Luis Patricio Mora Amigo",
    y tres de los seis quedaban con cero jornadas.

    Se exigen DOS palabras en comun y un unico candidato. Con una sola palabra,
    "Amigo" calzaria con tres personas distintas y "Juan" con Juan Parada,
    cuando en la cuadrilla hay cuatro Juanes y ninguno es el.
    """
    k = _clave(nombre)
    if not k:
        return (None, None)
    if k in personas:
        return personas[k]
    palabras = set(k.split())
    if len(palabras) < 2:
        return (None, None)
    candidatos = [v for kp, v in personas.items()
                  if palabras <= set(kp.split())]
    return candidatos[0] if len(candidatos) == 1 else (None, None)


def _jornadas_por_persona(filas, personas) -> tuple:
    """(jh por (rut, mes), jh total por rut, labores por rut, jh de planta por mes).

    Una fila de 2 jornadas con dos personas es UNA jornada de cada una.
    """
    jh_mes, jh_total, labores, jh_planta_mes = {}, {}, {}, {}
    for f in filas:
        if not f["fecha"] or not f["personas"]:
            continue
        mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
        parte = f["jornadas"] / len(f["personas"])
        for p in f["personas"]:
            rut, _ = _persona_de(p, personas)
            if not rut:
                continue
            jh_mes[(rut, mes)] = jh_mes.get((rut, mes), 0) + parte
            jh_total[rut] = jh_total.get(rut, 0) + parte
            jh_planta_mes[mes] = jh_planta_mes.get(mes, 0) + parte
            labores.setdefault(rut, set()).add(grupo_de(f["actividad"]))
    return jh_mes, jh_total, labores, jh_planta_mes


def costo_por_trabajador(desde=None, hasta=None, path=None) -> list:
    """Por persona: jornadas, pagado, previred imputado, costo y costo/jornada.

    Sin pagos cargados el costo queda en None y la pantalla dice "sin datos".
    Nunca cero: un cero se lee como "salio gratis".
    """
    personas = _personas_del_master(path)
    pagos = pagos_por_mes(path)
    filas = [f for f in _leer_bitacora(path) if _en_rango(f["fecha"], desde, hasta)]
    jh_mes, jh_total, labores, jh_planta_mes = _jornadas_por_persona(filas, personas)

    salida = []
    for rut, nombre in personas.values():
        if not rut:
            continue
        meses = pagos["planta"].get(rut, {})
        # Solo la plata de los meses que caen en el rango pedido.
        meses = {m: v for m, v in meses.items() if not filas or m in jh_planta_mes}
        pagado = sum(meses.values())
        previred = 0.0
        for mes, total_mes in pagos["previred"].items():
            mias, todas = jh_mes.get((rut, mes), 0), jh_planta_mes.get(mes, 0)
            if todas and mias:
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
            "costo_jornada": ((costo / jornadas)
                              if costo and jornadas >= MIN_JORNADAS else None),
            "labores": sorted(labores.get(rut, [])),
        })
    return sorted(salida, key=lambda x: -(x["costo"] or 0))


# ── Costo de cada fila, para repartirlo por labor, cultivo y mes ───────────


def _tabla_costos(personas, pagos, filas) -> dict:
    """Costo por jornada de cada persona y mes, y de la cuadrilla por mes.

    {"planta": {(rut, mes): $/jornada}, "temporal": {mes: $/jornada}}

    Un mes sin pagos NO entra: quien consulte recibe None y muestra "sin datos",
    nunca cero. Un cero se lee como "salio gratis" y es mentira.
    """
    jh_mes, _, _, jh_planta_mes = _jornadas_por_persona(filas, personas)
    jh_cuadrilla = {}
    for f in filas:
        if not f["fecha"] or not f["personas"]:
            continue
        mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
        parte = f["jornadas"] / len(f["personas"])
        for p in f["personas"]:
            if not _persona_de(p, personas)[0]:
                jh_cuadrilla[mes] = jh_cuadrilla.get(mes, 0) + parte

    planta = {}
    for (rut, mes), jh in jh_mes.items():
        base = pagos["planta"].get(rut, {}).get(mes)
        if not base or not jh:
            continue
        prev = pagos["previred"].get(mes, 0)
        cuota = prev * jh / jh_planta_mes[mes] if jh_planta_mes.get(mes) else 0
        planta[(rut, mes)] = (base + cuota) / jh

    temporal = {}
    for mes, jh in jh_cuadrilla.items():
        total = pagos["temporal"].get(mes)
        if total and jh:
            temporal[mes] = total / jh
    return {"planta": planta, "temporal": temporal}


def _costo_fila(f, personas, tabla):
    """Costo de una fila de bitacora, o None si no hay con que calcularlo."""
    if not f["fecha"] or not f["personas"]:
        return None
    mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
    parte = f["jornadas"] / len(f["personas"])
    total, visto = 0.0, False
    for p in f["personas"]:
        rut = _persona_de(p, personas)[0]
        cj = tabla["planta"].get((rut, mes)) if rut else tabla["temporal"].get(mes)
        if cj is None:
            continue
        total += cj * parte
        visto = True
    return total if visto else None


def _con_costo(desde, hasta, path):
    """(filas del rango, personas, tabla de costos). Base de las tres vistas."""
    personas = _personas_del_master(path)
    pagos = pagos_por_mes(path)
    filas = [f for f in _leer_bitacora(path) if _en_rango(f["fecha"], desde, hasta)]
    return filas, personas, _tabla_costos(personas, pagos, filas)


def _sumar(acc, clave, plantilla, f, costo):
    d = acc.setdefault(clave, dict(plantilla, jornadas=0.0, costo=None,
                                   _dias=set()))
    d["jornadas"] += f["jornadas"]
    d["_dias"].add(f["fecha"])
    if costo is not None:
        d["costo"] = (d["costo"] or 0) + costo
    return d


def por_cultivo_sector(desde=None, hasta=None, path=None) -> list:
    """Jornadas, dias y costo abiertos por cultivo y sector."""
    filas, personas, tabla = _con_costo(desde, hasta, path)
    acc = {}
    for f in filas:
        k = (f["cultivo"] or "SIN CULTIVO", f["sector"])
        _sumar(acc, k, {"cultivo": k[0], "sector": k[1]}, f,
               _costo_fila(f, personas, tabla))
    salida = [dict(d, dias=len(d.pop("_dias"))) for d in acc.values()]
    return sorted(salida, key=lambda x: -x["jornadas"])


def evolucion_mensual(desde=None, hasta=None, path=None) -> list:
    """Jornadas y costo por mes, abiertos por labor."""
    filas, personas, tabla = _con_costo(desde, hasta, path)
    acc = {}
    for f in filas:
        if not f["fecha"]:
            continue
        mes = "%04d-%02d" % (f["fecha"].year, f["fecha"].month)
        d = _sumar(acc, mes, {"mes": mes, "por_labor": {}}, f,
                   _costo_fila(f, personas, tabla))
        g = grupo_de(f["actividad"])
        d["por_labor"][g] = d["por_labor"].get(g, 0) + f["jornadas"]
    salida = [dict(d, dias=len(d.pop("_dias"))) for d in acc.values()]
    return sorted(salida, key=lambda x: x["mes"])
