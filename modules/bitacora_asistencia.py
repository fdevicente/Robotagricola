"""modules/bitacora_asistencia.py — Desglosa los partes de asistencia de Juan.

Juan reporta a diario en formato:
    Asistencia 23 julio 2026
    Felicito amigo : mantencion riego
    Ramiro amigo : mantencion riego
    Patricio Mora : mantencion maquinaria

Esto genera UNA fila por ACTIVIDAD (no una sola fila fusionada), que es lo que
permite medir jornadas-hombre por labor.
"""
import re
import unicodedata

from modules.bitacora_extractor import ALIAS, TRABAJADORES_CONOCIDOS

# Nombres de pila reconocidos → nombre canónico
_PILA = {}
for _alias, _canon in ALIAS.items():
    _PILA[_alias] = _canon
for _n in TRABAJADORES_CONOCIDOS:
    # setdefault, NO asignación: dos personas comparten el nombre de pila
    # ("Richard Padilla" y su hijo "Richard Padilla Crespo") y el último de la
    # lista pisaba el alias curado. Para los nombres de pila ambiguos manda
    # ALIAS, que apunta al padre.
    _PILA.setdefault(_n.split()[0].lower(), _n)
_PILA.setdefault("javier", "Javier Gonzalez")
_PILA.setdefault("gonzales", "Javier Gonzalez")
_PILA.setdefault("gonzalez", "Javier Gonzalez")

_LINEA = re.compile(r"^\s*([^:\n]{3,40}?)\s*:\s*(.+?)\s*$")


def _sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn").lower()


# Tokens de cada nombre canónico, para poder elegir el MÁS específico.
# Hace falta porque padre e hijo se llaman igual: "Richard Padilla" y
# "Richard Padilla Crespo". Quedarse con el primer nombre de pila que calzara
# borraba al hijo y subcontaba su jornada.
_TOKENS = {n: set(_sin_tildes(n).split()) for n in TRABAJADORES_CONOCIDOS}
_TOKENS.setdefault("Javier Gonzalez", {"javier", "gonzalez"})


def _canonico(nombre_crudo: str):
    """Devuelve el nombre canónico si la línea empieza con un trabajador.

    Gana el nombre MÁS específico: si el texto trae todos los tokens de
    "Richard Padilla" y además los de "Richard Padilla Crespo", es el hijo.
    """
    txt = _sin_tildes(nombre_crudo)
    palabras = [p for p in re.split(r"[^a-z]+", txt) if p]
    if not palabras:
        return None
    sueltas = set(palabras)

    # 1) Nombre completo: de los que estén contenidos enteros, el más largo.
    completos = [n for n, toks in _TOKENS.items() if toks and toks <= sueltas]
    if completos:
        return max(completos, key=lambda n: len(_TOKENS[n]))

    # 2) Por alias/nombre de pila. Si distintas palabras apuntan a personas
    #    distintas ("Richard crespo"), gana igual la más específica.
    candidatos = [_PILA[p] for p in palabras if p in _PILA]
    if candidatos:
        return max(candidatos, key=lambda n: len(_TOKENS.get(n, {n})))
    return None


def _norm_actividad(act: str) -> str:
    a = " ".join(act.split()).strip(" .,;")
    a = re.sub(r"^mantencion\b", "Mantención", a, flags=re.IGNORECASE)
    a = re.sub(r"^barvecho\b", "Barbecho", a, flags=re.IGNORECASE)
    if not a:
        return a
    return a[0].upper() + a[1:]


def _clave(act: str) -> str:
    """Clave para agrupar actividades equivalentes (sin tildes ni mayúsculas)."""
    return _sin_tildes(_norm_actividad(act)).strip()


def parsear_asistencia(texto: str):
    """Devuelve [{actividad, trabajadores, jornadas_hombre}] o None si no aplica.

    Solo aplica cuando hay 2+ líneas 'Trabajador : actividad'.
    """
    grupos = {}   # clave → {"actividad": str, "trabajadores": [..]}
    for linea in (texto or "").splitlines():
        m = _LINEA.match(linea)
        if not m:
            continue
        crudo, actividad = m.group(1), m.group(2)
        trabajador = _canonico(crudo)
        if not trabajador:
            continue
        act = _norm_actividad(actividad)
        if not act:
            continue
        g = grupos.setdefault(_clave(act), {"actividad": act, "trabajadores": []})
        if trabajador not in g["trabajadores"]:
            g["trabajadores"].append(trabajador)

    total = sum(len(g["trabajadores"]) for g in grupos.values())
    if total < 2:
        return None

    salida = []
    for g in grupos.values():
        # Ausencias: no suman jornadas-hombre. Se toleran las faltas de
        # ortografía frecuentes de Juan ("aucente", "vacasiones", "licensia").
        no_trabajada = bool(re.search(
            r"au[cs]+ente|vaca[cs]ion|licen[cs]ia|permiso|inasist|falt|no vino",
            _sin_tildes(g["actividad"])))
        salida.append({
            "actividad": g["actividad"],
            "trabajadores": sorted(g["trabajadores"]),
            "jornadas_hombre": None if no_trabajada else len(g["trabajadores"]),
        })
    salida.sort(key=lambda x: -(x["jornadas_hombre"] or 0))
    return salida


# ── Varios días en un mismo mensaje ─────────────────────────────────────────
_MESES = {
    "enero": 1, "ene": 1, "febrero": 2, "feb": 2, "marzo": 3, "mar": 3,
    "abril": 4, "abr": 4, "mayo": 5, "may": 5, "junio": 6, "jun": 6,
    "julio": 7, "jul": 7, "agosto": 8, "ago": 8, "septiembre": 9,
    "setiembre": 9, "sep": 9, "sept": 9, "octubre": 10, "oct": 10,
    "noviembre": 11, "nov": 11, "diciembre": 12, "dic": 12,
}
_RE_FECHA_TXT = re.compile(r"\b(\d{1,2})\s*(?:de\s+)?([a-z]{3,10})\.?"
                            r"(?:\s+(?:de\s+|del\s+)?(\d{4}))?\b")
_RE_FECHA_NUM = re.compile(r"\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b")


def fecha_de_linea(linea: str, hoy=None):
    """Extrae la fecha de un encabezado tipo 'Asistencia 3 de agosto 2026'.

    Devuelve `date` o None. Si no trae año usa el de hoy, y si eso da una fecha
    muy futura (ej: 'asistencia 28 de diciembre' escrito en enero) resta un año.
    """
    from datetime import date as _date
    hoy = hoy or _date.today()
    txt = _sin_tildes(linea)

    dia = mes = anio = None
    m = _RE_FECHA_TXT.search(txt)
    if m and m.group(2) in _MESES:
        dia, mes = int(m.group(1)), _MESES[m.group(2)]
        anio = int(m.group(3)) if m.group(3) else None
    else:
        m = _RE_FECHA_NUM.search(txt)
        if m:
            dia, mes = int(m.group(1)), int(m.group(2))
            if m.group(3):
                anio = int(m.group(3))
                if anio < 100:
                    anio += 2000
    if dia is None or not (1 <= mes <= 12):
        return None

    sin_anio = anio is None
    try:
        f = _date(anio or hoy.year, mes, dia)
    except ValueError:
        return None
    if sin_anio and (f - hoy).days > 7:
        try:
            f = _date(f.year - 1, mes, dia)
        except ValueError:
            return None
    return f


def _es_encabezado_fecha(linea: str) -> bool:
    """True si la línea es un encabezado de día y NO una línea de trabajador."""
    m = _LINEA.match(linea)
    if m and _canonico(m.group(1)):
        return False
    return fecha_de_linea(linea) is not None


def parsear_asistencia_multi(texto: str, hoy=None):
    """Desglosa un mensaje que puede traer VARIOS días.

    Devuelve [{"fecha": date|None, "grupos": [...]}] — un elemento por día.
    Si el mensaje trae un solo día equivale a `parsear_asistencia`.
    Devuelve None si no es un parte de asistencia.
    """
    bloques, actual = [], {"fecha": None, "lineas": []}
    for linea in (texto or "").splitlines():
        if _es_encabezado_fecha(linea):
            if actual["lineas"]:
                bloques.append(actual)
            actual = {"fecha": fecha_de_linea(linea, hoy), "lineas": []}
        else:
            actual["lineas"].append(linea)
    if actual["lineas"] or actual["fecha"]:
        bloques.append(actual)

    salida = []
    for b in bloques:
        grupos = parsear_asistencia("\n".join(b["lineas"]))
        if grupos:
            salida.append({"fecha": b["fecha"], "grupos": grupos})
    return salida or None


def cultivo_de(act: str) -> str:
    a = _sin_tildes(act)
    if "nogal" in a:
        return "NOGALES"
    if "avellano" in a or "barbecho" in a:
        return "AVELLANOS"
    if "cerezo" in a:
        return "CEREZOS"
    return "GENERAL"


def tipo_de(act: str) -> str:
    a = _sin_tildes(act)
    if "aplicacion" in a:
        return "APLICACION"
    if "vacacion" in a or "ausente" in a:
        return "OTRO"
    if "riego" in a or "desagu" in a:
        return "RIEGO"
    return "LABOR"
