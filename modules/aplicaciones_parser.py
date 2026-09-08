# -*- coding: utf-8 -*-
"""Lee los partes de aplicacion que manda Juan.

POR QUE EXISTE
Juan mando 15 partes de aplicacion entre junio y agosto de 2026 y la hoja
`Aplicaciones` tiene 3 filas. Once aplicaciones de agroquimico --producto, dosis,
cultivo y kilos o litros gastados-- nunca quedaron anotadas. Eso no es solo
costo: es el registro de que se aplico, en que dosis y en que cuartel.

El formato que usa es regular, asi que NO hace falta IA:

    Aplicacion fungicida
    Miercoles 17 de junio 2026
    Producto : nordox super 75 wp
    Dosis : 180 gramos por 100
    Cerezos produccion
    Mojamiento 1500 litros por ha
    Total ha : 2
    Total Mojamiento : 3000 litros
    Total producto : 5.4 kilos

OJO CON EL "TOTAL" QUE SE ELIGE: hay hasta tres. "Total litros" y "Total
Mojamiento" son AGUA; el que interesa es el del PRODUCTO, que Juan escribe como
"Total producto" o como "Total <nombre del producto>". Confundirlos mete 3000
litros de agua donde van 5,4 kilos de fungicida.
"""
import re
import unicodedata

CULTIVOS = ("CEREZOS", "AVELLANOS", "NOGALES")

_TIPO = re.compile(r"^\s*aplicaci[oó]n\s+(?P<tipo>herbicida|fungicida|insecticida|foliar)",
                   re.I)
_PRODUCTO = re.compile(r"^\s*producto\s*:?\s*(?P<v>.+?)\s*$", re.I)
_DOSIS = re.compile(r"^\s*dosis\s*:?\s*(?P<v>.+?)\s*$", re.I)
_NUM = r"\d+(?:[.,]\d+)?"
_UNIDADES = {"litro": "L", "litros": "L", "l": "L",
             "kilo": "kg", "kilos": "kg", "kg": "kg",
             "gramo": "g", "gramos": "g", "gr": "g", "g": "g"}
# Lineas de "Total ..." que son AGUA o superficie, no producto.
_TOTAL_AJENO = re.compile(r"^\s*total\s+(litros|mojamiento|ha|hect)", re.I)
_TOTAL = re.compile(r"^\s*total\b", re.I)
_CANTIDAD = re.compile(rf"(?P<num>{_NUM})\s*(?P<uni>litros?|kilos?|kg|gramos?|gr|l|g)\b",
                      re.I)
_ETIQUETADA = re.compile(r"^\s*(aplicaci[oó]n|producto|dosis|mojamiento|total|equipo"
                         r"|sector|hora)", re.I)


def _sin_tildes(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s or ""))
                   if unicodedata.category(c) != "Mn").upper()


def es_aplicacion(texto) -> bool:
    return bool(_TIPO.match(str(texto or "")))


def _cantidad_de_producto(lineas, producto):
    """El total del PRODUCTO, no el del agua. Devuelve (cantidad, unidad)."""
    prod = _sin_tildes(producto)
    candidatas = []
    for l in lineas:
        if not _TOTAL.match(l) or _TOTAL_AJENO.match(l):
            continue
        cuerpo = _sin_tildes(l)
        # "Total producto : 5.4 kilos" o "Total ripper full: 15 litros"
        if "PRODUCTO" in cuerpo or (prod and prod.split()[0] in cuerpo):
            candidatas.append(l)
    if not candidatas:
        # "Total litros producto : 20 litros ripper full" cae aqui: empieza con
        # "Total litros" pero SI habla del producto.
        candidatas = [l for l in lineas
                      if _TOTAL.match(l) and "PRODUCTO" in _sin_tildes(l)]
    for l in candidatas:
        m = _CANTIDAD.search(l)
        if m:
            num = float(m.group("num").replace(",", "."))
            return num, _UNIDADES.get(m.group("uni").lower(), m.group("uni"))
    return None, ""


def _cultivo_y_sector(lineas, fecha_de_linea):
    """El cultivo y el cuartel salen de la linea que NO lleva etiqueta.

    OJO: la linea de la FECHA tampoco lleva etiqueta ("Lunes 15 de junio 2026"),
    asi que hay que saltarla o el sector termina siendo la fecha.
    """
    for l in lineas:
        if _ETIQUETADA.match(l) or fecha_de_linea(l):
            continue
        limpio = l.strip()
        if not limpio:
            continue
        arriba = _sin_tildes(limpio)
        for c in CULTIVOS:
            if c in arriba:
                resto = re.sub(c, "", arriba, flags=re.I).strip(" .:-")
                return c, resto.lower()
        return "GENERAL", limpio          # "Canales y sercos", "Cssa y orilleros"
    return "GENERAL", ""


def parsear(texto: str, fecha_de_linea) -> dict | None:
    """Devuelve los campos de la hoja Aplicaciones, o None si no es una.

    `fecha_de_linea` se inyecta (viene de modules.bitacora_asistencia) para no
    duplicar el parseo de fechas en castellano.
    """
    texto = str(texto or "")
    tipo = _TIPO.match(texto)
    if not tipo:
        return None
    lineas = [l for l in texto.splitlines() if l.strip()]
    if len(lineas) < 3:
        return None                       # "Aplicación herbicida" a secas

    fecha = None
    for l in lineas[:3]:
        f = fecha_de_linea(l)
        if f:
            fecha = f.strftime("%Y-%m-%d")
            break

    producto = dosis = ""
    for l in lineas:
        m = _PRODUCTO.match(l)
        if m and not producto:
            producto = m.group("v").strip()
        m = _DOSIS.match(l)
        if m and not dosis:
            dosis = m.group("v").strip()
    if not producto:
        return None

    cantidad, unidad = _cantidad_de_producto(lineas, producto)
    cultivo, sector = _cultivo_y_sector(lineas, fecha_de_linea)
    return {
        "fecha": fecha, "producto": producto.title(), "cantidad": cantidad,
        "unidad": unidad, "cultivo": cultivo, "sector": sector,
        "tipo": tipo.group("tipo").lower(), "dosis": dosis,
        "texto_original": texto,
    }
