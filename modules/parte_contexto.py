# -*- coding: utf-8 -*-
"""Lo que el bot ya sabe y la IA necesita para normalizar.

Se arma una vez por mensaje y se le pasa IGUAL al lector y al juez: si midieran
contra vocabularios distintos, el juez marcaria como duda lo que la IA hizo bien.

OJO CON LOS TRABAJADORES: no salen solo de la hoja Personal. Medido el
2-sep-2026, Personal tiene 6 filas con el nombre legal completo ("Felicito
Amigo Soto") y no incluye a Richard Padilla ni a su hijo, mientras la columna
Trabajadores de la bitacora usa los 8 nombres canonicos que el bot viene usando.

Y NO SE PUEDEN METER LOS DOS. Medido el 3-sep-2026: de las 6 filas de Personal,
5 son el nombre legal de alguien que la bitacora ya conoce por su canonico, y
NINGUNA es gente nueva. Meter ambos le da a la IA dos nombres validos para la
misma persona; el que elija se escribe en la hoja y parte el historial, porque
bitacora_asistencia solo cuenta jornadas de los canonicos.
"""
import logging
import unicodedata

logger = logging.getLogger(__name__)

BITACORA_SHEET = "Bitácora"
PERSONAL_SHEET = "Personal"


def _clave(nombre) -> str:
    """Clave de comparacion: sin tildes, minuscula y espacios colapsados.

    Sin esto "Ramiro Amigo", "ramiro amigo" y "Ramiro  Amigo" son tres personas.
    Importa porque la columna Trabajadores ahora la escribe la IA y el ciclo es
    cerrado: lo que emita una vez se lee de vuelta y se queda en el vocabulario.
    """
    sin = "".join(c for c in unicodedata.normalize("NFD", str(nombre or ""))
                  if unicodedata.category(c) != "Mn")
    return " ".join(sin.lower().split())


def _encabezado(ws) -> list:
    """Primera fila, o [] si la hoja esta vacia.

    `next()` a secas sobre una hoja sin filas lanza StopIteration, y mas arriba
    eso se comia la hoja Maquinaria entera sin dejar rastro.
    """
    fila = next(ws.iter_rows(min_row=1, max_row=1), None)
    return [c.value for c in fila] if fila else []


def construir(excel_path: str | None = None) -> dict:
    """Devuelve {"trabajadores": [...], "alias": {...}, "maquinas": [...]}."""
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    from modules.bitacora_asistencia import canonico_por_nombre_completo
    from modules.bitacora_extractor import ALIAS, TRABAJADORES_CONOCIDOS
    from modules.maquinaria import maquinas_conocidas

    ruta = excel_path or EXCEL_PATH
    vistos, orden = {}, []            # clave normalizada -> posicion en orden

    def _sumar(n, canonico=False):
        n = str(n or "").strip()
        k = _clave(n)
        if not k:
            return
        if k not in vistos:
            vistos[k] = len(orden)
            orden.append(n)
        elif canonico:
            # La posicion la fija quien llego primero, pero la GRAFIA la gana el
            # canonico: la de la bitacora puede ser una variante que escribio la
            # IA, y esa no puede desplazar al nombre bueno en el prompt.
            orden[vistos[k]] = n

    def _leer(hoja, saca):
        try:
            wb = load_workbook(ruta, read_only=True, data_only=True)
            try:
                if hoja in wb.sheetnames:
                    saca(wb[hoja])
            finally:
                wb.close()
        except Exception as e:                # un Excel raro no puede voltear esto
            # repr y no str: StopIteration no tiene mensaje y el log salia vacio
            logger.warning("parte_contexto: no pude leer %s de %s: %r",
                           hoja, ruta, e)

    def _de_bitacora(ws):
        enc = _encabezado(ws)
        if "Trabajadores" not in enc:
            return
        i = enc.index("Trabajadores")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) > i and row[i]:
                for n in str(row[i]).split(","):
                    _sumar(n)

    def _de_personal(ws):
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if not row or not row[0]:
                continue
            # Solo la gente que NO conocemos ya. Se usa la version estricta, por
            # tokens completos: _canonico entero calza por nombre de pila y
            # descartaria a un "Juan Soto Rivera" recien contratado.
            if canonico_por_nombre_completo(row[0]) is None:
                _sumar(row[0])

    # El ORDEN importa: es el orden en que la lista se le muestra al modelo.
    _leer(BITACORA_SHEET, _de_bitacora)       # 1 — el vocabulario que ya se usa
    for n in TRABAJADORES_CONOCIDOS:          # 2 — los de siempre
        _sumar(n, canonico=True)
    _leer(PERSONAL_SHEET, _de_personal)       # 3 — solo los que no conocemos

    try:
        maquinas = maquinas_conocidas(ruta)
    except Exception as e:
        logger.warning("parte_contexto: no pude leer las máquinas: %r", e)
        maquinas = []

    return {"trabajadores": orden, "alias": dict(ALIAS), "maquinas": maquinas}
