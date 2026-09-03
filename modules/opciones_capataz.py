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
    con_fecha = [m for m in maquinas if m.get("fecha") is not None]
    sin_fecha = [m for m in maquinas if m.get("fecha") is None]
    # str() porque la fecha viene como date o como texto segun de donde salga
    con_fecha.sort(key=lambda m: str(m["fecha"]), reverse=True)
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
            for row in wb[BITACORA_SHEET].iter_rows(min_row=2, values_only=True):
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
