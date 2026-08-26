# -*- coding: utf-8 -*-
"""Vacía la cola de subidas. El archivo local NUNCA se borra acá.

Que el archivo siga en disco después de subir es a propósito: Drive no puede
ser el único lugar donde vivió un documento.
"""
import logging
import os

logger = logging.getLogger(__name__)


def procesar_cola(cola, drive, carpetas, excel_path: str = None) -> dict:
    """Vacía la cola y, por cada subida, deja el enlace en el Master.

    `excel_path` es el Excel donde `_enlazar` escribe el enlace. Por
    defecto es None y `_enlazar` usa el MASTER real (config.EXCEL_PATH):
    en producción no hace falta pasarlo. Los tests SIEMPRE deben pasar
    una ruta explícita a un Excel de prueba — de lo contrario terminan
    escribiendo en el Master real, que es justo lo que no puede pasar.
    """
    subidos = fallidos = 0
    for item in cola.pendientes():
        ruta = item["ruta_local"]
        try:
            if not os.path.exists(ruta):
                raise FileNotFoundError("ya no está en disco: %s" % ruta)
            cid = carpetas.id_de(item["carpeta"])
            existente = drive.buscar_archivo(item["nombre"], cid)
            if existente:
                # Juan reenvía cosas; no duplicar
                cola.marcar_ok(item["id"], existente)
                _enlazar(item, existente, excel_path)
                subidos += 1
                continue
            file_id = drive.subir(ruta, cid, item["nombre"])
            cola.marcar_ok(item["id"], file_id)
            _enlazar(item, file_id, excel_path)
            subidos += 1
            logger.info("Drive: subido %s -> %s", item["nombre"], item["carpeta"])
        except Exception as e:
            cola.marcar_error(item["id"], str(e))
            fallidos += 1
            logger.warning("Drive: falló %s (%s)", item["nombre"], e)
    return {"subidos": subidos, "fallidos": fallidos}


def _enlazar(item: dict, file_id: str, excel_path: str = None) -> None:
    """Deja el enlace en la fila de la factura. Nunca lanza.

    Perder el enlace es molesto; hacer fallar la subida por eso sería peor.

    `excel_path` permite aislar el Excel en las pruebas. Si no se pasa
    (None), se usa el MASTER real de config.EXCEL_PATH — el comportamiento
    de producción.
    """
    import re
    try:
        m = re.search(r"_(\d+)\.[A-Za-z0-9]+$", item["nombre"])
        if not m:
            return                       # respaldos y otros no llevan enlace
        if excel_path is None:
            from config import EXCEL_PATH as excel_path
        from modules.drive.enlaces import guardar_enlace
        guardar_enlace(excel_path, m.group(1), file_id)
    except Exception as e:
        logger.warning("No pude enlazar %s: %s", item["nombre"], e)
