# -*- coding: utf-8 -*-
"""Vacía la cola de subidas. El archivo local NUNCA se borra acá.

Que el archivo siga en disco después de subir es a propósito: Drive no puede
ser el único lugar donde vivió un documento.
"""
import logging
import os

logger = logging.getLogger(__name__)


def procesar_cola(cola, drive, carpetas) -> dict:
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
                subidos += 1
                continue
            file_id = drive.subir(ruta, cid, item["nombre"])
            cola.marcar_ok(item["id"], file_id)
            subidos += 1
            logger.info("Drive: subido %s -> %s", item["nombre"], item["carpeta"])
        except Exception as e:
            cola.marcar_error(item["id"], str(e))
            fallidos += 1
            logger.warning("Drive: falló %s (%s)", item["nombre"], e)
    return {"subidos": subidos, "fallidos": fallidos}
