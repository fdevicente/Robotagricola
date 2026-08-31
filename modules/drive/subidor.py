# -*- coding: utf-8 -*-
"""Vacía la cola de subidas. El archivo local NUNCA se borra acá.

Que el archivo siga en disco después de subir es a propósito: Drive no puede
ser el único lugar donde vivió un documento.
"""
import logging
import os
import re

# `PROVEEDOR_NRO_20260826_130058` -> se le saca el sello de fecha y hora
_SELLO = re.compile(r"^(.*?)_\d{8}_\d{6}$")
# ...y de lo que queda, el numero es la ultima tira de digitos tras un '_'
_PROV_Y_NUMERO = re.compile(r"^(.+)_(\d+)$")
_TIENE_LETRA = re.compile(r"[^\W\d_]", re.UNICODE)

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


def _partes_del_nombre(nombre: str):
    """(proveedor, número) del nombre del archivo, o None si no es una factura.

    Los nombres los arma `handlers.facturas._renombrar_archivo` como
    `PROVEEDOR_NRO.ext`, y cuando ese nombre ya existía le agrega un sello de
    fecha y hora: `PROVEEDOR_NRO_20260826_130058.ext`. Hay que sacarle el sello
    ANTES de buscar el número, porque si no el `130058` se hace pasar por el
    número de factura — así se perdieron 124 enlaces.

    Exige que haya algo delante del número: `20260826_130058.jpg` sin proveedor
    no identifica ninguna factura, y los respaldos del Master
    (`2026-08-26_16-40.xlsx`) no tienen que colar.
    """
    tallo = os.path.splitext(nombre)[0]
    sello = _SELLO.match(tallo)
    if sello:
        tallo = sello.group(1)
    m = _PROV_Y_NUMERO.match(tallo)
    if not m or not _TIENE_LETRA.search(m.group(1)):
        return None                  # un proveedor tiene letras; una fecha no
    return m.group(1), m.group(2)


def _enlazar(item: dict, file_id: str, excel_path: str = None) -> None:
    """Deja el enlace en la fila de la factura. Nunca lanza.

    Perder el enlace es molesto; hacer fallar la subida por eso sería peor.

    `excel_path` permite aislar el Excel en las pruebas. Si no se pasa
    (None), se usa el MASTER real de config.EXCEL_PATH — el comportamiento
    de producción.
    """
    try:
        partes = _partes_del_nombre(item["nombre"])
        if partes is None:
            return                       # respaldos y otros no llevan enlace
        proveedor, numero = partes
        if excel_path is None:
            from config import EXCEL_PATH as excel_path
        from modules.drive.enlaces import guardar_enlace
        if not guardar_enlace(excel_path, numero, file_id, proveedor=proveedor):
            # Sin fila que calce no hay nada que escribir, pero que se vea:
            # es la señal de que el proveedor del Master está escrito distinto.
            logger.info("Drive: %s subido, sin fila para %s Nº%s",
                        item["nombre"], proveedor, numero)
    except Exception as e:
        logger.warning("No pude enlazar %s: %s", item["nombre"], e)
