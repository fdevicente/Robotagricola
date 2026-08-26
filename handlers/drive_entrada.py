# -*- coding: utf-8 -*-
"""Revisa `_Entrada/` en Drive y procesa lo que aparezca.

Mover el archivo a su carpeta definitiva ES la marca de procesado: no hace
falta llevar una lista aparte, y un reinicio a mitad de camino no duplica nada.
"""
import logging

logger = logging.getLogger(__name__)

ENTRADA = "_Entrada"
SIN_PROCESAR = "_Entrada/Sin procesar"


def carpeta_para(datos: dict) -> str:
    """Carpeta destino según el tipo. Lanza si no lo reconoce.

    Lanzar es deliberado: `revisar_entrada` manda a 'Sin procesar' lo que no
    supo clasificar, en vez de adivinar una carpeta.
    """
    from datetime import date
    tipo = str(datos.get("tipo") or "").lower()
    if tipo == "honorarios":
        return "Boletas Honorarios"
    if tipo == "boleta":
        # Boleta de COMPRA (supermercado, ferretería). Va aparte de las
        # facturas porque el bot ya las separa en el PC: config.BOLETAS_DIR
        # es distinto de DOWNLOAD_DIR.
        return "Boletas"
    if tipo == "guia":
        return "Guías de Despacho"
    if tipo == "factura":
        anio = str(datos.get("fecha") or "")[:4]
        if not anio.isdigit():
            anio = str(date.today().year)
        return "Facturas Recibidas/%s" % anio
    raise ValueError("tipo de documento desconocido: %r" % tipo)


def revisar_entrada(drive, carpetas, procesar) -> dict:
    """`procesar(archivo)` devuelve la carpeta destino, o lanza si no pudo.

    Recibe el dict completo del archivo ({id, nombre}) porque quien lo procesa
    de verdad necesita el id para descargarlo.
    """
    entrada_id = carpetas.id_de(ENTRADA)
    sin_procesar_id = carpetas.id_de(SIN_PROCESAR)
    procesados = fallidos = 0
    for archivo in drive.listar(entrada_id):
        try:
            destino = procesar(archivo)
            drive.mover(archivo["id"], carpetas.id_de(destino))
            procesados += 1
            logger.info("Drive entrada: %s -> %s", archivo["nombre"], destino)
        except Exception as e:
            drive.mover(archivo["id"], sin_procesar_id)
            fallidos += 1
            logger.warning("Drive entrada: no pude con %s (%s)",
                            archivo["nombre"], e)
    return {"procesados": procesados, "sin_procesar": fallidos}
