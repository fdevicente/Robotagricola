# -*- coding: utf-8 -*-
"""Traduce rutas legibles ('Facturas Recibidas/2026') a IDs de carpeta de Drive.

Cachea los IDs en memoria: resolver la misma ruta en cada factura serían dos
llamadas a la API por documento.
"""
import logging

logger = logging.getLogger(__name__)


class Carpetas:
    def __init__(self, drive, raiz_id: str):
        self.drive = drive
        self.raiz_id = raiz_id
        self._cache = {"": raiz_id}

    def id_de(self, ruta: str) -> str:
        """ID de la carpeta, creando los tramos que falten."""
        ruta = (ruta or "").strip("/")
        if ruta in self._cache:
            return self._cache[ruta]
        padre = self.raiz_id
        acumulada = []
        for tramo in ruta.split("/"):
            if not tramo:
                continue
            acumulada.append(tramo)
            clave = "/".join(acumulada)
            if clave in self._cache:
                padre = self._cache[clave]
                continue
            cid = self.drive.buscar_carpeta(tramo, padre)
            if cid is None:
                cid = self.drive.crear_carpeta(tramo, padre)
                logger.info("Drive: carpeta creada %s", clave)
            self._cache[clave] = cid
            padre = cid
        return padre
