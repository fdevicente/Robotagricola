# -*- coding: utf-8 -*-
"""Envoltura fina de la API de Google Drive.

Se mantiene chica a propósito: todo lo que el robot necesita son siete
operaciones. Con una interfaz así de acotada, `tests/drive_falso.py` puede
sustituirla y las pruebas nunca tocan la red.
"""
import logging
import os

logger = logging.getLogger(__name__)

CARPETA_MIME = "application/vnd.google-apps.folder"


class DriveCliente:
    def __init__(self, servicio=None):
        """`servicio` se inyecta en pruebas; en producción lo construye auth."""
        if servicio is None:
            from modules.drive.auth import construir_servicio
            servicio = construir_servicio()
        self._s = servicio

    def subir(self, ruta_local, carpeta_id, nombre):
        from googleapiclient.http import MediaFileUpload
        media = MediaFileUpload(ruta_local, resumable=True)
        meta = {"name": nombre, "parents": [carpeta_id]}
        res = self._s.files().create(body=meta, media_body=media,
                                     fields="id").execute()
        return res["id"]

    def crear_carpeta(self, nombre, padre_id):
        meta = {"name": nombre, "mimeType": CARPETA_MIME, "parents": [padre_id]}
        return self._s.files().create(body=meta, fields="id").execute()["id"]

    def buscar_carpeta(self, nombre, padre_id):
        q = ("name = '%s' and mimeType = '%s' and '%s' in parents "
             "and trashed = false" % (nombre.replace("'", "\'"),
                                       CARPETA_MIME, padre_id))
        r = self._s.files().list(q=q, fields="files(id)", pageSize=1).execute()
        f = r.get("files") or []
        return f[0]["id"] if f else None

    def buscar_archivo(self, nombre, carpeta_id):
        q = ("name = '%s' and '%s' in parents and trashed = false"
             % (nombre.replace("'", "\'"), carpeta_id))
        r = self._s.files().list(q=q, fields="files(id)", pageSize=1).execute()
        f = r.get("files") or []
        return f[0]["id"] if f else None

    def listar(self, carpeta_id):
        """Solo documentos: EXCLUYE las subcarpetas.

        En Drive una carpeta es un archivo más, así que sin este filtro
        `_Entrada/Sin procesar` salía listada como si fuera un documento: el
        robot intentaba descargarla, fallaba, y después la movía dentro de sí
        misma (HttpError 400 con addParents == fileId). Pasó en producción el
        26-ago-2026, cada 15 minutos.
        """
        q = ("'%s' in parents and trashed = false and mimeType != '%s'"
             % (carpeta_id, CARPETA_MIME))
        r = self._s.files().list(q=q, fields="files(id,name)",
                                 pageSize=1000).execute()
        return [{"id": a["id"], "nombre": a["name"]} for a in r.get("files") or []]

    def mover(self, file_id, carpeta_destino_id):
        actual = self._s.files().get(fileId=file_id,
                                     fields="parents").execute()
        previos = ",".join(actual.get("parents") or [])
        self._s.files().update(fileId=file_id, addParents=carpeta_destino_id,
                               removeParents=previos, fields="id").execute()

    def cuota(self):
        r = self._s.about().get(fields="storageQuota").execute()
        q = r["storageQuota"]
        return {"usado": int(q.get("usage", 0)),
                "total": int(q.get("limit", 0)) or None}

    def descargar(self, file_id, ruta_local):
        from googleapiclient.http import MediaIoBaseDownload
        pedido = self._s.files().get_media(fileId=file_id)
        with open(ruta_local, "wb") as fh:
            bajada = MediaIoBaseDownload(fh, pedido)
            listo = False
            while not listo:
                _, listo = bajada.next_chunk()
        return ruta_local
