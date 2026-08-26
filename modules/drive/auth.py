# -*- coding: utf-8 -*-
"""Credenciales OAuth para la cuenta de Drive del robot.

POR QUÉ OAUTH Y NO UNA CUENTA DE SERVICIO
Con un Gmail común, los archivos que sube una cuenta de servicio quedan a
nombre de ella, y las cuentas de servicio no tienen cuota de Drive en cuentas
de consumidor: la subida falla. Ese camino solo sirve con Google Workspace y
unidades compartidas.

El permiso es `drive` completo (no `drive.file`) porque la carpeta de entrada
necesita leer archivos que el robot NO creó. Eso hace que Google muestre una
advertencia de "app no verificada" en la primera autorización: se acepta una
vez, es una herramienta interna de un solo usuario.
"""
import logging
import os

logger = logging.getLogger(__name__)

ALCANCES = ["https://www.googleapis.com/auth/drive"]


class FaltaAutorizacion(RuntimeError):
    """No hay token utilizable. Trae el paso a paso para arreglarlo."""


def cargar_credenciales(token_path: str = None, client_secret_path: str = None):
    """Devuelve credenciales válidas, refrescándolas si hace falta."""
    from config import DRIVE_TOKEN_PATH, DRIVE_CLIENT_SECRET
    token_path = token_path or DRIVE_TOKEN_PATH
    client_secret_path = client_secret_path or DRIVE_CLIENT_SECRET

    if not os.path.exists(token_path):
        raise FaltaAutorizacion(
            "Falta autorizar el Google Drive del robot.\n"
            "Corre una vez, desde la carpeta Robot:\n"
            "  %LOCALAPPDATA%\Python\bin\python3.11.exe "
            "scripts/autorizar_drive.py\n"
            "Se abre el navegador, aceptas la advertencia de app no verificada "
            "y queda listo.")

    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    creds = Credentials.from_authorized_user_file(token_path, ALCANCES)
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open(token_path, "w", encoding="utf-8") as fh:
                fh.write(creds.to_json())
        except Exception as e:
            raise FaltaAutorizacion(
                "El permiso de Drive dejó de servir (%s).\n"
                "Vuelve a autorizar: python scripts/autorizar_drive.py" % e)
    if not creds or not creds.valid:
        raise FaltaAutorizacion(
            "El token de Drive no sirve. Corre: "
            "python scripts/autorizar_drive.py")
    return creds


def construir_servicio():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=cargar_credenciales(),
                 cache_discovery=False)
