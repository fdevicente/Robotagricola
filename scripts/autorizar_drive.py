# -*- coding: utf-8 -*-
"""Autorización única del Drive del robot. Se corre a mano, abre el navegador."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google_auth_oauthlib.flow import InstalledAppFlow

from config import DRIVE_CLIENT_SECRET, DRIVE_TOKEN_PATH
from modules.drive.auth import ALCANCES

if not os.path.exists(DRIVE_CLIENT_SECRET):
    print("Falta el archivo de credenciales:", DRIVE_CLIENT_SECRET)
    print("Bájalo de Google Cloud Console > Credenciales > ID de cliente OAuth")
    print("(tipo: Aplicación de escritorio) y guárdalo con ese nombre.")
    raise SystemExit(1)

flow = InstalledAppFlow.from_client_secrets_file(DRIVE_CLIENT_SECRET, ALCANCES)
creds = flow.run_local_server(port=0)
with open(DRIVE_TOKEN_PATH, "w", encoding="utf-8") as fh:
    fh.write(creds.to_json())
print("Listo. Token guardado en", DRIVE_TOKEN_PATH)
