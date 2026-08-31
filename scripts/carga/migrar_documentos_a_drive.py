# -*- coding: utf-8 -*-
"""Migración única de los documentos locales a Drive.

Se corre por CARPETA, no todo de una: el riesgo no es el espacio (201 MB de 15 GB)
sino perder la pista de un archivo, así que conviene verificar conteos
antes de dar por buena cada tanda.

Uso:
    python scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas" --simular
    python scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas"
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS, EXCEL_PATH
from modules.drive.cola import Cola

# Los documentos NO viven dentro de Robot/, sino en la carpeta que lo contiene —
# la misma donde está el Master. Anclarlo a EXCEL_PATH respeta el override por
# entorno y evita contar mal los niveles de dirname.
BASE = os.path.dirname(EXCEL_PATH)

# carpeta local -> carpeta en Drive
DESTINOS = {
    "Facturas Recibidas": "Facturas Recibidas",
    "Facturas Enviadas": "Facturas Enviadas",
    "BH": "Boletas Honorarios",
    "Guias de Despacho": "Guías de Despacho",
    "Rendiciones": "Rendiciones",
    "Legal": "Legal",
    "Carpeta Tributaria": "Tributario",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("carpeta", choices=sorted(DESTINOS))
    ap.add_argument("--simular", action="store_true")
    args = ap.parse_args()

    origen = os.path.join(BASE, args.carpeta)
    destino = DESTINOS[args.carpeta]
    if not os.path.isdir(origen):
        print("No existe:", origen)
        raise SystemExit(1)

    archivos = [os.path.join(dp, f)
                for dp, _, fs in os.walk(origen) for f in fs]
    total_mb = sum(os.path.getsize(a) for a in archivos) / 1024 ** 2
    print("%s: %d archivos, %.1f MB -> Drive:%s"
          % (args.carpeta, len(archivos), total_mb, destino))

    if args.simular:
        for a in archivos[:10]:
            print("   ", os.path.basename(a))
        if len(archivos) > 10:
            print("    ... y %d más" % (len(archivos) - 10))
        print("\n(simulación: no se encoló nada)")
        return

    cola = Cola(DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS)
    for a in archivos:
        # las facturas se parten por año usando la fecha del archivo
        carpeta = destino
        if args.carpeta == "Facturas Recibidas":
            import datetime as dt
            anio = dt.date.fromtimestamp(os.path.getmtime(a)).year
            carpeta = "%s/%d" % (destino, anio)
        cola.encolar(a, carpeta, os.path.basename(a))
    print("Encolados %d archivos. El job del bot los va a ir subiendo."
          % len(archivos))


if __name__ == "__main__":
    main()
