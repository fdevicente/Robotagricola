# -*- coding: utf-8 -*-
"""Migración única de los documentos locales a Drive.

Se corre por CARPETA, no todo de una: el riesgo no es el espacio (195 MB de
15 GB) sino perder la pista de un archivo, así que conviene verificar conteos
antes de dar por buena cada tanda.

Uso:
    python scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas" --simular
    python scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas"

`Facturas Recibidas` se reparte por AÑO DE EMISION leído del Master; las demás
conservan sus subcarpetas. El porqué de las dos reglas está en
`modules/drive/migracion.py`.
"""
import argparse
import collections
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS, EXCEL_PATH
from modules.drive.cola import Cola
from modules.drive.migracion import (DESTINOS, POR_ANIO, anio_de, destino_de,
                                     es_basura, indice_de_anios)

# Los documentos NO viven dentro de Robot/, sino en la carpeta que lo contiene —
# la misma donde está el Master. Anclarlo a EXCEL_PATH respeta el override por
# entorno y evita contar mal los niveles de dirname.
BASE = os.path.dirname(EXCEL_PATH)


def planificar(carpeta):
    """[(ruta_local, carpeta_drive)], sin tocar Drive ni la cola."""
    origen = os.path.join(BASE, carpeta)
    if not os.path.isdir(origen):
        print("No existe:", origen)
        raise SystemExit(1)

    indice = indice_de_anios(EXCEL_PATH) if carpeta == POR_ANIO else {}
    plan, basura = [], 0
    for dp, _, fs in os.walk(origen):
        for f in sorted(fs):
            if es_basura(f):
                basura += 1
                continue
            ruta = os.path.join(dp, f)
            anio = anio_de(indice, f) if carpeta == POR_ANIO else None
            plan.append((ruta, destino_de(carpeta, ruta, origen, anio)))
    return plan, basura


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("carpeta", choices=sorted(DESTINOS))
    ap.add_argument("--simular", action="store_true")
    args = ap.parse_args()

    plan, basura = planificar(args.carpeta)
    total_mb = sum(os.path.getsize(r) for r, _ in plan) / 1024 ** 2
    print("%s: %d archivos, %.1f MB (%d de basura, se saltan)"
          % (args.carpeta, len(plan), total_mb, basura))

    reparto = collections.Counter(c for _, c in plan)
    for destino in sorted(reparto):
        print("   %-40s %4d" % (destino, reparto[destino]))

    if args.simular:
        print("\n(simulación: no se encoló nada)")
        return

    cola = Cola(DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS)
    for ruta, destino in plan:
        cola.encolar(ruta, destino, os.path.basename(ruta))
    print("\nEncolados %d archivos. El job del bot los va a ir subiendo."
          % len(plan))


if __name__ == "__main__":
    main()
