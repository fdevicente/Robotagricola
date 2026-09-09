# -*- coding: utf-8 -*-
"""Borra tres partes que quedaron anotados dos veces en la hoja Bitácora.

QUE SE ENCONTRO (8-sep-2026, auditando el Master real)
  1. El 27-jul entero, dos veces: 10 jornadas-hombre donde hubo 5. El mismo dia
     entro por el bot ("Asistencia 2026-07-27") y otra vez con el texto crudo
     del mensaje ("Lunes 27 de julio 2026 ...").
  2. El parte del "Martes 28 julio 2028" --el año mal tecleado-- anotado bajo
     DOS fechas: 28-jul y 4-ago. El texto dice martes 28 de julio y el
     28-jul-2026 ES martes, asi que la copia del 4-ago sobra. 6 jornadas de mas.
  3. La lectura del MASSEY FERGUSON 6711 del 18-ago (2041 → 2043), anotada dos
     veces palabra por palabra, incluido el typo "/massey". Esta ademas hacia
     RETROCEDER el horometro --12-ago 2039, 18-ago 2043, 19-ago 2041-- y las
     horas se calculan por diferencia, asi que la secuencia quedaba mal.

COMO SE ELIGE QUE FILA MUERE
No por numero de fila: por el TEXTO FUENTE. Y antes de borrar cada duplicado se
comprueba que **el gemelo que se queda existe de verdad**. Es la leccion que
costo dos filas duplicadas el 7-sep: la comprobacion de "¿ya esta?" tiene que ir
contra el dato con el que uno se va a quedar, no contra el que traia el mensaje.

USO
    python scripts/carga/borrar_duplicados_bitacora_sep2026.py --simular
    python scripts/carga/borrar_duplicados_bitacora_sep2026.py
"""
import os
import shutil
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))

from openpyxl import load_workbook

HOJA = "Bitácora"


def _txt(v):
    return " ".join(str(v or "").split())


# (nombre, criterio de la que SE BORRA, cuantas, criterio del gemelo que SE QUEDA, cuantas)
CASOS = [
    ("27-jul duplicado entero",
     lambda f: f["fecha"] == "2026-07-27"
     and f["registro"].startswith("Lunes 27 de julio 2026"), 4,
     lambda f: f["fecha"] == "2026-07-27"
     and f["registro"].startswith("Asistencia 2026-07-27"), 4),

    ("parte del 28-jul copiado al 4-ago",
     lambda f: f["fecha"] == "2026-08-04"
     and f["registro"].startswith("Martes 28 julio 2028"), 2,
     lambda f: f["fecha"] == "2026-07-28"
     and f["registro"].startswith("Martes 28 julio 2028"), 2),

    ("lectura del 6711 del 18-ago",
     lambda f: f["fecha"] == "2026-08-18" and f["odometro"] == 2043
     and "6711" in f["maquina"], 1,
     lambda f: f["fecha"] == "2026-08-19" and f["odometro"] == 2043
     and "6711" in f["maquina"], 1),
]


def _leer(ws):
    enc = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {n: k for k, n in enumerate(enc)}

    def col(r, n):
        return r[idx[n]] if n in idx and len(r) > idx[n] else None

    filas = {}
    for n, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not r or not any(r):
            continue
        odo = col(r, "Odómetro")
        filas[n] = {
            "fecha": str(col(r, "Fecha"))[:10],
            "tipo": _txt(col(r, "Tipo")),
            "actividad": _txt(col(r, "Actividad")),
            "trabajadores": _txt(col(r, "Trabajadores")),
            "jh": col(r, "Jornadas Hombre"),
            "maquina": _txt(col(r, "Máquina")),
            "odometro": float(odo) if odo not in (None, "") else None,
            "registro": _txt(col(r, "Registro")),
        }
    return filas


def main(simular: bool) -> int:
    from config import EXCEL_PATH

    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    filas = _leer(wb[HOJA])
    wb.close()
    print("Hoja %s: %d filas de datos\n" % (HOJA, len(filas)))

    a_borrar, jh_extra = [], 0.0
    for nombre, cond_borra, n_borra, cond_queda, n_queda in CASOS:
        muere = [n for n, f in filas.items() if cond_borra(f)]
        queda = [n for n, f in filas.items() if cond_queda(f)]
        print("=" * 78)
        print(nombre)
        print("=" * 78)
        if len(queda) != n_queda:
            print("  ABORTA: esperaba %d fila(s) que se quedan y encontré %d %s"
                  % (n_queda, len(queda), queda))
            return 2
        if len(muere) != n_borra:
            print("  ABORTA: esperaba %d fila(s) a borrar y encontré %d %s"
                  % (n_borra, len(muere), muere))
            return 2
        for n in sorted(queda):
            f = filas[n]
            print("  SE QUEDA  fila %-4d %s %-11s jh=%-5s odo=%-6s %s"
                  % (n, f["fecha"], f["tipo"][:11], f["jh"], f["odometro"],
                     (f["trabajadores"] or f["actividad"])[:34]))
        for n in sorted(muere):
            f = filas[n]
            print("  SE BORRA  fila %-4d %s %-11s jh=%-5s odo=%-6s %s"
                  % (n, f["fecha"], f["tipo"][:11], f["jh"], f["odometro"],
                     (f["trabajadores"] or f["actividad"])[:34]))
            jh_extra += float(f["jh"] or 0)
        a_borrar += muere
        print()

    print("=" * 78)
    print("TOTAL: %d filas a borrar · %g jornadas-hombre de más se van"
          % (len(a_borrar), jh_extra))
    print("       la hoja queda en %d filas" % (len(filas) - len(a_borrar)))
    print("=" * 78)

    if simular:
        print("\n--simular: NO se tocó el archivo.")
        return 0

    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(os.path.dirname(EXCEL_PATH),
                           "MASTER Agricola Santa Elisa_bak_antes_dups_%s.xlsx" % marca)
    shutil.copy2(EXCEL_PATH, destino)
    print("\nRespaldo local: %s" % os.path.basename(destino))
    try:
        from infrastructure.backups import backup_master
        backup_master("borrar duplicados bitácora")
        print("Respaldo a Dropbox/Drive: ok")
    except Exception as e:
        print("Respaldo a Dropbox/Drive: falló (%s) — el local sí quedó" % e)

    # data_only=False a proposito: guardar con data_only=True reemplazaria por
    # valores todas las formulas del Master.
    wb = load_workbook(EXCEL_PATH)
    ws = wb[HOJA]
    for n in sorted(a_borrar, reverse=True):     # de abajo hacia arriba
        ws.delete_rows(n, 1)
    wb.save(EXCEL_PATH)
    wb.close()
    print("Borradas %d filas." % len(a_borrar))
    return 0


if __name__ == "__main__":
    sys.exit(main("--simular" in sys.argv))
