# -*- coding: utf-8 -*-
"""Reingresa los tres partes que /bitacora colapso el 11-sep-2026.

QUE PASO
Juan mando tres partes usando /bitacora. El texto libre en modo capataz pasa por
`auto_guardar_bitacora`, que usa el parser determinista y escribe UNA FILA POR
LABOR; pero `/bitacora` abre un flujo y su texto iba a la IA, que RESUME. Los
tres quedaron mal:

  2-sep : UNA fila "Sacar restos poda" con 11 jornadas   (son 3 labores)
  10-sep: NADA                                            (son 4 labores)
  11-sep: UNA fila tipo OTRO, "labores varias"            (son 4 labores)

El parser lee los tres perfecto: el problema no era leerlos, era por donde
entraban. Ya esta arreglado en handle_text_bitacora.

Y DE PASO, el "Ripper full" de las 19:22 UTC: contesto "Salido 21 litros
avellanos" a la pregunta del producto, asi que quedo como "Insumo no
especificado" y los 21 L nunca se descontaron. Se le pone nombre y se descuentan.

USO
    python scripts/carga/recuperar_partes_11sep2026.py --simular
    python scripts/carga/recuperar_partes_11sep2026.py
"""
import os
import shutil
import sys
from datetime import datetime

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, RAIZ)

# Los mensajes, con un prefijo para comprobar que son los que creemos.
PERDIDOS = [
    (3298, "Asistencia miércoles 2 de septiembre 2026"),
    (3304, "Asistencia jueves 10 de septiembre 2026"),
    (3310, "Asistencia viernes 11 de septiembre 2026"),
]
QUIEN = "Juan Parada"

# La fila de aplicacion que quedo sin nombre de producto.
INSUMO_MALO = "Insumo no especificado"
INSUMO_BUENO = "Ripper Full"
LITROS = 21


def _mensajes():
    from modules.telegram_backup import leer_mes
    por_id = {f.get("message_id"): f for f in leer_mes("2026-09") if f.get("text")}
    salida = []
    for mid, prefijo in PERDIDOS:
        fila = por_id.get(mid)
        if fila is None:
            raise SystemExit("No encontré el mensaje %s en el respaldo crudo." % mid)
        if not fila["text"].startswith(prefijo):
            raise SystemExit("El mensaje %s no es el esperado: %r"
                             % (mid, fila["text"][:60]))
        salida.append(fila)
    return salida


def _filas_de(texto):
    from modules.bitacora_asistencia import (cultivo_de, parsear_asistencia_multi,
                                             tipo_de)
    filas = []
    for d in parsear_asistencia_multi(texto):
        for g in d["grupos"]:
            filas.append({
                "fecha": d["fecha"].strftime("%Y-%m-%d") if d["fecha"] else "",
                "tipo": tipo_de(g["actividad"]), "actividad": g["actividad"],
                "cultivo": cultivo_de(g["actividad"]), "sector": "",
                "jornadas_hombre": g["jornadas_hombre"],
                "trabajadores": g["trabajadores"], "insumo": "", "cantidad": None,
                "unidad": "", "maquina": "", "odometro": None,
                "superficie_ha": None, "texto_original": texto,
            })
    return filas


def main(simular):
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    from excel_manager import _save_wb

    nuevas = []
    for m in _mensajes():
        nuevas.extend(_filas_de(m["text"]))
    nuevas.sort(key=lambda f: f["fecha"])
    fechas = {f["fecha"] for f in nuevas}

    wb = load_workbook(EXCEL_PATH)
    ws = wb["Bitácora"]
    enc = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    i = {n: k + 1 for k, n in enumerate(enc)}

    # Las filas colapsadas que hay que sacar: son de esas fechas, tienen
    # trabajadores, y su texto original es uno de los partes.
    borrar = []
    for n in range(2, ws.max_row + 1):
        f = str(ws.cell(n, i["Fecha"]).value)[:10]
        if f not in fechas or not ws.cell(n, i["Trabajadores"]).value:
            continue
        reg = str(ws.cell(n, i["Registro"]).value or "")
        if reg.startswith("Asistencia"):
            borrar.append((n, f, str(ws.cell(n, i["Actividad"]).value)[:44],
                           ws.cell(n, i["Jornadas Hombre"]).value))

    print("BORRAR LAS FILAS COLAPSADAS (%d)" % len(borrar))
    for n, f, act, jh in borrar:
        print("  fila %-4d %s  %-44s %s jh" % (n, f, act, jh))
    if not borrar:
        print("  ninguna")

    print("\nREINGRESAR (%d filas, %g jornadas)"
          % (len(nuevas), sum(f["jornadas_hombre"] or 0 for f in nuevas)))
    for f in nuevas:
        print("  %s  %-11s %-28s %4s jh  %s"
              % (f["fecha"], f["tipo"], f["actividad"][:28],
                 f["jornadas_hombre"] or 0, ", ".join(f["trabajadores"])[:44]))

    # La aplicación sin nombre de producto.
    fila_app = next((n for n in range(2, ws.max_row + 1)
                     if str(ws.cell(n, i["Insumo"]).value or "") == INSUMO_MALO), None)
    print("\nLA APLICACIÓN SIN PRODUCTO")
    if fila_app:
        print("  bitácora fila %-4d insumo: %r → %r" % (fila_app, INSUMO_MALO,
                                                        INSUMO_BUENO))
        print("  descontar %g L de %s del inventario" % (LITROS, INSUMO_BUENO))
    else:
        print("  ya está corregida")

    if simular:
        print("\n--simular: no se escribió nada.")
        return 0

    resp = shutil.copy2(EXCEL_PATH, EXCEL_PATH.replace(
        ".xlsx", "_bak_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S")))
    print("\nRespaldo: %s" % os.path.basename(resp))

    if fila_app:
        ws.cell(fila_app, i["Insumo"]).value = INSUMO_BUENO
    for n, _, _, _ in sorted(borrar, reverse=True):
        ws.delete_rows(n, 1)
    _save_wb(wb)
    wb.close()

    from bitacora_manager import registrar_bitacora_estructurada
    for f in nuevas:
        registrar_bitacora_estructurada(f, QUIEN)
    print("%d filas reingresadas, %d colapsadas borradas." % (len(nuevas), len(borrar)))

    if fila_app:
        from inventario_manager import registrar_uso
        r = registrar_uso(INSUMO_BUENO, LITROS, "AVELLANOS", "", QUIEN,
                          "Uso reingresado: quedó como 'Insumo no especificado'",
                          "2026-09-11", "L")
        print("Inventario: %s queda en %s L" % (INSUMO_BUENO, r.get("stock_restante")))
    return 0


if __name__ == "__main__":
    sys.exit(main("--simular" in sys.argv))
