# -*- coding: utf-8 -*-
"""Reingresa las aplicaciones y las lecturas de horometro que se perdieron.

POR QUE EXISTE
Cruzando el export del chat de Juan (7-sep-2026) contra el Master aparecieron dos
agujeros que no se habian mirado:

  - APLICACIONES: mando 15 partes entre junio y agosto y la hoja `Aplicaciones`
    tiene 3 filas. Once aplicaciones de agroquimico sin registrar: que producto,
    en que dosis y en que cuartel.
  - MAQUINARIA: 3 dias (12, 13 y 18 de agosto) con lecturas de horometro que no
    quedaron en ninguna parte.

POR QUE NO SE USA registrar_uso()
Esa funcion sirve para el uso del dia, no para recuperar historia: DESCUENTA del
inventario --descontar hoy un consumo de junio dejaria el stock mal, porque el
stock actual ya se conto de verdad-- y ademas escribe la fecha de HOY en vez de
la del trabajo. Aca se escribe directo en la hoja, con la fecha correcta y sin
tocar el inventario.

LAS LECTURAS VAN CON forzar=True
Se estan insertando lecturas VIEJAS cuando la maquina ya tiene otras posteriores.
`validar_odometro` compara contra la ultima por fecha, que aca es mas nueva, asi
que rechazaria un numero mas chico por "salto imposible". Se fuerza para dejar la
lectura en el registro; las horas del dia quedan sin calcular a proposito, porque
restarlas contra una lectura posterior daria un numero sin sentido.

USO
    python scripts/carga/recuperar_aplicaciones_y_maquinaria.py --simular
    python scripts/carga/recuperar_aplicaciones_y_maquinaria.py
"""
import html
import os
import re
import sys
from datetime import datetime, timedelta

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, RAIZ)

EXPORT = os.path.join(RAIZ, "files", "telegram", "export_juan_20260907.html")
QUIEN = "Juan Parada"
APLICACIONES_SHEET = "Aplicaciones"

# El export viene renderizado en UTC+8 y el campo es Chile (UTC-4): 12 horas.
HORAS_DESFASE = 12

DIAS_HOROMETRO = ("2026-08-12", "2026-08-13", "2026-08-18")

_BLOQUE = re.compile(
    r'<div class="message default clearfix(?: joined)?"[^>]*>(?P<cuerpo>.*?)'
    r'(?=<div class="message |\Z)', re.S)
_FECHA = re.compile(r'class="pull_right date details"\s+title="([^"]+)"')
_AUTOR = re.compile(r'class="from_name">\s*(.*?)\s*<', re.S)
_TEXTO = re.compile(r'<div class="text">(.*?)</div>', re.S)


def _limpiar(bruto):
    return html.unescape(re.sub(r"<[^>]+>", "", re.sub(r"<br\s*/?>", "\n", bruto))).strip()


def _mensajes():
    """[(fecha_envio_chile, texto)] de Juan, solo los de varias lineas."""
    with open(EXPORT, encoding="utf-8") as fh:
        crudo = fh.read()
    out, autor = [], ""
    for m in _BLOQUE.finditer(crudo):
        cuerpo = m.group("cuerpo")
        a = _AUTOR.search(cuerpo)
        autor = _limpiar(a.group(1)) if a else autor
        t = _TEXTO.search(cuerpo)
        f = _FECHA.search(cuerpo)
        if autor != QUIEN or not t:
            continue
        texto = _limpiar(t.group(1))
        if "\n" not in texto:
            continue
        enviado = ""
        if f:
            try:
                d = datetime.strptime(f.group(1)[:19], "%d.%m.%Y %H:%M:%S")
                enviado = (d - timedelta(hours=HORAS_DESFASE)).strftime("%Y-%m-%d")
            except ValueError:
                pass
        out.append((enviado, texto))
    return out


def _fecha_del_texto(texto):
    from modules.bitacora_asistencia import fecha_de_linea
    for linea in texto.splitlines()[:3]:
        f = fecha_de_linea(linea)
        if f:
            return f.strftime("%Y-%m-%d")
    return None


# ── Aplicaciones ────────────────────────────────────────────────────────────

def _aplicaciones_a_escribir(mensajes):
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    from modules.aplicaciones_parser import es_aplicacion, parsear
    from modules.bitacora_asistencia import fecha_de_linea

    ya = set()
    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    try:
        if APLICACIONES_SHEET in wb.sheetnames:
            for row in wb[APLICACIONES_SHEET].iter_rows(min_row=2, values_only=True):
                if row and row[0]:
                    ya.add((str(row[0])[:10], str(row[1] or "").lower(),
                            float(row[2] or 0)))
    finally:
        wb.close()

    filas, sin_fecha = [], 0
    for enviado, texto in mensajes:
        if not es_aplicacion(texto):
            continue
        d = parsear(texto, fecha_de_linea)
        if d is None or d["cantidad"] is None:
            continue
        obs = d["texto_original"]
        if not d["fecha"]:
            d["fecha"] = enviado
            sin_fecha += 1
            obs = "[fecha tomada del envío, el texto no la traía] " + obs
        if not d["fecha"]:
            continue
        clave = (d["fecha"], d["producto"].lower(), float(d["cantidad"]))
        if clave in ya:
            continue
        ya.add(clave)
        filas.append({**d, "observaciones": obs})
    return filas, sin_fecha


def _escribir_aplicaciones(filas):
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    wb = load_workbook(EXCEL_PATH)
    ws = wb[APLICACIONES_SHEET]
    for f in filas:
        ws.append([f["fecha"], f["producto"], f["cantidad"], f["unidad"],
                   f["cultivo"], f["sector"], QUIEN, f["observaciones"][:600]])
    wb.save(EXCEL_PATH)
    wb.close()


# ── Horometros ──────────────────────────────────────────────────────────────

def _lecturas_a_escribir(mensajes):
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    from modules.maquinaria import (detectar_maquina, extraer_odometro,
                                    maquinas_conocidas)

    ya = set()
    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    try:
        for row in wb["Bitácora"].iter_rows(min_row=2, values_only=True):
            if row and row[0] and len(row) > 14 and row[14] not in (None, ""):
                ya.add((str(row[0])[:10], str(row[13] or "").upper(),
                        float(row[14])))
    finally:
        wb.close()

    conocidas = maquinas_conocidas()
    filas = []
    for _, texto in mensajes:
        bajo = texto.lower()
        if "horometro" not in bajo and "horómetro" not in bajo:
            continue
        fecha = _fecha_del_texto(texto)
        if fecha not in DIAS_HOROMETRO:
            continue
        maquina = detectar_maquina(texto, conocidas)
        odo = extraer_odometro(texto)
        if not maquina or odo is None:
            continue
        clave = (fecha, maquina.upper(), float(odo))
        if clave in ya:
            continue
        ya.add(clave)
        labor = ""
        for l in texto.splitlines():
            m = re.match(r"^\s*labor\s*:?\s*(.+)$", l, re.I)
            if m:
                labor = m.group(1).strip()
                break
        filas.append({
            "fecha": fecha, "tipo": "MAQUINARIA",
            "actividad": labor or "Lectura de horómetro", "cultivo": "GENERAL",
            "sector": "", "jornadas_hombre": None, "trabajadores": [],
            "insumo": "", "cantidad": None, "unidad": "", "maquina": maquina,
            "odometro": odo, "superficie_ha": None, "texto_original": texto,
        })
    return filas


def main(simular):
    mensajes = _mensajes()
    apls, sin_fecha = _aplicaciones_a_escribir(mensajes)
    lecturas = _lecturas_a_escribir(mensajes)

    print("APLICACIONES a escribir: %d" % len(apls))
    for f in apls:
        print("  %s  %-20s %7s %-3s %-10s %s"
              % (f["fecha"], f["producto"][:20], f["cantidad"], f["unidad"],
                 f["cultivo"], f["sector"][:24]))
    if sin_fecha:
        print("  (%d sin fecha en el texto: se tomó la del envío)" % sin_fecha)

    print()
    print("LECTURAS DE HORÓMETRO a escribir: %d" % len(lecturas))
    for f in lecturas:
        print("  %s  %-30s odo=%-8g %s"
              % (f["fecha"], f["maquina"], f["odometro"], f["actividad"][:28]))

    if simular:
        print()
        print("--simular: no se escribió nada.")
        return

    if apls:
        _escribir_aplicaciones(apls)
    from bitacora_manager import registrar_bitacora_estructurada
    for f in lecturas:
        registrar_bitacora_estructurada(f, QUIEN, forzar=True)
    print()
    print("%d aplicaciones y %d lecturas escritas" % (len(apls), len(lecturas)))


if __name__ == "__main__":
    main("--simular" in sys.argv)
