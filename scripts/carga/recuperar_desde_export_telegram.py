# -*- coding: utf-8 -*-
"""Reingresa partes de asistencia que se perdieron, leidos del export de Telegram.

POR QUE EXISTE
El respaldo crudo (files/telegram/*.jsonl) solo empieza el 26-ago-2026. El dueno
aporto el 7-sep un export HTML del chat de Juan que cubre desde el 9-jun, con 344
mensajes suyos. Cruzandolo contra la hoja Bitacora aparecieron 79 jornadas-hombre
sin anotar en 10 dias, incluidos DOS que no se conocian:

  - 2026-08-14, fuera de la ventana del flujo trabado que ya se habia diagnosticado.
  - Un parte rotulado "Martes 28 julio 2028": Juan tecleo mal el ano. La fecha no
    existia, asi que el parte nunca se anoto en ninguna parte. Se envio el 28-jul
    de 2026 (y de nuevo el 5-ago), asi que la fecha real es 2026-07-28.

QUE REINGRESA Y QUE NO
Solo los partes que `parsear_asistencia` lee COMPLETOS, comprobando que la
cantidad de personas leidas iguale la de lineas del mensaje. Los otros 6 dias
(25 a 28-ago, 31-ago y 1-sep) vienen sin dos puntos y el parser los lee a medias
o no los lee: esos esperan el plan de lectura con IA. Antes que escribir un parte
cojo en el Master, se deja pendiente.

USO
    python scripts/carga/recuperar_desde_export_telegram.py --simular
    python scripts/carga/recuperar_desde_export_telegram.py
"""
import html
import os
import re
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, RAIZ)

EXPORT = os.path.join(RAIZ, "files", "telegram", "export_juan_20260907.html")
QUIEN = "Juan Parada"

# Dia de trabajo declarado en el parte -> fecha con la que se guarda.
# Solo los que `parsear_asistencia` lee ENTEROS.
DIAS = {
    "2026-08-14": "2026-08-14",
    # Los seis que venian sin dos puntos. Se pudieron leer desde que
    # parsear_asistencia aprendio a partir la linea por el nombre del trabajador
    # y, para la cuadrilla de temporada, por las actividades del propio mensaje.
    "2026-08-25": "2026-08-25",
    "2026-08-26": "2026-08-26",
    "2026-08-27": "2026-08-27",
    "2026-08-28": "2026-08-28",
    "2026-08-31": "2026-08-31",
    "2026-09-01": "2026-09-01",
}

# ⚠️ NO agregar aqui "2028-07-28". Ese parte trae el ano tecleado mal, pero el
# dia REAL (2026-07-28) ya estaba anotado correctamente desde julio, con sus 6
# jornadas. Se detecto comparando contra la fecha inexistente --que logicamente
# daba cero filas-- en vez de contra la fecha corregida, y reingresarlo duplico
# dos filas en el Master el 7-sep-2026. Por eso existe el guard de abajo.

_BLOQUE = re.compile(
    r'<div class="message default clearfix(?: joined)?"[^>]*>(?P<cuerpo>.*?)'
    r'(?=<div class="message |\Z)', re.S)
_FECHA = re.compile(r'class="pull_right date details"\s+title="([^"]+)"')
_AUTOR = re.compile(r'class="from_name">\s*(.*?)\s*<', re.S)
_TEXTO = re.compile(r'<div class="text">(.*?)</div>', re.S)
_APLICACION = re.compile(r"^\s*aplicaci[oó]n\s+(herbicida|fungicida|insecticida|foliar)",
                         re.I)


def _limpiar(bruto):
    t = re.sub(r"<br\s*/?>", "\n", bruto)
    return html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def _mensajes_de_juan():
    with open(EXPORT, encoding="utf-8") as fh:
        crudo = fh.read()
    out, autor = [], ""
    for m in _BLOQUE.finditer(crudo):
        cuerpo = m.group("cuerpo")
        a = _AUTOR.search(cuerpo)
        autor = _limpiar(a.group(1)) if a else autor
        t = _TEXTO.search(cuerpo)
        if autor == QUIEN and t:
            texto = _limpiar(t.group(1))
            if "\n" in texto:
                out.append((_FECHA.search(cuerpo).group(1) if _FECHA.search(cuerpo)
                            else "", texto))
    return out


def _fecha_declarada(texto):
    from modules.bitacora_asistencia import fecha_de_linea
    for linea in texto.splitlines()[:3]:
        f = fecha_de_linea(linea)
        if f:
            return f.strftime("%Y-%m-%d")
    return None


def _lineas_persona(texto):
    n = 0
    for linea in texto.splitlines()[1:]:
        s = linea.strip()
        if not s or len(s.split()) < 2:
            continue
        if re.match(r"^(horometro|horas|total|labor|equipo|sector|trabajo)", s.lower()):
            continue
        n += 1
    return n


def _filas(dia_declarado, fecha_real, texto):
    """Filas de bitacora del parte, o (None, motivo) si no se lee entero."""
    from modules.bitacora_asistencia import (cultivo_de, parsear_asistencia_multi,
                                             tipo_de)
    dias = parsear_asistencia_multi(texto)
    if not dias:
        return None, "el parser no lo lee (las líneas van sin ':')"
    filas, personas = [], 0
    for d in dias:
        for g in d["grupos"]:
            personas += len(g["trabajadores"])
            filas.append({
                "fecha": fecha_real, "tipo": tipo_de(g["actividad"]),
                "actividad": g["actividad"], "cultivo": cultivo_de(g["actividad"]),
                "sector": "", "jornadas_hombre": g["jornadas_hombre"],
                "trabajadores": g["trabajadores"], "insumo": "", "cantidad": None,
                "unidad": "", "maquina": "", "odometro": None, "superficie_ha": None,
                "texto_original": texto if dia_declarado == fecha_real else
                                  ("[fecha corregida de %s a %s] %s"
                                   % (dia_declarado, fecha_real, texto)),
            })
    esperadas = _lineas_persona(texto)
    if personas < esperadas:
        return None, "se leyeron %d de %d líneas" % (personas, esperadas)
    return filas, None


def _ya_estan(filas):
    """Las que ya existen en el Master, comparando fecha + actividad + gente.

    Existe porque el 7-sep-2026 se duplicaron dos filas: el parte venia rotulado
    con un ano mal tecleado, se comparo contra esa fecha inexistente --que daba
    cero-- y el dia real ya estaba anotado. Comparar contra la fecha CON LA QUE
    SE VA A ESCRIBIR es lo unico que lo caza.
    """
    from openpyxl import load_workbook

    from config import EXCEL_PATH
    existentes = set()
    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    try:
        for row in wb["Bitácora"].iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            existentes.add((str(row[0])[:10], str(row[3] or ""), str(row[7] or "")))
    finally:
        wb.close()
    return [f for f in filas
            if (f["fecha"], f["actividad"], ", ".join(f["trabajadores"])) in existentes]


def main(simular):
    mensajes = _mensajes_de_juan()
    filas, pendientes = [], []
    for declarado, real in sorted(DIAS.items(), key=lambda kv: kv[1]):
        cands = [t for _, t in mensajes
                 if _fecha_declarada(t) == declarado and not _APLICACION.match(t)
                 and "horometro" not in t.lower()]
        if not cands:
            pendientes.append((declarado, "no encontré el parte en el export"))
            continue
        texto = max(cands, key=_lineas_persona)     # el más completo si se repite
        f, motivo = _filas(declarado, real, texto)
        if motivo:
            pendientes.append((declarado, motivo))
        else:
            filas.extend(f)

    print("%d filas a reingresar" % len(filas))
    print()
    for f in filas:
        print("  %s  %-11s %-34s %2s JH  %s"
              % (f["fecha"], f["tipo"], f["actividad"][:34],
                 f["jornadas_hombre"] or 0, ", ".join(f["trabajadores"])))
    if pendientes:
        print()
        print("PENDIENTES (siguen enteros en el export):")
        for dia, motivo in pendientes:
            print("  %s  %s" % (dia, motivo))

    repetidas = _ya_estan(filas) if filas else []
    if repetidas:
        print()
        print("YA ESTABAN, se saltan: %d fila(s)" % len(repetidas))
        for f in repetidas:
            print("   %s  %s  %s" % (f["fecha"], f["actividad"],
                                     ", ".join(f["trabajadores"])))
        clave = {(f["fecha"], f["actividad"], ", ".join(f["trabajadores"]))
                 for f in repetidas}
        filas = [f for f in filas
                 if (f["fecha"], f["actividad"], ", ".join(f["trabajadores"]))
                 not in clave]
        print()
        print("quedan %d filas nuevas" % len(filas))
    if not filas:
        print()
        print("nada nuevo que escribir.")
        return

    if simular:
        print()
        print("--simular: no se escribió nada.")
        return

    from bitacora_manager import registrar_bitacora_estructurada
    ok = 0
    print()
    for f in filas:
        registrar_bitacora_estructurada(f, QUIEN)
        ok += 1
    print("%d filas escritas" % ok)


if __name__ == "__main__":
    main("--simular" in sys.argv)
