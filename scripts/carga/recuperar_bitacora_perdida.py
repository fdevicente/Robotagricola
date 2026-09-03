# -*- coding: utf-8 -*-
"""Reingresa a la bitacora los partes de Juan que un flujo trabado se comio.

QUE PASO
Del 28-ago al 2-sep-2026 el flujo /deposito de Juan quedo abierto (escribio
"/ cancelar" CON ESPACIO, que Telegram no marca como comando, asi que nunca
ejecuto cmd_cancelar). Como es el primero de la fila en handlers/chat.py, cada
texto suyo se leyo como "el monto del deposito", fallo, y corto ahi: no llegaba
ni a maquinaria ni a la bitacora, y no dejaba linea en el log. La hoja Bitacora
quedo congelada desde el 21-ago.

Los mensajes NO se perdieron: estan crudos en files/telegram/*.jsonl, que
justamente nacio el 25-ago para poder diagnosticar esto.

QUE HACE
Replica el mismo camino que habria seguido el bot:
  - parte de asistencia -> parsear_asistencia_multi() -> una fila por labor
  - lectura de horometro -> registrar_bitacora_estructurada() tipo MAQUINARIA
Los reingresa en ORDEN CRONOLOGICO por fecha de trabajo, que es lo que necesita
el calculo de horas por diferencia de odometro.

DIFERENCIA A PROPOSITO CON EL BOT: el bot le pone al horometro la fecha de HOY;
aca se usa la fecha que dice el propio texto ("Lunes 31 de agosto 2026"), porque
se esta reconstruyendo historia, no anotando lo del dia.

LO QUE NO REINGRESA
parsear_asistencia exige "Nombre : actividad" y descarta EN SILENCIO las lineas
sin dos puntos. Juan dejo de poner los dos puntos: hay partes donde se leerian
2 jornadas de 7. Esos NO se escriben — quedan pendientes en el respaldo crudo,
que es reversible; una fila coja en el Master no lo es.

USO
    python scripts/carga/recuperar_bitacora_perdida.py --simular
    python scripts/carga/recuperar_bitacora_perdida.py
"""
import os
import sys
from datetime import date

RAIZ = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, RAIZ)

# Los 14 partes que se comio el flujo trabado, con un prefijo para comprobar
# que el mensaje es el que creemos antes de escribir nada.
PERDIDOS = [
    (2933, "Asistencia lunes 24 de agosto"),
    (2934, "Asistencia Martes 25 de agosto"),
    (2992, "Asistencia miércoles 26 agosto"),
    (3000, "Asistencia jueves 27 de agosto"),
    (3003, "Asistencia viernes 28 de agosto"),
    (3106, "Lunes 31 de agosto 2026"),
    (3122, "Martes 1 de septiembre 2026"),
    (3130, "Lunes 31 de agosto 2026"),
    (3143, "Lunes 31 de agosto 2026"),
    (3146, "Lunes 31 de agosto 2026"),
    (3149, "Martes 1 de septiembre 2026"),
    (3152, "Martes 1 de septiembre 2026"),
    (3155, "Martes 1 de septiembre 2026"),
    (3158, "1 de septiembre 2026"),
]
QUIEN = "Juan Parada"


def _mensajes():
    from modules.telegram_backup import leer_mes
    por_id = {}
    for mes in ("2026-08", "2026-09"):
        for fila in leer_mes(mes):
            if fila.get("text"):
                por_id[fila.get("message_id")] = fila
    salida = []
    for mid, prefijo in PERDIDOS:
        fila = por_id.get(mid)
        if fila is None:
            raise SystemExit("No encontre el mensaje %s en el respaldo crudo." % mid)
        if not fila["text"].startswith(prefijo):
            raise SystemExit("El mensaje %s no es el esperado: %r"
                             % (mid, fila["text"][:60]))
        salida.append(fila)
    return salida


def _fecha_del_texto(texto):
    """Fecha de trabajo declarada en la primera linea del parte."""
    from modules.bitacora_asistencia import fecha_de_linea
    for linea in texto.splitlines()[:2]:
        f = fecha_de_linea(linea)
        if f:
            return f
    return None


def _filas_de(fila):
    """Traduce un mensaje crudo a filas de bitacora.

    Devuelve (filas, motivo). Si `motivo` viene, el parte NO se reingresa.
    """
    from modules.bitacora_asistencia import (parsear_asistencia_multi,
                                             cultivo_de, tipo_de)
    from modules.maquinaria import (detectar_maquina, extraer_odometro,
                                    maquinas_conocidas)
    texto = fila["text"]
    fecha_txt = _fecha_del_texto(texto)
    lineas = [l for l in texto.splitlines()[1:] if l.strip()]

    dias = parsear_asistencia_multi(texto)
    if dias:
        filas, personas = [], 0
        for d in dias:
            f = d["fecha"] or fecha_txt
            for g in d["grupos"]:
                personas += len(g["trabajadores"])
                filas.append({
                    "fecha": f.strftime("%Y-%m-%d") if f else "",
                    "tipo": tipo_de(g["actividad"]), "actividad": g["actividad"],
                    "cultivo": cultivo_de(g["actividad"]), "sector": "",
                    "jornadas_hombre": g["jornadas_hombre"],
                    "trabajadores": g["trabajadores"], "insumo": "",
                    "cantidad": None, "unidad": "", "maquina": "",
                    "odometro": None, "superficie_ha": None,
                    "texto_original": texto,
                })
        if personas < len(lineas):
            return [], ("se leyeron %d de %d lineas (las otras van sin ':')"
                        % (personas, len(lineas)))
        return filas, None

    maquina = detectar_maquina(texto, maquinas_conocidas())
    odo = extraer_odometro(texto)
    if maquina and odo is not None:
        return [{
            "fecha": (fecha_txt or date.today()).strftime("%Y-%m-%d"),
            "tipo": "MAQUINARIA", "actividad": "Lectura de horómetro",
            "cultivo": "GENERAL", "sector": "", "jornadas_hombre": None,
            "trabajadores": [], "insumo": "", "cantidad": None, "unidad": "",
            "maquina": maquina, "odometro": odo, "superficie_ha": None,
            "texto_original": texto[:200],
        }], None

    return [], "ninguna linea trae ':' — el parser de asistencia no lo lee"


def main(simular):
    filas, pendientes = [], []
    for fila in _mensajes():
        f, motivo = _filas_de(fila)
        if motivo:
            pendientes.append((fila, motivo))
        filas.extend(f)
    filas.sort(key=lambda f: (f["fecha"], f["tipo"] != "MAQUINARIA"))

    print("%d filas a reingresar" % len(filas))
    print()
    for f in filas:
        extra = ("  %s odo=%g" % (f["maquina"], f["odometro"]) if f["maquina"]
                 else "  %s JH  %s" % (f["jornadas_hombre"] or 0,
                                       ", ".join(f["trabajadores"])))
        print("  %s  %-10s %-38s%s" % (f["fecha"], f["tipo"],
                                       f["actividad"][:38], extra))

    if pendientes:
        print()
        print("PENDIENTES — %d partes que NO se reingresan "
              "(siguen enteros en el respaldo crudo):" % len(pendientes))
        for fila, motivo in pendientes:
            print("  mid=%s  %-42s %s" % (fila["message_id"],
                                          fila["text"].splitlines()[0][:42], motivo))

    if simular:
        print()
        print("--simular: no se escribio nada.")
        return

    from bitacora_manager import registrar_bitacora_estructurada
    ok = rechazadas = 0
    print()
    for f in filas:
        res = registrar_bitacora_estructurada(f, QUIEN)
        if isinstance(res, dict) and res.get("error_odometro"):
            rechazadas += 1
            print("  RECHAZADA %s %s: %s" % (f["fecha"], f["maquina"],
                                             res["error_odometro"]))
        else:
            ok += 1
    print()
    print("%d filas escritas%s" % (ok, ", %d rechazadas" % rechazadas if rechazadas else ""))


if __name__ == "__main__":
    main("--simular" in sys.argv)
