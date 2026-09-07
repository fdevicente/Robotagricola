# -*- coding: utf-8 -*-
"""Flujo guiado para anotar una lectura de horometro.

POR QUE GUIADO Y NO LEIDO POR IA
Son tres datos y la maquina sale de una lista cerrada, asi que no se puede
inventar. Y el numero de inicio se contrasta contra la ultima lectura guardada
EN EL MOMENTO, con Juan todavia frente a la maquina. Hoy un error de tipeo entra
callado y descuadra las horas de todas las lecturas siguientes de esa maquina.

Este modulo es PURO: no toca Telegram ni Excel. Recibe el user_data y el texto,
y devuelve que responder. Asi se prueba entero en memoria.
"""
import logging
from datetime import date

logger = logging.getLogger(__name__)


class PASOS:
    MAQUINA = "esperando_maquina"
    INICIO = "esperando_inicio"
    TERMINO = "esperando_termino"
    LABOR = "esperando_labor"


def iniciar(user_data) -> None:
    user_data["horo_state"] = PASOS.MAQUINA
    user_data["horo_data"] = {}


def _cerrar(user_data) -> None:
    user_data["horo_state"] = None
    user_data["horo_data"] = None


def _numero(texto):
    """El numero que escribio Juan, o None. Acepta '5.239' y '5239'."""
    limpio = str(texto or "").strip().replace(".", "").replace(",", ".")
    try:
        return float(limpio)
    except ValueError:
        return None


def _ultima_de(maquina, ctx):
    for m in (ctx.get("maquinas") or []):
        if str(m.get("maquina", "")).upper() == str(maquina).upper():
            return m.get("ultimo_odometro")
    return None


def _miles(n) -> str:
    return f"{float(n):,.0f}".replace(",", ".")


def revisar_inicio(inicio, ultima) -> str | None:
    """Devuelve el aviso si el inicio no calza con la ultima lectura, o None.

    Esta es la razon de ser del flujo: cazar el tipeo mientras Juan mira la
    maquina, no tres dias despues cuando ya nadie se acuerda.
    """
    if ultima is None:
        return None                            # nunca se leyo: nada con que comparar
    if float(inicio) == float(ultima):
        return None
    return ("⚠️ La última que tengo es *%s*, y me pusiste *%s*.\n¿Está bien?"
            % (_miles(ultima), _miles(inicio)))


def avanzar(user_data, texto, ctx) -> dict:
    """Procesa un paso. Devuelve {"ok", "mensaje", "aviso", "campos"}.

    `campos` solo viene en el ultimo paso, listo para
    registrar_bitacora_estructurada.
    """
    paso = user_data.get("horo_state")
    datos = user_data.setdefault("horo_data", {})
    texto = str(texto or "").strip()

    if paso == PASOS.MAQUINA:
        conocidas = {str(m["maquina"]).upper(): m["maquina"]
                     for m in (ctx.get("maquinas") or [])}
        if texto.upper() not in conocidas:
            return {"ok": False, "campos": None,
                    "mensaje": "No conozco esa máquina. Elige una de la lista."}
        datos["maquina"] = conocidas[texto.upper()]
        user_data["horo_state"] = PASOS.INICIO
        ultima = _ultima_de(datos["maquina"], ctx)
        datos["ultima"] = ultima
        pista = ("La última que tengo es *%s*.\n" % _miles(ultima)
                 if ultima is not None else "")
        return {"ok": True, "campos": None,
                "mensaje": pista + "¿En cuánto *partió* hoy?"}

    if paso == PASOS.INICIO:
        n = _numero(texto)
        if n is None:
            return {"ok": False, "campos": None,
                    "mensaje": "Necesito el número del horómetro."}
        datos["inicio"] = n
        user_data["horo_state"] = PASOS.TERMINO
        return {"ok": True, "campos": None, "mensaje": "¿Y en cuánto *terminó*?",
                "aviso": revisar_inicio(n, datos.get("ultima"))}

    if paso == PASOS.TERMINO:
        n = _numero(texto)
        if n is None:
            return {"ok": False, "campos": None,
                    "mensaje": "Necesito el número del horómetro."}
        if n < datos.get("inicio", 0):
            return {"ok": False, "campos": None,
                    "mensaje": "El término no puede ser menor que el inicio "
                               "(%g). ¿Cuál es?" % datos["inicio"]}
        datos["termino"] = n
        user_data["horo_state"] = PASOS.LABOR
        return {"ok": True, "campos": None, "mensaje": "¿Qué *labor*?"}

    if paso == PASOS.LABOR:
        campos = {
            "fecha": date.today().strftime("%Y-%m-%d"),
            "tipo": "MAQUINARIA", "actividad": texto or "Lectura de horómetro",
            "cultivo": "GENERAL", "sector": "", "jornadas_hombre": None,
            "trabajadores": [], "insumo": "", "cantidad": None, "unidad": "",
            "maquina": datos.get("maquina", ""), "odometro": datos.get("termino"),
            "superficie_ha": None,
            "texto_original": "Horómetro guiado: %s %g → %g · %s"
                              % (datos.get("maquina", ""), datos.get("inicio", 0),
                                 datos.get("termino", 0), texto),
        }
        _cerrar(user_data)
        return {"ok": True, "mensaje": "", "campos": campos}

    return {"ok": False, "mensaje": "", "campos": None}
