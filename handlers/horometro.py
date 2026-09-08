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
import unicodedata
from datetime import date

logger = logging.getLogger(__name__)


class PASOS:
    MAQUINA = "esperando_maquina"
    INICIO = "esperando_inicio"
    CONFIRMA = "esperando_confirma"
    TERMINO = "esperando_termino"
    LABOR = "esperando_labor"


# El aviso del inicio raro es una PREGUNTA y hay que esperar la respuesta.
# Medido en el telefono el 8-sep-2026: se mandaba el aviso y en el mismo aliento
# la pregunta del termino, porque el paso ya habia avanzado. Asi, "no" caia en el
# termino y le contestaba "Necesito el numero del horometro" --lo retaba por
# responder lo que le acababan de preguntar-- y el 1.950 mal tecleado se quedaba
# sin forma de corregirse. Misma forma que el "/ cancelar" que nunca le funciono.
BOTON_SI = "✅ Sí, está bien"
BOTON_NO = "✏️ No, lo corrijo"


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


def _clave_texto(texto) -> str:
    """Minusculas, sin tildes y con los espacios colapsados.

    OJO: el filtro de marcas se lleva tambien el selector de variacion (U+FE0F)
    que arrastran emoji como ✏️, asi que los botones NO se pueden comparar
    contra un literal tipeado a mano. Por eso las constantes pasan por aca
    tambien: los dos lados se normalizan igual y da lo mismo como se escriban.
    """
    t = "".join(c for c in unicodedata.normalize("NFD", str(texto or "").lower())
                if unicodedata.category(c) != "Mn")
    return " ".join(t.split()).strip(" .,;:!¡?¿")


def _si_o_no(texto):
    """True si dijo que si, False si dijo que no, None si no se entiende.

    Se aceptan los dos botones y el si/no escrito a mano, con o sin tilde: Juan
    escribe tanto como aprieta. Lo que no se entiende NO se adivina, se vuelve a
    preguntar; adivinar aca es dar por bueno un horometro que nadie confirmo.
    """
    t = _clave_texto(texto)
    if t in (_clave_texto(BOTON_SI), "si", "si esta bien", "esta bien"):
        return True
    if t in (_clave_texto(BOTON_NO), "no", "no lo corrijo"):
        return False
    return None


def _pedir_inicio(ultima) -> str:
    pista = ("La última que tengo es *%s*.\n" % _miles(ultima)
             if ultima is not None else "")
    return pista + "¿En cuánto *partió* hoy?"


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


def _anotar_inicio(user_data, datos, n) -> dict:
    """Guarda el inicio y decide si hay que preguntar antes de seguir."""
    datos["inicio"] = n
    aviso = revisar_inicio(n, datos.get("ultima"))
    if aviso:
        user_data["horo_state"] = PASOS.CONFIRMA
        return {"ok": True, "campos": None, "mensaje": aviso,
                "opciones": [BOTON_SI, BOTON_NO]}
    user_data["horo_state"] = PASOS.TERMINO
    return {"ok": True, "campos": None, "mensaje": "¿Y en cuánto *terminó*?"}


def avanzar(user_data, texto, ctx) -> dict:
    """Procesa un paso. Devuelve {"ok", "mensaje", "opciones", "campos"}.

    `opciones` viene cuando el paso PREGUNTA algo con respuestas cerradas, y la
    capa de Telegram las pinta como teclado.

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
        return {"ok": True, "campos": None, "mensaje": _pedir_inicio(ultima)}

    if paso == PASOS.INICIO:
        n = _numero(texto)
        if n is None:
            return {"ok": False, "campos": None,
                    "mensaje": "Necesito el número del horómetro."}
        return _anotar_inicio(user_data, datos, n)

    if paso == PASOS.CONFIRMA:
        dijo = _si_o_no(texto)
        if dijo is True:                        # el numero raro era el bueno
            user_data["horo_state"] = PASOS.TERMINO
            return {"ok": True, "campos": None, "mensaje": "¿Y en cuánto *terminó*?"}
        if dijo is False:
            user_data["horo_state"] = PASOS.INICIO
            return {"ok": True, "campos": None,
                    "mensaje": _pedir_inicio(datos.get("ultima"))}
        n = _numero(texto)
        if n is not None:
            # Contesto con un numero: esta arreglando el inicio que ve citado, no
            # adelantando el termino. Tomarlo como termino escribiria un dato que
            # nadie pidio; tomarlo como inicio se puede volver a corregir.
            return _anotar_inicio(user_data, datos, n)
        # Sin Markdown a proposito: la rama de "no ok" manda el texto crudo.
        return {"ok": False, "campos": None,
                "mensaje": "No te entendí. Apreta el botón de arriba: Sí si el "
                           "número está bien, No para corregirlo."}

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
