# -*- coding: utf-8 -*-
"""Qué hacer cuando algo revienta adentro de un handler o un job.

Sin esto, `python-telegram-bot` escribe *"No error handlers are registered"* en
el log y sigue como si nada: la excepción muere ahí y **nadie se entera**. Pasó
en producción el 1-sep-2026 a las 10:22 con un timeout de Drive, y es la
explicación más probable del parte de Juan que se perdió en silencio el 24-ago.

Dos reglas que no se negocian:

1. **Nunca puede lanzar.** Un error handler que falla deja al bot peor que sin
   error handler: se pierde el error original *y* el aviso.
2. **Nunca puede spamear.** El error que lo destapó se repite cada 15 minutos.
   Sin ventana serían ~96 mensajes al día por una sola falla de red, que es la
   forma más rápida de que el dueño silencie el bot y deje de leerlo.
"""
import logging
import time
import traceback

logger = logging.getLogger(__name__)

# Un mismo error se avisa una vez por ventana. 6 h: suficiente para no spamear
# con una falla que se repite sola, corto para que si sigue mañana se vuelva a
# ver — que persista un día entero es información nueva, no ruido.
VENTANA_AVISO = 6 * 60 * 60

_PREFIJO = "_error_visto_"


def _identidad(error) -> str:
    """Qué error es y DÓNDE se produjo.

    El tipo solo no alcanza: dos `TimeoutError` en dos jobs distintos son dos
    problemas distintos y los dos merecen su aviso.
    """
    tb = getattr(error, "__traceback__", None)
    ultimo = None
    while tb is not None:
        ultimo = tb
        tb = tb.tb_next
    if ultimo is None:
        return type(error).__name__
    marco = ultimo.tb_frame.f_code
    return "%s@%s:%s" % (type(error).__name__, marco.co_filename,
                         ultimo.tb_lineno)


def _donde(error) -> str:
    """`archivo.py:123, en nombre_de_la_funcion`, para el mensaje."""
    import os
    tb = getattr(error, "__traceback__", None)
    ultimo = None
    while tb is not None:
        ultimo = tb
        tb = tb.tb_next
    if ultimo is None:
        return "origen desconocido"
    marco = ultimo.tb_frame.f_code
    return "%s:%s, en %s" % (os.path.basename(marco.co_filename),
                             ultimo.tb_lineno, marco.co_name)


def _ya_se_aviso(bot_data, clave) -> bool:
    """True si este mismo error ya se avisó dentro de la ventana."""
    try:
        antes = bot_data.get(_PREFIJO + clave)
        ahora = time.time()
        if antes is not None and (ahora - antes) < VENTANA_AVISO:
            return True
        bot_data[_PREFIJO + clave] = ahora
        return False
    except Exception:                       # bot_data raro: mejor avisar
        return False


async def manejar_error(update, context) -> None:
    """Loguea el traceback completo y le manda al dueño un resumen corto."""
    try:
        error = getattr(context, "error", None)
        if error is None:
            return

        logger.error("Excepción no atrapada: %s\n%s", error,
                     "".join(traceback.format_exception(
                         type(error), error, error.__traceback__)))

        if _ya_se_aviso(context.bot_data, _identidad(error)):
            return                          # ya avisado, no spamear

        chat_id = (context.bot_data.get("owner_chat_id")
                   or context.bot_data.get("banco_chat_id"))
        if not chat_id:
            from config import TELEGRAM_CHAT_ID
            chat_id = TELEGRAM_CHAT_ID
        if not chat_id:
            return                          # no hay a quién avisarle

        texto = ("⚠️ Algo falló adentro del bot\n\n"
                 "%s: %s\n"
                 "Dónde: %s\n\n"
                 "El bot sigue andando. El detalle completo está en bot.log.\n"
                 "Si se repite, no te voy a avisar de nuevo por %d horas."
                 % (type(error).__name__, error, _donde(error),
                    VENTANA_AVISO // 3600))
        # Sin Markdown a propósito: el texto del error puede traer guiones
        # bajos o asteriscos y romper el formato, y ahí se pierde el aviso.
        await context.bot.send_message(chat_id=int(chat_id), text=texto)
    except Exception as e:                  # noqa: BLE001
        # Acá se termina el camino: si el aviso falla, queda en el log y punto.
        logger.warning("El error handler no pudo avisar: %s", e)
