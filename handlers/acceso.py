# -*- coding: utf-8 -*-
"""Quien puede hablarle al bot.

🔴 POR QUE EXISTE
El 12-sep-2026 aparecio un usuario llamado "prince" (user_id 6934038077),
escribio /start y el bot le respondio con el MENU COMPLETO: /saldo,
/inventario, /reporte, /personal, /dashboard... y tambien los que ESCRIBEN,
/deposito, /pagado, /bitacora, /uso y subir facturas.

No habia ningun control de acceso. Los unicos chequeos por user_id que existian
--AUTO_SAVE_USERS-- decidian si mostrar el teclado del capataz, no si la persona
tenia permiso. Un bot de Telegram es publico: cualquiera que sepa su nombre le
puede escribir, y este maneja la plata y los datos del campo.

COMO FUNCIONA
Un solo guardia, en un grupo de handlers ANTERIOR a todos los demas, que levanta
ApplicationHandlerStop. Se eligio asi y no un filtro por handler porque hay 31
comandos: uno nuevo se agrega sin acordarse del filtro, y el agujero vuelve.

Va DESPUES del espejo (mirror_update) a proposito: el intento queda guardado en
el respaldo crudo aunque se rechace. Sin eso, este incidente no se habria podido
reconstruir.
"""
import logging

logger = logging.getLogger(__name__)

# Un aviso por desconocido, no uno por mensaje: si alguien insiste, no se llena
# el chat del dueño.
_AVISADOS = "acceso_avisados"


def leer_autorizados(crudo) -> set:
    """Los user_id de una cadena tipo "123,456". Lo que no sea numero se ignora."""
    partes = str(crudo or "").replace(";", ",").split(",")
    return {int(p.strip()) for p in partes if p.strip().isdigit()}


def autorizado(user_id, autorizados) -> bool:
    """Si ese usuario puede usar el bot.

    ⚠️ Una lista vacia NO autoriza a nadie. Si la configuracion se rompe, el bot
    se cierra entero: es molesto, pero abrirse a todos no es una opcion.
    """
    if not autorizados or user_id in (None, ""):
        return False
    try:
        uid = int(user_id)
    except (TypeError, ValueError):
        return False
    return uid in {int(a) for a in autorizados}


async def guardia(update, context) -> None:
    """Corta todo si quien escribe no esta autorizado."""
    from telegram.ext import ApplicationHandlerStop

    from config import USUARIOS_AUTORIZADOS
    u = update.effective_user if update else None
    if autorizado(u.id if u else None, USUARIOS_AUTORIZADOS):
        return

    uid = getattr(u, "id", None)
    nombre = getattr(u, "full_name", None) or "?"
    logger.warning("Acceso denegado: %s (user_id %s)", nombre, uid)

    vistos = context.bot_data.setdefault(_AVISADOS, [])
    if uid not in vistos:
        vistos.append(uid)
        # Al dueño le importa saberlo la primera vez. Al desconocido se le dice
        # que no, sin detalles: no hace falta contarle que maneja.
        dueno = context.bot_data.get("owner_chat_id")
        if dueno:
            try:
                await context.bot.send_message(
                    chat_id=dueno,
                    text=("🔒 Alguien de fuera le escribió al bot y lo bloqueé.\n"
                          "*%s* · user id `%s`\n\n"
                          "Si es alguien tuyo, agrégalo a `TELEGRAM_USUARIOS` "
                          "en el `.env` y reinicia." % (nombre, uid)),
                    parse_mode="Markdown")
            except Exception as e:                       # pragma: no cover
                logger.warning("No pude avisarle al dueño: %r", e)
        if update and getattr(update, "message", None):
            try:
                await update.message.reply_text(
                    "Este bot es de uso interno y no tienes acceso.")
            except Exception as e:                       # pragma: no cover
                logger.warning("No pude responderle al desconocido: %r", e)

    raise ApplicationHandlerStop
