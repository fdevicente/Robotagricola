# -*- coding: utf-8 -*-
"""handlers/drive_jobs.py — Job que vacía la cola de Drive y comando /drive.

Comandos: /drive
Callbacks: cb_drive_reintentar (botón "Reintentar" de /drive)
Jobs: job_drive_cola (cada 10 min)

CONTEXTO: tras 5 fallos un documento pasa a "rendido" y el subidor deja de
mirarlo — queda a salvo en el PC pero sin forma de llegar a Drive.
`Cola.reintentar_rendidos()` los devuelve a pendientes, pero nadie lo dispara
solo: el comando /drive es lo que cierra ese ciclo.
"""
import asyncio
import logging

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

from config import (DRIVE_RAIZ, DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS,
                     DRIVE_UMBRAL_AVISO, TELEGRAM_CHAT_ID)
from modules.drive.auth import FaltaAutorizacion
from modules.drive.cola import Cola

logger = logging.getLogger(__name__)

_MARCA_AVISO_AUTORIZACION = "drive_aviso_autorizacion"
_MARCA_AVISO_CUOTA = "drive_aviso_cuota"


def hay_que_avisar_cuota(drive, umbral: float = 0.80) -> bool:
    """True si el uso de Drive alcanzó el umbral. Nunca revienta."""
    try:
        q = drive.cuota()
    except Exception as e:
        logger.warning("Drive: no pude leer la cuota (%s)", e)
        return False
    usado, total = q.get("usado", 0), q.get("total", 0)
    if not total:
        return False
    return (usado / total) >= umbral


def resumen_cola(cola) -> dict:
    """Cuenta pendientes y rendidos, y trae el último error para diagnosticar."""
    pendientes = cola.pendientes()
    rendidos = cola.rendidos()
    ultimo_error = ""
    for i in rendidos:
        if i.get("ultimo_error"):
            ultimo_error = i["ultimo_error"]
    return {"pendientes": len(pendientes), "rendidos": len(rendidos),
            "ultimo_error": ultimo_error}


def _raiz_id(drive, nombre_raiz: str) -> str:
    """ID de la carpeta raíz, creándola la primera vez."""
    cid = drive.buscar_carpeta(nombre_raiz, "root")
    return cid or drive.crear_carpeta(nombre_raiz, "root")


def _chat_destino(context) -> str:
    return (context.bot_data.get("owner_chat_id")
            or context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID)


async def job_drive_cola(context):
    """Vacía la cola de subidas a Drive. Corre cada 10 min.

    Si la cola está vacía (ni pendientes ni rendidos) no autentica: correrlo
    seguido tiene que ser barato, si no serían 144 autenticaciones al día
    para nada.
    """
    cola = Cola(DRIVE_COLA_PATH, max_intentos=DRIVE_MAX_INTENTOS)
    if not cola.pendientes() and not cola.rendidos():
        return

    chat_id = _chat_destino(context)

    try:
        from modules.drive.cliente import DriveCliente
        drive = await asyncio.to_thread(DriveCliente)
    except FaltaAutorizacion as e:
        if not context.bot_data.get(_MARCA_AVISO_AUTORIZACION):
            context.bot_data[_MARCA_AVISO_AUTORIZACION] = True
            if chat_id:
                await context.bot.send_message(
                    chat_id=int(chat_id),
                    text="⚠️ *Drive sin autorizar*\n\n%s" % e,
                    parse_mode="Markdown")
        return

    from modules.drive.carpetas import Carpetas
    from modules.drive.subidor import procesar_cola
    raiz_id = await asyncio.to_thread(_raiz_id, drive, DRIVE_RAIZ)
    carpetas = Carpetas(drive, raiz_id)
    await asyncio.to_thread(procesar_cola, cola, drive, carpetas)

    rendidos = cola.rendidos()
    if rendidos and chat_id:
        await context.bot.send_message(
            chat_id=int(chat_id),
            text=("⚠️ *Drive: %d subida(s) no se pudieron enviar*\n\n"
                  "Siguen guardadas en el PC, nada se perdió.\n"
                  "Usa /drive para reintentarlas." % len(rendidos)),
            parse_mode="Markdown")

    if hay_que_avisar_cuota(drive, umbral=DRIVE_UMBRAL_AVISO):
        if not context.bot_data.get(_MARCA_AVISO_CUOTA):
            context.bot_data[_MARCA_AVISO_CUOTA] = True
            if chat_id:
                await context.bot.send_message(
                    chat_id=int(chat_id),
                    text="⚠️ *Drive se está llenando* (pasó el %.0f%% de uso)."
                         % (DRIVE_UMBRAL_AVISO * 100),
                    parse_mode="Markdown")


async def cmd_drive(update, context):
    """Muestra cuántas subidas hay pendientes y rendidas, con opción a reintentar."""
    cola = Cola(DRIVE_COLA_PATH, max_intentos=DRIVE_MAX_INTENTOS)
    r = await asyncio.to_thread(resumen_cola, cola)

    if not r["pendientes"] and not r["rendidos"]:
        await update.message.reply_text(
            "✅ *Drive al día*\n\nNo hay subidas pendientes.",
            parse_mode="Markdown")
        return

    texto = "📁 *Drive*\n\n"
    texto += "⏳ Pendientes: *%d*\n" % r["pendientes"]
    texto += "🔴 Rendidos (agotaron reintentos): *%d*\n" % r["rendidos"]
    if r["ultimo_error"]:
        texto += "\nÚltimo error: `%s`\n" % r["ultimo_error"]

    if r["rendidos"]:
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔄 Reintentar", callback_data="drive_reintentar")]])
        await update.message.reply_text(texto, parse_mode="Markdown", reply_markup=kb)
    else:
        await update.message.reply_text(texto, parse_mode="Markdown")


async def cb_drive_reintentar(update, context):
    """Callback del botón de /drive: devuelve los rendidos a la cola."""
    query = update.callback_query
    await query.answer()
    cola = Cola(DRIVE_COLA_PATH, max_intentos=DRIVE_MAX_INTENTOS)
    n = await asyncio.to_thread(cola.reintentar_rendidos)
    if n:
        texto = ("🔄 *%d subida(s) devuelta(s) a la cola.*\n\n"
                  "Se van a intentar en la próxima pasada (cada 10 min)." % n)
    else:
        texto = "No había nada rendido para reintentar."
    try:
        await query.edit_message_text(texto, parse_mode="Markdown")
    except Exception:
        await query.edit_message_text(texto)
