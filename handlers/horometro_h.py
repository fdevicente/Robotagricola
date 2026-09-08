# -*- coding: utf-8 -*-
"""La capa de Telegram del flujo de horometro. La logica esta en horometro.py."""
import asyncio
import logging

from telegram import ReplyKeyboardMarkup

from handlers.horometro import PASOS, avanzar, iniciar
from handlers.teclado import BOTON_HOROMETRO, teclado_capataz, texto_de_ayuda

logger = logging.getLogger(__name__)


def _menu(opciones):
    """Teclado de una columna con las opciones, mas Otra..."""
    filas = [[o] for o in opciones] + [["Otra…"]]
    return ReplyKeyboardMarkup(filas, resize_keyboard=True, is_persistent=True)


async def atender_boton(update, context, boton):
    """Responde a un boton ya reconocido por teclado.preparar()."""
    ayuda = texto_de_ayuda(boton)
    if ayuda:                                   # Asistencia y Factura solo instruyen
        await update.message.reply_text(ayuda, parse_mode="Markdown",
                                        reply_markup=teclado_capataz())
        return
    if boton != BOTON_HOROMETRO:
        return
    from modules.opciones_capataz import maquinas_para_botones
    from modules.parte_contexto import construir
    ctx = await asyncio.to_thread(construir)
    context.user_data["horo_ctx"] = ctx
    iniciar(context.user_data)
    await update.message.reply_text(
        "🚜 ¿Qué *máquina*?", parse_mode="Markdown",
        reply_markup=_menu(await asyncio.to_thread(maquinas_para_botones)))


async def atender_paso(update, context) -> bool:
    ctx = context.user_data.get("horo_ctx") or {}
    r = avanzar(context.user_data, update.message.text, ctx)

    if not r["ok"]:
        await update.message.reply_text(r["mensaje"])
        return True

    if r["campos"] is None:
        teclado = teclado_capataz()
        if r.get("opciones"):
            # Un paso que pregunta trae sus propias respuestas. Sin "Otra…":
            # aca las opciones son todas las que hay.
            teclado = ReplyKeyboardMarkup([[o] for o in r["opciones"]],
                                          resize_keyboard=True, is_persistent=True)
        elif context.user_data.get("horo_state") == PASOS.LABOR:
            from modules.opciones_capataz import labores_frecuentes
            teclado = _menu(await asyncio.to_thread(labores_frecuentes))
        await update.message.reply_text(r["mensaje"], parse_mode="Markdown",
                                        reply_markup=teclado)
        return True

    from bitacora_manager import registrar_bitacora_estructurada
    quien = update.effective_user.full_name if update.effective_user else ""
    try:
        res = await asyncio.to_thread(registrar_bitacora_estructurada,
                                      r["campos"], quien)
    except Exception as e:
        logger.error("Horómetro guiado: no pude guardar: %s", e)
        await update.message.reply_text("❌ No pude guardarla: %s" % str(e)[:120],
                                        reply_markup=teclado_capataz())
        return True

    if isinstance(res, dict) and res.get("error_odometro"):
        await update.message.reply_text("🤔 No la guardé: %s" % res["error_odometro"],
                                        reply_markup=teclado_capataz())
        return True

    horas = res.get("horas_dia") if isinstance(res, dict) else None
    msg = "✅ *%s* — %g" % (r["campos"]["maquina"], r["campos"]["odometro"])
    if horas is not None:
        msg += "\n🕐 *%g h* desde la lectura anterior" % horas
    await update.message.reply_text(msg, parse_mode="Markdown",
                                    reply_markup=teclado_capataz())
    return True
