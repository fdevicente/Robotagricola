"""handlers/chat.py — Dispatcher central de mensajes de texto.

Revisa flujos conversacionales activos y delega al handler correcto.
Si no hay flujo, usa el chat inteligente.
"""
import logging

from chat_inteligente import responder_chat

logger = logging.getLogger(__name__)


async def handle_text(update, context):
    """Punto de entrada para TODOS los mensajes de texto (no comandos)."""
    # Imports diferidos para evitar import-time circularidad
    from handlers.finanzas import handle_text_deposito, handle_text_pagado
    from handlers.tareas import handle_text_tarea
    from handlers.bitacora import handle_text_bitacora
    from handlers.inventario_h import handle_text_uso
    from handlers.personal import handle_text_vacacion, handle_text_trabajador
    from handlers.vencimientos import handle_text_vencimiento
    from handlers.facturas import handle_text_edit_factura

    # ── Flujos activos (orden de prioridad) ──
    if await handle_text_deposito(update, context):
        return
    if await handle_text_pagado(update, context):
        return
    if await handle_text_tarea(update, context):
        return
    if await handle_text_bitacora(update, context):
        return
    if await handle_text_vencimiento(update, context):
        return
    if await handle_text_uso(update, context):
        return
    if await handle_text_vacacion(update, context):
        return
    if await handle_text_trabajador(update, context):
        return
    if await handle_text_edit_factura(update, context):
        return

    # ── Maquinaria: horómetros, mantenciones y fichas ──
    # Va ANTES de la bitácora automática: "al 5085 le cambiaron aceite" es una
    # mantención, no una labor del día.
    from handlers.maquinaria import (modo_activo, parece_maquinaria,
                                      procesar_texto_maquinaria)
    if modo_activo(context) or parece_maquinaria(update.message.text or ""):
        if await procesar_texto_maquinaria(update, context):
            return

    # ── Modo capataz: texto libre = bitácora automática (sin /bitacora ni confirmar) ──
    from config import AUTO_SAVE_USERS
    if update.effective_user and update.effective_user.id in AUTO_SAVE_USERS:
        from handlers.bitacora import auto_guardar_bitacora
        await auto_guardar_bitacora(update, context)
        return

    # ── Sin flujo activo → Chat inteligente ──
    texto = update.message.text.strip()
    if len(texto) > 2:
        msg = await update.message.reply_text("💬 Pensando...")
        respuesta = await responder_chat(texto)
        try:
            await msg.edit_text(respuesta, parse_mode="Markdown")
        except Exception:
            await msg.edit_text(respuesta)
