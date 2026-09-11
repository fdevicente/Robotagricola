"""handlers/chat.py — Dispatcher central de mensajes de texto.

Revisa flujos conversacionales activos y delega al handler correcto.
Si no hay flujo, usa el chat inteligente.
"""
import logging

from chat_inteligente import responder_chat
from modules.flujos import (MINUTOS_VIDA, comando_con_espacio,
                            revisar_flujos)


logger = logging.getLogger(__name__)


def _comandos_registrados(context) -> set:
    """Los comandos que el bot tiene de verdad, sacados de PTB.

    Se leen de los handlers y no de una lista escrita a mano: una lista
    aparte envejece, y ofrecerle a Juan un comando que ya no existe es
    peor que no decirle nada.
    """
    from telegram.ext import CommandHandler
    nombres = set()
    try:
        for grupo in context.application.handlers.values():
            for h in grupo:
                if isinstance(h, CommandHandler):
                    nombres |= {str(c).lower() for c in h.commands}
    except Exception:                       # pragma: no cover
        pass
    return nombres


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

    # ── Los botones fijos se atienden antes que nada ──
    # "Asistencia" a secas cae en es_mensaje_sin_contenido y se trataria como
    # basura --con razon: Juan escribia la palabra sola antes del parte-- asi
    # que el boton no haria nada si no se mira primero. Y apretar un boton
    # cierra lo que haya quedado a medias: es la salida que le faltaba.
    from handlers.teclado import preparar
    _boton = preparar(context.user_data, update.message.text)
    if _boton:
        from handlers.botones import atender_boton
        await atender_boton(update, context, _boton)
        return

    # ── "/ comando" con espacio ──
    # Telegram solo marca como comando lo que va PEGADO a la barra, asi que
    # "/ uso" llega como texto normal. Va ANTES de repartir: si no, se lo come
    # el primer flujo abierto, que es justo lo que paso el 28-ago con
    # "/ cancelar" y el 10-sep con "/ uso".
    _cmd = comando_con_espacio(update.message.text, _comandos_registrados(context))
    if _cmd:
        if _cmd == "cancelar":
            from modules.flujos import limpiar_flujos
            limpiar_flujos(context.user_data)
            await update.message.reply_text(
                "🧹 Listo, cancelé lo que estaba a medias.\n"
                "Ojo: escribiste */ cancelar* con un espacio y Telegram no lo "
                "toma como comando. La próxima, */cancelar* pegado.",
                parse_mode="Markdown")
        else:
            await update.message.reply_text(
                "Escribiste */ %s* con un espacio y Telegram no lo toma como "
                "comando, así que no hice nada.\nEscribilo pegado: */%s*"
                % (_cmd, _cmd), parse_mode="Markdown")
        return

    # ── Flujos vencidos: se sueltan ANTES de repartir ──
    # Un flujo a medias se queda con todo el texto que llegue. El 28-ago-2026
    # un /deposito sin cerrar se comio 12 dias de partes de Juan en silencio.
    #
    # 🔴 Y VA ANTES DEL HOROMETRO, no despues. Cuando el flujo guiado se agrego
    # el 7-sep quedo POR ENCIMA de esta linea, o sea inmune a caducar: era el
    # unico. Juan dejo un horometro a medias probandolo el 8-sep a las 16:23 y
    # al dia siguiente mando CINCO partes --3, 4, 7, 8 y 9 de septiembre-- que
    # se comio ese flujo uno por uno, sin una linea en el log. El mismo bug del
    # /deposito, contra el guardia construido para evitarlo.
    descartado = revisar_flujos(context.user_data)
    if descartado:
        await update.message.reply_text(
            f"🧹 Cancelé el {descartado} que quedó a medias "
            f"(más de {MINUTOS_VIDA} min sin terminar). Sigo con tu mensaje.")

    # ── Flujo guiado de horómetro en curso ──
    if context.user_data.get("horo_state"):
        from handlers.horometro_h import atender_paso
        if await atender_paso(update, context):
            return

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

    # ── Aplicaciones: producto, dosis y cuartel ──
    # Va ANTES del portón de maquinaria: medido sobre los 15 partes reales de
    # Juan, ese portón se llevaba 2 por delante (traen números y palabras que
    # le parecen máquina). De los 15 que mandó entre junio y agosto solo 2
    # quedaron anotados; los otros 11 hubo que recuperarlos desde su chat.
    from handlers.aplicaciones import procesar_aplicacion
    if await procesar_aplicacion(update, context):
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
