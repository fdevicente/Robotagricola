"""handlers/tareas.py — Comandos de tareas y bitacora.

Comandos: /tarea, /tareas, /hecho, /bitacora
Callbacks: cb_tarea_prioridad (prior_*)
Flujos texto: handle_text_tarea, handle_text_bitacora
"""
import asyncio
import logging

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

from tareas_manager import (crear_tarea, listar_tareas, actualizar_tarea,
                              registrar_bitacora)
from utils.formatting import esc

logger = logging.getLogger(__name__)


async def cmd_tarea(update, context):
    context.user_data["tarea_state"] = "esperando_desc"
    await update.message.reply_text(
        "📝 *Nueva tarea*\n\nEscribe la *descripción* de la tarea:\n(o /cancelar)",
        parse_mode="Markdown")


async def cmd_tareas(update, context):
    tareas = await asyncio.to_thread(listar_tareas)
    if not tareas:
        await update.message.reply_text("✅ No hay tareas pendientes.")
        return
    texto = "📋 *Tareas pendientes:*\n\n"
    for t in tareas[:20]:
        icon = "🔴" if t["prioridad"] == "Alta" else "🟡" if t["prioridad"] == "Media" else "🟢"
        estado_icon = "🔄" if t["estado"] == "En Progreso" else "⏳"
        texto += f"{estado_icon} {icon} *#{t['id']}* — {esc(t['descripcion'])}\n"
        if t["responsable"]:
            texto += f"    👤 {esc(t['responsable'])}"
        if t["fecha_limite"]:
            texto += f"  📅 {t['fecha_limite']}"
        texto += "\n"
    try:
        await update.message.reply_text(texto, parse_mode="Markdown")
    except Exception:
        await update.message.reply_text(texto)


async def cmd_hecho(update, context):
    context.user_data["tarea_state"] = "esperando_id_hecho"
    await update.message.reply_text(
        "✅ *Completar tarea*\n\nEscribe el *número de tarea* (ej: 3):\n(o /cancelar)",
        parse_mode="Markdown")


async def cmd_bitacora(update, context):
    context.user_data["bitacora_state"] = "esperando_registro"
    await update.message.reply_text(
        "📓 *Bitácora diaria*\n\nEscribe lo que quieras registrar:\n(o /cancelar)",
        parse_mode="Markdown")


async def cb_tarea_prioridad(update, context):
    """Callback para seleccionar prioridad de tarea."""
    query = update.callback_query
    await query.answer()
    prioridad = query.data.replace("prior_", "")
    desc = context.user_data.get("tarea_desc", "")
    result = await asyncio.to_thread(crear_tarea, desc, prioridad)
    context.user_data["tarea_state"] = None
    await query.edit_message_text(
        f"✅ *Tarea #{result['id']} creada*\n\n"
        f"📝 {esc(result['descripcion'])}\n"
        f"{'🔴' if prioridad == 'Alta' else '🟡' if prioridad == 'Media' else '🟢'} Prioridad: {prioridad}",
        parse_mode="Markdown")


async def handle_text_tarea(update, context) -> bool:
    """Procesa texto si hay flujo de tarea activo. Devuelve True si lo manejó."""
    tarea_state = context.user_data.get("tarea_state")
    if not tarea_state:
        return False
    texto = update.message.text.strip()

    if tarea_state == "esperando_desc":
        context.user_data["tarea_desc"] = texto
        context.user_data["tarea_state"] = "esperando_prioridad"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🔴 Alta", callback_data="prior_Alta"),
             InlineKeyboardButton("🟡 Media", callback_data="prior_Media"),
             InlineKeyboardButton("🟢 Baja", callback_data="prior_Baja")]])
        await update.message.reply_text(
            f"📝 *Tarea:* {esc(texto)}\n\n¿Qué *prioridad*?",
            parse_mode="Markdown", reply_markup=kb)
        return True

    if tarea_state == "esperando_id_hecho":
        try:
            task_id = int(texto)
        except ValueError:
            await update.message.reply_text("❌ Escribe solo el número de tarea.")
            return True
        context.user_data["tarea_state"] = "esperando_obs_hecho"
        context.user_data["tarea_id_hecho"] = task_id
        await update.message.reply_text(
            f"✅ Tarea *#{task_id}*\n\nEscribe una *observación* de lo que se hizo\n(o escribe *ok* para omitir):",
            parse_mode="Markdown")
        return True

    if tarea_state == "esperando_obs_hecho":
        task_id = context.user_data.get("tarea_id_hecho", 0)
        obs = texto if texto.lower() != "ok" else ""
        ok = await asyncio.to_thread(actualizar_tarea, task_id, "Hecho", obs)
        context.user_data["tarea_state"] = None
        if ok:
            await update.message.reply_text(f"✅ Tarea *#{task_id}* completada.", parse_mode="Markdown")
        else:
            await update.message.reply_text(f"❌ No encontré tarea #{task_id}.")
        return True

    return False


async def handle_text_bitacora(update, context) -> bool:
    """Procesa texto si hay flujo de bitácora. Devuelve True si lo manejó."""
    if context.user_data.get("bitacora_state") != "esperando_registro":
        return False
    texto = update.message.text.strip()
    await asyncio.to_thread(registrar_bitacora, texto)
    context.user_data["bitacora_state"] = None
    await update.message.reply_text(f"📓 *Registrado en bitácora:*\n{esc(texto)}",
                                     parse_mode="Markdown")
    return True
