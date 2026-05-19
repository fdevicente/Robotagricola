"""handlers/personal.py — Comandos de personal y vacaciones.

Comandos: /personal, /vacaciones, /agregar_trabajador
Flujos texto: handle_text_vacacion, handle_text_trabajador
Jobs: job_vacaciones_mensuales
"""
import asyncio
import logging
from datetime import datetime

from telegram.ext import ContextTypes

from config import TELEGRAM_CHAT_ID
from vacaciones_manager import (listar_personal, registrar_vacacion,
                                  agregar_trabajador, actualizar_dias_mensuales)
from utils.formatting import esc

logger = logging.getLogger(__name__)


def _parse_fecha_dmy(texto: str) -> str | None:
    """Convierte DD/MM/YYYY o YYYY-MM-DD a YYYY-MM-DD."""
    from dateutil import parser as date_parser
    try:
        if "/" in texto:
            parts = texto.split("/")
            if len(parts) == 3 and len(parts[0]) <= 2:
                return datetime.strptime(texto, "%d/%m/%Y").strftime("%Y-%m-%d")
        return date_parser.parse(texto).strftime("%Y-%m-%d")
    except Exception:
        return None


async def cmd_personal(update, context):
    personal = await asyncio.to_thread(listar_personal)
    if not personal:
        await update.message.reply_text(
            "👥 No hay personal registrado.\nUsa /agregar\\_trabajador para agregar.",
            parse_mode="Markdown")
        return
    texto = "👥 *Personal — Agrícola Santa Elisa:*\n\n"
    for p in personal:
        texto += f"• *{esc(p['nombre'])}*"
        if p['cargo']:
            texto += f" — {esc(p['cargo'])}"
        texto += f"\n  📅 Pendientes: *{p['dias_pendientes']:.0f} días*"
        if p['ultima_vacacion']:
            texto += f" | Última: {p['ultima_vacacion'][:10]}"
        texto += "\n"
    try:
        await update.message.reply_text(texto, parse_mode="Markdown")
    except Exception:
        await update.message.reply_text(texto)


async def cmd_vacaciones(update, context):
    context.user_data["vacacion_state"] = "esperando_nombre"
    context.user_data["vacacion_data"] = {}
    personal = await asyncio.to_thread(listar_personal)
    if personal:
        nombres = "\n".join(f"  • {p['nombre']}" for p in personal)
        await update.message.reply_text(
            f"🏖️ *Registrar vacaciones*\n\nPersonal:\n{nombres}\n\n"
            "Escribe el *nombre del trabajador*:\n(o /cancelar)",
            parse_mode="Markdown")
    else:
        await update.message.reply_text(
            "🏖️ *Registrar vacaciones*\n\nNo hay personal registrado.\n"
            "Usa /agregar\\_trabajador primero.", parse_mode="Markdown")


async def cmd_agregar_trabajador(update, context):
    context.user_data["trabajador_state"] = "esperando_nombre"
    await update.message.reply_text(
        "👤 *Agregar trabajador*\n\nEscribe el *nombre completo*:\n(o /cancelar)",
        parse_mode="Markdown")


async def handle_text_vacacion(update, context) -> bool:
    """Procesa texto flujo vacaciones. Devuelve True si lo manejó."""
    vac_state = context.user_data.get("vacacion_state")
    if not vac_state:
        return False
    texto = update.message.text.strip()
    data = context.user_data.get("vacacion_data", {})

    if vac_state == "esperando_nombre":
        data["nombre"] = texto
        context.user_data["vacacion_data"] = data
        context.user_data["vacacion_state"] = "esperando_inicio"
        await update.message.reply_text(
            f"🏖️ *Trabajador:* {esc(texto)}\n\n📅 Fecha *inicio* vacaciones (DD/MM/YYYY):",
            parse_mode="Markdown")
        return True

    if vac_state == "esperando_inicio":
        fecha = _parse_fecha_dmy(texto)
        if not fecha:
            await update.message.reply_text("❌ Formato no válido. Usa DD/MM/YYYY")
            return True
        data["inicio"] = fecha
        context.user_data["vacacion_data"] = data
        context.user_data["vacacion_state"] = "esperando_fin"
        await update.message.reply_text(
            f"📅 Inicio: *{fecha}*\n\n📅 Fecha *término* vacaciones (DD/MM/YYYY):",
            parse_mode="Markdown")
        return True

    if vac_state == "esperando_fin":
        fecha = _parse_fecha_dmy(texto)
        if not fecha:
            await update.message.reply_text("❌ Formato no válido. Usa DD/MM/YYYY")
            return True
        result = await asyncio.to_thread(
            registrar_vacacion, data.get("nombre", ""),
            data.get("inicio", ""), fecha)
        context.user_data["vacacion_state"] = None
        context.user_data["vacacion_data"] = {}
        await update.message.reply_text(
            f"🏖️ *Vacaciones registradas*\n\n"
            f"👤 {esc(result['nombre'])}\n"
            f"📅 {result['inicio']} al {result['fin']}\n"
            f"📊 *{result['dias_habiles']} días hábiles* ({result['dias_corridos']} corridos)",
            parse_mode="Markdown")
        return True

    return False


async def handle_text_trabajador(update, context) -> bool:
    """Procesa texto flujo agregar trabajador. Devuelve True si lo manejó."""
    trab_state = context.user_data.get("trabajador_state")
    if not trab_state:
        return False
    texto = update.message.text.strip()

    if trab_state == "esperando_nombre":
        context.user_data["trabajador_nombre"] = texto
        context.user_data["trabajador_state"] = "esperando_cargo"
        await update.message.reply_text(
            f"👤 *{esc(texto)}*\n\nEscribe el *cargo* (o *-* para omitir):",
            parse_mode="Markdown")
        return True

    if trab_state == "esperando_cargo":
        nombre = context.user_data.get("trabajador_nombre", "")
        cargo = texto if texto != "-" else ""
        ok = await asyncio.to_thread(agregar_trabajador, nombre, "", cargo)
        context.user_data["trabajador_state"] = None
        if ok:
            await update.message.reply_text(
                f"✅ *{esc(nombre)}* agregado al personal.", parse_mode="Markdown")
        else:
            await update.message.reply_text(
                f"⚠️ *{esc(nombre)}* ya existe.", parse_mode="Markdown")
        return True

    return False


async def job_vacaciones_mensuales(context: ContextTypes.DEFAULT_TYPE):
    """Job mensual: acumula días de vacaciones a cada trabajador."""
    try:
        result = await asyncio.to_thread(actualizar_dias_mensuales)
        chat_id = context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID
        if chat_id:
            await context.bot.send_message(
                chat_id=int(chat_id),
                text=f"🏖️ *Vacaciones actualizadas*\n\n"
                     f"📊 {result['actualizados']} trabajadores\n"
                     f"📅 +{result['incremento']} días acumulados por persona",
                parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Job vacaciones falló: {e}")
