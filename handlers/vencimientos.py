"""handlers/vencimientos.py — Registro diferido de fechas de vencimiento.

Flujo:
  Al confirmar una factura de insumos → sus productos quedan PENDIENTES.
  Recordatorio suave: "Tienes N productos sin fecha de vencimiento".
  /vencimientos → pregunta producto por producto: fecha / No vence / Saltar.
Alertas (50%, 10%, vencido) se calculan y salen en el reporte mensual.
"""
import asyncio
import logging
from datetime import datetime

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

logger = logging.getLogger(__name__)


def _esc(t):
    if t is None:
        return ""
    for ch in r"_*[]()~`>#+-=|{}.!":
        t = str(t).replace(ch, "\\" + ch)
    return t


def _kb_item():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("🚫 No vence", callback_data="venc_novence"),
        InlineKeyboardButton("⏭️ Saltar", callback_data="venc_skip"),
        InlineKeyboardButton("❌ Terminar", callback_data="venc_stop"),
    ]])


def _parse_fecha(texto: str):
    texto = texto.strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d/%m/%y"):
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            pass
    return None


async def _preguntar_siguiente(enviar, context):
    """Pregunta el siguiente pendiente. `enviar` es una corutina send/edit."""
    pendientes = context.user_data.get("venc_pendientes", [])
    idx = context.user_data.get("venc_idx", 0)
    if idx >= len(pendientes):
        context.user_data["venc_state"] = None
        await enviar("✅ *Listo.* Fechas de vencimiento registradas.\n"
                     "Las alertas saldrán en el reporte mensual.",
                     parse_mode="Markdown")
        return
    p = pendientes[idx]
    total = len(pendientes)
    await enviar(
        f"📅 *Vencimiento {idx+1}/{total}*\n\n"
        f"📦 *{_esc(p['producto'])}*\n"
        f"🏢 {_esc(p['proveedor'])} · F{_esc(p['nro_factura'])}\n"
        f"🛒 Comprado: {p['fecha_compra'] or '—'}\n\n"
        f"¿Cuándo vence? Escribe la fecha _(DD\\-MM\\-AAAA)_\n"
        f"o usa los botones:",
        parse_mode="MarkdownV2", reply_markup=_kb_item())


async def cmd_vencimientos(update, context):
    """Inicia el registro de fechas de los productos pendientes."""
    from vencimientos_manager import listar_pendientes
    pendientes = await asyncio.to_thread(listar_pendientes)
    if not pendientes:
        await update.message.reply_text(
            "✅ No hay productos pendientes de fecha de vencimiento.")
        return
    context.user_data["venc_pendientes"] = pendientes
    context.user_data["venc_idx"] = 0
    context.user_data["venc_state"] = "esperando_fecha"
    await update.message.reply_text(
        f"📅 *{len(pendientes)} producto(s)* sin fecha de vencimiento.\n"
        f"Vamos uno por uno:", parse_mode="Markdown")
    await _preguntar_siguiente(update.message.reply_text, context)


async def handle_text_vencimiento(update, context) -> bool:
    """Captura la fecha escrita si hay flujo de vencimientos activo."""
    if context.user_data.get("venc_state") != "esperando_fecha":
        return False
    fecha = _parse_fecha(update.message.text)
    pendientes = context.user_data.get("venc_pendientes", [])
    idx = context.user_data.get("venc_idx", 0)
    if idx >= len(pendientes):
        context.user_data["venc_state"] = None
        return True
    if not fecha:
        await update.message.reply_text(
            "⚠️ No entendí la fecha. Usa formato DD-MM-AAAA "
            "(o el botón 🚫 No vence / ⏭️ Saltar).")
        return True
    from vencimientos_manager import registrar_vencimiento, calcular_estado
    p = pendientes[idx]
    await asyncio.to_thread(registrar_vencimiento, p["fila"], fecha)
    estado, pct, vida = calcular_estado(p.get("fecha_compra"), fecha)
    nota = ""
    if estado in ("ALERTA 10%", "VENCIDO"):
        nota = f" ⚠️ {estado}"
    elif estado == "ALERTA 50%":
        nota = " 🟡 mitad de vida útil"
    await update.message.reply_text(
        f"✅ {_esc(p['producto'])}: vence {fecha.strftime('%d-%m-%Y')}{nota}",
        parse_mode="Markdown")
    context.user_data["venc_idx"] = idx + 1
    await _preguntar_siguiente(update.message.reply_text, context)
    return True


# ── Callbacks ───────────────────────────────────────────

async def cb_venc_novence(update, context):
    query = update.callback_query
    await query.answer()
    from vencimientos_manager import registrar_vencimiento
    pendientes = context.user_data.get("venc_pendientes", [])
    idx = context.user_data.get("venc_idx", 0)
    if idx < len(pendientes):
        p = pendientes[idx]
        await asyncio.to_thread(registrar_vencimiento, p["fila"], None, True)
        await query.edit_message_text(f"🚫 {_esc(p['producto'])}: marcado como *no vence*",
                                       parse_mode="Markdown")
        context.user_data["venc_idx"] = idx + 1
    await _preguntar_siguiente(
        lambda *a, **k: context.bot.send_message(query.message.chat_id, *a, **k),
        context)


async def cb_venc_skip(update, context):
    query = update.callback_query
    await query.answer()
    pendientes = context.user_data.get("venc_pendientes", [])
    idx = context.user_data.get("venc_idx", 0)
    if idx < len(pendientes):
        p = pendientes[idx]
        await query.edit_message_text(f"⏭️ {_esc(p['producto'])}: lo dejo pendiente",
                                       parse_mode="Markdown")
        context.user_data["venc_idx"] = idx + 1
    await _preguntar_siguiente(
        lambda *a, **k: context.bot.send_message(query.message.chat_id, *a, **k),
        context)


async def cb_venc_stop(update, context):
    query = update.callback_query
    await query.answer()
    context.user_data["venc_state"] = None
    restantes = len(context.user_data.get("venc_pendientes", [])) - context.user_data.get("venc_idx", 0)
    msg = "✅ Terminado."
    if restantes > 0:
        msg += f" Quedan {restantes} pendientes (usa /vencimientos para seguir)."
    await query.edit_message_text(msg)


async def recordatorio_pendientes(context, chat_id):
    """Aviso suave si hay productos pendientes de vencimiento."""
    from vencimientos_manager import listar_pendientes
    pendientes = await asyncio.to_thread(listar_pendientes)
    if not pendientes:
        return
    await context.bot.send_message(
        chat_id=chat_id,
        text=f"📅 Tienes *{len(pendientes)}* producto(s) sin fecha de vencimiento. "
             f"Usa /vencimientos cuando quieras registrarlas.",
        parse_mode="Markdown")
