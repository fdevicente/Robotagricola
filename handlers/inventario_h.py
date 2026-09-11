"""handlers/inventario_h.py — Comandos de inventario y registro de uso.
(Sufijo _h para no colisionar con inventario_manager.py)

Comandos: /inventario, /uso
Callbacks: cb_cultivo (cult_*)
Flujos texto: handle_text_uso
"""
import asyncio
import logging

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

from inventario_manager import consultar_inventario, registrar_uso
from utils.formatting import esc

logger = logging.getLogger(__name__)


async def cmd_inventario(update, context):
    items = await asyncio.to_thread(consultar_inventario)
    if not items:
        await update.message.reply_text(
            "📦 Inventario vacío. Los insumos se agregan automáticamente al procesar facturas.")
        return
    texto = "📦 *Inventario de insumos:*\n\n"
    for it in items[:30]:
        alerta = " ⚠️" if it["alerta"] else ""
        texto += f"• *{esc(it['producto'])}*\n"
        texto += f"  {it['categoria']} — Stock: {it['stock']:g} {it['unidad']}{alerta}\n"
    bajo = [i for i in items if i["alerta"]]
    if bajo:
        texto += f"\n⚠️ *{len(bajo)} producto(s) con stock bajo*"
    try:
        await update.message.reply_text(texto, parse_mode="Markdown")
    except Exception:
        await update.message.reply_text(texto)


async def cmd_uso(update, context):
    context.user_data["uso_state"] = "esperando_producto"
    context.user_data["uso_data"] = {}
    await update.message.reply_text(
        "🧪 *Registrar uso de insumo*\n\nEscribe el *nombre del producto*:\n(o /cancelar)",
        parse_mode="Markdown")


async def cb_cultivo(update, context):
    """Callback para seleccionar cultivo en uso de insumo."""
    query = update.callback_query
    await query.answer()
    cultivo = query.data.replace("cult_", "")
    data = context.user_data.get("uso_data", {})
    result = await asyncio.to_thread(
        registrar_uso, data.get("producto", ""), data.get("cantidad", 0),
        cultivo, data.get("sector", ""))
    # Un nombre ambiguo NO se resuelve al azar: se pregunta. El flujo se queda
    # abierto esperando el nombre bueno.
    if result.get("error") == "ambiguo":
        opciones = "\n".join("• %s" % esc(c) for c in result["candidatos"])
        await query.edit_message_text(
            "🤔 *%s* calza con más de un producto y no quiero elegir por ti:\n\n"
            "%s\n\nEscribe el nombre completo del que usaste."
            % (esc(result["producto"]), opciones), parse_mode="Markdown")
        context.user_data["uso_state"] = "esperando_producto"
        return

    if result.get("error") == "sin_producto":
        await query.edit_message_text("🤔 No entendí qué producto usaste. "
                                      "Empieza de nuevo con /uso.")
        context.user_data["uso_state"] = None
        context.user_data["uso_data"] = {}
        return

    context.user_data["uso_state"] = None
    context.user_data["uso_data"] = {}

    # El stock puede venir en None: el producto no esta en el inventario y ya
    # no se le inventa una fila en negativo.
    if result.get("producto_desconocido"):
        stock = ("⚠️ *%s* no está en el inventario, así que anoté el uso pero no "
                 "pude descontarlo. Revisa el nombre." % esc(result["producto"]))
    elif result.get("stock_negativo"):
        stock = ("⚠️ *Stock en rojo:* queda en %g %s. O falta cargar una compra, "
                 "o la cantidad no es esa."
                 % (result["stock_restante"], result["unidad"]))
    else:
        alerta = "  ⚠️ *bajo!*" if result.get("alerta_bajo") else ""
        stock = "📦 Stock restante: %g %s%s" % (result["stock_restante"],
                                                result["unidad"], alerta)

    await query.edit_message_text(
        f"✅ *Uso registrado*\n\n"
        f"🧪 {esc(result['producto'])} — {result['cantidad']:g} {result['unidad']}\n"
        f"🌳 Cultivo: {cultivo}\n"
        f"{stock}",
        parse_mode="Markdown")


async def handle_text_uso(update, context) -> bool:
    """Procesa texto si hay flujo de uso activo. Devuelve True si lo manejó."""
    uso_state = context.user_data.get("uso_state")
    if not uso_state:
        return False
    texto = update.message.text.strip()
    data = context.user_data.get("uso_data", {})

    if uso_state == "esperando_producto":
        data["producto"] = texto
        context.user_data["uso_data"] = data
        context.user_data["uso_state"] = "esperando_cantidad"
        await update.message.reply_text(
            f"🧪 *Producto:* {esc(texto)}\n\nEscribe la *cantidad* utilizada (ej: 10):",
            parse_mode="Markdown")
        return True

    if uso_state == "esperando_cantidad":
        try:
            n = texto.replace(",", ".")
            data["cantidad"] = float(n)
        except ValueError:
            await update.message.reply_text("❌ No es un número válido.")
            return True
        context.user_data["uso_data"] = data
        context.user_data["uso_state"] = "esperando_sector"
        await update.message.reply_text(
            f"📍 ¿En qué *sector* del campo? (o escribe *-* para omitir):",
            parse_mode="Markdown")
        return True

    if uso_state == "esperando_sector":
        data["sector"] = texto if texto != "-" else ""
        context.user_data["uso_data"] = data
        context.user_data["uso_state"] = "esperando_cultivo"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🌰 Nogales", callback_data="cult_Nogales"),
             InlineKeyboardButton("🍒 Cerezos", callback_data="cult_Cerezos"),
             InlineKeyboardButton("🌿 Avellanos", callback_data="cult_Avellanos")]])
        await update.message.reply_text(
            f"🌳 ¿En qué *cultivo*?", parse_mode="Markdown", reply_markup=kb)
        return True

    return False
