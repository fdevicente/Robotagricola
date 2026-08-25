"""handlers/conciliacion.py — /conciliar: conciliación bancaria con IA.

Flujo: /conciliar [días] → análisis (solo lectura) + IA para dudosos →
reporte con [✅ Aplicar] [❌ Cancelar]. Nada se escribe sin confirmar.
"""
import asyncio
import logging

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

logger = logging.getLogger(__name__)


def _kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Aplicar", callback_data="conc_apply"),
        InlineKeyboardButton("❌ Cancelar", callback_data="conc_cancel"),
    ]])


async def cmd_conciliar(update, context):
    """/conciliar [días]  (default 90)"""
    dias = 90
    if context.args:
        try:
            dias = max(7, min(400, int(context.args[0])))
        except ValueError:
            pass
    status = await update.message.reply_text(
        f"🏦 Conciliando banco ↔ facturas (últimos {dias} días)…")
    from modules.conciliador import analizar, resolver_dudosos_ia, formato_resumen
    try:
        res = await asyncio.to_thread(analizar, dias)
        ia_ok = await asyncio.to_thread(resolver_dudosos_ia, res["dudosos"])
    except Exception as e:
        logger.error(f"Conciliación: {e}")
        await status.edit_text(f"❌ Error en la conciliación: {e}")
        return

    matches = res["confirmados"] + ia_ok
    context.user_data["conc_matches"] = matches
    texto = formato_resumen(res, ia_ok)
    if len(texto) > 3900:
        texto = texto[:3900] + "\n…"
    if matches:
        texto += (f"\n\n¿Aplico los {len(matches)} enlaces? "
                  f"(escribe el link banco↔factura y completa fechas de pago)")
        await status.edit_text(texto, reply_markup=_kb())
    else:
        await status.edit_text(texto + "\n\nNada que aplicar.")


async def cb_conc_apply(update, context):
    query = update.callback_query
    await query.answer()
    matches = context.user_data.get("conc_matches") or []
    if not matches:
        await query.edit_message_text("⚠️ No hay conciliación pendiente.")
        return
    await query.edit_message_text(f"💾 Aplicando {len(matches)} enlaces…")
    from modules.conciliador import aplicar_conciliacion
    try:
        r = await asyncio.to_thread(aplicar_conciliacion, matches)
        context.user_data["conc_matches"] = None
        await query.edit_message_text(
            f"✅ Conciliación aplicada\n"
            f"🔗 {r['links']} cargos enlazados a su factura\n"
            f"📅 {r['fechas']} líneas con fecha de pago completada")
    except Exception as e:
        logger.error(f"Aplicar conciliación: {e}")
        await query.edit_message_text(
            f"❌ No pude aplicar: {e}\n¿Está abierto el Excel?")


async def cb_conc_cancel(update, context):
    query = update.callback_query
    await query.answer()
    context.user_data["conc_matches"] = None
    await query.edit_message_text("🚫 Conciliación descartada (no se escribió nada).")
