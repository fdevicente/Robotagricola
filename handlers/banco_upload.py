"""handlers/banco_upload.py — Subir la cartola del banco por Telegram.

El usuario descarga la cartola del portal (CSV/TXT/Excel) y se la manda al bot.
El bot la revisa, muestra qué agregaría y solo escribe tras confirmar.
"""
import asyncio
import logging
import os
from datetime import datetime

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

logger = logging.getLogger(__name__)

EXT_CARTOLA = (".csv", ".txt", ".xlsx", ".xls", ".xlsm")

# Palabras que delatan una cartola bancaria por el nombre del archivo. Sirven
# para atajar los PDF: si una cartola en PDF se cuela al flujo de facturas, el
# extractor la interpreta como factura y crea un registro falso.
PISTAS_CARTOLA = ("cartola", "movimiento", "saldo", "estado_cuenta",
                  "estadocuenta", "typedesc")


def es_archivo_cartola(filename: str) -> bool:
    return os.path.splitext(filename or "")[1].lower() in EXT_CARTOLA


def parece_cartola_por_nombre(filename: str) -> bool:
    """True si el nombre sugiere una cartola aunque la extensión no sirva."""
    base = os.path.basename(filename or "").lower()
    return any(p in base for p in PISTAS_CARTOLA)


def _kb():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Importar", callback_data="cart_import"),
        InlineKeyboardButton("❌ Cancelar", callback_data="cart_cancel"),
    ]])


async def procesar_cartola(update, context, file_path: str, status_msg):
    """Analiza la cartola descargada y muestra el preview."""
    from modules.banco_import import analizar_cartola, formato_resumen
    try:
        res = await asyncio.to_thread(analizar_cartola, file_path)
    except Exception as e:
        logger.error(f"Cartola: error analizando: {e}")
        await status_msg.edit_text(
            f"❌ No pude leer la cartola: {str(e)[:150]}\n\n"
            "Formatos soportados: CSV, TXT o Excel con columnas "
            "Fecha, Descripción, Cargos, Abonos, Saldo.")
        return

    if not res["nuevos"]:
        await status_msg.edit_text(
            f"✅ La cartola no trae movimientos nuevos.\n"
            f"({res['total_archivo']} en el archivo, todos ya estaban en el Master)")
        return

    context.user_data["cartola_path"] = file_path
    texto = formato_resumen(res)
    if len(texto) > 3900:
        texto = texto[:3900] + "\n…"
    await status_msg.edit_text(texto + "\n\n¿Los agrego al Master?",
                                reply_markup=_kb())


async def cb_cart_import(update, context):
    query = update.callback_query
    await query.answer()
    path = context.user_data.get("cartola_path")
    if not path:
        await query.edit_message_text("⚠️ No hay cartola pendiente.")
        return
    await query.edit_message_text("💾 Importando movimientos…")
    from modules.banco_import import importar_cartola
    try:
        r = await asyncio.to_thread(importar_cartola, path)
    except Exception as e:
        logger.error(f"Cartola: error importando: {e}")
        await query.edit_message_text(
            f"❌ No pude guardar: {str(e)[:150]}\n¿Está abierto el Excel?")
        return

    context.user_data["cartola_path"] = None
    saldo = r.get("saldo_archivo")
    txt = (f"✅ Cartola importada\n"
           f"📥 {r['agregados']} movimientos nuevos\n"
           f"⏭️ {r['duplicados']} ya estaban (omitidos)\n")
    if saldo:
        txt += f"💰 Saldo según cartola: ${saldo:,.0f}\n"
    txt += "\nQuedan sin categorizar — revísalos en /banco/revisar o con /conciliar."
    await query.edit_message_text(txt)


async def cb_cart_cancel(update, context):
    query = update.callback_query
    await query.answer()
    context.user_data["cartola_path"] = None
    await query.edit_message_text("🚫 Cartola descartada (no se escribió nada).")
