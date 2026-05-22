"""
main.py — Bot Agrícola Santa Elisa
"""
import os
import re
import sys
import logging
import asyncio
from datetime import datetime, time as dtime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler,
    MessageHandler, CallbackQueryHandler,
    filters, ContextTypes
)

from config import TELEGRAM_TOKEN, DOWNLOAD_DIR, BOLETAS_DIR, EXCEL_PATH, TELEGRAM_CHAT_ID
from processors.extractor import process_file
from excel_manager import (append_to_excel, delete_last_rows, buscar_factura, registrar_pago,
                          append_boleta, delete_last_boletas, registrar_deposito_caja, consultar_saldo_caja,
                          reporte_diario, reporte_semanal, reporte_mensual, crear_hoja_banco,
                          guardar_movimientos_banco, obtener_resumen_banco)
from tareas_manager import (crear_hojas_tareas, crear_tarea, listar_tareas, actualizar_tarea,
                            registrar_bitacora, listar_bitacora, resumen_tareas_semana)
from inventario_manager import (crear_hojas_inventario, registrar_uso, consultar_inventario,
                                productos_bajo_stock, consumo_por_cultivo, agregar_stock_desde_factura)
from vacaciones_manager import (crear_hojas_vacaciones, agregar_trabajador, registrar_vacacion,
                                listar_personal, vacaciones_pendientes, actualizar_dias_mensuales,
                                ultimas_vacaciones)
from chat_inteligente import responder_chat

_LOG_FMT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
logging.basicConfig(format=_LOG_FMT, level=logging.INFO)
logger = logging.getLogger(__name__)

# FileHandler además del stdout para que warnings/errors persistan
# (ej: fallas al renombrar archivos por handles abiertos en Windows).
try:
    _LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot.log")
    _fh = logging.FileHandler(_LOG_FILE, encoding="utf-8")
    _fh.setLevel(logging.INFO)
    _fh.setFormatter(logging.Formatter(_LOG_FMT))
    logging.getLogger().addHandler(_fh)
except Exception as _e:
    logger.warning(f"No se pudo crear FileHandler de log: {_e}")

from handlers.facturas import (
    CAMPOS_EDITABLES, _save_path, _save_path_boleta,
    _renombrar_archivo, _es_boleta,
)

# Imports de utils/ (deduplicados desde main.py)
from utils.formatting import esc as _esc, format_date as _format_date, calc_vencimiento as _calc_vencimiento


from handlers.facturas import _registrar_correccion


from handlers.facturas import _build_preview

from utils.keyboards import (
    main_keyboard as _main_keyboard,
    edit_keyboard as _edit_keyboard,
    proveedor_nuevo_keyboard as _proveedor_nuevo_keyboard,
)

from handlers.facturas import (
    CAMPOS_COMUNES, CAMPOS_POR_ITEM,
    _rut_existe, _agregar_proveedor,
)

# ── COMANDOS ──────────────────────────────────
async def cmd_start(update, context):
    # Guardar chat_id para jobs automáticos
    context.bot_data["banco_chat_id"] = update.effective_chat.id
    await update.message.reply_text(
        "👋 ¡Hola! Soy el bot de *Agrícola Santa Elisa*.\n\n"
        "📲 Envíame una *foto* o *PDF* de una factura.\n"
        "💬 También puedes escribirme consultas agrícolas.\n\n"
        "*Finanzas:*\n"
        "  /pagado — Registrar pago de factura\n"
        "  /deposito — Depositar en caja chica\n"
        "  /saldo — Saldo caja chica\n"
        "  /banco — Sincronizar Scotiabank\n"
        "  /reporte — Reportes diario/semanal/mensual\n\n"
        "*Inventario:*\n"
        "  /inventario — Ver stock de insumos\n"
        "  /uso — Registrar uso de insumo\n\n"
        "*Tareas y Bitácora:*\n"
        "  /tarea — Crear tarea nueva\n"
        "  /tareas — Ver tareas pendientes\n"
        "  /hecho — Marcar tarea completada\n"
        "  /bitacora — Registrar actividad diaria\n\n"
        "*Personal:*\n"
        "  /personal — Ver personal y vacaciones\n"
        "  /vacaciones — Registrar vacaciones\n"
        "  /agregar\\_trabajador — Agregar trabajador\n\n"
        "*Dashboard:*\n"
        "  /dashboard — Abrir panel de control web\n\n"
        "*Otros:*\n"
        "  /deshacer — Eliminar ultima factura\n"
        "  /cancelar — Cancelar operacion\n"
        "  /ayuda — Esta ayuda",
        parse_mode="Markdown")

async def cmd_ayuda(update, context): await cmd_start(update, context)

from handlers.finanzas import (
    cmd_pagado, cmd_deposito, cmd_saldo, cmd_dashboard, cmd_reporte, cmd_banco,
    cb_reporte, cb_medio_pago, cb_calce_verificacion,
    job_sync_banco, handle_text_pagado, handle_text_deposito,
)

from handlers.personal import (
    cmd_personal, cmd_vacaciones, cmd_agregar_trabajador,
    handle_text_vacacion, handle_text_trabajador, job_vacaciones_mensuales,
)


async def cmd_cancelar(update, context):
    for key in ("pagado_state", "pagado_nro", "deposito_state", "editing_field",
                "editing_item_idx", "tarea_state", "uso_state", "uso_data",
                "vacacion_state", "vacacion_data", "trabajador_state", "bitacora_state"):
        context.user_data[key] = None
    await update.message.reply_text("🚫 Operación cancelada.")

async def cmd_deshacer(update, context):
    last_file = context.user_data.get("last_invoice_file")
    last_rows = context.user_data.get("last_invoice_rows", 0)
    if not last_file or last_rows <= 0:
        await update.message.reply_text("⚠️ No hay factura reciente para deshacer."); return
    msg = await update.message.reply_text("⏳ Deshaciendo…")
    try:
        fue_boleta = context.user_data.get("last_invoice_boleta", False)
        if fue_boleta:
            success = await asyncio.to_thread(delete_last_boletas, last_rows)
        else:
            success = await asyncio.to_thread(delete_last_rows, last_rows)
    except Exception as e: logger.error(e); success = False
    if success:
        try:
            if os.path.exists(last_file): os.remove(last_file)
        except: pass
        context.user_data["last_invoice_file"] = None
        context.user_data["last_invoice_rows"]  = 0
        await msg.edit_text(f"🗑️ Listo. {last_rows} fila(s) eliminadas del Excel.")
    else:
        await msg.edit_text("❌ Error al borrar del Excel.")

# ── TAREAS Y BITÁCORA ────────────────────────
from handlers.tareas import (
    cmd_tarea, cmd_tareas, cmd_hecho, cmd_bitacora,
    cb_tarea_prioridad, handle_text_tarea, handle_text_bitacora,
)

# ── INVENTARIO ───────────────────────────────
from handlers.inventario_h import (
    cmd_inventario, cmd_uso, cb_cultivo, handle_text_uso,
)

# ── VACACIONES Y PERSONAL ────────────────────
# ── ARCHIVOS ──────────────────────────────────
from handlers.facturas import _download_with_retry

from handlers.facturas import (
    _show_preview, handle_photo, handle_document,
    _process_and_reply, _guardar_excel,
)

# ── CALLBACKS ─────────────────────────────────
from handlers.facturas import (
    cb_confirm_save, cb_add_proveedor_yes, cb_add_proveedor_no, cb_cancel_save,
    cb_edit_menu, cb_edit_field, cb_select_item, cb_edit_total_factura,
    cb_back_preview, cb_add_item, cb_del_item_menu, cb_del_item_confirm,
)

from handlers.chat import handle_text as handle_text_edit

# ── ARRANQUE ──────────────────────────────────
def main():
    if not TELEGRAM_TOKEN:
        logger.error("TELEGRAM_TOKEN no configurado"); return
    app = (ApplicationBuilder().token(TELEGRAM_TOKEN)
           .read_timeout(120).write_timeout(120).connect_timeout(120).pool_timeout(120).build())

    # Comandos
    app.add_handler(CommandHandler("start",    cmd_start))
    app.add_handler(CommandHandler("ayuda",    cmd_ayuda))
    app.add_handler(CommandHandler("deshacer", cmd_deshacer))
    app.add_handler(CommandHandler("pagado",   cmd_pagado))
    app.add_handler(CommandHandler("deposito", cmd_deposito))
    app.add_handler(CommandHandler("saldo",    cmd_saldo))
    app.add_handler(CommandHandler("reporte",  cmd_reporte))
    app.add_handler(CommandHandler("dashboard", cmd_dashboard))
    app.add_handler(CommandHandler("banco",    cmd_banco))
    app.add_handler(CommandHandler("cancelar", cmd_cancelar))
    # Cash Flow (Fase 1)
    from handlers.cash_flow_cmds import cmd_proyeccion, cmd_categoria, cosecha_conv
    app.add_handler(CommandHandler("proyeccion", cmd_proyeccion))
    app.add_handler(CommandHandler("categoria",  cmd_categoria))
    app.add_handler(cosecha_conv)
    # Tareas y Bitácora
    app.add_handler(CommandHandler("tarea",    cmd_tarea))
    app.add_handler(CommandHandler("tareas",   cmd_tareas))
    app.add_handler(CommandHandler("hecho",    cmd_hecho))
    app.add_handler(CommandHandler("bitacora", cmd_bitacora))
    # Inventario
    app.add_handler(CommandHandler("inventario", cmd_inventario))
    app.add_handler(CommandHandler("uso",      cmd_uso))
    # Personal y Vacaciones
    app.add_handler(CommandHandler("personal",  cmd_personal))
    app.add_handler(CommandHandler("vacaciones", cmd_vacaciones))
    app.add_handler(CommandHandler("agregar_trabajador", cmd_agregar_trabajador))
    # Mensajes
    app.add_handler(MessageHandler(filters.Document.ALL,            handle_document))
    app.add_handler(MessageHandler(filters.PHOTO,                   handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text_edit))
    # Callbacks
    app.add_handler(CallbackQueryHandler(cb_confirm_save,      pattern="^confirm_save$"))
    app.add_handler(CallbackQueryHandler(cb_cancel_save,       pattern="^cancel_save$"))
    app.add_handler(CallbackQueryHandler(cb_edit_menu,         pattern="^edit_menu$"))
    app.add_handler(CallbackQueryHandler(cb_back_preview,      pattern="^back_preview$"))
    app.add_handler(CallbackQueryHandler(cb_add_proveedor_yes, pattern="^add_proveedor_yes$"))
    app.add_handler(CallbackQueryHandler(cb_add_proveedor_no,  pattern="^add_proveedor_no$"))
    app.add_handler(CallbackQueryHandler(cb_add_item,          pattern="^add_item$"))
    app.add_handler(CallbackQueryHandler(cb_del_item_menu,     pattern="^del_item$"))
    app.add_handler(CallbackQueryHandler(cb_del_item_confirm,  pattern="^delitem_"))
    app.add_handler(CallbackQueryHandler(cb_edit_total_factura, pattern="^edit_total_factura$"))
    app.add_handler(CallbackQueryHandler(cb_medio_pago,        pattern="^pago_"))
    app.add_handler(CallbackQueryHandler(cb_calce_verificacion, pattern="^calce_"))
    app.add_handler(CallbackQueryHandler(cb_reporte,           pattern="^rep_"))
    app.add_handler(CallbackQueryHandler(cb_tarea_prioridad,   pattern="^prior_"))
    app.add_handler(CallbackQueryHandler(cb_cultivo,           pattern="^cult_"))
    app.add_handler(CallbackQueryHandler(cb_select_item,       pattern="^selitem_"))
    app.add_handler(CallbackQueryHandler(cb_edit_field,        pattern="^edit_"))

    # Crear hojas si no existen
    try: crear_hoja_banco()
    except: pass
    try: crear_hojas_tareas()
    except: pass
    try: crear_hojas_inventario()
    except: pass
    try: crear_hojas_vacaciones()
    except: pass

    # Programar sincronización bancaria automática: 8:00 y 18:00 hora Chile
    try:
        from zoneinfo import ZoneInfo
    except ImportError:
        from backports.zoneinfo import ZoneInfo
    tz_chile = ZoneInfo("America/Santiago")
    app.job_queue.run_daily(job_sync_banco, time=dtime(hour=8,  minute=0, tzinfo=tz_chile), name="banco_8am")
    app.job_queue.run_daily(job_sync_banco, time=dtime(hour=18, minute=0, tzinfo=tz_chile), name="banco_6pm")
    # Actualizar días de vacaciones el 1ro de cada mes
    app.job_queue.run_monthly(job_vacaciones_mensuales, when=dtime(hour=7, minute=0, tzinfo=tz_chile),
                              day=1, name="vacaciones_mensuales")
    # Resumen semanal cash flow: lunes 08:00
    from handlers.cash_flow_jobs import job_resumen_semanal
    app.job_queue.run_daily(job_resumen_semanal,
                              time=dtime(hour=8, minute=0, tzinfo=tz_chile),
                              days=(0,), name="resumen_semanal")
    logger.info("⏰ Jobs programados: banco 08:00/18:00, vacaciones día 1, resumen lunes 08:00")

    logger.info("✅ Bot iniciado. Esperando mensajes...")
    app.run_polling()

if __name__ == "__main__":
    main()
