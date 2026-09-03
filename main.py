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
    MessageHandler, CallbackQueryHandler, TypeHandler,
    PicklePersistence, filters, ContextTypes
)
from infrastructure import bot_state

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

# ── HEARTBEAT / RECONEXIÓN ────────────────────
async def _track_activity(update, context):
    """Registra el último update procesado (heartbeat). Corre en group=-1.

    Debe dejar pasar el update a los demás handlers, por eso no detiene la
    propagación (no levanta ApplicationHandlerStop).

    Acá también se respalda el mensaje CRUDO, antes de que ningún handler lo
    interprete: es el único registro de lo que realmente llegó. Sin esto, un
    mensaje que se pierde (pasó el 24-ago-2026) no deja rastro para diagnosticar.
    """
    from modules.telegram_backup import guardar_update
    guardar_update(update)          # nunca lanza; si falla, solo lo loguea

    try:
        chat_id = update.effective_chat.id if update.effective_chat else None
        # Resumen breve del contenido para el log de estado
        resumen = ""
        msg = update.effective_message
        if msg:
            if msg.document:
                resumen = f"documento: {msg.document.file_name}"
            elif msg.photo:
                resumen = "foto/imagen"
            elif msg.text:
                resumen = msg.text[:60]
        bot_state.guardar_actividad(update_id=update.update_id,
                                     chat_id=chat_id, resumen=resumen)
        # Recordar último chat activo para avisos de reconexión
        if chat_id:
            context.bot_data["ultimo_chat_activo"] = chat_id
    except Exception as e:
        logger.warning(f"track_activity: {e}")


async def job_latido(context):
    """Deja constancia de que el proceso está vivo (cada 5 min, en silencio)."""
    bot_state.guardar_latido()


# Menú que Telegram le muestra a CUALQUIERA que escriba "/" en el chat.
# Sin esto no aparece ninguna sugerencia y hay que saberse los comandos de
# memoria: por eso a Juan "no le salía" /personal. Se registra en cada arranque.
COMANDOS_MENU = [
    ("bitacora", "Registrar la actividad del día"),
    ("maquinaria", "Horómetro, mantención o ficha de máquina"),
    ("inventario", "Ver stock de insumos"),
    ("uso", "Registrar uso de un insumo"),
    ("vencimientos", "Insumos por vencer"),
    ("bodega", "Chequeo de bodega"),
    ("tarea", "Crear una tarea"),
    ("tareas", "Ver tareas pendientes"),
    ("hecho", "Marcar una tarea como lista"),
    ("personal", "Ver personal y días de vacaciones"),
    ("vacaciones", "Registrar vacaciones de alguien"),
    ("saldo", "Saldo de caja chica"),
    ("deposito", "Depositar en caja chica"),
    ("pagado", "Registrar el pago de una factura"),
    ("reporte", "Reporte diario, semanal o mensual"),
    ("drive", "Ver y reintentar las subidas a Drive"),
    ("cancelar", "Cancelar lo que se está haciendo"),
    ("ayuda", "Ver todo lo que puedo hacer"),
]


async def _post_init(app):
    """Tras inicializar el bot: registrar el menú y avisar si estuvo caído.

    El corte se mide con el latido, no con el último mensaje recibido: un fin
    de semana sin que nadie escriba no es una caída.
    Telegram entrega los pendientes (<24h) porque run_polling usa
    drop_pending_updates=False; este aviso solo informa.
    """
    try:
        from telegram import BotCommand
        await app.bot.set_my_commands(
            [BotCommand(c, d) for c, d in COMANDOS_MENU])
        logger.info(f"Menú de comandos registrado ({len(COMANDOS_MENU)}).")
    except Exception as e:
        logger.warning(f"No pude registrar el menú de comandos: {e}")

    msg = bot_state.mensaje_reconexion()
    bot_state.guardar_latido()      # arranca el latido de inmediato
    if not msg:
        return
    # Destino: dueño si está fijado (/soydueno); si no, último chat activo o banco
    chat_id = (app.bot_data.get("owner_chat_id")
               or app.bot_data.get("ultimo_chat_activo")
               or app.bot_data.get("banco_chat_id")
               or TELEGRAM_CHAT_ID)
    if not chat_id:
        logger.info(f"Reconexión sin chat destino. {msg}")
        return
    try:
        await app.bot.send_message(chat_id=chat_id, text=msg, parse_mode="Markdown")
    except Exception as e:
        logger.warning(f"No pude enviar aviso de reconexión: {e}")


# ── COMANDOS ──────────────────────────────────
async def cmd_start(update, context):
    # Guardar chat_id para jobs automáticos SOLO si aún no está fijado
    # (evita que cualquier usuario nuevo secuestre el destino de los avisos;
    #  para re-apuntarlo a propósito: correr /banco en el chat deseado)
    if not context.bot_data.get("banco_chat_id"):
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
# Vencimientos de insumos (fechas diferidas + alertas)
from handlers.vencimientos import (
    cmd_vencimientos, handle_text_vencimiento,
    cb_venc_novence, cb_venc_skip, cb_venc_stop,
)
# Monitoreo del bot + mirror de actividad
from handlers.monitoreo import (
    cmd_soydueno, cmd_estado, mirror_update, job_heartbeat,
    cmd_bodega, job_bodega_check, cmd_correlativos,
    cmd_basedatos, job_sync_db,
)
# Conciliación bancaria con IA
from handlers.conciliacion import (
    cmd_conciliar, cb_conc_apply, cb_conc_cancel,
)
# Carga manual de la cartola del banco
from handlers.banco_upload import cb_cart_import, cb_cart_cancel


async def cmd_cancelar(update, context):
    # La lista vive en modules/flujos.py: si se agrega un flujo nuevo y aca
    # quedara una copia, /cancelar dejaria ese flujo abierto para siempre.
    from modules.flujos import limpiar_flujos
    limpiar_flujos(context.user_data)
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
    cmd_tarea, cmd_tareas, cmd_hecho,
    cb_tarea_prioridad, handle_text_tarea,
)
# Bitácora estructurada con IA (reemplaza el flujo viejo de handlers.tareas)
from handlers.bitacora import (
    cmd_bitacora, handle_text_bitacora,
    cb_bita_save, cb_bita_edit, cb_bita_cancel,
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
    # Persistencia: mantiene user_data (cola de facturas, factura en preview),
    # chat_data y bot_data (banco_chat_id) entre reinicios del bot.
    _persist_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".bot_persistence.pickle")
    persistence = PicklePersistence(filepath=_persist_path)

    app = (ApplicationBuilder().token(TELEGRAM_TOKEN)
           .persistence(persistence)
           .post_init(_post_init)
           .read_timeout(120).write_timeout(120).connect_timeout(120).pool_timeout(120).build())

    # Sin esto, PTB escribe "No error handlers are registered" y la excepción
    # muere en el log: NADIE se entera. Pasó el 1-sep-2026 con un timeout de
    # Drive, y es lo más probable detrás del parte de Juan que se perdió en
    # silencio el 24-ago. Va primero para que cubra todo lo que se registre
    # después.
    from handlers.errores import manejar_error
    app.add_error_handler(manejar_error)

    # Tracker de actividad: corre PRIMERO para cada update (group=-1) y registra
    # el último mensaje procesado (heartbeat persistente).
    app.add_handler(TypeHandler(Update, _track_activity), group=-1)
    # Mirror: reenvía al dueño todo lo que mandan los demás.
    # OJO: en PTB solo corre UN handler por grupo → debe ir en un grupo propio
    # (en group=-1 junto al tracker jamás se ejecutaría).
    app.add_handler(TypeHandler(Update, mirror_update), group=-2)

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
    # Monitoreo / visibilidad
    app.add_handler(CommandHandler("soydueno", cmd_soydueno))
    app.add_handler(CommandHandler("estado",   cmd_estado))
    app.add_handler(CommandHandler("bodega",   cmd_bodega))
    app.add_handler(CommandHandler("correlativos", cmd_correlativos))
    app.add_handler(CommandHandler("basedatos", cmd_basedatos))
    # Conciliación bancaria
    app.add_handler(CommandHandler("conciliar", cmd_conciliar))
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
    # Maquinaria: horómetros, fichas y mantenciones
    from handlers.maquinaria import cb_maquinaria_fin, cmd_maquinaria
    app.add_handler(CommandHandler("maquinaria", cmd_maquinaria))
    app.add_handler(CommandHandler("maquina",    cmd_maquinaria))
    app.add_handler(CallbackQueryHandler(cb_maquinaria_fin, pattern="^maq_fin$"))
    # Inventario
    app.add_handler(CommandHandler("inventario", cmd_inventario))
    app.add_handler(CommandHandler("uso",      cmd_uso))
    # Personal y Vacaciones
    app.add_handler(CommandHandler("personal",  cmd_personal))
    app.add_handler(CommandHandler("vacaciones", cmd_vacaciones))
    app.add_handler(CommandHandler("vencimientos", cmd_vencimientos))
    app.add_handler(CommandHandler("agregar_trabajador", cmd_agregar_trabajador))
    # Google Drive: ver cola y reintentar lo rendido
    from handlers.drive_jobs import cmd_drive, cb_drive_reintentar
    app.add_handler(CommandHandler("drive", cmd_drive))
    app.add_handler(CallbackQueryHandler(cb_drive_reintentar,
                                          pattern="^drive_reintentar$"))
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
    app.add_handler(CallbackQueryHandler(cb_bita_save,         pattern="^bita_save$"))
    app.add_handler(CallbackQueryHandler(cb_bita_edit,         pattern="^bita_edit$"))
    app.add_handler(CallbackQueryHandler(cb_bita_cancel,       pattern="^bita_cancel$"))
    app.add_handler(CallbackQueryHandler(cb_cart_import,       pattern="^cart_import$"))
    app.add_handler(CallbackQueryHandler(cb_cart_cancel,       pattern="^cart_cancel$"))
    app.add_handler(CallbackQueryHandler(cb_conc_apply,        pattern="^conc_apply$"))
    app.add_handler(CallbackQueryHandler(cb_conc_cancel,       pattern="^conc_cancel$"))
    app.add_handler(CallbackQueryHandler(cb_venc_novence,      pattern="^venc_novence$"))
    app.add_handler(CallbackQueryHandler(cb_venc_skip,         pattern="^venc_skip$"))
    app.add_handler(CallbackQueryHandler(cb_venc_stop,         pattern="^venc_stop$"))
    app.add_handler(CallbackQueryHandler(cb_cultivo,           pattern="^cult_"))
    app.add_handler(CallbackQueryHandler(cb_select_item,       pattern="^selitem_"))
    app.add_handler(CallbackQueryHandler(cb_edit_field,        pattern="^edit_"))

    # Crear hojas si no existen
    try: crear_hoja_banco()
    except: pass
    try: crear_hojas_tareas()
    except: pass
    try:
        from bitacora_manager import crear_hoja_bitacora
        crear_hoja_bitacora()  # migra al esquema ampliado si es necesario
    except Exception as e:
        logger.warning(f"crear_hoja_bitacora: {e}")
    try: crear_hojas_inventario()
    except: pass
    try:
        from vencimientos_manager import crear_hoja_vencimientos
        crear_hoja_vencimientos()
    except Exception as e:
        logger.warning(f"crear_hoja_vencimientos: {e}")
    try: crear_hojas_vacaciones()
    except: pass

    # Sincronización bancaria: VIERNES 08:00 hora Chile, una vez por semana.
    # Antes corría 08:00 y 18:00 todos los días. El scraper es frágil (Akamai lo
    # tumbó en agosto-2026), así que en vez de insistir se intenta una vez y, si
    # falla, el bot pide la cartola — que es la vía que siempre funciona.
    # ⚠️ En PTB >= 20 `days` va de 0=domingo a 6=sábado: VIERNES ES 5.
    #    Con el mapeo viejo (0=lunes) days=(4,) era viernes; hoy sería jueves.
    try:
        from zoneinfo import ZoneInfo
    except ImportError:
        from backports.zoneinfo import ZoneInfo
    tz_chile = ZoneInfo("America/Santiago")
    app.job_queue.run_daily(job_sync_banco,
                            time=dtime(hour=8, minute=0, tzinfo=tz_chile),
                            days=(5,), name="banco_viernes")
    # Actualizar días de vacaciones el 1ro de cada mes
    app.job_queue.run_monthly(job_vacaciones_mensuales, when=dtime(hour=7, minute=0, tzinfo=tz_chile),
                              day=1, name="vacaciones_mensuales")
    # Resumen semanal cash flow: LUNES 08:00.
    # ⚠️ days=(1,) es lunes. Estuvo en days=(0,) — DOMINGO — desde el salto a
    # PTB 20+, que invirtió el mapeo. Corregido 2026-08-24.
    from handlers.cash_flow_jobs import job_resumen_semanal, job_reporte_mensual
    app.job_queue.run_daily(job_resumen_semanal,
                              time=dtime(hour=8, minute=0, tzinfo=tz_chile),
                              days=(1,), name="resumen_semanal")
    # Reporte mensual PDF: día 1 de cada mes, 08:00
    app.job_queue.run_monthly(job_reporte_mensual,
                              when=dtime(hour=8, minute=0, tzinfo=tz_chile),
                              day=1, name="reporte_mensual")
    # Heartbeat diario: 20:00 Chile → aviso "sigo vivo" + resumen al dueño
    app.job_queue.run_daily(job_heartbeat,
                            time=dtime(hour=20, minute=0, tzinfo=tz_chile),
                            name="heartbeat")
    # Latido silencioso cada 5 min: deja constancia de que el proceso vive.
    # Sin esto, un fin de semana sin mensajes se reportaba como "62h apagado".
    app.job_queue.run_repeating(job_latido, interval=300, first=10,
                                name="latido")
    # Vaciar la cola de subidas a Drive cada 10 min. Es barato: si la cola
    # está vacía, ni siquiera se autentica.
    from handlers.drive_jobs import job_drive_cola
    app.job_queue.run_repeating(job_drive_cola, interval=600, first=60,
                                name="drive_cola")
    # Revisar _Entrada de Drive cada 15 min: clasifica y mueve lo que llegue.
    from handlers.drive_jobs import job_drive_entrada
    app.job_queue.run_repeating(job_drive_entrada, interval=900, first=120,
                                name="drive_entrada")
    # Chequeo semanal del Excel de bodega vs Master: LUNES 08:30 (avisa solo si
    # NO calza). ⚠️ Mismo caso que el resumen: estuvo cayendo en DOMINGO por el
    # days=(0,) heredado de PTB < 20. Corregido 2026-08-24.
    app.job_queue.run_daily(job_bodega_check,
                            time=dtime(hour=8, minute=30, tzinfo=tz_chile),
                            days=(1,), name="bodega_check")
    # Sync diaria Excel → base de datos (modo paralelo): 21:00, avisa si no calza
    app.job_queue.run_daily(job_sync_db,
                            time=dtime(hour=21, minute=0, tzinfo=tz_chile),
                            name="sync_db")
    logger.info("⏰ Jobs programados: banco VIERNES 08:00, vacaciones día 1, "
                "resumen LUNES 08:00, reporte mensual día 1, heartbeat 20:00, "
                "bodega LUNES 08:30")

    logger.info("✅ Bot iniciado. Esperando mensajes...")
    # drop_pending_updates=False → procesa los mensajes que llegaron mientras
    # el bot estuvo apagado (Telegram los retiene hasta 24h con polling).
    app.run_polling(drop_pending_updates=False,
                    allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
