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

async def handle_text_edit(update, context):
    # ── Flujo /deposito ──
    if await handle_text_deposito(update, context):
        return

    # ── Flujo /pagado ──
    if await handle_text_pagado(update, context):
        return

    # ── Flujo /tarea ──
    if await handle_text_tarea(update, context):
        return

    # ── Flujo /bitacora ──
    if await handle_text_bitacora(update, context):
        return

    # ── Flujo /uso (inventario) ──
    if await handle_text_uso(update, context):
        return

    # ── Flujo /vacaciones ──
    if await handle_text_vacacion(update, context):
        return

    # ── Flujo /agregar_trabajador ──
    if await handle_text_trabajador(update, context):
        return

    # ── Edición de factura pendiente ──
    campo = context.user_data.get("editing_field")
    if not campo:
        # ── Chat inteligente (fallback) ──
        texto = update.message.text.strip()
        if len(texto) > 2:
            msg = await update.message.reply_text("💬 Pensando...")
            respuesta = await responder_chat(texto)
            try:
                await msg.edit_text(respuesta, parse_mode="Markdown")
            except Exception:
                await msg.edit_text(respuesta)
        return
    nuevo = update.message.text.strip()
    items = context.user_data.get("pending_items", [])
    if not items: await update.message.reply_text("⚠️ No hay factura pendiente."); return

    idx = context.user_data.get("editing_item_idx")  # None = todos, int = solo ese

    # Caso especial: editar total de factura
    if campo == "_total_factura":
        try:
            n = nuevo.replace("$", "").replace(" ", "")
            if "," in n:
                n = n.replace(".", "").replace(",", ".")
            elif n.count(".") > 1:
                n = n.replace(".", "")
            total_nuevo = float(n)
        except ValueError:
            await update.message.reply_text(f"❌ '{nuevo}' no es un número válido. Intenta de nuevo."); return
        # Registrar corrección del total factura
        valor_orig_tf = items[0].get("Total Factura") or items[0].get("Monto / TOTAL")
        _registrar_correccion(items[0], "Total Factura", valor_orig_tf, total_nuevo)
        # Detectar si es exenta (sin IVA)
        doc = str(items[0].get("Documento") or "").lower()
        exenta = any(k in doc for k in ("exenta", "exento", "no afecta", "no afecto"))
        iva_factor = 1.0 if exenta else 1.19

        # Cálculo inverso: Total → Neto → Unitario
        neto_nuevo = total_nuevo / iva_factor
        total_neto_actual = sum(float(i.get("Valor unitario") or 0) * float(i.get("Cantidad") or 1) for i in items)
        if total_neto_actual > 0:
            factor = neto_nuevo / total_neto_actual
            for item in items:
                unit = float(item.get("Valor unitario") or 0)
                qty = float(item.get("Cantidad") or 1)
                nuevo_unit = round(unit * factor, 2)
                item["Valor unitario"] = nuevo_unit
                item["Monto / TOTAL"] = round(nuevo_unit * qty * iva_factor)
        else:
            # No hay unitarios: poner todo el total
            items[0]["Monto / TOTAL"] = total_nuevo
            items[0]["Valor unitario"] = round(neto_nuevo)
        # Actualizar Total Factura en todos los ítems para que _build_preview lo refleje
        for item in items:
            item["Total Factura"] = round(total_nuevo)
        context.user_data["total_override"] = total_nuevo
        context.user_data["editing_field"] = None
        context.user_data["editing_item_idx"] = None
        await update.message.reply_text(f"✅ *Total factura* actualizado a `${total_nuevo:,.0f}`", parse_mode="Markdown")
        try:
            await update.message.reply_text(_build_preview(items), parse_mode="Markdown", reply_markup=_main_keyboard())
        except Exception:
            await update.message.reply_text(_build_preview(items), reply_markup=_main_keyboard())
        return

    targets = [items[idx]] if idx is not None else items

    NUMERICOS = {"Valor unitario", "Cantidad", "Monto / TOTAL", "TOTAL NETO"}
    for item in targets:
        valor_original = item.get(campo)  # capturar antes de sobrescribir
        if campo in NUMERICOS:
            try:
                # Formato chileno: punto = miles, coma = decimal
                # "7.164,32" → 7164.32 | "105,612" → 105.612 | "500000" → 500000
                n = nuevo.replace("$", "").replace(" ", "")
                if "," in n:
                    n = n.replace(".", "").replace(",", ".")
                else:
                    if n.count(".") > 1:
                        n = n.replace(".", "")
                val = float(n)
                item[campo] = val
                if campo == "Monto / TOTAL":
                    # Cálculo inverso: Total ítem (con IVA) → Valor unitario
                    qty = float(item.get("Cantidad") or 1)
                    imp_esp = float(item.get("Impuesto Especifico") or 0)
                    doc = str(item.get("Documento") or "").lower()
                    sin_iva = any(k in doc for k in ("exenta", "exento", "no afecta", "no afecto", "boleta de honorario"))
                    iva_factor = 1.0 if sin_iva else 1.19
                    neto_nuevo = (val - imp_esp) / iva_factor
                    item["Valor unitario"] = neto_nuevo / qty if qty else 0
                elif campo == "TOTAL NETO":
                    # El usuario ingresó el neto de línea (precio × cantidad, sin IVA)
                    qty = float(item.get("Cantidad") or 1)
                    imp_esp = float(item.get("Impuesto Especifico") or 0)
                    doc = str(item.get("Documento") or "").lower()
                    sin_iva = any(k in doc for k in ("exenta", "exento", "no afecta", "no afecto", "boleta de honorario"))
                    iva_factor = 1.0 if sin_iva else 1.19
                    item["Valor unitario"] = val / qty if qty else 0
                    item["Monto / TOTAL"] = round(val * iva_factor + imp_esp)
            except ValueError:
                await update.message.reply_text(f"❌ '{nuevo}' no es un número válido. Intenta de nuevo."); return
        else:
            item[campo] = nuevo
        # Registrar corrección solo si el valor cambió
        if str(valor_original) != str(item.get(campo)):
            _registrar_correccion(item, campo, valor_original, item.get(campo))

    # Si fue edición numérica, propagar a Monto/TOTAL y Total Factura para que el preview calce
    if campo in NUMERICOS:
        doc0 = str(items[0].get("Documento") or "").lower()
        sin_iva0 = any(k in doc0 for k in ("exenta", "exento", "no afecta", "no afecto", "boleta de honorario"))
        iva_factor0 = 1.0 if sin_iva0 else 1.19
        # Si el usuario editó Valor unitario o Cantidad del/los ítem(s), refrescar su Monto/TOTAL
        if campo not in ("Monto / TOTAL", "TOTAL NETO"):
            for it in targets:
                u = float(it.get("Valor unitario") or 0)
                q = float(it.get("Cantidad") or 1)
                imp = float(it.get("Impuesto Especifico") or 0)
                it["Monto / TOTAL"] = round(u * q * iva_factor0 + imp)
        # Sumar Monto/TOTAL de todos los ítems → nuevo Total Factura
        nuevo_tf = round(sum(float(it.get("Monto / TOTAL") or 0) for it in items))
        if nuevo_tf > 0:
            for it in items:
                it["Total Factura"] = nuevo_tf

    context.user_data["editing_field"] = None
    context.user_data["editing_item_idx"] = None
    label = context.user_data.get("editing_field_label", campo)
    suffix = f" (Ítem {idx+1})" if idx is not None else ""
    await update.message.reply_text(f"✅ *{label}*{suffix} actualizado.", parse_mode="Markdown")
    try:
        await update.message.reply_text(_build_preview(items), parse_mode="Markdown", reply_markup=_main_keyboard())
    except Exception:
        await update.message.reply_text(_build_preview(items), reply_markup=_main_keyboard())

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
