"""handlers/finanzas.py — Comandos financieros y sync banco.

Comandos: /pagado, /deposito, /saldo, /dashboard, /reporte, /banco
Callbacks: cb_reporte (rep_*), cb_medio_pago (pago_*), cb_calce_verificacion (calce_*)
Jobs: job_sync_banco
Flujos texto: handle_text_pagado, handle_text_deposito
"""
import asyncio
import logging
import os
import sys
from datetime import date, datetime

from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes

from config import TELEGRAM_CHAT_ID
from scotiabank_scraper import CaptchaRequerido
from excel_manager import (
    buscar_factura, registrar_pago, registrar_deposito_caja,
    consultar_saldo_caja, reporte_diario, reporte_semanal, reporte_mensual,
    guardar_movimientos_banco, obtener_resumen_banco,
)
from utils.formatting import esc
from utils.parsing import parsear_fecha, parsear_monto

logger = logging.getLogger(__name__)


# ── Comandos ─────────────────────────────────────

async def cmd_pagado(update, context):
    context.user_data["pagado_state"] = "esperando_nro"
    context.user_data["editing_field"] = None
    await update.message.reply_text(
        "💳 *Registrar pago de factura*\n\n"
        "Escribe el *número de factura* que fue pagada\n"
        "(o /cancelar para salir):",
        parse_mode="Markdown")


async def cmd_deposito(update, context):
    context.user_data["deposito_state"] = "esperando_monto"
    context.user_data["editing_field"] = None
    context.user_data["pagado_state"] = None
    await update.message.reply_text(
        "💰 *Depósito a Caja Chica*\n\n"
        "Escribe el *monto* del depósito (ej: 700000):\n"
        "(o /cancelar para salir)",
        parse_mode="Markdown")


async def cmd_saldo(update, context):
    info = await asyncio.to_thread(consultar_saldo_caja)
    saldo = info["saldo"]
    alerta = "\n⚠️ *Saldo bajo, considerar nuevo depósito*" if saldo < 100000 else ""
    await update.message.reply_text(
        f"💰 *Estado Caja Chica*\n\n"
        f"📊 Saldo actual: *${saldo:,.0f}*{alerta}\n\n"
        f"📥 Total depositado: ${info['total_ingresos']:,.0f}\n"
        f"📤 Total gastado: ${info['total_egresos']:,.0f}\n"
        f"🧾 Gastos registrados: {info['n_gastos']}\n"
        f"📅 Último depósito: {info.get('ultimo_deposito') or '—'}",
        parse_mode="Markdown")


async def cmd_dashboard(update, context):
    """Inicia el dashboard web y envía el link."""
    import subprocess
    import urllib.request
    port = 5000
    try:
        urllib.request.urlopen(f"http://localhost:{port}/", timeout=2)
        running = True
    except Exception:
        running = False

    if not running:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        dashboard_script = os.path.join(base_dir, "src", "dashboard.py")
        subprocess.Popen([sys.executable, dashboard_script],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                          creationflags=0x00000008 if os.name == 'nt' else 0)
        await asyncio.sleep(2)

    await update.message.reply_text(
        f"📊 *Dashboard Agricola Santa Elisa*\n\n"
        f"🌐 Abrir en el navegador:\n"
        f"`http://localhost:{port}`\n\n"
        f"Pestanas disponibles:\n"
        f"• General — Resumen del campo\n"
        f"• Facturas — Estado de pagos\n"
        f"• Costos por Cultivo — Nogales, Cerezos, Avellanos\n"
        f"• Exportaciones — Espana, comparacion anual\n"
        f"• Finanzas — Banco y Caja Chica\n"
        f"• Conciliación — `http://localhost:{port}/conciliacion`\n"
        f"• Inventario — Stock de insumos\n"
        f"• Personal — Vacaciones pendientes\n"
        f"• Tareas — Bitacora y seguimiento\n"
        f"• Flujo de Caja — Proyeccion + simulador replante",
        parse_mode="Markdown")


async def cmd_reporte(update, context):
    kb = InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Diario", callback_data="rep_diario"),
         InlineKeyboardButton("📊 Semanal", callback_data="rep_semanal")],
        [InlineKeyboardButton("📈 Mensual", callback_data="rep_mensual")]])
    await update.message.reply_text(
        "📊 *Reportes*\n\nSelecciona el tipo de reporte:",
        parse_mode="Markdown", reply_markup=kb)


async def cmd_banco(update, context):
    """Sincroniza movimientos de Scotiabank y los guarda en Cuenta Banco."""
    chat_id = update.effective_chat.id
    context.bot_data["banco_chat_id"] = chat_id
    msg = await update.message.reply_text(
        "🏦 Conectando con Scotiabank...\nEsto puede tomar 30-60 segundos.")
    try:
        texto = await _sync_banco_core()
        try:
            await msg.edit_text(texto, parse_mode="Markdown")
        except Exception:
            await msg.edit_text(texto)
    except CaptchaRequerido:
        # El bot no resuelve CAPTCHAs: se explica la vía manual, que sí funciona.
        await msg.edit_text(
            "🔐 *El banco pide un CAPTCHA para entrar*\n\n"
            "No puedo pasarlo: esa verificación está puesta para impedir el "
            "acceso automático y saltarla arriesga que bloqueen la cuenta.\n\n"
            "*Hagámoslo así:*\n"
            "1️⃣ Entra al portal y descarga la cartola\n"
            "2️⃣ Mándame el archivo por acá (CSV, TXT o Excel)\n"
            "3️⃣ Te muestro un preview y lo importo sin duplicar",
            parse_mode="Markdown")
    except ImportError:
        await msg.edit_text("❌ Playwright no está instalado.\n"
                             "Ejecuta: pip install playwright && playwright install chromium")
    except RuntimeError as e:
        await msg.edit_text(f"❌ Error: {str(e)}")
    except Exception as e:
        logger.error(f"Error sync banco: {e}")
        await msg.edit_text(f"❌ Error al sincronizar con el banco:\n{str(e)[:200]}")


# ── Core sync banco ─────────────────────────────

async def _sync_banco_core() -> str:
    """Logica compartida de sincronizacion bancaria. Retorna texto del resultado."""
    from scotiabank_scraper import sync_scotiabank_movements
    movimientos = await asyncio.to_thread(sync_scotiabank_movements)
    if not movimientos:
        return "🏦 Conexión exitosa pero no se encontraron movimientos nuevos."
    result = await asyncio.to_thread(guardar_movimientos_banco, movimientos)
    resumen = await asyncio.to_thread(obtener_resumen_banco)
    texto = f"🏦 *Scotiabank — Sincronización completa*\n\n"
    texto += f"📥 Movimientos obtenidos: *{result['total']}*\n"
    texto += f"✅ Nuevos guardados: *{result['nuevos']}*\n"
    texto += f"⏭️ Duplicados omitidos: *{result['duplicados']}*\n\n"
    if resumen["ultimo_saldo"] is not None:
        texto += f"💰 Último saldo: *${resumen['ultimo_saldo']:,.0f}*\n"
    texto += f"📊 Total en hoja: *{resumen['n_movimientos']}* movimientos\n"
    texto += f"   Cargos: ${resumen['total_cargos']:,.0f}\n"
    texto += f"   Abonos: ${resumen['total_abonos']:,.0f}"
    return texto


def _aviso_banco_manual(hora: str, motivo: str = "error",
                         detalle: str = "") -> str:
    """Mensaje para cuando el scraper no pudo actualizar el banco.

    Pase lo que pase termina pidiendo la cartola: es la vía que sí funciona, y
    la única salida cuando el banco pone verificación anti-robot. Antes el
    camino del error genérico solo reportaba la falla y dejaba al dueño sin
    saber qué hacer.
    """
    detalle = (detalle or "")[:400]
    low = detalle.lower()
    if motivo == "captcha":
        cabecera = (
            f"🔐 *El banco pidió verificación anti-robot* ({hora})\n\n"
            "No la voy a saltar: existe justamente para impedir el acceso "
            "automático, y forzarla arriesga que te *bloqueen la cuenta*.")
        pista = ""
    else:
        cabecera = (f"🔴 *No pude actualizar el banco solo* ({hora})\n\n"
                    f"`{detalle}`")
        if any(k in low for k in ("no encontré el campo", "selector", "timeout")):
            pista = ("\n\n💡 Puede que el banco haya cambiado su *página*. "
                     "Avísame y reviso los selectores.")
        elif "credenciales" in low or "login" in low:
            pista = "\n\n💡 Revisa las credenciales del banco en el `.env`."
        elif any(k in low for k in ("executable", "playwright", "browser")):
            pista = ("\n\n💡 Falta el navegador: corre "
                     "`python -m playwright install chromium`.")
        else:
            pista = ""

    return (
        f"{cabecera}{pista}\n\n"
        "*Mándame la cartola y la subo yo:*\n"
        "1️⃣ Entra al portal y descarga la cartola\n"
        "2️⃣ Mándame el archivo por acá (Excel, CSV o TXT)\n"
        "3️⃣ Te muestro un preview y la importo sin duplicar nada\n\n"
        "_Vuelvo a intentarlo solo el próximo viernes._")


async def job_sync_banco(context: ContextTypes.DEFAULT_TYPE):
    """Job SEMANAL (viernes am): sincroniza el banco y avisa al dueño.

    Corre una vez por semana porque el scraper es frágil: si falla, el camino
    bueno es que el dueño mande la cartola, no que el bot siga reintentando.
    """
    chat_id = (context.bot_data.get("owner_chat_id")
               or context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID)
    if not chat_id:
        logger.warning("Job banco: no hay chat_id configurado.")
        return
    chat_id = int(chat_id)
    hora = datetime.now().strftime("%H:%M")
    try:
        texto = await _sync_banco_core()
        texto = f"⏰ *Sincronización automática ({hora})*\n\n" + texto
        try:
            await context.bot.send_message(chat_id=chat_id, text=texto,
                                            parse_mode="Markdown")
        except Exception:
            await context.bot.send_message(chat_id=chat_id, text=texto)
    except CaptchaRequerido as e:
        # El banco puso verificación anti-robot: el scraper no va a volver a
        # funcionar solo. No se reintenta; se pide la cartola, que sí funciona.
        logger.warning(f"Job banco: CAPTCHA en el login — {e}")
        await _avisar_caida(context, chat_id, hora, "captcha", str(e))
    except Exception as e:
        # Si la sincronización falla hay que ENTERARSE, no solo dejarlo en el log:
        # así fue como pasó desapercibido que el banco cambió su página de login.
        logger.error(f"Job banco falló: {e}")
        await _avisar_caida(context, chat_id, hora, "error", str(e))


async def _avisar_caida(context, chat_id: int, hora: str,
                         motivo: str, detalle: str) -> None:
    """Manda el aviso pidiendo la cartola, con fallback sin Markdown."""
    aviso = _aviso_banco_manual(hora, motivo=motivo, detalle=detalle)
    try:
        await context.bot.send_message(chat_id=chat_id, text=aviso,
                                        parse_mode="Markdown")
    except Exception:
        # Markdown roto por el detalle del error: mandarlo plano igual, porque
        # quedarse callado es peor que mandarlo feo.
        await context.bot.send_message(
            chat_id=chat_id,
            text=("No pude actualizar el banco solo. Descarga la cartola del "
                  "portal y mándamela por acá para importarla."))


# ── Callback /reporte ────────────────────────────

async def cb_reporte(update, context):
    query = update.callback_query
    await query.answer()
    tipo = query.data

    if tipo == "rep_diario":
        await query.edit_message_text("⏳ Generando reporte diario...")
        r = await asyncio.to_thread(reporte_diario)
        texto = f"📋 *REPORTE DIARIO — {r['fecha']}*\n\n"
        if r["vencen_hoy"]:
            texto += f"⚠️ *Vencen HOY ({len(r['vencen_hoy'])}):*\n"
            for f in r["vencen_hoy"][:10]:
                texto += (f"  {esc(f.get('proveedor'))} — Nº {esc(f.get('nro_factura'))} "
                           f"— ${float(f.get('total') or 0):,.0f}\n")
            texto += f"  *Total: ${r['total_hoy']:,.0f}*\n\n"
        else:
            texto += "✅ No hay facturas que venzan hoy.\n\n"
        if r["vencidas"]:
            texto += f"🔴 *Vencidas sin pagar ({len(r['vencidas'])}):*\n"
            for f in r["vencidas"][:15]:
                texto += (f"  {esc(f.get('proveedor'))} — Nº {esc(f.get('nro_factura'))} "
                           f"— ${float(f.get('total') or 0):,.0f}\n")
            texto += f"  *Total vencido: ${r['total_vencido']:,.0f}*\n"
        else:
            texto += "✅ No hay facturas vencidas pendientes."

    elif tipo == "rep_semanal":
        await query.edit_message_text("⏳ Generando reporte semanal...")
        r = await asyncio.to_thread(reporte_semanal)
        texto = f"📊 *REPORTE SEMANAL — {r['periodo']}*\n\n"
        texto += f"✅ *Pagadas esta semana:* {len(r['pagadas'])} — ${r['total_pagado']:,.0f}\n"
        if r["pagadas"]:
            for f in r["pagadas"][:10]:
                texto += f"  {esc(f.get('proveedor'))} — ${float(f.get('total') or 0):,.0f}\n"
        texto += f"\n🔴 *Vencidas sin pagar:* {len(r['vencidas'])} — ${r['total_vencido']:,.0f}\n"
        if r["vencidas"]:
            for f in r["vencidas"][:10]:
                texto += f"  {esc(f.get('proveedor'))} — Nº {esc(f.get('nro_factura'))}\n"
        texto += f"\n💵 *Caja Chica esta semana:* {len(r['gastos_caja'])} gastos — ${r['total_caja']:,.0f}\n"
        if r["gastos_caja"]:
            for g in r["gastos_caja"][:10]:
                texto += f"  {esc(g.get('comercio') or g.get('detalle'))} — ${g.get('monto', 0):,.0f}\n"

    elif tipo == "rep_mensual":
        await query.edit_message_text("⏳ Generando reporte mensual...")
        r = await asyncio.to_thread(reporte_mensual)
        texto = f"📈 *REPORTE MENSUAL — {r['periodo']}*\n\n"
        texto += f"📄 Facturas recibidas: *{r['total_emitidas']}* — ${r['monto_emitido']:,.0f}\n"
        texto += f"✅ Facturas pagadas: *{r['total_pagadas']}* — ${r['monto_pagado']:,.0f}\n"
        texto += f"🔴 Vencidas sin pagar: *{r['vencidas_sin_pago']}* — ${r['monto_vencido']:,.0f}\n"
        texto += f"💵 Caja chica: *{r['n_gastos_caja']}* gastos — ${r['total_caja_mes']:,.0f}\n"
        if r["top_proveedores"]:
            texto += f"\n🏢 *Top proveedores por gasto:*\n"
            for i, (prov, monto) in enumerate(r["top_proveedores"], 1):
                texto += f"  {i}. {esc(prov)} — ${monto:,.0f}\n"
    else:
        texto = "❌ Tipo de reporte no reconocido."

    try:
        await query.edit_message_text(texto, parse_mode="Markdown")
    except Exception:
        await query.edit_message_text(texto)


# ── Callbacks medio pago + calce ────────────────

async def cb_medio_pago(update, context):
    """Callback para seleccionar medio de pago (Banco/Caja Chica)."""
    query = update.callback_query
    await query.answer()
    medio = "Banco" if query.data == "pago_banco" else "Caja Chica"
    nro = context.user_data.get("pagado_nro")
    fecha = context.user_data.get("pagado_fecha")
    resultado = await asyncio.to_thread(registrar_pago, nro, fecha, medio)
    context.user_data["pagado_state"] = None
    context.user_data["pagado_nro"] = None
    context.user_data["pagado_fecha"] = None

    actualizadas = resultado.get("actualizadas", 0) if isinstance(resultado, dict) else resultado
    calce = resultado.get("calce") if isinstance(resultado, dict) else None

    if actualizadas <= 0:
        await query.edit_message_text("❌ Error al actualizar el Excel. ¿Está abierto?")
        return

    icono = "🏦" if medio == "Banco" else "💵"
    texto = (
        f"✅ *Pago registrado*\n\n"
        f"📄 Factura Nº {esc(nro)}\n"
        f"📅 Fecha: {fecha}\n"
        f"{icono} Medio: {medio}\n"
        f"📊 {actualizadas} fila(s) actualizada(s)")

    if calce:
        if calce["exacto"]:
            texto += (
                f"\n\n🔗 *Calce bancario encontrado:*\n"
                f"📅 {calce['fecha_banco']}\n"
                f"📝 {esc(calce['descripcion_banco'])}\n"
                f"💰 ${calce['cargo_banco']:,.0f} ✅ Monto exacto")
        else:
            texto += (
                f"\n\n🔗 *Posible calce bancario:*\n"
                f"📅 {calce['fecha_banco']}\n"
                f"📝 {esc(calce['descripcion_banco'])}\n"
                f"💰 ${calce['cargo_banco']:,.0f}\n"
                f"⚠️ Calce parcial — verifica manualmente")
            if calce["total_candidatos"] > 1:
                texto += f"\n📊 {calce['total_candidatos']} movimientos similares encontrados"
        context.user_data["ultimo_calce"] = calce
        context.user_data["ultimo_calce_nro"] = nro
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Calce correcto", callback_data="calce_ok"),
             InlineKeyboardButton("❌ No es este", callback_data="calce_no")]])
        try:
            await query.edit_message_text(texto, parse_mode="Markdown", reply_markup=kb)
        except Exception:
            await query.edit_message_text(texto, reply_markup=kb)
    else:
        if medio == "Banco":
            texto += "\n\n🔍 No se encontró calce bancario automático."
        try:
            await query.edit_message_text(texto, parse_mode="Markdown")
        except Exception:
            await query.edit_message_text(texto)


async def cb_calce_verificacion(update, context):
    """Callback para confirmar o rechazar calce bancario."""
    query = update.callback_query
    await query.answer()
    calce = context.user_data.get("ultimo_calce")
    nro = context.user_data.get("ultimo_calce_nro", "?")

    if query.data == "calce_ok":
        msg = (f"✅ *Calce confirmado*\n\n"
               f"📄 Factura Nº {esc(nro)} vinculada a:\n"
               f"🏦 {esc(calce['descripcion_banco'] if calce else '—')}\n"
               f"💰 ${calce['cargo_banco']:,.0f}" if calce else "✅ Calce confirmado")
        await query.edit_message_text(msg, parse_mode="Markdown")
    else:
        await query.edit_message_text(
            f"📝 *Calce descartado*\n\n"
            f"Factura Nº {esc(nro)} marcada como pagada.\n"
            f"El calce bancario no se aplicó — puedes revisarlo manualmente.",
            parse_mode="Markdown")

    context.user_data["ultimo_calce"] = None
    context.user_data["ultimo_calce_nro"] = None


# ── Flujos texto ────────────────────────────────

async def handle_text_deposito(update, context) -> bool:
    """Flujo /deposito (monto -> fecha). Devuelve True si lo manejo."""
    dep_state = context.user_data.get("deposito_state")
    if not dep_state:
        return False
    texto = update.message.text.strip()

    if dep_state == "esperando_monto":
        monto = parsear_monto(texto)
        if monto is None:
            await update.message.reply_text("❌ No es un monto válido. Intenta de nuevo:")
            return True
        context.user_data["deposito_monto"] = monto
        context.user_data["deposito_state"] = "esperando_fecha_dep"
        await update.message.reply_text(
            f"💰 Monto: *${monto:,.0f}*\n\n"
            f"📅 Escribe la *fecha del depósito* (DD/MM/YYYY)\n"
            f"o escribe *hoy* para usar la fecha de hoy:",
            parse_mode="Markdown")
        return True

    if dep_state == "esperando_fecha_dep":
        fecha = parsear_fecha(texto)
        if not fecha:
            await update.message.reply_text(
                "❌ Formato no válido. Usa DD/MM/YYYY o escribe *hoy*:",
                parse_mode="Markdown")
            return True
        monto = context.user_data.get("deposito_monto", 0)
        nuevo_saldo = await asyncio.to_thread(registrar_deposito_caja, fecha, monto)
        context.user_data["deposito_state"] = None
        context.user_data["deposito_monto"] = None
        if nuevo_saldo >= 0:
            await update.message.reply_text(
                f"✅ *Depósito registrado*\n\n"
                f"💰 Monto: ${monto:,.0f}\n"
                f"📅 Fecha: {fecha}\n"
                f"📊 *Nuevo saldo caja chica: ${nuevo_saldo:,.0f}*",
                parse_mode="Markdown")
        else:
            await update.message.reply_text("❌ Error al registrar. ¿Está el Excel abierto?")
        return True

    return False


async def handle_text_pagado(update, context) -> bool:
    """Flujo /pagado (nro -> fecha -> medio). Devuelve True si lo manejo."""
    pagado_state = context.user_data.get("pagado_state")
    if not pagado_state:
        return False
    texto = update.message.text.strip()

    if pagado_state == "esperando_nro":
        resultados = await asyncio.to_thread(buscar_factura, texto)
        if not resultados:
            await update.message.reply_text(
                f"❌ No encontré factura Nº *{esc(texto)}* en el Excel.\n"
                "Verifica el número e intenta de nuevo, o escribe /cancelar",
                parse_mode="Markdown")
            return True
        context.user_data["pagado_nro"] = texto
        context.user_data["pagado_state"] = "esperando_fecha"
        r = resultados[0]
        ya_pagada = (f"\n⚠️ *Ya tiene fecha de pago:* {r['fecha_pago']}"
                      if r.get("fecha_pago") else "")
        await update.message.reply_text(
            f"📄 *Factura encontrada:*\n\n"
            f"🏢 {esc(r.get('proveedor') or '—')}\n"
            f"🪪 RUT: {esc(r.get('rut') or '—')}\n"
            f"📄 {esc(r.get('documento') or '—')} Nº {esc(r.get('nro_factura'))}\n"
            f"📦 {esc(r.get('glosa') or '—')}\n"
            f"💰 Total: ${float(r.get('total') or 0):,.0f}\n"
            f"({len(resultados)} fila(s) en Excel){ya_pagada}\n\n"
            f"📅 Escribe la *fecha de pago* (DD/MM/YYYY o YYYY-MM-DD):",
            parse_mode="Markdown")
        return True

    if pagado_state == "esperando_fecha":
        fecha = parsear_fecha(texto)
        if not fecha:
            await update.message.reply_text(
                "❌ No entendí la fecha. Escribe en formato *DD/MM/YYYY* (ej: 15/03/2026):",
                parse_mode="Markdown")
            return True
        context.user_data["pagado_fecha"] = fecha
        context.user_data["pagado_state"] = "esperando_medio"
        kb = InlineKeyboardMarkup([
            [InlineKeyboardButton("🏦 Banco", callback_data="pago_banco"),
             InlineKeyboardButton("💵 Caja Chica", callback_data="pago_caja_chica")]])
        await update.message.reply_text(
            f"📅 Fecha: *{fecha}*\n\n💳 *¿Medio de pago?*",
            parse_mode="Markdown", reply_markup=kb)
        return True

    return False
