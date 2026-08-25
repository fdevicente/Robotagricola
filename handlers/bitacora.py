"""handlers/bitacora.py — Registro de bitácora con IA + confirmación.

Flujo:
  /bitacora <texto>   → la IA estructura → preview → [Confirmar][Corregir][Cancelar]
  /bitacora (solo)    → pide el texto
Al confirmar: guarda en hoja Bitácora y, si hay insumo+cantidad, descuenta
del inventario (registra aplicación con sector/responsable).
"""
import asyncio
import logging
from datetime import date

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

logger = logging.getLogger(__name__)


def _esc(t):
    if t is None:
        return ""
    for ch in r"_*[]()~`>#+-=|{}.!":
        t = str(t).replace(ch, "\\" + ch)
    return t


def _kb_confirmar():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Confirmar", callback_data="bita_save"),
        InlineKeyboardButton("✏️ Corregir", callback_data="bita_edit"),
        InlineKeyboardButton("❌ Cancelar", callback_data="bita_cancel"),
    ]])


def _build_preview(campos: dict) -> str:
    ICON = {"LABOR": "🛠️", "APLICACION": "🧪", "RIEGO": "💧",
            "MAQUINARIA": "🚜", "EVENTO": "🌧️", "OTRO": "📝"}
    icon = ICON.get(campos.get("tipo", "OTRO"), "📝")
    lines = [f"📓 *Entendí esto:*", ""]
    lines.append(f"{icon} *{_esc(campos.get('actividad') or campos.get('tipo'))}*"
                 f"{' · ' + _esc(campos['cultivo']) if campos.get('cultivo') and campos['cultivo'] != 'GENERAL' else ''}")
    if campos.get("sector"):
        lines.append(f"📍 Sector: {_esc(campos['sector'])}")
    jh = campos.get("jornadas_hombre")
    trab = campos.get("trabajadores") or []
    if jh:
        det = f" ({', '.join(_esc(t) for t in trab)})" if trab else ""
        lines.append(f"👷 *{jh} jornada{'s' if jh != 1 else ''}\\-hombre*{det}")
    elif trab:
        lines.append(f"👷 {', '.join(_esc(t) for t in trab)}")
    if campos.get("maquina"):
        odo = campos.get("odometro")
        odo_txt = f" · odómetro {odo:g}" if odo not in (None, "") else ""
        lines.append(f"🚜 Máquina: {_esc(campos['maquina'])}{odo_txt}")
    if campos.get("superficie_ha"):
        lines.append(f"📐 Superficie: {campos['superficie_ha']:g} ha")
    if campos.get("insumo"):
        cant = campos.get("cantidad")
        uni = campos.get("unidad") or ""
        cant_txt = f" — {cant:g} {_esc(uni)}" if cant else ""
        lines.append(f"🧪 Insumo: {_esc(campos['insumo'])}{cant_txt} "
                     f"_\\(descontará del inventario\\)_" if cant else
                     f"🧪 Insumo: {_esc(campos['insumo'])}")
    conf = campos.get("confianza", 0)
    if conf < 0.6:
        lines.append("")
        lines.append("⚠️ _No estoy muy seguro, revisa bien._")
    lines.append("")
    lines.append("¿Confirmo el registro?")
    return "\n".join(lines)


async def _procesar_texto_bitacora(update_or_msg, context, texto, registrado_por):
    """Extrae con IA y muestra el preview. Recibe un objeto con .reply_text/.edit_text."""
    from modules.bitacora_extractor import extraer_bitacora
    status = await update_or_msg.reply_text("📓 Interpretando el registro…")
    try:
        campos = await asyncio.to_thread(
            extraer_bitacora, texto, date.today().strftime("%Y-%m-%d"))
    except Exception as e:
        logger.error(f"Error extrayendo bitácora: {e}")
        await status.edit_text("❌ No pude interpretar el registro. Intenta de nuevo.")
        return
    if campos.get("error"):
        await status.edit_text("❌ La IA no está disponible ahora. Intenta más tarde.")
        return

    context.user_data["bitacora_pending"] = campos
    context.user_data["bitacora_registrado_por"] = registrado_por
    context.user_data["bitacora_state"] = None
    try:
        await status.edit_text(_build_preview(campos), parse_mode="MarkdownV2",
                                reply_markup=_kb_confirmar())
    except Exception:
        # fallback sin markdown
        await status.edit_text(_build_preview(campos).replace("\\", ""),
                                reply_markup=_kb_confirmar())


async def auto_guardar_bitacora(update, context):
    """MODO CAPATAZ: guarda el texto como bitácora directo, sin confirmar.

    Si la IA no está disponible, guarda el texto crudo como OTRO para NO perderlo.
    Descuenta inventario si detecta insumo+cantidad y calcula horas si es maquinaria.
    """
    texto = (update.message.text or "").strip()
    if len(texto) < 3:
        return
    # Un mensaje que es SOLO un comando ("Bitácora", "Personal", "Asistencia /")
    # no es un registro. Antes pasaba el filtro de largo y se guardaba como OTRO
    # con la propia palabra como actividad: se limpiaron a mano el 10 y el
    # 18-ago-2026 y volvieron 3 al día siguiente.
    from modules.bitacora_extractor import es_mensaje_sin_contenido
    if es_mensaje_sin_contenido(texto):
        await update.message.reply_text(
            "📓 Te leo, pero ahí no viene nada que anotar.\n\n"
            "Mándame el parte completo, por ejemplo:\n"
            "*Martes 18 de agosto 2026*\n"
            "Felicito amigo : poda nogales\n"
            "Ramiro amigo : sacar restos\n\n"
            "O una lectura: *Tractor 6711, horómetro término 2041*",
            parse_mode="Markdown")
        return
    registrado_por = update.effective_user.full_name if update.effective_user else ""
    status = await update.message.reply_text("📓 Anotando en bitácora…")

    from modules.bitacora_extractor import extraer_bitacora
    campos = None
    try:
        campos = await asyncio.to_thread(
            extraer_bitacora, texto, date.today().strftime("%Y-%m-%d"))
    except Exception as e:
        logger.error(f"Auto-bitácora: error extrayendo: {e}")
    if not campos or campos.get("error"):
        # Nunca perder la entrada: guardar crudo como OTRO
        campos = {"tipo": "OTRO", "actividad": texto[:40], "cultivo": "GENERAL",
                  "sector": "", "jornadas_hombre": None, "trabajadores": [],
                  "insumo": "", "cantidad": None, "unidad": "",
                  "maquina": "", "odometro": None, "superficie_ha": None,
                  "confianza": 0.0, "texto_original": texto}

    # ── Parte de asistencia (varios trabajadores) → una fila por ACTIVIDAD ──
    # Un mismo mensaje puede traer VARIOS días (Juan a veces reporta la semana
    # junta); cada bloque conserva la fecha de su encabezado.
    from modules.bitacora_asistencia import (parsear_asistencia_multi,
                                              cultivo_de, tipo_de)
    dias = parsear_asistencia_multi(texto)
    if dias:
        from bitacora_manager import registrar_bitacora_estructurada as _reg
        fecha_ia = campos.get("fecha") or ""
        partes, jh_global, n_labores = [], 0, 0
        for d in dias:
            fecha = (d["fecha"].strftime("%Y-%m-%d") if d["fecha"]
                     else (fecha_ia if len(dias) == 1 else ""))
            guardadas, jh_dia = [], 0
            for g in d["grupos"]:
                sub = {
                    "fecha": fecha, "tipo": tipo_de(g["actividad"]),
                    "actividad": g["actividad"], "cultivo": cultivo_de(g["actividad"]),
                    "sector": "", "jornadas_hombre": g["jornadas_hombre"],
                    "trabajadores": g["trabajadores"], "insumo": "", "cantidad": None,
                    "unidad": "", "maquina": "", "odometro": None, "superficie_ha": None,
                    "texto_original": texto,
                }
                try:
                    await asyncio.to_thread(_reg, sub, registrado_por)
                    guardadas.append(g)
                    jh_dia += g["jornadas_hombre"] or 0
                except Exception as e:
                    logger.error(f"Auto-bitácora asistencia: {e}")
            if not guardadas:
                continue
            jh_global += jh_dia
            n_labores += len(guardadas)
            f_txt = fecha or date.today().strftime("%Y-%m-%d")
            det = "\n".join(
                f"• {g['actividad']}: {g['jornadas_hombre'] or 0} "
                f"({', '.join(n.split()[0] for n in g['trabajadores'])})"
                for g in guardadas)
            partes.append(f"📅 *{f_txt}* — {jh_dia} JH\n{det}"
                          if len(dias) > 1 else f"{f_txt}|{det}")
        if partes:
            if len(dias) > 1:
                cuerpo = "\n\n".join(partes).replace("*", "")
                msg = (f"✅ Asistencia anotada — {len(partes)} días\n"
                       f"👷 {jh_global} jornadas-hombre en {n_labores} labores\n\n{cuerpo}")
            else:
                f_txt, det = partes[0].split("|", 1)
                msg = (f"✅ Asistencia anotada — {f_txt}\n"
                       f"👷 {jh_global} jornadas-hombre en {n_labores} labores\n\n{det}")
            try:
                await status.edit_text(msg)
            except Exception:
                await status.edit_text(f"✅ Asistencia anotada ({jh_global} JH)")
            return

    from bitacora_manager import registrar_bitacora_estructurada
    try:
        res = await asyncio.to_thread(registrar_bitacora_estructurada, campos, registrado_por)
    except Exception as e:
        # Excel bloqueado u otro fallo: NUNCA perder la entrada → respaldo a archivo
        logger.error(f"Auto-bitácora: fallo guardando en Excel: {e}")
        try:
            import os
            fb_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                  "files", "logs")
            os.makedirs(fb_dir, exist_ok=True)
            from datetime import datetime as _dt
            with open(os.path.join(fb_dir, "bitacora_fallback.txt"), "a", encoding="utf-8") as fh:
                fh.write(f"{_dt.now().isoformat()} | {registrado_por} | {texto}\n")
        except Exception as e2:
            logger.error(f"Auto-bitácora: fallo también el respaldo: {e2}")
        await status.edit_text(
            "⚠️ No pude guardar en el Excel ahora (¿está abierto?). "
            "Tu registro quedó respaldado y se puede recuperar — no se perdió.")
        return

    # Descontar inventario si hay insumo + cantidad
    msg_inv = ""
    if campos.get("insumo") and campos.get("cantidad"):
        try:
            from inventario_manager import registrar_uso
            r = await asyncio.to_thread(
                registrar_uso, campos["insumo"], float(campos["cantidad"]),
                campos.get("cultivo", "GENERAL"), campos.get("sector", ""),
                registrado_por, texto[:60])
            msg_inv = f"\n📦 {_esc(campos['insumo'])} \\-{campos['cantidad']:g} (quedan {r.get('stock_restante', 0):g})"
        except Exception as e:
            logger.warning(f"Auto-bitácora: no pude descontar inventario: {e}")

    # Mensaje de maquinaria (horas por odómetro)
    msg_maq = ""
    if campos.get("maquina") and isinstance(res, dict):
        if res.get("es_baseline"):
            msg_maq = f"\n🚜 {_esc(campos['maquina'])}: odómetro inicial registrado"
        elif res.get("horas_dia") is not None:
            msg_maq = f"\n🚜 {_esc(campos['maquina'])}: *{res['horas_dia']:g} h* este período"

    act = campos.get("actividad") or campos.get("tipo") or "registro"
    cult = campos.get("cultivo") or "GENERAL"
    baja_conf = "\n⚠️ _Interpretación dudosa, revisa la bitácora si algo quedó raro._" if campos.get("confianza", 1) < 0.4 else ""
    try:
        await status.edit_text(
            f"✅ *Anotado en bitácora*\n{_esc(act)} · {_esc(cult)}{msg_maq}{msg_inv}{baja_conf}",
            parse_mode="MarkdownV2")
    except Exception:
        await status.edit_text(f"✅ Anotado en bitácora: {act} · {cult}")


async def cmd_bitacora(update, context):
    """/bitacora <texto>  o  /bitacora solo (pide el texto)."""
    registrado_por = update.effective_user.full_name if update.effective_user else ""
    # Texto en el mismo mensaje (autocontenido → robusto a apagones)
    texto = ""
    if context.args:
        texto = " ".join(context.args).strip()
    if texto:
        # Si ya hay un registro esperando confirmación, encolar
        if context.user_data.get("bitacora_pending"):
            await encolar_o_avisar(context, update.effective_chat.id, texto, registrado_por)
            return
        await _procesar_texto_bitacora(update.message, context, texto, registrado_por)
        return
    # Modo conversacional
    context.user_data["bitacora_state"] = "esperando_registro"
    context.user_data["bitacora_registrado_por"] = registrado_por
    await update.message.reply_text(
        "📓 *Bitácora del día*\n\n"
        "Escribe lo que se hizo (ej: _\"hoy poda en avellanos, todos: "
        "Felicito, Patricio, Ramiro, Richard y Jorge\"_)\n\n"
        "_O /cancelar_",
        parse_mode="Markdown")


async def handle_text_bitacora(update, context) -> bool:
    """Si hay flujo de bitácora esperando texto, lo procesa. Devuelve True si lo manejó."""
    if context.user_data.get("bitacora_state") != "esperando_registro":
        return False
    texto = update.message.text.strip()
    registrado_por = context.user_data.get("bitacora_registrado_por") or (
        update.effective_user.full_name if update.effective_user else "")
    await _procesar_texto_bitacora(update.message, context, texto, registrado_por)
    return True


# ── Callbacks ───────────────────────────────────────────

async def cb_bita_save(update, context):
    query = update.callback_query
    await query.answer()
    campos = context.user_data.get("bitacora_pending")
    if not campos:
        await query.edit_message_text("⚠️ No hay registro pendiente.")
        return
    registrado_por = context.user_data.get("bitacora_registrado_por", "")

    from bitacora_manager import registrar_bitacora_estructurada
    res_bita = await asyncio.to_thread(registrar_bitacora_estructurada, campos, registrado_por)
    ok = res_bita is not False and res_bita is not None

    # Descontar inventario si hay insumo + cantidad
    msg_inv = ""
    if ok and campos.get("insumo") and campos.get("cantidad"):
        try:
            from inventario_manager import registrar_uso
            res = await asyncio.to_thread(
                registrar_uso, campos["insumo"], float(campos["cantidad"]),
                campos.get("cultivo", "GENERAL"), campos.get("sector", ""),
                registrado_por, campos.get("texto_original", "")[:60])
            alerta = " ⚠️ *stock bajo*" if res.get("alerta_bajo") else ""
            msg_inv = (f"\n📦 Inventario: {_esc(campos['insumo'])} "
                       f"\\-{campos['cantidad']:g} {_esc(res.get('unidad',''))} "
                       f"\\(quedan {res.get('stock_restante', 0):g}\\){alerta}")
        except Exception as e:
            logger.warning(f"No pude descontar inventario: {e}")
            msg_inv = "\n📦 _No pude descontar el inventario \\(revisar\\)_"

    # Mensaje de maquinaria (horas calculadas por odómetro)
    msg_maq = ""
    if ok and campos.get("maquina"):
        info = res_bita if isinstance(res_bita, dict) else {}
        if info.get("es_baseline"):
            msg_maq = (f"\n🚜 {_esc(campos['maquina'])}: odómetro inicial "
                       f"registrado \\(base para el próximo cálculo\\)")
        elif info.get("horas_dia") is not None:
            h = info["horas_dia"]
            msg_maq = (f"\n🚜 {_esc(campos['maquina'])}: *{h:g} h* este período "
                       f"\\(odómetro {info.get('odo_previo'):g} → "
                       f"{campos.get('odometro'):g}\\)")
        elif campos.get("odometro") not in (None, ""):
            msg_maq = f"\n🚜 {_esc(campos['maquina'])}: odómetro {campos['odometro']:g}"

    # Limpiar estado y procesar siguiente de cola si hubiera
    context.user_data["bitacora_pending"] = None
    jh = campos.get("jornadas_hombre")
    jh_txt = f"\n👷 {jh} JH registradas" if jh else ""
    if ok:
        await query.edit_message_text(
            f"✅ *Registrado en bitácora*\n"
            f"{_esc(campos.get('actividad'))} · {_esc(campos.get('cultivo'))}"
            f"{jh_txt}{msg_maq}{msg_inv}",
            parse_mode="MarkdownV2")
    else:
        await query.edit_message_text("❌ No pude guardar el registro.")
    await _procesar_siguiente_bita(context, query.message.chat_id)


async def cb_bita_edit(update, context):
    query = update.callback_query
    await query.answer()
    context.user_data["bitacora_state"] = "esperando_registro"
    await query.edit_message_text(
        "✏️ Reescribe el registro completo y lo vuelvo a interpretar:")


async def cb_bita_cancel(update, context):
    query = update.callback_query
    await query.answer()
    context.user_data["bitacora_pending"] = None
    context.user_data["bitacora_state"] = None
    await query.edit_message_text("🚫 Registro descartado.")
    await _procesar_siguiente_bita(context, query.message.chat_id)


# ── Cola de bitácoras (si llegan varias antes de confirmar) ──

async def encolar_o_avisar(context, chat_id, texto, registrado_por):
    """Si hay un registro en preview, encola el nuevo texto."""
    cola = context.user_data.setdefault("cola_bitacora", [])
    cola.append({"texto": texto, "por": registrado_por})
    await context.bot.send_message(
        chat_id=chat_id,
        text=f"📓 Anoté otro registro. Está en cola (#{len(cola)}). "
             f"Lo proceso apenas confirmes el actual.")


async def _procesar_siguiente_bita(context, chat_id):
    cola = context.user_data.get("cola_bitacora", [])
    if not cola:
        return
    item = cola.pop(0)
    msg = await context.bot.send_message(
        chat_id=chat_id, text="📓 Interpretando siguiente registro de la cola…")
    from modules.bitacora_extractor import extraer_bitacora
    try:
        campos = await asyncio.to_thread(
            extraer_bitacora, item["texto"], date.today().strftime("%Y-%m-%d"))
        context.user_data["bitacora_pending"] = campos
        context.user_data["bitacora_registrado_por"] = item["por"]
        await msg.edit_text(_build_preview(campos), parse_mode="MarkdownV2",
                            reply_markup=_kb_confirmar())
    except Exception as e:
        logger.error(f"Error procesando bitácora en cola: {e}")
        await msg.edit_text("❌ Error con un registro en cola. Reenvíalo.")
        await _procesar_siguiente_bita(context, chat_id)
