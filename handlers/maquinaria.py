"""handlers/maquinaria.py — Recibir horómetros, fichas y mantenciones.

Problema que resuelve: TODA foto que llega al bot se procesa como factura. Si
Juan manda fotos de los horómetros, el bot las lee como facturas y crea
registros basura. Acá se detecta que el mensaje es de maquinaria —por el modo
activo o por lo que diga el texto— y se desvía antes de llegar al extractor.
"""
import logging
import re
from datetime import date, datetime, timedelta

from telegram import InlineKeyboardButton, InlineKeyboardMarkup

logger = logging.getLogger(__name__)

# Cuánto dura el "modo maquinaria" tras activarlo con /maquinaria
MINUTOS_MODO = 30

# Palabras que delatan un mensaje de maquinaria aunque el modo esté apagado
PISTAS = ("horometro", "horómetro", "odometro", "odómetro", "kilometraje",
          "horas maquina", "horas máquina", "mantencion", "mantención",
          "cambio de aceite", "engrase", "patente", "n de serie",
          "numero de serie", "número de serie")


def _ud(context):
    return context.user_data.setdefault("maq", {})


def modo_activo(context) -> bool:
    d = _ud(context)
    hasta = d.get("hasta")
    return bool(hasta and datetime.now() < hasta)


def activar_modo(context, minutos: int = MINUTOS_MODO) -> None:
    _ud(context)["hasta"] = datetime.now() + timedelta(minutes=minutos)


def apagar_modo(context) -> None:
    _ud(context).pop("hasta", None)


def parece_maquinaria(texto: str) -> bool:
    """True si el texto habla de horómetros, mantenciones o fichas."""
    t = (texto or "").lower()
    if any(p in t for p in PISTAS):
        return True
    # "JD 5085 3200" / "Massey 6711: 1980" → nombre de máquina + número
    from modules.maquinaria import detectar_maquina, extraer_odometro
    try:
        if detectar_maquina(t) and extraer_odometro(t) is not None:
            return True
    except Exception:
        pass
    return False


# ── /maquinaria ──────────────────────────────────────────────────────────

async def cmd_maquinaria(update, context):
    """Muestra el estado de cada máquina y activa el modo de captura."""
    from modules.maquinaria import (campos_faltantes, listar_fichas,
                                     maquinas_conocidas)
    import asyncio

    activar_modo(context)
    maquinas = await asyncio.to_thread(maquinas_conocidas)
    fichas = {f["maquina"]: f for f in await asyncio.to_thread(listar_fichas)}

    if not maquinas:
        await update.message.reply_text(
            "🚜 Todavía no hay máquinas registradas.\n\n"
            "Mándame una foto del horómetro con el nombre de la máquina, "
            "o escríbelo: «John Deere 5085 horómetro 3200».")
        return

    hoy = date.today()
    lineas = ["🚜 *Maquinaria*", ""]
    for m in maquinas:
        u = m["unidad"]
        if m["ultimo_odometro"] is not None:
            f = m["fecha"]
            dias = (hoy - f).days if isinstance(f, date) else None
            atraso = f" · hace {dias} días" if dias and dias > 20 else ""
            lect = f"{m['ultimo_odometro']:,.1f} {u}{atraso}"
        else:
            lect = "⚠️ sin lectura"
        falta = campos_faltantes(fichas.get(m["maquina"], {}))
        ficha = f"\n     falta: {', '.join(falta)}" if falta else ""
        lineas.append(f"• *{m['maquina']}* — {lect}{ficha}")

    lineas += [
        "",
        "Escríbeme (los horómetros van escritos, no por foto):",
        "✍️ «MF 6711 horómetro 1980»",
        "🔧 «Al John Deere 5085 le cambiaron aceite y filtros el 20 de julio "
        "a las 3100 horas, lo hizo Álamos»",
        "🪪 «El 5085 es John Deere 2018, patente ABCD12»",
    ]
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Terminar", callback_data="maq_fin")]])
    await update.message.reply_text("\n".join(lineas),
                                     parse_mode="Markdown", reply_markup=kb)


async def cb_maquinaria_fin(update, context):
    q = update.callback_query
    await q.answer()
    apagar_modo(context)
    await q.edit_message_reply_markup(reply_markup=None)
    await q.message.reply_text("🚜 Listo, modo maquinaria apagado.")


# ── Foto de horómetro ────────────────────────────────────────────────────

async def procesar_foto_horometro(update, context) -> bool:
    """Ataja la foto para que NO entre al lector de facturas.

    Por decisión del dueño (10-ago) las lecturas se escriben, no se sacan de la
    foto: menos piezas que puedan fallar. Acá solo se evita el daño —una foto
    de horómetro leída como factura crea un registro basura— y se le pide el
    número por escrito.
    """
    import asyncio

    from modules.maquinaria import detectar_maquina

    caption = (update.message.caption or "").strip()
    maquina = await asyncio.to_thread(detectar_maquina, caption) if caption else None

    if maquina:
        await update.message.reply_text(
            f"📷 Veo que es de *{maquina}*, pero el número no lo leo de la foto.\n"
            f"Escríbemelo así: «{maquina} 3200»", parse_mode="Markdown")
    else:
        await update.message.reply_text(
            "📷 Los horómetros mándamelos escritos, no por foto.\n"
            "_Ejemplo: «John Deere 5085 horómetro 3200»_",
            parse_mode="Markdown")
    return True


# ── Texto de maquinaria ──────────────────────────────────────────────────

async def procesar_texto_maquinaria(update, context) -> bool:
    """Interpreta el mensaje y lo guarda. True si lo tomó."""
    import asyncio

    from modules.maquinaria import (detectar_maquina, extraer_odometro,
                                     maquinas_conocidas)

    texto = (update.message.text or "").strip()
    if not texto:
        return False

    conocidas = await asyncio.to_thread(maquinas_conocidas)
    maquina = await asyncio.to_thread(detectar_maquina, texto, conocidas)
    odo = extraer_odometro(texto)

    # Lectura simple: "MF 6711 horómetro 1980"
    if maquina and odo is not None and not _parece_mantencion(texto):
        return await _guardar_lectura(update, context, maquina, odo)

    # Mantención o ficha → lo estructura la IA
    if _parece_mantencion(texto) or _parece_ficha(texto):
        return await _guardar_con_ia(update, context, texto, conocidas)

    return False


def _parece_mantencion(t: str) -> bool:
    t = t.lower()
    return any(p in t for p in ("mantencion", "mantención", "aceite", "filtro",
                                 "engrase", "repar", "cambio de", "taller",
                                 "revision", "revisión", "neumatico", "neumático"))


def _parece_ficha(t: str) -> bool:
    t = t.lower()
    return any(p in t for p in ("patente", "serie", "modelo", "marca",
                                 "arrendad", "propia", "año "))


async def _guardar_lectura(update, context, maquina, odometro) -> bool:
    """Registra la lectura en la bitácora (así entra al cálculo de horas)."""
    import asyncio

    from bitacora_manager import registrar_bitacora_estructurada
    from modules.maquinaria import unidad_de

    u = unidad_de(maquina)
    campos = {
        "fecha": date.today().strftime("%Y-%m-%d"),
        "tipo": "MAQUINARIA", "actividad": "Lectura de horómetro",
        "cultivo": "GENERAL", "sector": "", "jornadas_hombre": None,
        "trabajadores": [], "insumo": "", "cantidad": None, "unidad": "",
        "maquina": maquina, "odometro": odometro, "superficie_ha": None,
        "texto_original": (update.message.text or "")[:200],
    }
    quien = update.effective_user.full_name if update.effective_user else ""
    try:
        res = await asyncio.to_thread(registrar_bitacora_estructurada, campos, quien)
    except Exception as e:
        logger.error(f"Lectura de horómetro: {e}")
        await update.message.reply_text(f"❌ No pude guardarla: {str(e)[:120]}")
        return True

    # Lectura imposible: no se guarda, se le pregunta
    if isinstance(res, dict) and res.get("error_odometro"):
        previo = res.get("odo_previo")
        await update.message.reply_text(
            f"🤔 No guardé la lectura de *{maquina}*: {res['error_odometro']}\n\n"
            f"Revisa el número y mándamelo de nuevo"
            + (f" (la última que tengo es {previo:,.1f} {u})." if previo else ".")
            + "\n_Ojo con dejar espacio entre el modelo y el horómetro._",
            parse_mode="Markdown")
        return True

    msg = f"✅ *{maquina}* — {odometro:,.1f} {u}"
    if isinstance(res, dict):
        if res.get("es_baseline"):
            msg += "\n_Primera lectura: desde aquí cuento las horas._"
        elif res.get("horas_dia") is not None:
            msg += f"\n🕐 *{res['horas_dia']:g} {u}* desde la lectura anterior"
    await update.message.reply_text(msg, parse_mode="Markdown")
    return True


async def _guardar_con_ia(update, context, texto, conocidas) -> bool:
    """Extrae mantención o ficha del texto libre con la IA."""
    import asyncio

    from modules.maquinaria_extractor import extraer

    status = await update.message.reply_text("🔧 Anotando…")
    try:
        datos = await asyncio.to_thread(
            extraer, texto, [c["maquina"] for c in conocidas],
            date.today().strftime("%Y-%m-%d"))
    except Exception as e:
        logger.error(f"IA maquinaria: {e}")
        await status.edit_text(
            "⚠️ No pude interpretarlo. Escríbelo más simple, por ejemplo:\n"
            "«Al JD 5085 le cambiaron aceite el 20 de julio a las 3100 horas»")
        return True

    quien = update.effective_user.full_name if update.effective_user else ""
    partes = []

    for m in datos.get("mantenciones", []):
        try:
            from modules.maquinaria import registrar_mantencion
            await asyncio.to_thread(registrar_mantencion, m, quien)
            pend = str(m.get("estado") or "").upper() == "PENDIENTE"
            icono = "⏳" if pend else "🔧"
            det = f"{icono} *{m.get('maquina')}* — {m.get('descripcion') or m.get('tipo')}"
            if pend:
                det += "\n     _pendiente de hacer_"
            if m.get("fecha"):
                det += f"\n     {m['fecha']}"
            if m.get("odometro"):
                det += f" · {float(m['odometro']):,.0f} h"
            if m.get("proveedor"):
                det += f" · {m['proveedor']}"
            partes.append(det)
        except Exception as e:
            logger.error(f"Guardando mantención: {e}")

    for f in datos.get("fichas", []):
        try:
            from modules.maquinaria import guardar_ficha
            nombre = await asyncio.to_thread(guardar_ficha, f)
            campos = [f"{k}: {v}" for k, v in f.items()
                      if k != "maquina" and v not in (None, "")]
            partes.append(f"🪪 *{nombre}*\n     " + " · ".join(campos))
        except Exception as e:
            logger.error(f"Guardando ficha: {e}")

    if not partes:
        await status.edit_text(
            "🤔 No entendí de qué máquina hablas. Dime el nombre o el modelo "
            "(ej: «John Deere 5085», «MF 6711»).")
        return True

    await status.edit_text("\n\n".join(partes), parse_mode="Markdown")
    return True
