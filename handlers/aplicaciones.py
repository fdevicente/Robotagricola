# -*- coding: utf-8 -*-
"""Anota los partes de aplicacion que manda Juan.

POR QUE EXISTE
Juan mando 15 partes de aplicacion entre junio y agosto de 2026 y solo 2 se
anotaron: los que la IA de bitacora pillo de casualidad porque traian insumo y
cantidad. Los otros 11 se perdieron enteros --producto, dosis y cuartel-- y hubo
que recuperarlos a mano desde el export de su chat.

El formato de Juan es regular, asi que aca no hay IA: lo lee
`modules/aplicaciones_parser`, que esta probado contra sus 15 mensajes reales.
"""
import asyncio
import logging

logger = logging.getLogger(__name__)

FALTA = (
    "🧪 Te leo la aplicación, pero me falta el *producto* o el *total*.\n\n"
    "Mándamela así:\n"
    "_Aplicación fungicida_\n"
    "_Miércoles 17 de junio 2026_\n"
    "_Producto : nordox super 75 wp_\n"
    "_Dosis : 180 gramos por 100_\n"
    "_Cerezos producción_\n"
    "_Total producto : 5.4 kilos_")


async def procesar_aplicacion(update, context) -> bool:
    """Anota el parte si es una aplicacion. True si lo tomo."""
    from modules.aplicaciones_parser import es_aplicacion, parsear
    from modules.bitacora_asistencia import fecha_de_linea

    texto = (update.message.text or "").strip()
    if not es_aplicacion(texto):
        return False

    datos = await asyncio.to_thread(parsear, texto, fecha_de_linea)
    if datos is None or datos["cantidad"] is None:
        logger.info("Aplicación incompleta, se le pide a %s el producto/total",
                    update.effective_user.full_name if update.effective_user else "?")
        await update.message.reply_text(FALTA, parse_mode="Markdown")
        return True

    # Sin fecha en el texto, la del dia. Es lo unico que se puede suponer, y
    # queda dicho en la propia fila para que se note si estaba equivocada.
    from datetime import date
    obs = texto
    if not datos["fecha"]:
        datos["fecha"] = date.today().strftime("%Y-%m-%d")
        obs = "[sin fecha en el texto, se usó la del día] " + texto

    quien = update.effective_user.full_name if update.effective_user else ""
    from inventario_manager import registrar_uso
    try:
        res = await asyncio.to_thread(
            registrar_uso, datos["producto"], float(datos["cantidad"]),
            datos["cultivo"], datos["sector"], quien, obs[:600],
            datos["fecha"], datos["unidad"])
    except Exception as e:
        logger.error("Aplicación: no pude anotarla: %s", e)
        await update.message.reply_text(
            "❌ No pude anotar la aplicación: %s" % str(e)[:120])
        return True

    logger.info("Aplicación: %s %s%s en %s (%s)", datos["producto"],
                datos["cantidad"], datos["unidad"], datos["cultivo"],
                datos["fecha"])
    linea = "✅ *%s* — %g %s\n📅 %s · %s" % (
        datos["producto"], datos["cantidad"], datos["unidad"],
        datos["fecha"], datos["cultivo"])
    if datos["sector"]:
        linea += " · %s" % datos["sector"]
    resto = res.get("stock_restante")
    if resto is not None:
        linea += "\n📦 Quedan %g %s" % (resto, datos["unidad"])
        if res.get("alerta_bajo"):
            linea += "  ⚠️ stock bajo"
    await update.message.reply_text(linea, parse_mode="Markdown")
    return True
