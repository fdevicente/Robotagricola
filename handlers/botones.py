# -*- coding: utf-8 -*-
"""Que hace cada boton del teclado fijo de Juan.

Estaba dentro de horometro_h.py, que era su sitio cuando el unico boton con
flujo propio era el horometro. Con siete botones ya no: vive aparte.

⚠️ TODO BOTON TIENE QUE HACER ALGO. Un boton que no esta cableado no falla ni
avisa: se queda mudo, y Juan vuelve a escribir comandos rotos. `test_botones.py`
recorre BOTONES y exige que cada uno tenga destino, para que agregar uno y
olvidar cablearlo rompa la suite en vez de romperse en el campo.
"""
import logging

from handlers.teclado import (BOTON_AYUDA, BOTON_CANCELAR, BOTON_HOROMETRO,
                              BOTON_INSUMO, BOTON_MAQUINARIA, teclado_capataz,
                              texto_de_ayuda)

logger = logging.getLogger(__name__)


def destino(boton) -> str:
    """Como se atiende ese boton: 'instruye', 'flujo' o None si no esta cableado."""
    b = str(boton or "").strip()
    if b in (BOTON_HOROMETRO, BOTON_INSUMO, BOTON_MAQUINARIA,
             BOTON_CANCELAR, BOTON_AYUDA):
        return "flujo"
    return "instruye" if texto_de_ayuda(b) else None


async def atender_boton(update, context, boton) -> None:
    """Responde a un boton ya reconocido por teclado.preparar().

    OJO: `preparar()` ya cerro cualquier flujo a medias antes de llegar aca. Esa
    es la salida que Juan buscaba escribiendo "/ cancelar" con espacio.
    """
    boton = str(boton or "").strip()

    if boton == BOTON_CANCELAR:
        # preparar() ya limpio; aca solo se confirma, para que no quede la duda.
        await update.message.reply_text(
            "❌ Listo, cancelé lo que estuviera a medias.\n"
            "Puedes empezar de nuevo con cualquier botón.",
            reply_markup=teclado_capataz())
        return

    if boton == BOTON_AYUDA:
        from main import cmd_ayuda
        await cmd_ayuda(update, context)
        return

    if boton == BOTON_INSUMO:
        from handlers.inventario_h import cmd_uso
        await cmd_uso(update, context)
        return

    if boton == BOTON_MAQUINARIA:
        from handlers.maquinaria import cmd_maquinaria
        await cmd_maquinaria(update, context)
        return

    if boton == BOTON_HOROMETRO:
        from handlers.horometro_h import abrir_horometro
        await abrir_horometro(update, context)
        return

    ayuda = texto_de_ayuda(boton)
    if ayuda:                                # Asistencia y Factura solo instruyen
        await update.message.reply_text(ayuda, parse_mode="Markdown",
                                        reply_markup=teclado_capataz())
        return

    logger.warning("Botón sin cablear: %r", boton)
