# -*- coding: utf-8 -*-
"""Los botones fijos que Juan tiene siempre sobre el teclado.

POR QUE, Y POR QUE ESTOS
Salen de lo que INTENTA, no de lo que suponemos. Medido el 11-sep-2026 sobre
sus 85 mensajes: 24 intentos de comando --/bitacora 7, /maquinaria 3, /uso 3,
/ayuda 2, /cancelar 2, "/" a secas 2, y uno cada uno de /deposito, /tareas,
/asistencia, /start e /inventario--, 15 fotos y 46 mensajes de texto libre.
Busca un menu y no lo encuentra.

Los tres primeros botones (asistencia, horometro, factura) cubrian lo mas
frecuente. Faltaban JUSTO los dos que se rompieron: /uso --que el 10-sep dejo
el inventario con un producto fantasma en -5-- y un cancelar explicito, que el
escribia "/ cancelar" con espacio y Telegram no marca como comando.

Se usa ReplyKeyboardMarkup y no botones inline a proposito: el teclado queda
fijo sobre el de Telegram, siempre visible, sin gastar un mensaje ni obligar a
preguntarle "quieres ingresar datos?" antes de cada cosa.

ASISTENCIA Y FACTURA NO ABREN FLUJO. Solo le dicen que mandar. Cada estado
conversacional nuevo es una forma mas de que se trabe --un /deposito sin cerrar
se comio 12 dias de partes-- y esos dos no lo necesitan.
"""
from telegram import KeyboardButton, ReplyKeyboardMarkup

BOTON_ASISTENCIA = "📋 Asistencia"
BOTON_HOROMETRO = "🚜 Horómetro"
BOTON_FACTURA = "🧾 Factura"
BOTON_INSUMO = "🧪 Insumo"
BOTON_MAQUINARIA = "🔧 Maquinaria"
BOTON_CANCELAR = "❌ Cancelar"
BOTON_AYUDA = "❓ Ayuda"

BOTONES = (BOTON_ASISTENCIA, BOTON_HOROMETRO, BOTON_FACTURA,
           BOTON_INSUMO, BOTON_MAQUINARIA, BOTON_CANCELAR, BOTON_AYUDA)

_AYUDA = {
    BOTON_ASISTENCIA: (
        "📋 Dale, mándame el *parte* del día.\n\n"
        "Escríbelo como te salga, una línea por persona:\n"
        "_Lunes 1 de septiembre 2026_\n"
        "_Felicito amigo sacar restos poda nogales_\n"
        "_Patricio Mora aplicación herbicida_"),
    BOTON_FACTURA: (
        "🧾 Mándame la *foto* de la factura.\n\n"
        "Si quieres, escríbele al lado de qué es "
        "(_«compra petróleo campo»_) y lo anoto."),
}


def teclado_capataz() -> ReplyKeyboardMarkup:
    """El teclado fijo. Queda puesto hasta que alguien lo reemplace.

    Dos por fila como maximo: en un telefono, tres quedan ilegibles.
    """
    return ReplyKeyboardMarkup(
        [[KeyboardButton(BOTON_ASISTENCIA)],
         [KeyboardButton(BOTON_HOROMETRO), KeyboardButton(BOTON_FACTURA)],
         [KeyboardButton(BOTON_INSUMO), KeyboardButton(BOTON_MAQUINARIA)],
         [KeyboardButton(BOTON_CANCELAR), KeyboardButton(BOTON_AYUDA)]],
        resize_keyboard=True, is_persistent=True)


def es_boton(texto) -> bool:
    return str(texto or "").strip() in BOTONES


def texto_de_ayuda(boton) -> str | None:
    """Que responderle. None para los botones que abren un flujo propio."""
    return _AYUDA.get(str(boton or "").strip())


def preparar(user_data, texto) -> str | None:
    """Cierra lo que haya quedado a medias y devuelve el boton apretado, o None.

    Los botones son la SALIDA cuando un flujo quedo abierto. Juan intentaba
    salir escribiendo "/ cancelar" --con espacio, que Telegram no marca como
    comando, asi que nunca ejecutaba cmd_cancelar-- y ese /deposito abierto se
    comio 12 dias de partes en silencio. Apretar un boton ahora si lo cierra.
    """
    from modules.flujos import limpiar_flujos
    if not es_boton(texto):
        return None
    limpiar_flujos(user_data)
    return str(texto).strip()
