# -*- coding: utf-8 -*-
"""Los tres botones fijos que Juan tiene siempre sobre el teclado.

POR QUE
Se midio que manda: 15 intentos de comando --muchos rotos, "/ cancelar",
"/ bitacora", "/ Asistencia", "/" a secas--, 8 fotos de factura, 8 partes de
horometro y 7 de asistencia. Busca un menu y no lo encuentra.

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

BOTONES = (BOTON_ASISTENCIA, BOTON_HOROMETRO, BOTON_FACTURA)

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
    """El teclado fijo. Queda puesto hasta que alguien lo reemplace."""
    return ReplyKeyboardMarkup(
        [[KeyboardButton(BOTON_ASISTENCIA)],
         [KeyboardButton(BOTON_HOROMETRO), KeyboardButton(BOTON_FACTURA)]],
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
