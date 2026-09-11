# -*- coding: utf-8 -*-
"""Los tres botones fijos de Juan.

Se midio que manda: 15 intentos de comando (muchos rotos, "/ cancelar",
"/ Asistencia"), 8 fotos de factura, 8 partes de horometro, 7 de asistencia.
Busca un menu y no lo encuentra.

Asistencia y Factura NO abren flujo: solo instruyen. Cada estado conversacional
nuevo es una forma mas de que se trabe, y esos dos no lo necesitan.
"""
from handlers.teclado import (BOTON_ASISTENCIA, BOTON_FACTURA, BOTON_HOROMETRO,
                              es_boton, teclado_capataz, texto_de_ayuda)


def test_el_teclado_trae_los_tres_botones_de_siempre():
    """Los tres originales siguen estando; el resto los cubre otro test."""
    filas = teclado_capataz().keyboard
    textos = {b.text for fila in filas for b in fila}
    assert {BOTON_ASISTENCIA, BOTON_HOROMETRO, BOTON_FACTURA} <= textos


def test_el_teclado_es_persistente_y_no_se_esconde():
    kb = teclado_capataz()
    assert kb.resize_keyboard is True
    assert getattr(kb, "one_time_keyboard", False) is not True


def test_reconoce_los_botones():
    assert es_boton(BOTON_ASISTENCIA)
    assert es_boton(BOTON_HOROMETRO)
    assert es_boton(BOTON_FACTURA)


def test_no_confunde_un_parte_con_un_boton():
    """El parte de Juan empieza con la fecha, no puede parecerse a un boton."""
    assert not es_boton("Lunes 31 de agosto 2026\nFelicito amigo poda")
    assert not es_boton("")
    assert not es_boton(None)


def test_reconoce_el_boton_aunque_venga_con_espacios():
    assert es_boton("  " + BOTON_ASISTENCIA + " ")


def test_asistencia_y_factura_solo_instruyen():
    """No abren flujo: el texto tiene que decir QUE mandar."""
    assert "parte" in texto_de_ayuda(BOTON_ASISTENCIA).lower()
    assert "foto" in texto_de_ayuda(BOTON_FACTURA).lower()


def test_el_horometro_no_tiene_texto_de_ayuda():
    """Ese si abre flujo, lo maneja handlers/horometro.py."""
    assert texto_de_ayuda(BOTON_HOROMETRO) is None


# ── El teclado sale de lo que Juan INTENTA, no de lo que suponemos ─────────
# Medido el 11-sep-2026 sobre sus 85 mensajes: 24 intentos de comando.
#   /bitacora 7 · /maquinaria 3 · /uso 3 · /ayuda 2 · /cancelar 2 · "/" roto 2
#   /deposito, /tareas, /asistencia, /start, /inventario: 1 cada uno.
# Los tres botones que habia cubrian asistencia, horometro y factura. Faltaban
# JUSTO los dos que se rompieron: /uso (el stock fantasma del 10-sep) y el
# cancelar explicito, que el escribia "/ cancelar" con espacio.

from handlers.teclado import (BOTON_AYUDA, BOTON_CANCELAR, BOTON_INSUMO,
                              BOTON_MAQUINARIA, BOTONES, es_boton,
                              teclado_capataz)


def test_estan_los_botones_de_lo_que_juan_mas_intenta():
    for b in (BOTON_INSUMO, BOTON_MAQUINARIA, BOTON_CANCELAR, BOTON_AYUDA):
        assert b in BOTONES
        assert es_boton(b)


def test_el_teclado_los_muestra_todos():
    filas = teclado_capataz().keyboard
    en_pantalla = {b.text for fila in filas for b in fila}
    assert en_pantalla == set(BOTONES)


def test_ninguna_fila_tiene_mas_de_dos_botones():
    """En un teléfono, tres por fila quedan ilegibles."""
    assert all(len(fila) <= 2 for fila in teclado_capataz().keyboard)


def test_no_hay_dos_botones_con_el_mismo_texto():
    assert len(set(BOTONES)) == len(BOTONES)
