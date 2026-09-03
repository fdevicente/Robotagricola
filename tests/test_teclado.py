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


def test_el_teclado_trae_los_tres_botones():
    filas = teclado_capataz().keyboard
    textos = [b.text for fila in filas for b in fila]
    assert set(textos) == {BOTON_ASISTENCIA, BOTON_HOROMETRO, BOTON_FACTURA}


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
