# -*- coding: utf-8 -*-
"""Ningun boton del teclado de Juan puede quedar mudo.

Un boton sin cablear no falla ni avisa: no pasa nada y Juan vuelve a escribir
comandos rotos, que es de donde venimos. Este test recorre BOTONES y exige que
cada uno tenga destino, para que agregar uno y olvidar cablearlo rompa la suite
en vez de romperse en el campo.
"""
import inspect

import handlers.botones as botones
from handlers.teclado import (BOTON_AYUDA, BOTON_CANCELAR, BOTON_FACTURA,
                              BOTON_HOROMETRO, BOTON_INSUMO, BOTON_MAQUINARIA,
                              BOTONES)


def test_todos_los_botones_tienen_destino():
    mudos = [b for b in BOTONES if botones.destino(b) is None]
    assert mudos == [], "botones sin cablear: %s" % mudos


def test_los_que_abren_flujo_y_los_que_solo_instruyen():
    assert botones.destino(BOTON_HOROMETRO) == "flujo"
    assert botones.destino(BOTON_INSUMO) == "flujo"
    assert botones.destino(BOTON_MAQUINARIA) == "flujo"
    assert botones.destino(BOTON_CANCELAR) == "flujo"
    assert botones.destino(BOTON_AYUDA) == "flujo"
    assert botones.destino(BOTON_FACTURA) == "instruye"


def test_un_texto_cualquiera_no_es_un_boton():
    assert botones.destino("Lunes 7 de septiembre") is None
    assert botones.destino("") is None
    assert botones.destino(None) is None


def test_cada_boton_con_flujo_esta_realmente_cableado_en_el_despacho():
    """`destino` podría mentir: se comprueba contra el código del despacho."""
    fuente = inspect.getsource(botones.atender_boton)
    for nombre in ("BOTON_CANCELAR", "BOTON_AYUDA", "BOTON_INSUMO",
                   "BOTON_MAQUINARIA", "BOTON_HOROMETRO"):
        assert nombre in fuente, "%s no se atiende en atender_boton" % nombre


def test_el_boton_de_insumo_llama_al_flujo_de_uso():
    """Es el que faltaba: por no tenerlo, Juan escribió "/ uso" con espacio y
    el inventario quedó con un producto fantasma en -5."""
    fuente = inspect.getsource(botones.atender_boton)
    assert "cmd_uso" in fuente
