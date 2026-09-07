# -*- coding: utf-8 -*-
"""El ruteo de los botones.

Se atienden ANTES que cualquier otra cosa: "Asistencia" a secas cae en
es_mensaje_sin_contenido, que la trata como basura --con razon, porque Juan
escribia la palabra sola antes del parte-- y el boton no haria nada.

Y apretar un boton CIERRA lo que haya quedado a medias. Es la salida que Juan
buscaba cuando escribia "/ cancelar" --con espacio, que Telegram no marca como
comando-- y que no le funcionaba nunca: ese /deposito abierto se comio 12 dias
de partes.
"""
import inspect

import handlers.chat as chat
from handlers.teclado import BOTON_ASISTENCIA, BOTON_HOROMETRO, preparar


def test_apretar_un_boton_cierra_lo_que_quedo_a_medias():
    ud = {"deposito_state": "esperando_monto", "deposito_monto": 5000,
          "tarea_state": "esperando_desc"}
    assert preparar(ud, BOTON_ASISTENCIA) == BOTON_ASISTENCIA
    assert not ud.get("deposito_state")
    assert not ud.get("tarea_state")
    assert not ud.get("deposito_monto")


def test_un_texto_que_no_es_boton_no_toca_nada():
    ud = {"deposito_state": "esperando_monto"}
    assert preparar(ud, "Lunes 31 de agosto 2026\nFelicito amigo poda") is None
    assert ud["deposito_state"] == "esperando_monto"


def test_preparar_reconoce_el_boton_con_espacios():
    assert preparar({}, "  " + BOTON_HOROMETRO + " ") == BOTON_HOROMETRO


def test_los_botones_se_atienden_antes_que_los_flujos():
    """Compara contra la LLAMADA, no contra el nombre: `handle_text_deposito`
    tambien aparece en la linea de imports diferidos, que va siempre primero."""
    fuente = inspect.getsource(chat.handle_text)
    assert "preparar" in fuente
    assert fuente.index("preparar") < fuente.index("await handle_text_deposito")


def test_el_flujo_de_horometro_se_atiende_en_el_dispatcher():
    assert "horo_state" in inspect.getsource(chat.handle_text)
