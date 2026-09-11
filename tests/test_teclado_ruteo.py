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


# ── El horómetro NO puede saltarse la caducidad de flujos ──────────────────
# 🔴 Pasó el 9-sep-2026. El flujo guiado de horómetro se agregó el 7-sep POR
# ENCIMA de `revisar_flujos`, así que quedó siendo el único flujo inmune a
# caducar. Juan dejó un horómetro a medias probándolo el 8-sep a las 16:23, y al
# día siguiente mandó CINCO partes --3, 4, 7, 8 y 9 de septiembre-- que se los
# comió ese flujo, uno por uno, sin una línea en el log. En el pickle quedó
# `horo_state='esperando_termino'` y `flujo_ts` SIN CREAR: prueba de que
# `revisar_flujos` nunca llegó a correr.
#
# Es exactamente el bug del 28-ago con el /deposito, otra vez, contra el guardia
# que se construyó para evitarlo.


def test_los_flujos_vencidos_se_sueltan_ANTES_de_atender_el_horometro():
    fuente = inspect.getsource(chat.handle_text)
    assert fuente.index("revisar_flujos(") < fuente.index('get("horo_state")'), (
        "revisar_flujos tiene que correr antes del paso de horómetro; si no, un "
        "horómetro a medias es inmune a caducar y se come todo lo que llegue")


def test_un_horometro_vencido_no_se_come_el_parte_del_dia_siguiente():
    """El caso real: horómetro abierto ayer, parte de asistencia hoy."""
    import time
    from modules.flujos import revisar_flujos
    ud = {"horo_state": "esperando_termino", "horo_data": {"inicio": 1950},
          "flujo_ts": time.time() - 20 * 3600}          # abierto hace 20 horas
    assert revisar_flujos(ud) == "horo"
    assert not ud.get("horo_state")
    assert not ud.get("horo_data")


def test_un_horometro_recien_abierto_sigue_vivo():
    """No se puede cerrar el flujo de alguien que está contestando."""
    import time
    from modules.flujos import revisar_flujos
    ud = {"horo_state": "esperando_termino", "flujo_ts": time.time() - 60}
    assert revisar_flujos(ud) is None
    assert ud["horo_state"] == "esperando_termino"


def test_el_comando_con_espacio_se_mira_antes_de_repartir():
    """Si se mirara después, el primer flujo abierto se lo come — que es
    exactamente lo que pasó el 28-ago con "/ cancelar" y el 10-sep con "/ uso"."""
    fuente = inspect.getsource(chat.handle_text)
    assert (fuente.index("comando_con_espacio(")
            < fuente.index("revisar_flujos(")
            < fuente.index("await handle_text_deposito"))
