# -*- coding: utf-8 -*-
"""Un flujo a medias no puede quedarse abierto para siempre comiendose el texto.

EL CASO REAL (28-ago a 2-sep-2026, 12 dias de partes perdidos)
Juan escribio /deposito y quiso salir con "/ cancelar" -- CON ESPACIO. Telegram
solo marca como comando lo que va pegado, asi que ese mensaje no ejecuto
cmd_cancelar: entro como texto y se lo comio el propio flujo de deposito.

Desde ahi `deposito_state='esperando_monto'` quedo abierto en el pickle (que
sobrevive a los reinicios) y como es el PRIMERO de la fila en handlers/chat.py,
cada parte de Juan se interpretaba como "el monto del deposito", fallaba con
"No es un monto valido" y cortaba ahi. No llegaba a maquinaria ni a la bitacora,
y no dejaba ni una linea en el log. La hoja Bitacora quedo congelada 12 dias.

LA TRAMPA A EVITAR: el reloj NO se puede reiniciar con cada mensaje. Juan mandaba
un parte cada ~80 s; un timeout "desde el ultimo mensaje" nunca habria vencido.
El reloj corre desde que el flujo se abrio.
"""
import pytest

from modules.flujos import (CLAVES_ESTADO, MINUTOS_VIDA, flujos_abiertos,
                            limpiar_flujos, revisar_flujos)

MIN = 60.0


def test_sin_flujo_abierto_no_pasa_nada():
    ud = {"deposito_state": None, "otra_cosa": "se queda"}
    assert revisar_flujos(ud, ahora=0) is None
    assert ud["otra_cosa"] == "se queda"
    assert "flujo_ts" not in ud


def test_flujo_recien_abierto_sobrevive_y_queda_fechado():
    ud = {"deposito_state": "esperando_monto"}
    assert revisar_flujos(ud, ahora=1000.0) is None
    assert ud["deposito_state"] == "esperando_monto"
    assert ud["flujo_ts"] == 1000.0


def test_flujo_joven_sigue_vivo():
    ud = {"deposito_state": "esperando_monto", "flujo_ts": 1000.0}
    assert revisar_flujos(ud, ahora=1000.0 + 5 * MIN) is None
    assert ud["deposito_state"] == "esperando_monto"


def test_el_reloj_no_se_reinicia_con_cada_mensaje():
    """La trampa: mirar el flujo no puede refrescarle la fecha."""
    ud = {"deposito_state": "esperando_monto"}
    revisar_flujos(ud, ahora=1000.0)
    revisar_flujos(ud, ahora=1000.0 + 5 * MIN)
    revisar_flujos(ud, ahora=1000.0 + 10 * MIN)
    assert ud["flujo_ts"] == 1000.0


def test_flujo_vencido_se_descarta():
    ud = {"deposito_state": "esperando_monto", "deposito_monto": 5000,
          "flujo_ts": 1000.0}
    assert revisar_flujos(ud, ahora=1000.0 + (MINUTOS_VIDA + 1) * MIN) == "deposito"
    assert not ud.get("deposito_state")
    assert not ud.get("deposito_monto")
    assert "flujo_ts" not in ud


def test_el_caso_de_juan_un_parte_cada_80_segundos():
    """Reproduce el bug: mensajes seguidos NO pueden mantener vivo el flujo."""
    ud = {"deposito_state": "esperando_monto"}
    t = 0.0
    descartado = None
    for _ in range(40):                       # 40 partes, uno cada 80 s
        descartado = revisar_flujos(ud, ahora=t) or descartado
        t += 80.0
    assert descartado == "deposito"
    assert not ud.get("deposito_state"), "el flujo se comio los partes para siempre"


def test_se_descartan_todos_los_flujos_abiertos_no_solo_el_primero():
    """El pickle de Juan tenia CUATRO abiertos a la vez."""
    ud = {"deposito_state": "esperando_monto", "tarea_state": "esperando_desc",
          "bitacora_state": "esperando_registro", "uso_state": "esperando_producto",
          "uso_data": {"algo": 1}, "flujo_ts": 0.0}
    assert revisar_flujos(ud, ahora=(MINUTOS_VIDA + 1) * MIN) is not None
    assert flujos_abiertos(ud) == []
    assert not ud.get("uso_data")


def test_limpiar_no_toca_lo_que_no_es_flujo():
    ud = {"deposito_state": "esperando_monto", "cola_facturas": [1, 2],
          "last_invoice_file": "x.jpg", "maq": {"hasta": "luego"}}
    limpiar_flujos(ud)
    assert ud["cola_facturas"] == [1, 2]
    assert ud["last_invoice_file"] == "x.jpg"
    assert ud["maq"] == {"hasta": "luego"}


def test_vencimientos_tambien_es_un_flujo():
    """venc_state no estaba en la lista de /cancelar: tambien puede trabarse."""
    assert "venc_state" in CLAVES_ESTADO


@pytest.mark.parametrize("clave", list(CLAVES_ESTADO))
def test_cualquier_flujo_abierto_vence(clave):
    ud = {clave: "en_curso", "flujo_ts": 0.0}
    assert revisar_flujos(ud, ahora=(MINUTOS_VIDA + 1) * MIN) is not None
    assert not ud.get(clave)
