# -*- coding: utf-8 -*-
"""El flujo guiado de horometro, paso a paso.

Lo que compra frente a leerlo con IA: la maquina sale de una lista cerrada (no
se puede inventar) y el numero de inicio se contrasta contra la ultima lectura
EN EL MOMENTO, con Juan parado frente al tractor. Hoy un error de tipeo entra
callado y descuadra las horas de todas las lecturas siguientes de esa maquina.
"""
import pytest

from handlers.horometro import PASOS, avanzar, iniciar, revisar_inicio

CTX = {"maquinas": [
    {"maquina": "TRACTOR MASSEY FERGUSON 4292", "ultimo_odometro": 5239.0,
     "fecha": "2026-09-01", "unidad": "h"},
    {"maquina": "TRACTOR JOHN DEERE 5085", "ultimo_odometro": 3265.0,
     "fecha": "2026-09-01", "unidad": "h"},
]}


def test_iniciar_deja_el_flujo_esperando_la_maquina():
    ud = {}
    iniciar(ud)
    assert ud["horo_state"] == PASOS.MAQUINA


def test_elegir_maquina_pasa_a_pedir_el_inicio():
    ud = {}
    iniciar(ud)
    r = avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    assert ud["horo_state"] == PASOS.INICIO
    assert ud["horo_data"]["maquina"] == "TRACTOR MASSEY FERGUSON 4292"
    assert r["ok"] is True


def test_una_maquina_que_no_existe_no_avanza():
    ud = {}
    iniciar(ud)
    r = avanzar(ud, "TRACTOR FANTASMA", CTX)
    assert ud["horo_state"] == PASOS.MAQUINA
    assert r["ok"] is False


def test_el_inicio_que_calza_con_la_ultima_lectura_pasa_derecho():
    assert revisar_inicio(5239, 5239.0) is None


def test_el_inicio_que_no_calza_avisa_con_los_dos_numeros():
    aviso = revisar_inicio(5137, 5239.0)
    assert aviso is not None
    assert "5.239" in aviso or "5239" in aviso


def test_una_diferencia_chica_tambien_avisa():
    """5237 vs 5239 es justo el error de tipeo que hay que cazar."""
    assert revisar_inicio(5237, 5239.0) is not None


def test_sin_lectura_previa_no_hay_con_que_contrastar():
    assert revisar_inicio(5239, None) is None


def test_un_termino_menor_que_el_inicio_no_avanza():
    ud = {}
    iniciar(ud)
    avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    avanzar(ud, "5239", CTX)
    r = avanzar(ud, "5230", CTX)
    assert r["ok"] is False
    assert ud["horo_state"] == PASOS.TERMINO


def test_un_numero_con_letras_no_avanza():
    ud = {}
    iniciar(ud)
    avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    r = avanzar(ud, "como cinco mil", CTX)
    assert r["ok"] is False
    assert ud["horo_state"] == PASOS.INICIO


def test_el_flujo_completo_deja_los_campos_listos_para_guardar():
    ud = {}
    iniciar(ud)
    avanzar(ud, "TRACTOR MASSEY FERGUSON 4292", CTX)
    avanzar(ud, "5239", CTX)
    avanzar(ud, "5242", CTX)
    r = avanzar(ud, "Sacar restos poda nogales", CTX)
    assert r["ok"] is True
    campos = r["campos"]
    assert campos["tipo"] == "MAQUINARIA"
    assert campos["maquina"] == "TRACTOR MASSEY FERGUSON 4292"
    assert campos["odometro"] == 5242
    assert campos["actividad"] == "Sacar restos poda nogales"
    assert ud.get("horo_state") in (None, "")


@pytest.mark.parametrize("clave", ["horo_state"])
def test_el_flujo_esta_registrado_para_caducar(clave):
    """Un flujo sin registrar no caduca nunca: es el bug que costo 12 dias."""
    from modules.flujos import CLAVES_ESTADO
    assert clave in CLAVES_ESTADO


def test_ningun_flujo_del_proyecto_queda_sin_registrar():
    """Guard contra volver a agregar un flujo y olvidar registrarlo."""
    import os
    import re

    from modules.flujos import CLAVES_ESTADO
    raiz = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "handlers")
    usadas = set()
    for nombre in os.listdir(raiz):
        if not nombre.endswith(".py"):
            continue
        with open(os.path.join(raiz, nombre), encoding="utf-8") as fh:
            usadas |= set(re.findall(r'user_data\.get\("(\w+_state)"\)', fh.read()))
    faltan = usadas - set(CLAVES_ESTADO)
    assert not faltan, "flujos sin registrar en modules/flujos.py: %s" % faltan
