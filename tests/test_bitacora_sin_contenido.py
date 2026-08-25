# -*- coding: utf-8 -*-
"""Un mensaje que es solo un comando NO puede generar fila en la bitácora.

Juan escribe seguido la palabra sola antes del parte: "Bitácora", "Personal",
"Asistencia /", "/ maquinaria". El guard era `len(texto) < 3`, asi que todas
pasaban y se guardaban como OTRO con actividad = el propio texto.

Se limpiaron a mano el 10-ago (2 filas) y el 18-ago (7 filas), y volvieron a
aparecer 3 al dia siguiente. Sin guard, esto no para.
"""
import pytest

from modules.bitacora_extractor import es_mensaje_sin_contenido as vacio


@pytest.mark.parametrize("texto", [
    "Bitácora", "bitacora", "BITÁCORA",
    "Personal", "personal",
    "Maquinaria", "maquinaria", "Maquina",
    "Asistencia", "asistencia",
    "/", "//", "/ ", " / ",
    "/bitacora", "/personal", "/maquinaria",
    "Asistencia /", "Bitácora /", "/ maquinaria",
    "Inventario", "Tareas", "Vacaciones",
    '"/"',
    "...", "-", ".",
])
def test_mensajes_que_no_deben_guardarse(texto):
    assert vacio(texto) is True, texto


@pytest.mark.parametrize("texto", [
    "Asistencia 18 de agosto 2026\nFelicito amigo : poda nogales",
    "Tractor massey ferguson 6711\nHorometro termino 2041",
    "Se acabó el petróleo",
    "Bitácora: hoy se podó el sector 3",
    "personal nuevo: llegó Richard",
    "regamos los nogales",
    "Maquinaria: al 5085 le cambiaron el aceite",
])
def test_mensajes_con_contenido_real_si_se_guardan(texto):
    assert vacio(texto) is False, texto


def test_el_caso_exacto_del_19_de_agosto():
    """Las 3 filas basura que aparecieron el dia despues de limpiar."""
    for t in ('"/"', "/ maquinaria", "Maquinaria"):
        assert vacio(t) is True, t


def test_texto_vacio_o_none():
    assert vacio("") is True
    assert vacio(None) is True
    assert vacio("   ") is True


def test_no_se_come_un_parte_que_empieza_con_la_palabra():
    """'Bitácora' sola es basura; con el parte detrás es un registro bueno."""
    assert vacio("Bitácora") is True
    assert vacio("Bitácora\nSe pasó rastra en el sector 2") is False
