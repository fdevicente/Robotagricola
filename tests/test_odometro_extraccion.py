# -*- coding: utf-8 -*-
"""Extraccion del horometro cuando Juan usa su plantilla con etiquetas.

CASO REAL 2026-08-18: Juan empezo a mandar el parte estructurado

    Tractor massey ferguson 6711
    Horometro inicio 2039
    Horometro termino 2041
    Total horas 2
    Labor sacar restos poda nogales
    Equipo 4
    Sector 1

y el bot extrajo **1.0** (el "Sector 1") en vez de 2041. Dos fallas encadenadas:
 1. los patrones exigian el numero PEGADO a la palabra clave, asi que
    "Horometro termino 2041" no matcheaba;
 2. sin match, el fallback tomaba el ULTIMO numero del mensaje, sin mirar a que
    etiqueta pertenecia.

La validacion de odometro hizo bien su trabajo (rechazo el retroceso 2.033,5 -> 1),
pero la lectura se perdia. El arreglo va en la extraccion, no en la validacion.
"""
import pytest

from modules.maquinaria import extraer_odometro, detectar_maquina

PARTE_DE_JUAN = """Tractor massey ferguson 6711
Horometro inicio 2039
Horometro termino 2041
Total horas 2
Labor sacar restos poda nogales
Equipo 4
Sector 1"""


def test_parte_de_juan_toma_el_horometro_de_termino():
    """El caso que rompio: debe dar 2041, nunca el 1 de 'Sector 1'."""
    assert extraer_odometro(PARTE_DE_JUAN) == 2041


def test_parte_de_juan_sigue_detectando_la_maquina():
    assert detectar_maquina(PARTE_DE_JUAN) == "TRACTOR MASSEY FERGUSON 6711"


@pytest.mark.parametrize("texto,esperado", [
    ("Horometro termino 2041", 2041),
    ("Horómetro término 2041", 2041),
    ("horometro final 2041", 2041),
    ("Horometro fin 2041", 2041),
    ("horómetro actual 3.261", 3261),
    ("Odometro llegada 126157", 126157),
])
def test_calificadores_de_lectura_final(texto, esperado):
    assert extraer_odometro(texto) == esperado


def test_si_solo_viene_el_inicio_se_usa_ese():
    assert extraer_odometro("Horometro inicio 2039") == 2039


def test_termino_le_gana_al_inicio_sin_importar_el_orden():
    assert extraer_odometro("Horometro termino 2041\nHorometro inicio 2039") == 2041


# ── El fallback ya no puede agarrar numeros de otras etiquetas ──────────────

@pytest.mark.parametrize("texto", [
    "MF 6711\nSector 1",
    "MF 6711\nEquipo 4",
    "MF 6711\nCuartel 3",
    "MF 6711\nHilera 7",
])
def test_fallback_ignora_etiquetas_que_no_son_horometro(texto):
    """Antes devolvia el numero del sector/equipo como si fuera el horometro."""
    assert extraer_odometro(texto) != 1
    assert extraer_odometro(texto) in (6711, None)


def test_fallback_ignora_las_horas_trabajadas():
    """'Total horas 2' son horas del dia, no la lectura del horometro."""
    assert extraer_odometro("Tractor 6711\nTotal horas 2") != 2


# ── Nada de lo que ya funcionaba puede romperse ─────────────────────────────

@pytest.mark.parametrize("texto,esperado", [
    ("Horómetro 3200", 3200),
    ("horometro 3.166", 3166),
    ("7240,7", 7240.7),
    ("1964 horas", 1964),
    ("MF 6711 1980", 1980),
    ("JD 5085: 3200", 3200),
    ("kilometraje 126157", 126157),
    ("massey ferguson 6711 2033,5", 2033.5),
])
def test_formatos_que_ya_andaban(texto, esperado):
    assert extraer_odometro(texto) == esperado


def test_sin_numeros_devuelve_none():
    assert extraer_odometro("el tractor no anduvo hoy") is None
    assert extraer_odometro("") is None
