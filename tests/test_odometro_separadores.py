# -*- coding: utf-8 -*-
"""Juan escribe el horometro de tres formas distintas, y una se leia 1000x mal.

Encontrado el 7-sep-2026 al reingresar lecturas del export de Telegram:

    "Horometro inicio 2,037 / termino 2,039"   ->  se leia 2.039
    "Horometro inicio 5.232 / termino 5.234"   ->  se leia 5234   OK
    "Horometro inicio 3261  / termino 3263"    ->  se leia 3263   OK

La coma se tomaba SIEMPRE como decimal, aunque para el punto ya existia la regla
de los tres digitos. Un 2,039 leido como 2.039 no es una fila mala: descuadra el
calculo de horas de TODAS las lecturas siguientes de esa maquina, porque las
horas del dia se sacan restando contra la anterior.

Y "Tractor /massey ferguson 6711" --con la barra que Juan mete a veces-- se
detectaba como "TRACTOR" a secas, la maquina generica, en vez del 6711.
"""
import pytest

from modules.maquinaria import _a_float, detectar_maquina, extraer_odometro

CONOCIDAS = [{"maquina": "TRACTOR"}, {"maquina": "TRACTOR MASSEY FERGUSON 6711"},
             {"maquina": "TRACTOR MASSEY FERGUSON 4292"},
             {"maquina": "TRACTOR JOHN DEERE 5085"}]


@pytest.mark.parametrize("crudo,esperado", [
    ("2,039", 2039),        # coma como separador de miles: el caso que fallaba
    ("2.039", 2039),        # punto como separador de miles
    ("5.234", 5234),
    ("3263", 3263),
    ("7240,7", 7240.7),     # coma decimal de verdad: un solo digito detras
    ("7240.7", 7240.7),
    ("7205,95", 7205.95),   # dos digitos detras tampoco son miles
    ("104,000", 104000),
])
def test_los_separadores_de_miles_no_se_leen_como_decimales(crudo, esperado):
    assert _a_float(crudo) == esperado


def test_el_caso_exacto_del_12_de_agosto():
    """Se leia 2.039 en vez de 2039: un error de mil veces."""
    texto = ("Miércoles 12 de agosto 2026\nTractor massey ferguson 6711\n"
             "Horometro inicio 2,037\nHorometro termino 2,039\nTotal horas 2\n"
             "Labor sacar restos poda nogales\nEquipo 4\nSector 2")
    assert extraer_odometro(texto) == 2039


def test_la_barra_no_convierte_al_6711_en_un_tractor_generico():
    """Juan escribe "Tractor /massey ferguson 6711" y hay que reconocerlo."""
    texto = ("Martes 18 de agosto 2026\nTractor /massey ferguson 6711\n"
             "Horometro inicio 2041\nHorometro termino 2043")
    assert detectar_maquina(texto, CONOCIDAS) == "TRACTOR MASSEY FERGUSON 6711"


@pytest.mark.parametrize("escrito", [
    "Tractor /massey ferguson 6711",
    "Tractor / massey ferguson 6711",
    "Tractor massey ferguson 6711",
    "MF 6711",
])
def test_da_igual_como_lo_escriba(escrito):
    assert detectar_maquina(escrito, CONOCIDAS) == "TRACTOR MASSEY FERGUSON 6711"


def test_el_tractor_generico_sigue_existiendo():
    """No hay que romper el caso en que de verdad no dice cual es."""
    assert detectar_maquina("Se uso el tractor todo el dia", CONOCIDAS) == "TRACTOR"
