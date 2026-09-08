# -*- coding: utf-8 -*-
"""Los partes de aplicacion de Juan, con los mensajes REALES que mando.

Juan mando 15 partes entre junio y agosto de 2026 y la hoja Aplicaciones tenia 3
filas. Once aplicaciones de agroquimico sin registrar: que producto, en que dosis
y en que cuartel.

Los textos de aca salen del export de su chat, tal cual los escribio.
"""
import pytest

from modules.aplicaciones_parser import es_aplicacion, parsear
from modules.bitacora_asistencia import fecha_de_linea

FUNGICIDA = """Aplicación fungicida
Miércoles 17 de junio 2026
Producto :nordox super 75 wp
Dosis :180 gramos por 100
Cerezos producción
Mojamiento 1500 litros por ha
Total ha : 2
Total Mojamiento : 3000 litros
Total producto : 5.4 kilos"""

HERBICIDA = """Aplicación herbicida
Martes 16 de junio 2026
Producto : aliado
Dosis :0.08 gramos por 100 litros agua
Canales y sercos
Total litros : 600
Total aliado: 0.048 gramos"""

SIN_FECHA = """Aplicación herbicida avellanos
Producto: ripper full
Dosis : 2.5 litros por 100 agua
Mojamiento 100 litros por ha
Total litros : 800
Total litros producto : 20 litros ripper full"""


def p(texto):
    return parsear(texto, fecha_de_linea)


def test_reconoce_una_aplicacion():
    assert es_aplicacion(FUNGICIDA)
    assert es_aplicacion(HERBICIDA)
    assert not es_aplicacion("Asistencia lunes 24 de agosto 2026\nFelicito : poda")
    assert not es_aplicacion("")


def test_el_total_que_vale_es_el_del_PRODUCTO_no_el_del_agua():
    """El parte trae tres totales: 2 ha, 3000 litros de agua y 5,4 kilos.

    Confundirlos mete 3000 litros de agua donde van 5,4 kilos de fungicida.
    """
    d = p(FUNGICIDA)
    assert d["cantidad"] == 5.4
    assert d["unidad"] == "kg"


def test_saca_producto_fecha_cultivo_y_sector():
    d = p(FUNGICIDA)
    assert d["producto"] == "Nordox Super 75 Wp"
    assert d["fecha"] == "2026-06-17"
    assert d["cultivo"] == "CEREZOS"
    assert d["sector"] == "produccion"


def test_el_sector_no_puede_ser_la_fecha():
    """La linea de la fecha tampoco lleva etiqueta: hay que saltarla."""
    for texto in (FUNGICIDA, HERBICIDA):
        d = p(texto)
        assert "junio" not in d["sector"].lower(), d["sector"]


def test_un_lugar_que_no_es_cultivo_queda_como_sector():
    d = p(HERBICIDA)
    assert d["cultivo"] == "GENERAL"
    assert d["sector"] == "Canales y sercos"
    assert d["cantidad"] == 0.048
    assert d["unidad"] == "g"


def test_el_total_del_producto_aunque_la_linea_empiece_con_total_litros():
    """«Total litros producto : 20 litros ripper full» SI es el producto."""
    d = p(SIN_FECHA)
    assert d["cantidad"] == 20.0
    assert d["unidad"] == "L"


def test_sin_fecha_en_el_texto_la_deja_en_None():
    """Que la resuelva quien llame; aca no se inventa."""
    assert p(SIN_FECHA)["fecha"] is None


def test_guarda_el_texto_original_completo():
    assert p(FUNGICIDA)["texto_original"] == FUNGICIDA


def test_distingue_herbicida_de_fungicida():
    assert p(FUNGICIDA)["tipo"] == "fungicida"
    assert p(HERBICIDA)["tipo"] == "herbicida"


@pytest.mark.parametrize("texto", [
    "Aplicación herbicida",                       # solo el encabezado
    "Aplicación herbicida\nMartes 16 de junio",   # sin producto
])
def test_un_parte_a_medias_no_devuelve_nada(texto):
    """Mejor no anotar que anotar una aplicacion sin producto."""
    assert p(texto) is None


def test_no_revienta_con_basura():
    assert p(None) is None
    assert p("cualquier cosa") is None
