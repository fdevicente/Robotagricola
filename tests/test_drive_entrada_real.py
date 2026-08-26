# -*- coding: utf-8 -*-
"""Qué carpeta le toca a cada documento que aparece en _Entrada."""
import pytest

from handlers.drive_entrada import carpeta_para


def test_una_factura_va_al_anio_de_su_emision():
    assert carpeta_para({"tipo": "factura",
                          "fecha": "2026-08-25"}) == "Facturas Recibidas/2026"


def test_una_boleta_de_honorarios_va_a_su_carpeta():
    assert carpeta_para({"tipo": "boleta", "fecha": "2026-08-25"}) == \
        "Boletas Honorarios"


def test_una_guia_va_a_guias():
    assert carpeta_para({"tipo": "guia", "fecha": "2026-08-25"}) == \
        "Guías de Despacho"


def test_sin_fecha_usa_el_anio_actual():
    from datetime import date
    assert carpeta_para({"tipo": "factura", "fecha": None}).endswith(
        str(date.today().year))


def test_una_fecha_ilegible_usa_el_anio_actual():
    from datetime import date
    assert carpeta_para({"tipo": "factura", "fecha": "ayer"}).endswith(
        str(date.today().year))


def test_un_tipo_desconocido_lanza_para_que_vaya_a_sin_procesar():
    with pytest.raises(ValueError):
        carpeta_para({"tipo": "quien sabe", "fecha": "2026-08-25"})


def test_sin_tipo_tambien_lanza():
    with pytest.raises(ValueError):
        carpeta_para({})
