# -*- coding: utf-8 -*-
"""Las boletas de compra tienen carpeta propia, aparte de las facturas.

El bot ya las separa en el PC (config.BOLETAS_DIR es distinto de DOWNLOAD_DIR),
asi que en Drive tienen que quedar igual de separadas. La primera version las
mandaba a "Facturas Recibidas/ANIO" junto con las facturas.

Son TRES cosas distintas:
  - boleta de honorarios -> "Boletas Honorarios"  (personas, con RUT)
  - boleta de compra     -> "Boletas"             (supermercado, ferreteria)
  - factura              -> "Facturas Recibidas/ANIO"
"""
import pytest

from handlers.drive_entrada import carpeta_para
from handlers.drive_jobs import _tipo_de_documento


# ── Las tres carpetas ──────────────────────────────────────────────────────

def test_boleta_de_honorarios_va_a_su_carpeta():
    assert carpeta_para({"tipo": "honorarios", "fecha": "2026-08-25"}) == \
        "Boletas Honorarios"


def test_boleta_de_compra_va_a_boletas():
    assert carpeta_para({"tipo": "boleta", "fecha": "2026-08-25"}) == "Boletas"


def test_la_factura_sigue_yendo_por_anio():
    assert carpeta_para({"tipo": "factura", "fecha": "2026-08-25"}) == \
        "Facturas Recibidas/2026"


def test_una_boleta_no_termina_entre_las_facturas():
    assert "Facturas" not in carpeta_para({"tipo": "boleta", "fecha": "2026-08-25"})
    assert "Facturas" not in carpeta_para({"tipo": "honorarios",
                                            "fecha": "2026-08-25"})


# ── Clasificar lo que devuelve el extractor ────────────────────────────────

@pytest.mark.parametrize("documento,esperado", [
    ("Boleta de Honorarios Electronica", "honorarios"),
    ("BOLETA DE HONORARIOS ELECTRONICA", "honorarios"),
    ("Boleta Electronica", "boleta"),
    ("Boleta Exenta", "boleta"),
    ("Factura Electronica", "factura"),
    ("Factura Exenta", "factura"),
    ("Nota de Credito", "factura"),
    ("Nota de Debito", "factura"),
    ("", "factura"),
])
def test_clasificacion_por_nombre_del_documento(documento, esperado):
    assert _tipo_de_documento(documento) == esperado


def test_honorarios_gana_sobre_boleta():
    """'Boleta de Honorarios' contiene la palabra 'boleta': no debe confundirse."""
    assert _tipo_de_documento("Boleta de Honorarios Electronica") == "honorarios"
