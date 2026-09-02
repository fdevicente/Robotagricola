# -*- coding: utf-8 -*-
"""Borrar filas de `Facturas` corre las referencias de `Conciliaciones`.

La hoja `Conciliaciones` guarda el vínculo banco↔factura **por número de fila**
(`Fila Doc`). Al borrar una fila de `Facturas`, todo lo que está debajo sube
uno y esas referencias quedan apuntando a la factura equivocada — en silencio,
que es lo peor.

Al 2-sep-2026 hay 25 vínculos entre las filas 991 y 2194, y hay que borrar 10
filas duplicadas repartidas entre medio. Sin este ajuste, el vínculo de la
factura 2777 de Jorge Bravo (hoy fila 2153) pasaría a apuntar a otra factura.

La cuenta es simple y por eso conviene fijarla: a cada referencia se le
descuenta **cuántas filas borradas quedaron por encima**.
"""
import pytest

from modules.filas import ajustar_referencia


BORRADAS = [501, 954, 2117, 2139, 2152, 2181, 2182, 2183, 2184, 2192]


class TestCuentaBasica:
    def test_una_fila_por_encima_descuenta_uno(self):
        assert ajustar_referencia(1000, [501]) == 999

    def test_dos_filas_por_encima_descuentan_dos(self):
        assert ajustar_referencia(1000, [501, 954]) == 998

    def test_una_fila_por_debajo_no_afecta(self):
        assert ajustar_referencia(500, [501, 954]) == 500

    def test_sin_borrados_no_cambia_nada(self):
        assert ajustar_referencia(2153, []) == 2153

    def test_el_orden_de_la_lista_da_igual(self):
        assert (ajustar_referencia(2153, [2152, 501, 954])
                == ajustar_referencia(2153, [501, 954, 2152]))


class TestCasosReales:
    """Los números concretos del Master al 2-sep-2026."""

    @pytest.mark.parametrize("antes,despues", [
        (991,  989),    # INDELEC: solo 501 y 954 por encima
        (1199, 1197),
        (1310, 1308),
        (2132, 2129),   # + 2117
        (2153, 2148),   # Jorge Bravo: + 2139 y 2152
        (2166, 2161),   # S-INVEST
        (2171, 2166),   # Ferreteria M y G 6950, la que se conserva
        (2179, 2174),
        (2193, 2183),   # CORA: las 10 por encima
        (2194, 2184),   # AGROCAMPO
    ])
    def test_las_referencias_del_master(self, antes, despues):
        assert ajustar_referencia(antes, BORRADAS) == despues


class TestBordes:
    def test_la_fila_borrada_misma_no_deberia_consultarse(self):
        """Si alguien pregunta por una fila que se borro, se avisa fuerte."""
        with pytest.raises(ValueError):
            ajustar_referencia(2152, BORRADAS)

    def test_la_fila_justo_debajo_de_una_borrada(self):
        assert ajustar_referencia(502, [501]) == 501

    def test_la_fila_justo_encima_de_una_borrada(self):
        assert ajustar_referencia(500, [501]) == 500
