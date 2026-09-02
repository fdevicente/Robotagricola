# -*- coding: utf-8 -*-
"""Decidir si un número viene con separador de miles se hace con ARITMÉTICA.

`_limpiar_items` traía una regla por FORMA: "float con exactamente 3 decimales
→ ×1000", pensada para el punto de miles chileno (771.784 → 771784).

Medido sobre el Master el 2-sep-2026, esa forma NO distingue:

    41 valores de total con exactamente 3 decimales
       6 la regla los ARREGLA (hoy no cuadran)
      35 la regla los ROMPERÍA (hoy SÍ cuadran)

Los 35 son residuo legítimo de multiplicar por 1,19:

    563.600 × 1,19 = 670.684,976   <- 3 decimales, y está PERFECTO
     85.714 × 1,19 = 102.000,136

O sea la forma del número no dice nada; lo que decide es si cuadra con el resto
de la factura. Un total de 670.684,976 con neto 563.600 está bien; el mismo
número con neto 563.600.000 estaría mil veces abajo.

Misma idea que `sanear_impuesto_especifico`: se prueban las dos hipótesis y gana
la que reconcilia. Sin referencia contra qué comparar se cae a la regla vieja,
que es el comportamiento histórico.
"""
import pytest

from processors.extractor import _elegir_escala


class TestConReferencia:
    def test_el_residuo_de_multiplicar_por_119_se_deja_como_esta(self):
        """563.600 x 1,19 = 670.684,976. Multiplicarlo por mil lo destruiria."""
        assert _elegir_escala(670684.976, referencia=670685.0) == 670685

    def test_el_separador_de_miles_se_corrige(self):
        """El modelo escribio 771.784 queriendo decir 771784."""
        assert _elegir_escala(771.784, referencia=771784.0) == 771784

    def test_el_caso_con_ceros_comidos_por_el_json(self):
        """518.600 se serializa como 518.6 y hay que reconstruirlo igual."""
        assert _elegir_escala(518.6, referencia=518600.0) == 518600

    def test_gana_la_hipotesis_mas_cercana_no_la_mas_grande(self):
        assert _elegir_escala(102000.136, referencia=102000.0) == 102000

    def test_un_entero_no_se_toca_aunque_haya_referencia(self):
        assert _elegir_escala(29998, referencia=29998.0) == 29998

    def test_tolera_que_la_referencia_no_sea_exacta(self):
        """La referencia es neto x IVA, que redondea distinto que el emisor."""
        assert _elegir_escala(116.463, referencia=116460.0) == 116463


class TestSinReferencia:
    """Sin con qué comparar se mantiene la regla vieja: 3 decimales → ×1000."""

    def test_tres_decimales_se_multiplica(self):
        assert _elegir_escala(771.784, referencia=None) == 771784

    def test_un_decimal_no_alcanza_para_decidir(self):
        assert _elegir_escala(518.6, referencia=None) == 519

    def test_el_entero_queda_igual(self):
        assert _elegir_escala(29998, referencia=None) == 29998

    def test_referencia_cero_es_como_no_tenerla(self):
        assert _elegir_escala(771.784, referencia=0) == 771784


def test_nunca_lanza_con_basura():
    for v in (None, "", "abc"):
        with pytest.raises((TypeError, ValueError)):
            _elegir_escala(v, referencia=1000.0)
