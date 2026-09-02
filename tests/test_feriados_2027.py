# -*- coding: utf-8 -*-
"""Feriados de Chile 2027, con los movibles ya resueltos.

Los años se cargan a mano porque varios feriados se corren de dia segun donde
caigan. Para 2027 las reglas dan:

  Ley 19.668 (San Pedro y San Pablo, Encuentro de Dos Mundos): si cae martes,
  miercoles o jueves se traslada al LUNES DE ESA MISMA SEMANA; si cae viernes,
  al lunes siguiente.
    29-jun-2027 es MARTES  -> lunes 28-jun
    12-oct-2027 es MARTES  -> lunes 11-oct

  Ley 20.299 (Iglesias Evangelicas): solo se mueve si el 31 de octubre cae
  martes (viernes anterior) o miercoles (viernes de esa semana).
    31-oct-2027 es DOMINGO -> se queda el 31

  Semana Santa: Pascua 2027 es el domingo 28 de marzo, asi que Viernes Santo
  es el 26 y Sabado Santo el 27.

  El 18 de septiembre cae SABADO, asi que no aplica el dia extra de la Ley
  20.215 (que solo suma cuando el 18 cae martes o miercoles).
"""
from datetime import date

import pytest

from modules.feriados import FERIADOS, anio_cubierto, es_feriado, nombre_feriado


def test_2027_esta_cargado():
    assert anio_cubierto(2027)


def test_no_se_perdio_2026():
    assert anio_cubierto(2026)
    assert es_feriado(date(2026, 7, 16))


@pytest.mark.parametrize("d,nombre", [
    (date(2027, 1, 1),   "Año Nuevo"),
    (date(2027, 3, 26),  "Viernes Santo"),
    (date(2027, 3, 27),  "Sábado Santo"),
    (date(2027, 5, 1),   "Día del Trabajo"),
    (date(2027, 5, 21),  "Glorias Navales"),
    (date(2027, 6, 21),  "Día de los Pueblos Indígenas"),
    (date(2027, 7, 16),  "Virgen del Carmen"),
    (date(2027, 8, 15),  "Asunción de la Virgen"),
    (date(2027, 9, 18),  "Independencia Nacional"),
    (date(2027, 9, 19),  "Glorias del Ejército"),
    (date(2027, 11, 1),  "Día de Todos los Santos"),
    (date(2027, 12, 8),  "Inmaculada Concepción"),
    (date(2027, 12, 25), "Navidad"),
])
def test_los_fijos(d, nombre):
    assert es_feriado(d), d
    assert nombre_feriado(d) == nombre


class TestMovibles:
    def test_san_pedro_se_corre_al_lunes_28(self):
        """El 29 cae martes: manda el lunes de esa semana."""
        assert es_feriado(date(2027, 6, 28))
        assert not es_feriado(date(2027, 6, 29))

    def test_dos_mundos_se_corre_al_lunes_11(self):
        """El 12 de octubre cae martes."""
        assert es_feriado(date(2027, 10, 11))
        assert not es_feriado(date(2027, 10, 12))

    def test_iglesias_evangelicas_NO_se_mueve(self):
        """Cae domingo, y la ley solo mueve martes y miercoles."""
        assert es_feriado(date(2027, 10, 31))


def test_un_dia_habil_cualquiera_no_es_feriado():
    assert not es_feriado(date(2027, 3, 25))
    assert not es_feriado(date(2027, 9, 17))


def test_2027_tiene_los_16_feriados():
    assert len(FERIADOS[2027]) == 16, sorted(FERIADOS[2027])
