# -*- coding: utf-8 -*-
"""Retención de respaldos: diarios 30 días, mensuales del año, anuales siempre.

Sin esto, un Master de ~530 KB diario suma ~190 MB al año de copias casi
idénticas.
"""
from datetime import date

from infrastructure.backups import cuales_borrar


def _snaps(fechas):
    return [{"nombre": "master_%s.xlsx" % f, "fecha": date.fromisoformat(f)}
            for f in fechas]


def test_conserva_todos_los_de_los_ultimos_30_dias():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2026-08-25", "2026-08-20", "2026-08-01"])
    assert cuales_borrar(snaps, hoy=hoy) == []


def test_de_un_mes_viejo_conserva_solo_uno():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2026-03-01", "2026-03-15", "2026-03-28"])
    borrar = cuales_borrar(snaps, hoy=hoy)
    assert len(borrar) == 2
    # se conserva el más reciente del mes
    assert "2026-03-28" not in " ".join(b["nombre"] for b in borrar)


def test_de_un_anio_viejo_conserva_solo_uno():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2024-01-10", "2024-06-10", "2024-12-10"])
    borrar = cuales_borrar(snaps, hoy=hoy)
    assert len(borrar) == 2
    assert "2024-12-10" not in " ".join(b["nombre"] for b in borrar)


def test_lista_vacia_no_falla():
    assert cuales_borrar([], hoy=date(2026, 8, 25)) == []


def test_nunca_borra_el_mas_reciente_de_todos():
    """Aunque sea de hace años, el último respaldo NO se puede perder."""
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2023-05-10"])
    assert cuales_borrar(snaps, hoy=hoy) == []


def test_un_mes_con_un_solo_respaldo_no_se_toca():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2026-02-14"])
    assert cuales_borrar(snaps, hoy=hoy) == []
