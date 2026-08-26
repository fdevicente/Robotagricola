# -*- coding: utf-8 -*-
"""Resolver rutas tipo 'Facturas Recibidas/2026' a IDs de Drive, creando lo que falte."""
import pytest

from modules.drive.carpetas import Carpetas
from tests.drive_falso import DriveFalso


@pytest.fixture
def carpetas():
    return Carpetas(DriveFalso(), raiz_id="raiz")


def test_crea_la_ruta_completa(carpetas):
    cid = carpetas.id_de("Facturas Recibidas/2026")
    assert cid is not None
    # la intermedia también existe
    assert carpetas.drive.buscar_carpeta("Facturas Recibidas", "raiz") is not None


def test_no_crea_dos_veces_la_misma_carpeta(carpetas):
    a = carpetas.id_de("Facturas Recibidas/2026")
    b = carpetas.id_de("Facturas Recibidas/2026")
    assert a == b
    assert len([c for c in carpetas.drive.carpetas.values()
                if c and c["nombre"] == "2026"]) == 1


def test_reutiliza_la_carpeta_padre_entre_anios(carpetas):
    carpetas.id_de("Facturas Recibidas/2025")
    carpetas.id_de("Facturas Recibidas/2026")
    padres = [c for c in carpetas.drive.carpetas.values()
              if c and c["nombre"] == "Facturas Recibidas"]
    assert len(padres) == 1


def test_una_sola_carpeta_sin_barras(carpetas):
    assert carpetas.id_de("Respaldos") is not None


def test_ruta_vacia_devuelve_la_raiz(carpetas):
    assert carpetas.id_de("") == "raiz"
