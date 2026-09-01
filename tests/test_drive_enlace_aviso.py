# -*- coding: utf-8 -*-
"""Cuando NO se escribe el enlace, el aviso tiene que decir POR QUE.

Migrando los 827 documentos el log dijo 192 veces "sin fila para X Nº N", pero
solo 69 eran de verdad facturas ausentes del Master. Las otras 122 eran la
SEGUNDA foto de una factura ya enlazada — `AGROCAMPO_184882.jpg` y
`AGROCAMPO_184882_20260423_093503.jpg` son el mismo documento — donde la fila
existe y ya tenia su enlace.

Un aviso que miente en 122 de 192 casos entrena a no mirarlo, y esto se repite
cada vez que Juan reenvia una foto. La diferencia se sabe dentro de
`guardar_enlace`, que es el unico que recorre las filas: el aviso va ahi.
"""
import logging

import openpyxl
import pytest

from modules.drive.enlaces import guardar_enlace

COL_PROV, COL_NUM, COL_DRIVE = 4, 7, 22


@pytest.fixture
def libro(tmp_path):
    ruta = tmp_path / "m.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append([None] * 21 + ["Archivo Drive"])
    f = [None] * 22
    f[COL_PROV - 1], f[COL_NUM - 1] = "AGROCAMPO", 184882
    ws.append(f)
    wb.save(ruta)
    wb.close()
    return str(ruta)


def test_avisa_que_la_fila_YA_TENIA_enlace(libro, caplog):
    """El caso de las 122: la segunda foto de una factura ya enlazada."""
    guardar_enlace(libro, "184882", "primero", proveedor="AGROCAMPO")
    with caplog.at_level(logging.INFO, logger="modules.drive.enlaces"):
        guardar_enlace(libro, "184882", "segundo", proveedor="AGROCAMPO")
    assert "ya tenía enlace" in caplog.text
    assert "sin fila" not in caplog.text


def test_avisa_que_NO_HAY_fila(libro, caplog):
    """El caso de las 69: la factura no esta en el Master."""
    with caplog.at_level(logging.INFO, logger="modules.drive.enlaces"):
        guardar_enlace(libro, "999999", "abc", proveedor="DESCONOCIDO")
    assert "sin fila" in caplog.text
    assert "ya tenía enlace" not in caplog.text


def test_cuando_escribe_no_avisa_nada(libro, caplog):
    with caplog.at_level(logging.INFO, logger="modules.drive.enlaces"):
        assert guardar_enlace(libro, "184882", "abc",
                              proveedor="AGROCAMPO") is True
    assert caplog.text == ""


def test_el_proveedor_equivocado_es_SIN_FILA_no_ya_enlazada(libro, caplog):
    """La fila del nº existe, pero es de otro proveedor: no hay fila para MI."""
    with caplog.at_level(logging.INFO, logger="modules.drive.enlaces"):
        guardar_enlace(libro, "184882", "abc", proveedor="OTRO PROVEEDOR")
    assert "sin fila" in caplog.text
