# -*- coding: utf-8 -*-
"""Un test que aisla sus rutas NO puede escribir en la cola de produccion.

`tests/test_backups.py` siempre paso `excel_path` y `backup_base` propios, como
manda la regla del proyecto. Pero el encolado hacia Drive leia DRIVE_COLA_PATH
de config, o sea la cola REAL: cada corrida completa de la suite dejaba ahi
entradas apuntando a carpetas temporales de pytest ya borradas.

Es el mismo patron que una vez destruyo el Master real — confiar en un default
dentro de algo que el test creia haber aislado.
"""
import openpyxl
import pytest

from infrastructure.backups import backup_master
from modules.drive.cola import Cola


@pytest.fixture
def master_falso(tmp_path):
    ruta = tmp_path / "master.xlsx"
    wb = openpyxl.Workbook()
    wb.save(ruta)
    wb.close()
    return str(ruta)


def test_con_cola_propia_encola_ahi(tmp_path, master_falso):
    cola_path = str(tmp_path / "cola_test.jsonl")
    backup_master(reason="test", excel_path=master_falso,
                  backup_base=str(tmp_path / "backups"), cola_path=cola_path)
    assert len(Cola(cola_path).pendientes()) == 1


def test_no_toca_la_cola_de_produccion(tmp_path, master_falso):
    """Lo que de verdad importa: la cola real queda como estaba."""
    from config import DRIVE_COLA_PATH
    antes = len(Cola(DRIVE_COLA_PATH).pendientes())

    backup_master(reason="test", excel_path=master_falso,
                  backup_base=str(tmp_path / "backups"),
                  cola_path=str(tmp_path / "cola_test.jsonl"))

    assert len(Cola(DRIVE_COLA_PATH).pendientes()) == antes


def test_lo_encolado_apunta_al_snapshot_que_existe(tmp_path, master_falso):
    import os
    cola_path = str(tmp_path / "cola_test.jsonl")
    backup_master(reason="test", excel_path=master_falso,
                  backup_base=str(tmp_path / "backups"), cola_path=cola_path)
    item = Cola(cola_path).pendientes()[0]
    assert os.path.exists(item["ruta_local"]), item["ruta_local"]
    assert item["carpeta"] == "Respaldos/Master"
