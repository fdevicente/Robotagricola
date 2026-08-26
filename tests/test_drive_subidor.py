# -*- coding: utf-8 -*-
"""El subidor vacía la cola sin perder nunca el archivo local."""
import pytest

from modules.drive.carpetas import Carpetas
from modules.drive.cola import Cola
from modules.drive.subidor import procesar_cola
from tests.drive_falso import DriveFalso


@pytest.fixture
def entorno(tmp_path):
    doc = tmp_path / "COPEVAL_1.pdf"
    doc.write_text("contenido", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    drive = DriveFalso()
    return cola, drive, Carpetas(drive, "raiz"), str(doc)


def test_sube_lo_pendiente_y_vacia_la_cola(entorno):
    cola, drive, carpetas, doc = entorno
    cola.encolar(doc, "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    res = procesar_cola(cola, drive, carpetas)
    assert res["subidos"] == 1
    assert cola.pendientes() == []
    cid = carpetas.id_de("Facturas Recibidas/2026")
    assert drive.buscar_archivo("COPEVAL_1.pdf", cid) is not None


def test_si_drive_falla_el_archivo_queda_encolado_y_en_disco(entorno):
    cola, drive, carpetas, doc = entorno
    drive.fallar_con = ConnectionError("sin internet")
    cola.encolar(doc, "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    res = procesar_cola(cola, drive, carpetas)
    assert res["subidos"] == 0
    assert res["fallidos"] == 1
    assert len(cola.pendientes()) == 1          # sigue pendiente
    import os
    assert os.path.exists(doc)                   # y el archivo NO se borró


def test_no_sube_dos_veces_el_mismo_documento(entorno):
    cola, drive, carpetas, doc = entorno
    cid = carpetas.id_de("Facturas Recibidas/2026")
    drive.subir(doc, cid, "COPEVAL_1.pdf")       # ya estaba en Drive
    cola.encolar(doc, "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    procesar_cola(cola, drive, carpetas)
    assert len(drive.listar(cid)) == 1


def test_un_archivo_que_ya_no_existe_no_bloquea_la_cola(entorno, tmp_path):
    cola, drive, carpetas, doc = entorno
    cola.encolar(str(tmp_path / "borrado.pdf"), "F/2026", "borrado.pdf")
    res = procesar_cola(cola, drive, carpetas)
    assert res["fallidos"] == 1
    assert cola.pendientes()[0]["intentos"] == 1


def test_cola_vacia_no_hace_nada(entorno):
    cola, drive, carpetas, _ = entorno
    assert procesar_cola(cola, drive, carpetas) == {"subidos": 0, "fallidos": 0}
