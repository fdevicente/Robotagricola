# -*- coding: utf-8 -*-
"""Los documentos que se rindieron tienen que poder volver a la cola.

Con el job cada 10 minutos, los 5 intentos se agotan en 50 minutos. Un corte de
internet mas largo que eso —nada raro en el campo— dejaba el documento en
"rendido" y el subidor NO lo miraba nunca mas: el archivo quedaba a salvo en el
PC pero no llegaba a Drive jamas, sin forma de reintentarlo.
"""
import pytest

from modules.drive.cola import Cola


@pytest.fixture
def cola_rendida(tmp_path):
    """Una cola con un item que ya agoto sus 5 intentos."""
    c = Cola(str(tmp_path / "cola.jsonl"))
    c.encolar("/docs/factura.pdf", "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    iid = c.pendientes()[0]["id"]
    for _ in range(5):
        c.marcar_error(iid, "sin internet")
    assert c.pendientes() == [] and len(c.rendidos()) == 1
    return c


def test_reintentar_devuelve_el_item_a_pendientes(cola_rendida):
    n = cola_rendida.reintentar_rendidos()
    assert n == 1
    assert len(cola_rendida.pendientes()) == 1
    assert cola_rendida.rendidos() == []


def test_el_contador_de_intentos_vuelve_a_cero(cola_rendida):
    cola_rendida.reintentar_rendidos()
    assert cola_rendida.pendientes()[0]["intentos"] == 0


def test_conserva_el_destino_y_el_nombre(cola_rendida):
    cola_rendida.reintentar_rendidos()
    item = cola_rendida.pendientes()[0]
    assert item["ruta_local"] == "/docs/factura.pdf"
    assert item["carpeta"] == "Facturas Recibidas/2026"
    assert item["nombre"] == "COPEVAL_1.pdf"


def test_sin_rendidos_no_hace_nada(tmp_path):
    c = Cola(str(tmp_path / "cola.jsonl"))
    c.encolar("a.pdf", "F/2026", "a.pdf")
    assert c.reintentar_rendidos() == 0
    assert len(c.pendientes()) == 1


def test_no_revive_los_que_ya_subieron(tmp_path):
    c = Cola(str(tmp_path / "cola.jsonl"))
    c.encolar("a.pdf", "F/2026", "a.pdf")
    c.marcar_ok(c.pendientes()[0]["id"], "file-1")
    assert c.reintentar_rendidos() == 0
    assert c.pendientes() == []


def test_el_reintento_tambien_sobrevive_un_reinicio(tmp_path):
    ruta = str(tmp_path / "cola.jsonl")
    c = Cola(ruta)
    c.encolar("a.pdf", "F/2026", "a.pdf")
    iid = c.pendientes()[0]["id"]
    for _ in range(5):
        c.marcar_error(iid, "falla")
    c.reintentar_rendidos()
    # proceso nuevo, misma ruta
    assert len(Cola(ruta).pendientes()) == 1


def test_puede_volver_a_rendirse_y_reintentarse(cola_rendida):
    """Un segundo corte largo no puede dejarlo varado otra vez."""
    cola_rendida.reintentar_rendidos()
    iid = cola_rendida.pendientes()[0]["id"]
    for _ in range(5):
        cola_rendida.marcar_error(iid, "sin internet de nuevo")
    assert len(cola_rendida.rendidos()) == 1
    assert cola_rendida.reintentar_rendidos() == 1
    assert len(cola_rendida.pendientes()) == 1
