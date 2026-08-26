# -*- coding: utf-8 -*-
"""La cola de subidas a Drive sobrevive reinicios.

El bot corre bajo watchdog y se reinicia seguido. Una cola en memoria perdería
las subidas pendientes en cada reinicio, y con ellas el vínculo entre el archivo
que ya está en disco y su lugar en Drive.
"""
import pytest

from modules.drive.cola import Cola


@pytest.fixture
def cola(tmp_path):
    return Cola(str(tmp_path / "cola.jsonl"))


def test_encolar_y_leer_pendientes(cola):
    cola.encolar("C:/docs/factura.pdf", "Facturas Recibidas/2026", "COPEVAL_123.pdf")
    p = cola.pendientes()
    assert len(p) == 1
    assert p[0]["ruta_local"] == "C:/docs/factura.pdf"
    assert p[0]["carpeta"] == "Facturas Recibidas/2026"
    assert p[0]["nombre"] == "COPEVAL_123.pdf"
    assert p[0]["intentos"] == 0


def test_la_cola_sobrevive_un_reinicio(tmp_path):
    ruta = str(tmp_path / "cola.jsonl")
    Cola(ruta).encolar("a.pdf", "Facturas Recibidas/2026", "a.pdf")
    # otro proceso, misma ruta
    assert len(Cola(ruta).pendientes()) == 1


def test_marcar_ok_saca_el_item(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    item = cola.pendientes()[0]
    cola.marcar_ok(item["id"], "drive-file-id-123")
    assert cola.pendientes() == []


def test_marcar_error_incrementa_intentos_y_lo_deja_pendiente(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    item = cola.pendientes()[0]
    cola.marcar_error(item["id"], "sin internet")
    p = cola.pendientes()
    assert len(p) == 1
    assert p[0]["intentos"] == 1
    assert p[0]["ultimo_error"] == "sin internet"


def test_tras_demasiados_intentos_deja_de_reintentar(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    item_id = cola.pendientes()[0]["id"]
    for _ in range(5):
        cola.marcar_error(item_id, "falla")
    assert cola.pendientes() == []
    rendidos = cola.rendidos()
    assert len(rendidos) == 1
    assert rendidos[0]["intentos"] == 5


def test_no_encola_dos_veces_el_mismo_archivo(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    assert len(cola.pendientes()) == 1


def test_una_linea_corrupta_no_inutiliza_la_cola(cola, tmp_path):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    with open(cola.ruta, "a", encoding="utf-8") as fh:
        fh.write("{no es json}\n")
    cola.encolar("b.pdf", "F/2026", "b.pdf")
    assert len(cola.pendientes()) == 2


def test_cola_vacia_no_falla(tmp_path):
    assert Cola(str(tmp_path / "no-existe.jsonl")).pendientes() == []
