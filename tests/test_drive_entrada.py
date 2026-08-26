# -*- coding: utf-8 -*-
"""La carpeta _Entrada: procesar y mover. El movimiento ES la marca de procesado."""
import pytest

from handlers.drive_entrada import revisar_entrada
from modules.drive.carpetas import Carpetas
from tests.drive_falso import DriveFalso


@pytest.fixture
def entorno():
    drive = DriveFalso()
    carpetas = Carpetas(drive, "raiz")
    entrada = carpetas.id_de("_Entrada")
    return drive, carpetas, entrada


def test_mueve_a_su_carpeta_lo_que_pudo_procesar(entorno):
    drive, carpetas, entrada = entorno
    fid = drive.subir("x", entrada, "COPEVAL_55.pdf")
    res = revisar_entrada(drive, carpetas,
                          procesar=lambda a: "Facturas Recibidas/2026")
    assert res["procesados"] == 1
    assert drive.archivos[fid]["carpeta_id"] == carpetas.id_de("Facturas Recibidas/2026")


def test_lo_que_no_pudo_leer_va_a_sin_procesar(entorno):
    drive, carpetas, entrada = entorno
    fid = drive.subir("x", entrada, "foto_borrosa.jpg")

    def falla(a):
        raise ValueError("no se pudo leer")

    res = revisar_entrada(drive, carpetas, procesar=falla)
    assert res["sin_procesar"] == 1
    assert drive.archivos[fid]["carpeta_id"] == carpetas.id_de("_Entrada/Sin procesar")


def test_no_vuelve_a_tomar_lo_que_ya_movio(entorno):
    drive, carpetas, entrada = entorno
    drive.subir("x", entrada, "A.pdf")
    revisar_entrada(drive, carpetas, procesar=lambda a: "Facturas Recibidas/2026")
    res = revisar_entrada(drive, carpetas, procesar=lambda a: "Facturas Recibidas/2026")
    assert res["procesados"] == 0


def test_ignora_la_subcarpeta_sin_procesar(entorno):
    drive, carpetas, entrada = entorno
    sp = carpetas.id_de("_Entrada/Sin procesar")
    drive.subir("x", sp, "viejo.pdf")
    res = revisar_entrada(drive, carpetas, procesar=lambda a: "F/2026")
    assert res["procesados"] == 0


def test_entrada_vacia_no_hace_nada(entorno):
    drive, carpetas, _ = entorno
    assert revisar_entrada(drive, carpetas, procesar=lambda a: "F/2026") == {
        "procesados": 0, "sin_procesar": 0}


def test_un_archivo_malo_no_impide_procesar_los_demas(entorno):
    """Si uno falla, los otros igual tienen que llegar a su carpeta."""
    drive, carpetas, entrada = entorno
    bueno1 = drive.subir("x", entrada, "BUENO_1.pdf")
    malo = drive.subir("x", entrada, "MALO.jpg")
    bueno2 = drive.subir("x", entrada, "BUENO_2.pdf")

    def procesar(a):
        if a["nombre"] == "MALO.jpg":
            raise ValueError("ilegible")
        return "Facturas Recibidas/2026"

    res = revisar_entrada(drive, carpetas, procesar=procesar)
    assert res == {"procesados": 2, "sin_procesar": 1}
    destino = carpetas.id_de("Facturas Recibidas/2026")
    assert drive.archivos[bueno1]["carpeta_id"] == destino
    assert drive.archivos[bueno2]["carpeta_id"] == destino
    assert drive.archivos[malo]["carpeta_id"] == carpetas.id_de("_Entrada/Sin procesar")


def test_procesar_recibe_el_archivo_completo_no_solo_el_nombre(entorno):
    """La Tarea 12 necesita el id para poder descargar el archivo."""
    drive, carpetas, entrada = entorno
    fid = drive.subir("x", entrada, "A.pdf")
    vistos = []

    def procesar(a):
        vistos.append(a)
        return "F/2026"

    revisar_entrada(drive, carpetas, procesar=procesar)
    assert vistos[0]["id"] == fid
    assert vistos[0]["nombre"] == "A.pdf"
