# -*- coding: utf-8 -*-
"""El cliente real y el falso exponen la MISMA interfaz.

Si divergen, las pruebas pasan contra el falso y la producción se rompe.
"""
import inspect

from modules.drive.cliente import DriveCliente
from tests.drive_falso import DriveFalso

METODOS = ["subir", "crear_carpeta", "buscar_carpeta", "buscar_archivo",
           "listar", "mover", "cuota", "descargar"]


def test_el_falso_implementa_todos_los_metodos_del_real():
    for m in METODOS:
        assert hasattr(DriveCliente, m), "falta %s en el real" % m
        assert hasattr(DriveFalso, m), "falta %s en el falso" % m


def test_las_firmas_coinciden():
    for m in METODOS:
        real = inspect.signature(getattr(DriveCliente, m))
        falso = inspect.signature(getattr(DriveFalso, m))
        assert list(real.parameters) == list(falso.parameters), \
            "%s: %s vs %s" % (m, list(real.parameters), list(falso.parameters))


def test_el_falso_sube_y_encuentra():
    d = DriveFalso()
    cid = d.crear_carpeta("Facturas Recibidas", "raiz")
    fid = d.subir("C:/x/COPEVAL_1.pdf", cid, "COPEVAL_1.pdf")
    assert d.buscar_archivo("COPEVAL_1.pdf", cid) == fid
    assert d.listar(cid) == [{"id": fid, "nombre": "COPEVAL_1.pdf"}]


def test_el_falso_puede_simular_una_caida():
    d = DriveFalso()
    d.fallar_con = ConnectionError("sin internet")
    try:
        d.subir("a.pdf", "dir-1", "a.pdf")
        assert False, "debió lanzar"
    except ConnectionError:
        pass
