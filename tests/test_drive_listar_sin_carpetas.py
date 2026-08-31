# -*- coding: utf-8 -*-
"""`listar()` devuelve ARCHIVOS, nunca carpetas.

BUG REAL 2026-08-26, primer arranque con la integracion viva:

    HttpError 400 ... files/1y8OM...?addParents=1y8OM...&removeParents=1zFy...

El id del archivo y el destino eran EL MISMO: el job intentaba mover una
carpeta dentro de si misma. `listar(_Entrada)` devolvia la subcarpeta
"Sin procesar" como si fuera un documento; el job trataba de descargarla,
fallaba, y la mandaba a Sin procesar — que es ella misma.

Los tests no lo cazaron porque `DriveFalso.listar` solo mira `self.archivos` y
nunca devuelve carpetas: el doble era MAS AMABLE que la realidad. La consulta
real (`'<id>' in parents`) devuelve todos los hijos directos, carpetas incluidas.
"""
from modules.drive.cliente import CARPETA_MIME, DriveCliente


class _ServicioEspia:
    """Captura la consulta que arma el cliente, sin tocar la red."""

    def __init__(self):
        self.q = None

    def files(self):
        return self

    def list(self, q=None, **kw):
        self.q = q
        return self

    def execute(self):
        return {"files": []}


def test_la_consulta_excluye_las_carpetas():
    espia = _ServicioEspia()
    DriveCliente(servicio=espia).listar("id-de-entrada")
    assert CARPETA_MIME in espia.q, espia.q
    assert "mimeType !=" in espia.q, espia.q


def test_la_consulta_sigue_filtrando_por_padre_y_papelera():
    espia = _ServicioEspia()
    DriveCliente(servicio=espia).listar("id-de-entrada")
    assert "'id-de-entrada' in parents" in espia.q
    assert "trashed = false" in espia.q


def test_el_doble_tampoco_devuelve_carpetas():
    """El falso tiene que ser igual de estricto que el real, no mas amable."""
    from tests.drive_falso import DriveFalso
    d = DriveFalso()
    entrada = d.crear_carpeta("_Entrada", "raiz")
    d.crear_carpeta("Sin procesar", entrada)          # subcarpeta
    d.subir("x", entrada, "FACTURA_1.pdf")            # archivo de verdad

    listado = d.listar(entrada)
    assert [a["nombre"] for a in listado] == ["FACTURA_1.pdf"], listado


def test_una_subcarpeta_nunca_llega_al_procesador():
    """La prueba de punta a punta del bug: revisar_entrada la ignora."""
    from handlers.drive_entrada import revisar_entrada
    from modules.drive.carpetas import Carpetas
    from tests.drive_falso import DriveFalso

    drive = DriveFalso()
    carpetas = Carpetas(drive, "raiz")
    carpetas.id_de("_Entrada/Sin procesar")           # crea ambas
    vistos = []

    def procesar(a):
        vistos.append(a["nombre"])
        return "Facturas Recibidas/2026"

    res = revisar_entrada(drive, carpetas, procesar=procesar)
    assert vistos == []                                # no vio la carpeta
    assert res == {"procesados": 0, "sin_procesar": 0}
