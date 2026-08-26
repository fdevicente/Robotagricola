# -*- coding: utf-8 -*-
"""Listar una carpeta NO puede devolver sus subcarpetas como si fueran documentos.

BUG REAL 2026-08-26, visto en produccion a los minutos de reiniciar el bot:

    HttpError 400 ... files/1y8OM3...?addParents=1y8OM3...&removeParents=1zFyR...

El fileId y el addParents eran EL MISMO id. `_Entrada/Sin procesar` es una
subcarpeta de `_Entrada`, asi que al listar `_Entrada` aparecia como un
"documento" mas: el robot intentaba descargarla, fallaba, y despues trataba de
moverla a Sin procesar — o sea, dentro de si misma. Cada 15 minutos.

El doble de pruebas no lo podia cazar porque `DriveFalso.listar` solo devolvia
archivos. En Drive real una carpeta ES un archivo, con mimeType de carpeta, y
`'<id>' in parents` la devuelve igual. El falso mentia.
"""
from modules.drive.carpetas import Carpetas
from tests.drive_falso import DriveFalso

CARPETA_MIME = "application/vnd.google-apps.folder"


def test_listar_no_devuelve_subcarpetas():
    d = DriveFalso()
    padre = d.crear_carpeta("_Entrada", "raiz")
    d.crear_carpeta("Sin procesar", padre)          # subcarpeta
    d.subir("x", padre, "factura.pdf")              # documento de verdad

    nombres = [a["nombre"] for a in d.listar(padre)]
    assert nombres == ["factura.pdf"], nombres


def test_el_falso_igual_sabe_de_sus_subcarpetas():
    """No se trata de que el falso las olvide, sino de que listar() las filtre."""
    d = DriveFalso()
    padre = d.crear_carpeta("_Entrada", "raiz")
    sub = d.crear_carpeta("Sin procesar", padre)
    assert d.buscar_carpeta("Sin procesar", padre) == sub


def test_el_cliente_real_excluye_carpetas_en_la_consulta():
    """La query tiene que traer el mimeType != carpeta, o el bug vuelve."""
    import inspect

    from modules.drive.cliente import DriveCliente
    fuente = inspect.getsource(DriveCliente.listar)
    assert "mimeType" in fuente, "listar() no filtra carpetas en la consulta"
    assert "!=" in fuente


def test_revisar_entrada_no_intenta_procesar_la_subcarpeta():
    """El caso exacto que reventó: _Entrada con Sin procesar adentro y nada mas."""
    from handlers.drive_entrada import revisar_entrada

    d = DriveFalso()
    carpetas = Carpetas(d, "raiz")
    carpetas.id_de("_Entrada")
    carpetas.id_de("_Entrada/Sin procesar")

    llamadas = []

    def procesar(archivo):
        llamadas.append(archivo)
        return "Facturas Recibidas/2026"

    res = revisar_entrada(d, carpetas, procesar=procesar)
    assert llamadas == [], "intento procesar la subcarpeta"
    assert res == {"procesados": 0, "sin_procesar": 0}
