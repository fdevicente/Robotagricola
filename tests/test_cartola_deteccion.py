"""Con el scraper bloqueado por detección de bots, la carga manual de cartola
pasó a ser la vía principal para actualizar el banco: tiene que ser robusta.
"""
import pytest

from handlers.banco_upload import (EXT_CARTOLA, es_archivo_cartola,
                                    parece_cartola_por_nombre)


@pytest.mark.parametrize("nombre", [
    "typeDesc.txt",                 # el que entrega el portal de Scotiabank
    "cartola.csv",
    "Movimientos_agosto.xlsx",
    "CARTOLA.XLS",
])
def test_reconoce_las_cartolas_soportadas(nombre):
    assert es_archivo_cartola(nombre)


@pytest.mark.parametrize("nombre", ["factura.pdf", "foto.jpg", "guia.docx", ""])
def test_no_confunde_otros_archivos(nombre):
    assert not es_archivo_cartola(nombre)


@pytest.mark.parametrize("nombre", [
    "Cartola_Julio_2026.pdf",
    "movimientos.pdf",
    "ESTADO_CUENTA.pdf",
    "typeDesc.pdf",
])
def test_detecta_cartola_en_pdf_por_el_nombre(nombre):
    """Una cartola PDF que llegue al extractor se guardaría como factura falsa."""
    assert parece_cartola_por_nombre(nombre)


@pytest.mark.parametrize("nombre", [
    "F107228 Llaneza.pdf",
    "Copeval ND506808.pdf",
    "boleta honorarios donoso.pdf",
])
def test_una_factura_normal_no_se_confunde_con_cartola(nombre):
    assert not parece_cartola_por_nombre(nombre)


def test_pdf_no_esta_entre_las_extensiones_importables():
    """Del PDF no se leen movimientos: debe rechazarse, no procesarse."""
    assert ".pdf" not in EXT_CARTOLA
