# -*- coding: utf-8 -*-
"""Al guardar una factura se encola su subida, sin bloquear la respuesta."""
import pytest

from modules.drive.cola import Cola
from handlers.facturas import encolar_documento


def test_encola_en_la_carpeta_del_anio(tmp_path):
    doc = tmp_path / "COPEVAL_123.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision="2026-08-25", cola=cola)
    p = cola.pendientes()
    assert len(p) == 1
    assert p[0]["carpeta"] == "Facturas Recibidas/2026"
    assert p[0]["nombre"] == "COPEVAL_123.pdf"


def test_sin_fecha_usa_el_anio_actual(tmp_path):
    from datetime import date
    doc = tmp_path / "X_1.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision=None, cola=cola)
    assert cola.pendientes()[0]["carpeta"].endswith(str(date.today().year))


def test_una_boleta_de_honorarios_va_a_su_carpeta(tmp_path):
    doc = tmp_path / "DONOSO_9.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision="2026-08-25", cola=cola,
                      tipo="boleta")
    assert cola.pendientes()[0]["carpeta"] == "Boletas Honorarios"


def test_si_falla_al_encolar_no_revienta_el_flujo(tmp_path):
    """Perder el enlace es malo; perder la factura es peor."""
    cola = Cola("Z:/ruta/que/no/existe/cola.jsonl")
    # no debe lanzar
    encolar_documento(str(tmp_path / "no-existe.pdf"),
                      fecha_emision="2026-08-25", cola=cola)


def test_una_fecha_en_otro_formato_no_rompe(tmp_path):
    """La IA a veces devuelve la fecha como datetime o con otro formato."""
    from datetime import date
    doc = tmp_path / "Y_2.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision="no es una fecha", cola=cola)
    assert cola.pendientes()[0]["carpeta"].endswith(str(date.today().year))
