# Robot/tests/test_ocr_dual.py
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from processors.extractor import _combinar_ocr


def test_ambos_textos_produce_etiquetas():
    resultado = _combinar_ocr("texto tess", "texto surya")
    assert "=== OCR Tesseract ===" in resultado
    assert "=== OCR Surya ===" in resultado
    assert "texto tess" in resultado
    assert "texto surya" in resultado


def test_solo_tesseract_sin_etiqueta():
    resultado = _combinar_ocr("solo tess", "")
    assert "===" not in resultado
    assert resultado == "solo tess"


def test_solo_surya_sin_etiqueta():
    resultado = _combinar_ocr("", "solo surya")
    assert "===" not in resultado
    assert resultado == "solo surya"


def test_ambos_vacios_retorna_vacio():
    assert _combinar_ocr("", "") == ""


def test_ambos_none_retorna_vacio():
    assert _combinar_ocr(None, None) == ""
