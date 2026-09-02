# -*- coding: utf-8 -*-
"""El prompt tiene que prohibir el separador de miles. Es JSON, no un recibo.

Medido el 1-sep-2026 con `qwen2.5vl` sobre 6 facturas reales del Master: el
modelo **acertó el total en las 6**, pero en 3 lo emitió con el punto chileno
como separador de miles, que en JSON es un decimal:

    Master 116.463 -> Ollama emitió  116.463   (float 116,463)
    Master  74.696 -> Ollama emitió   74.696
    Master 518.600 -> Ollama emitió  518.6     (JSON come los ceros finales)

O sea la lectura era correcta y lo que fallaba era el formato.

⚠️ **No se puede arreglar abajo con "en CLP no hay centavos"**: medido sobre el
Master, `Valor unitario` tiene decimales REALES en el **43,8%** de los casos
(7.465,88 · 27.250,849057) y `TOTAL FACTURA` en el 8,9%. Una regla que
multiplique por 1000 todo decimal habría corrompido 659 valores buenos. Por eso
se ataca en el origen: que el modelo no ponga el separador.

El prompt es el mismo que usa el Claude de emergencia, así que la regla sirve
para los dos motores.
"""
import re

from processors.extractor import PROMPT_SIMPLE, _prompt_ollama_con_ejemplos


def _texto():
    return PROMPT_SIMPLE.format(ocr_text="(sin OCR)")


def test_el_prompt_prohibe_el_separador_de_miles():
    t = _texto().lower()
    assert "separador de miles" in t or "separadores de miles" in t, \
        "el prompt no dice nada del separador de miles"


def test_el_prompt_muestra_el_caso_concreto():
    """Una regla abstracta no basta: hay que mostrarle el error exacto."""
    t = _texto()
    assert "116463" in t, "falta el ejemplo de como SI hay que escribirlo"
    assert "116.463" in t, "falta el ejemplo de como NO"


def test_la_regla_sobrevive_al_armado_del_prompt_de_ollama():
    """`_prompt_ollama_con_ejemplos` agrega few-shots; no puede perder la regla."""
    t = _prompt_ollama_con_ejemplos("(sin OCR)").lower()
    assert "separador de miles" in t or "separadores de miles" in t


def test_el_prompt_sigue_pidiendo_solo_json():
    assert "json" in _texto().lower()


def test_no_se_rompio_el_formato_del_prompt():
    """PROMPT_SIMPLE usa llaves dobles para el JSON de ejemplo: si alguien
    escribe una llave suelta, .format() revienta en producción."""
    PROMPT_SIMPLE.format(ocr_text="x")   # no debe lanzar
