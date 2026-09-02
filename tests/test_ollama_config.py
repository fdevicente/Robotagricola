# -*- coding: utf-8 -*-
"""El motor local tiene que tener modelo de visión y tiempo suficiente.

1-sep-2026. Ollama llevaba tiempo sin hacer nada y el log solo decía
"Ollama no disponible o tardó demasiado: 500 Server Error". Llamando a la API
a mano salió el motivo real:

    error loading model: unknown model architecture: 'mllama'

Ollama se actualizó a 0.33.2 y dejó de soportar `mllama`, la arquitectura de
`llama3.2-vision`. El modelo seguía en disco (7,3 GB) pero el motor ya no lo
cargaba: fallaba **hasta sin imagen**, así que nunca fue el prompt ni el tamaño.

Se reemplazó por `qwen2.5vl`, que sí carga. Medido contra una factura real
(FERRETERIA M Y G Nº3533): sacó neto 435.798 y total 518.600, **exactos** contra
el Master... en **55 segundos**. El timeout estaba en 60: por cinco segundos el
motor local iba a seguir “no disponible”, y el mensaje de error habría sido otra
vez el genérico de siempre.
"""
import re
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent


def test_el_modelo_por_defecto_no_es_el_roto():
    """`llama3.2-vision` es mllama: no lo carga ninguna Ollama moderna."""
    from config import OLLAMA_MODEL
    assert "llama3.2-vision" not in OLLAMA_MODEL, (
        "OLLAMA_MODEL apunta a un modelo con arquitectura mllama, que la "
        "Ollama instalada no puede cargar")


def test_hay_un_modelo_de_vision_configurado():
    from config import OLLAMA_MODEL
    assert OLLAMA_MODEL.strip(), "no hay modelo de Ollama configurado"


def test_el_timeout_da_para_una_factura_de_verdad():
    """Medido: 55 s en una factura real. Con 60 no alcanzaba."""
    from config import OLLAMA_TIMEOUT
    assert OLLAMA_TIMEOUT >= 120, (
        "OLLAMA_TIMEOUT=%s es muy corto: una factura real tardó 55 s y el "
        "modelo local es varias veces más lento que la API" % OLLAMA_TIMEOUT)


def test_el_extractor_usa_el_timeout_de_config_y_no_uno_escrito_a_mano():
    fuente = (RAIZ / "processors" / "extractor.py").read_text(encoding="utf-8")
    cuerpo = fuente[fuente.index("def _call_ollama"):]
    cuerpo = cuerpo[:cuerpo.index("\ndef ")]
    assert "OLLAMA_TIMEOUT" in cuerpo, \
        "_call_ollama tiene el timeout hardcodeado en vez de leerlo de config"
    assert not re.search(r"timeout\s*=\s*\d+", cuerpo), \
        "quedó un timeout numérico escrito a mano en _call_ollama"
