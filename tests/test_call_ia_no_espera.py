# -*- coding: utf-8 -*-
"""Si Claude ya contestó, no hay que quedarse esperando a Ollama.

REGRESIÓN INTRODUCIDA AL ARREGLAR OLLAMA (2-sep-2026). `_call_ia` lanza los dos
motores en paralelo pero después hace `futures["ollama"].result(timeout=90)`
SIEMPRE, aunque Claude ya haya devuelto un resultado bueno.

Antes eso era gratis: Ollama estaba roto y fallaba en 3,8 s.

    17:34:50  Llamando a claude / Llamando a Ollama
    17:34:54  Ollama no disponible (3,8 s)
    17:34:56  Claude OK en 5,6 s

Ahora Ollama anda y tarda ~30 s medidos, contra ~6 s de Claude. O sea cada
factura pasó a demorar 30 s en vez de 6: **arreglar Ollama hizo el bot cinco
veces más lento**, y para nada — cuando Claude contesta, lo de Ollama solo se
usaba para una línea de log comparativa.

Los ejemplos de few-shot se guardan del resultado de CLAUDE
(`_guardar_ejemplo(claude_items)`), así que no esperar a Ollama no le quita
nada al aprendizaje.

Ollama sigue siendo el fallback: si Claude falla, ahí sí se lo espera.
"""
import threading
import time

import pytest

from processors import extractor

ITEM = {"Nombre Factura / Proveedor": "X", "Total Factura": 1000,
        "Documento": "Factura Electronica",
        "Numero Factura / Nro Documento": "1"}


@pytest.fixture
def motores(monkeypatch):
    """Claude rápido, Ollama lento. Como en la vida real."""
    monkeypatch.setattr(extractor, "_ollama_disponible", lambda: True)
    monkeypatch.setattr("config.ANTHROPIC_API_KEY", "fake", raising=False)
    monkeypatch.setattr(extractor, "_guardar_ejemplo", lambda *a, **k: None)

    arranco = threading.Event()

    def ollama_lento(*a, **k):
        arranco.set()
        time.sleep(3)
        return [dict(ITEM, **{"Total Factura": 999})]

    monkeypatch.setattr(extractor, "_call_ollama", ollama_lento)
    return arranco


def test_no_espera_a_ollama_si_claude_contesto(motores, monkeypatch):
    """El caso que importa: Claude bien, Ollama lento."""
    monkeypatch.setattr(extractor, "_call_claude", lambda *a, **k: [dict(ITEM)])

    t = time.time()
    items = extractor._call_ia("x.jpg", "")
    tardó = time.time() - t

    assert items and items[0]["Total Factura"] == 1000, "tiene que ganar Claude"
    assert tardó < 2, ("se quedó esperando a Ollama %.1fs de más" % tardó)


def test_si_claude_falla_SI_espera_a_ollama(motores, monkeypatch):
    """Ollama sigue siendo el fallback; ahí la espera vale la pena."""
    monkeypatch.setattr(extractor, "_call_claude", lambda *a, **k: [])

    items = extractor._call_ia("x.jpg", "")
    assert items and items[0]["Total Factura"] == 999, "no uso el fallback"


def test_si_claude_revienta_tambien_usa_ollama(motores, monkeypatch):
    def revienta(*a, **k):
        raise RuntimeError("API caida")
    monkeypatch.setattr(extractor, "_call_claude", revienta)

    items = extractor._call_ia("x.jpg", "")
    assert items and items[0]["Total Factura"] == 999


def test_igual_se_lanza_ollama_para_que_pueda_ser_fallback(motores, monkeypatch):
    """No basta con no esperarlo: hay que haberlo lanzado.

    Se espera sobre un Event y no se mira un contador al toque: el hilo de
    Ollama puede no haber alcanzado a arrancar cuando `_call_ia` ya volvio,
    que es justo lo que queremos que pase.
    """
    monkeypatch.setattr(extractor, "_call_claude", lambda *a, **k: [dict(ITEM)])
    extractor._call_ia("x.jpg", "")
    assert motores.wait(timeout=5), "ni siquiera se lanzo Ollama"
