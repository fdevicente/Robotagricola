"""El aviso de "bot apagado" tiene que medir CAÍDAS, no silencio.

Caso real (10-ago-2026): el bot avisó "62h apagado" cuando había estado
corriendo todo el fin de semana — lo que llevaba 62 h sin pasar era que
alguien le escribiera. El dueño lo notó porque sí le habían llegado mensajes.
"""
import json
from datetime import datetime, timedelta, timezone

import pytest

from infrastructure import bot_state


@pytest.fixture
def estado(tmp_path, monkeypatch):
    """Redirige el archivo de estado a un temporal."""
    p = tmp_path / ".bot_state.json"
    monkeypatch.setattr(bot_state, "STATE_FILE", str(p))

    def escribir(**campos):
        p.write_text(json.dumps(campos), encoding="utf-8")

    def hace(horas):
        return (datetime.now(timezone.utc) - timedelta(hours=horas)).isoformat()

    escribir.hace = hace
    escribir.ruta = p
    return escribir


# ── El caso que falló ────────────────────────────────────────────────────

def test_fin_de_semana_sin_mensajes_no_es_una_caida(estado):
    """62 h sin mensajes entrantes, pero el proceso latiendo hace 2 min."""
    estado(ultima_actividad_utc=estado.hace(62),
           ultimo_latido_utc=estado.hace(0.03))
    assert bot_state.mensaje_reconexion() is None


def test_una_caida_real_si_avisa(estado):
    estado(ultima_actividad_utc=estado.hace(70),
           ultimo_latido_utc=estado.hace(62))
    msg = bot_state.mensaje_reconexion()
    assert msg and "62h" in msg
    assert "perdido" in msg          # advierte por el límite de 24h de Telegram


def test_caida_corta_avisa_sin_alarmar(estado):
    estado(ultimo_latido_utc=estado.hace(3))
    msg = bot_state.mensaje_reconexion()
    assert msg and "3h" in msg
    assert "no debería faltar nada" in msg


def test_reinicio_normal_no_molesta(estado):
    """Reiniciar el bot a mano no debe disparar aviso."""
    estado(ultimo_latido_utc=estado.hace(0.1))
    assert bot_state.mensaje_reconexion() is None


def test_sin_latido_previo_no_inventa_una_caida(estado):
    """Primer arranque tras la actualización: no se puede saber, no se avisa."""
    estado(ultima_actividad_utc=estado.hace(500))
    assert bot_state.mensaje_reconexion() is None


def test_sin_archivo_no_revienta(tmp_path, monkeypatch):
    monkeypatch.setattr(bot_state, "STATE_FILE", str(tmp_path / "no_existe.json"))
    assert bot_state.cargar_estado() == {}
    assert bot_state.mensaje_reconexion() is None
    assert bot_state.horas_apagado() is None


# ── Las dos marcas son independientes ────────────────────────────────────

def test_el_latido_no_pisa_la_ultima_actividad(estado):
    estado(ultima_actividad_utc=estado.hace(10), ultimo_resumen="hola")
    bot_state.guardar_latido()
    d = bot_state.cargar_estado()
    assert d["ultimo_resumen"] == "hola"
    assert bot_state.horas_desde_ultima_actividad() == pytest.approx(10, abs=0.1)
    assert bot_state.horas_apagado() < 0.01


def test_la_actividad_no_pisa_el_latido(estado):
    estado(ultimo_latido_utc=estado.hace(5))
    bot_state.guardar_actividad(update_id=1, chat_id=2, resumen="mensaje")
    assert bot_state.horas_apagado() == pytest.approx(5, abs=0.1)
    assert bot_state.horas_desde_ultima_actividad() < 0.01


def test_guardar_latido_crea_el_archivo_si_no_existe(tmp_path, monkeypatch):
    p = tmp_path / ".bot_state.json"
    monkeypatch.setattr(bot_state, "STATE_FILE", str(p))
    bot_state.guardar_latido()
    assert p.exists()
    assert bot_state.horas_apagado() < 0.01


def test_timestamp_corrupto_no_revienta(estado):
    estado(ultimo_latido_utc="no es una fecha")
    assert bot_state.horas_apagado() is None
    assert bot_state.mensaje_reconexion() is None
