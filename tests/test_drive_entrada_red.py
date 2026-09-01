# -*- coding: utf-8 -*-
"""Un corte de red no puede tumbar el job de la carpeta de entrada.

BUG REAL, 1-sep-2026 10:22:

    TimeoutError: The read operation timed out
      handlers/drive_jobs.py:178, job_drive_entrada

`job_drive_entrada` solo atrapaba `FaltaAutorizacion`, así que cualquier error
de red lo mataba. Durante la migración de los 960 documentos ese mismo timeout
apareció **109 veces** — es lo normal, no lo excepcional, y la cola de subidas
ya lo trata así (reintenta y sigue).

Lo que NO puede pasar es tragarse todo: un error de verdad tiene que llegar al
error handler para que el dueño se entere. Solo se perdonan los de red.
"""
import asyncio
import logging
import types

import pytest

from handlers import drive_jobs


class _Ctx:
    def __init__(self):
        self.bot_data = {"owner_chat_id": "42"}
        self.bot = types.SimpleNamespace(send_message=self._enviar)
        self.enviados = []

    async def _enviar(self, chat_id, text, **kw):
        self.enviados.append(text)


@pytest.fixture
def sin_drive_real(monkeypatch):
    """Evita tocar la red: el cliente y la raíz son de mentira."""
    monkeypatch.setattr(drive_jobs, "_raiz_id", lambda *a, **k: "raiz")
    return monkeypatch


def _reventar(excepcion):
    def _revienta(*a, **k):
        raise excepcion
    return _revienta


@pytest.mark.parametrize("error", [
    TimeoutError("The read operation timed out"),
    ConnectionError("conexión reseteada"),
    OSError("la red no responde"),
])
def test_un_error_de_RED_no_tumba_el_job(sin_drive_real, monkeypatch, error,
                                          caplog):
    """El caso exacto de producción: se anota y se sigue."""
    monkeypatch.setattr("modules.drive.cliente.DriveCliente",
                        lambda *a, **k: object())
    monkeypatch.setattr("handlers.drive_entrada.revisar_entrada",
                        _reventar(error))
    ctx = _Ctx()
    with caplog.at_level(logging.WARNING, logger="handlers.drive_jobs"):
        asyncio.run(drive_jobs.job_drive_entrada(ctx))   # no lanza
    assert "red" in caplog.text.lower() or "reintenta" in caplog.text.lower()


def test_un_error_DE_VERDAD_si_sube(sin_drive_real, monkeypatch):
    """Un bug del código tiene que llegar al error handler, no esconderse."""
    monkeypatch.setattr("modules.drive.cliente.DriveCliente",
                        lambda *a, **k: object())
    monkeypatch.setattr("handlers.drive_entrada.revisar_entrada",
                        _reventar(KeyError("nombre")))
    with pytest.raises(KeyError):
        asyncio.run(drive_jobs.job_drive_entrada(_Ctx()))
