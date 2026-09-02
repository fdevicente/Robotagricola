# -*- coding: utf-8 -*-
"""Un corte de red con Telegram no despierta al dueño; una racha sí.

El dueño recibió tres avisos del error handler en un día:

    ⚠️ Algo falló adentro del bot
    TimedOut: Timed out          — _httprequest.py:293, en do_request
    NetworkError: httpx.ReadError — _httprequest.py:300, en do_request

No son bugs: es el transporte HTTP de python-telegram-bot perdiendo la conexión
mientras hace `getUpdates`. PTB reintenta solo y sigue andando. Avisar de cada
uno entrena a ignorar los avisos, que es exactamente lo que el throttle de 6 h
quería evitar — y encima no los agrupaba, porque cada uno cae en una línea
distinta del mismo archivo y el identificador incluye archivo:línea.

Medido en `bot.log`: 1 o 2 por día durante meses, 5 el 2-sep. O sea son raros
pero constantes, y ninguno requirió jamás que alguien hiciera algo.

⚠️ OJO CON LA JERARQUÍA DE PTB: `BadRequest` **hereda de `NetworkError`**.
Silenciar `NetworkError` entero taparía errores de verdad — una petición mal
armada es un bug del bot. Solo se callan `TimedOut` y `NetworkError` exacto.

Lo que SÍ merece aviso es la racha: si la red se cae de verdad, el bot deja de
recibir mensajes y eso hay que saberlo.
"""
import asyncio
import logging
import types

import pytest
from telegram.error import BadRequest, Forbidden, NetworkError, TimedOut

from handlers.errores import UMBRAL_RACHA_RED, manejar_error


class _Bot:
    def __init__(self):
        self.enviados = []

    async def send_message(self, chat_id, text, **kw):
        self.enviados.append(text)


def _correr(error, bot, datos):
    ctx = types.SimpleNamespace(error=error, bot=bot, bot_data=datos)
    asyncio.run(manejar_error(None, ctx))


def _con_traceback(exc):
    try:
        raise exc
    except Exception as e:      # noqa: BLE001
        return e


@pytest.fixture
def bot():
    return _Bot()


@pytest.fixture
def datos():
    return {"owner_chat_id": "42"}


class TestNoMolesta:
    @pytest.mark.parametrize("exc", [
        TimedOut(),
        NetworkError("httpx.ReadError: "),
    ])
    def test_un_corte_de_red_suelto_no_avisa(self, exc, bot, datos):
        _correr(_con_traceback(exc), bot, datos)
        assert bot.enviados == []

    def test_igual_queda_en_el_log(self, bot, datos, caplog):
        with caplog.at_level(logging.INFO, logger="handlers.errores"):
            _correr(_con_traceback(TimedOut()), bot, datos)
        assert "red" in caplog.text.lower() or "TimedOut" in caplog.text


class TestSiEsUnBugSiAvisa:
    def test_BadRequest_avisa_aunque_herede_de_NetworkError(self, bot, datos):
        """La trampa: BadRequest ES NetworkError en PTB, pero es un bug."""
        _correr(_con_traceback(BadRequest("campo invalido")), bot, datos)
        assert len(bot.enviados) == 1

    def test_Forbidden_avisa(self, bot, datos):
        _correr(_con_traceback(Forbidden("bloqueado")), bot, datos)
        assert len(bot.enviados) == 1

    def test_un_error_del_codigo_avisa(self, bot, datos):
        _correr(_con_traceback(KeyError("nombre")), bot, datos)
        assert len(bot.enviados) == 1


class TestLaRachaSiAvisa:
    def test_muchos_cortes_seguidos_terminan_avisando(self, bot, datos):
        """Si la red se cae de verdad, el bot deja de recibir mensajes."""
        for _ in range(UMBRAL_RACHA_RED):
            _correr(_con_traceback(TimedOut()), bot, datos)
        assert len(bot.enviados) == 1
        # El mensaje habla de "conexión", no de "red": es lo que el dueño
        # puede ir a revisar. Que diga la palabra tecnica no importa.
        assert "conexi" in bot.enviados[0].lower()
        assert "telegram" in bot.enviados[0].lower()

    def test_avisa_UNA_sola_vez_por_racha(self, bot, datos):
        for _ in range(UMBRAL_RACHA_RED * 3):
            _correr(_con_traceback(TimedOut()), bot, datos)
        assert len(bot.enviados) == 1

    def test_el_aviso_dice_cuantos_fueron(self, bot, datos):
        for _ in range(UMBRAL_RACHA_RED):
            _correr(_con_traceback(TimedOut()), bot, datos)
        assert str(UMBRAL_RACHA_RED) in bot.enviados[0]
