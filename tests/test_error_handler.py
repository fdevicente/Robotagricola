# -*- coding: utf-8 -*-
"""Una excepción en un handler o un job tiene que AVISARLE al dueño.

BUG REAL, visto en producción el 1-sep-2026 10:22:

    telegram.ext.Application - ERROR - No error handlers are registered
    TimeoutError: The read operation timed out
      ... handlers/drive_jobs.py línea 178, job_drive_entrada

`main.py` no llamaba nunca a `add_error_handler`, así que CUALQUIER excepción
en cualquier handler o job moría en el log y nadie se enteraba. Es la
explicación más probable del parte de Juan que se perdió en silencio el
24-ago: el bot respondió dos veces, no guardó nada y no logueó ningún error.

⚠️ El aviso TIENE que estar throttleado. El error que lo destapó se repite cada
15 minutos: sin throttle serían ~96 mensajes al día por una sola falla de red,
que es la forma más rápida de que el dueño termine silenciando al bot.
"""
import asyncio
import logging
import types

from handlers.errores import VENTANA_AVISO, manejar_error


class _Bot:
    def __init__(self, falla=False):
        self.enviados = []
        self.falla = falla

    async def send_message(self, chat_id, text, **kw):
        if self.falla:
            raise RuntimeError("Telegram caído")
        self.enviados.append((chat_id, text))


def _ctx(error, bot=None, bot_data=None):
    return types.SimpleNamespace(
        error=error, bot=bot or _Bot(),
        bot_data={"owner_chat_id": "42"} if bot_data is None else bot_data)


def _correr(ctx, update=None):
    asyncio.run(manejar_error(update, ctx))


def _error(tipo=TimeoutError, mensaje="The read operation timed out"):
    """Una excepción con traceback de verdad, como la que llega en producción."""
    try:
        raise tipo(mensaje)
    except Exception as e:      # noqa: BLE001 - queremos el traceback poblado
        return e


# ── Avisa ──────────────────────────────────────────────────────────────────

def test_le_avisa_al_dueno():
    ctx = _ctx(_error())
    _correr(ctx)
    assert len(ctx.bot.enviados) == 1
    assert ctx.bot.enviados[0][0] == 42


def test_el_aviso_dice_QUE_error_fue():
    ctx = _ctx(_error())
    _correr(ctx)
    assert "TimeoutError" in ctx.bot.enviados[0][1]


def test_el_aviso_dice_DONDE_fue():
    """Sin el archivo y la línea el aviso no sirve para diagnosticar."""
    ctx = _ctx(_error())
    _correr(ctx)
    assert "test_error_handler.py" in ctx.bot.enviados[0][1]


def test_deja_el_traceback_completo_en_el_log(caplog):
    with caplog.at_level(logging.ERROR, logger="handlers.errores"):
        _correr(_ctx(_error()))
    assert "Traceback" in caplog.text


# ── No spamea ──────────────────────────────────────────────────────────────

def test_el_mismo_error_repetido_avisa_UNA_vez():
    """El caso real: un timeout de Drive cada 15 min son 96 al día."""
    datos = {"owner_chat_id": "42"}
    bot = _Bot()
    for _ in range(10):
        _correr(_ctx(_error(), bot=bot, bot_data=datos))
    assert len(bot.enviados) == 1


def test_un_error_DISTINTO_si_avisa():
    datos = {"owner_chat_id": "42"}
    bot = _Bot()
    _correr(_ctx(_error(), bot=bot, bot_data=datos))
    _correr(_ctx(_error(ValueError, "otra cosa"), bot=bot, bot_data=datos))
    assert len(bot.enviados) == 2


def test_pasada_la_ventana_vuelve_a_avisar():
    """Que se repita mañana es información nueva, no ruido."""
    datos = {"owner_chat_id": "42"}
    bot = _Bot()
    _correr(_ctx(_error(), bot=bot, bot_data=datos))
    for clave in list(datos):
        if clave.startswith("_error_visto_"):
            datos[clave] -= VENTANA_AVISO + 1
    _correr(_ctx(_error(), bot=bot, bot_data=datos))
    assert len(bot.enviados) == 2


# ── Nunca puede voltear el bot ─────────────────────────────────────────────

def test_si_no_puede_avisar_NO_lanza():
    """Un error handler que falla es peor que no tenerlo."""
    _correr(_ctx(_error(), bot=_Bot(falla=True)))


def test_sin_chat_conocido_NO_lanza():
    _correr(_ctx(_error(), bot_data={}))


def test_sin_error_en_el_contexto_NO_lanza():
    ctx = _ctx(None)
    _correr(ctx)
    assert ctx.bot.enviados == []


# ── Queda registrado en main.py ────────────────────────────────────────────

def test_main_registra_el_error_handler():
    """El bug era justamente que nadie lo registraba."""
    from pathlib import Path
    src = Path(__file__).resolve().parent.parent / "main.py"
    texto = src.read_text(encoding="utf-8")
    assert "add_error_handler" in texto, "main.py no registra el error handler"
    assert "manejar_error" in texto
