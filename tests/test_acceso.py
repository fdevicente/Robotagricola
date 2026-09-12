# -*- coding: utf-8 -*-
"""El bot solo le contesta a quien tiene que contestarle.

🔴 EL 12-SEP-2026 APARECIO UN DESCONOCIDO. Un usuario llamado "prince"
(user_id 6934038077) escribio /start y el bot le respondio con el MENU COMPLETO:
/saldo, /inventario, /reporte, /personal, /dashboard... y los que ESCRIBEN,
/deposito, /pagado, /bitacora, /uso y subir facturas.

No habia ningun control de acceso. Los unicos chequeos por user_id que existian
--AUTO_SAVE_USERS-- decidian si mostrar el teclado del capataz, no si la persona
tenia permiso. Un bot de Telegram es publico: cualquiera que sepa su nombre le
puede escribir.

Medido en el respaldo: en todo el historial solo hay DOS chats legitimos, el del
dueño y el de Juan. Ningun grupo.
"""
import asyncio
import inspect

import pytest

from handlers.acceso import autorizado, leer_autorizados

DUENO = 8684368429
JUAN = 8840816610
PRINCE = 6934038077
AUTORIZADOS = {DUENO, JUAN}


def test_el_dueno_y_juan_pasan():
    assert autorizado(DUENO, AUTORIZADOS)
    assert autorizado(JUAN, AUTORIZADOS)


def test_el_desconocido_no_pasa():
    assert not autorizado(PRINCE, AUTORIZADOS)


def test_un_update_sin_usuario_no_pasa():
    """Sin saber quien es, no se autoriza. Callar es mas barato que filtrar."""
    assert not autorizado(None, AUTORIZADOS)
    assert not autorizado("", AUTORIZADOS)


def test_el_id_como_texto_tambien_calza():
    """Telegram y las variables de entorno mezclan int y str."""
    assert autorizado(str(JUAN), AUTORIZADOS)
    assert autorizado(JUAN, {str(JUAN)})


def test_una_lista_vacia_no_autoriza_a_nadie():
    """⚠️ Si la lista se queda vacia por un error de configuracion, el bot se
    cierra entero. Es molesto, pero es el lado seguro: abrirse a todos no."""
    assert not autorizado(DUENO, set())
    assert not autorizado(DUENO, None)


@pytest.mark.parametrize("crudo,esperado", [
    ("8684368429,8840816610", {8684368429, 8840816610}),
    (" 8684368429 ; 8840816610 ", {8684368429, 8840816610}),
    ("8684368429", {8684368429}),
    ("", set()),
    ("basura,8840816610", {8840816610}),
])
def test_leer_la_lista_de_la_variable_de_entorno(crudo, esperado):
    assert leer_autorizados(crudo) == esperado


def test_el_guardia_corta_la_cadena_de_handlers():
    """Tiene que LEVANTAR ApplicationHandlerStop, no solo devolver.

    Si solo devolviera, los handlers siguientes atenderian al desconocido igual:
    el guardia seria decorativo.
    """
    from handlers import acceso
    fuente = inspect.getsource(acceso.guardia)
    assert "ApplicationHandlerStop" in fuente


def test_el_guardia_esta_registrado_antes_que_los_comandos():
    """En un grupo menor que 0, y despues del espejo para que el intento quede
    guardado en el respaldo crudo aunque se rechace."""
    import main
    fuente = inspect.getsource(main)
    assert "guardia" in fuente, "el guardia no está registrado en main.py"
    i_guardia = fuente.index("guardia")
    i_start = fuente.index('CommandHandler("start"')
    assert i_guardia < i_start


# ── El guardia corriendo de verdad ─────────────────────────────────────────


class _Usuario:
    def __init__(self, uid, nombre="X"):
        self.id, self.full_name = uid, nombre


class _Mensaje:
    def __init__(self):
        self.respuestas = []

    async def reply_text(self, texto, **kw):
        self.respuestas.append(texto)


class _Update:
    def __init__(self, uid):
        self.effective_user = _Usuario(uid) if uid else None
        self.message = _Mensaje()


class _Bot:
    def __init__(self):
        self.enviados = []

    async def send_message(self, chat_id, text, **kw):
        self.enviados.append((chat_id, text))


class _Ctx:
    def __init__(self):
        self.bot = _Bot()
        self.bot_data = {"owner_chat_id": DUENO}


def test_al_autorizado_lo_deja_pasar_sin_decir_nada():
    from handlers.acceso import guardia
    up, ctx = _Update(JUAN), _Ctx()
    asyncio.run(guardia(up, ctx))            # no levanta nada
    assert up.message.respuestas == []
    assert ctx.bot.enviados == []


def test_al_desconocido_lo_corta_y_le_avisa_al_dueno():
    from telegram.ext import ApplicationHandlerStop

    from handlers.acceso import guardia
    up, ctx = _Update(PRINCE), _Ctx()
    with pytest.raises(ApplicationHandlerStop):
        asyncio.run(guardia(up, ctx))
    assert len(ctx.bot.enviados) == 1
    destino, texto = ctx.bot.enviados[0]
    assert destino == DUENO
    assert str(PRINCE) in texto
    assert up.message.respuestas                       # y se le dice que no


def test_al_dueno_se_le_avisa_UNA_vez_por_desconocido():
    """Si alguien insiste, no se le llena el chat al dueño."""
    from telegram.ext import ApplicationHandlerStop

    from handlers.acceso import guardia
    ctx = _Ctx()
    for _ in range(4):
        with pytest.raises(ApplicationHandlerStop):
            asyncio.run(guardia(_Update(PRINCE), ctx))
    assert len(ctx.bot.enviados) == 1


def test_no_se_le_cuenta_al_desconocido_lo_que_maneja_el_bot():
    """El menú completo es justo lo que recibió 'prince' el 12-sep."""
    from telegram.ext import ApplicationHandlerStop

    from handlers.acceso import guardia
    up, ctx = _Update(PRINCE), _Ctx()
    with pytest.raises(ApplicationHandlerStop):
        asyncio.run(guardia(up, ctx))
    dicho = " ".join(up.message.respuestas).lower()
    for filtrado in ("/saldo", "/deposito", "/inventario", "/reporte",
                     "/dashboard", "/personal"):
        assert filtrado not in dicho


# ── El token no puede quedar en el log ─────────────────────────────────────


def test_httpx_no_logea_en_INFO():
    """httpx logea cada getUpdates y la URL LLEVA EL TOKEN.

    Medido el 12-sep-2026: 557.257 líneas con el token en un bot.log de 95 MB.
    Además de regalar el token a quien reciba el log, tapaban todo lo demás:
    diagnosticar el horómetro trabado del 9-sep costó justamente por eso.
    """
    import logging

    import main            # noqa: F401  — configura el logging al importarse
    for ruidoso in ("httpx", "httpcore"):
        assert logging.getLogger(ruidoso).level >= logging.WARNING, (
            "%s en INFO escribe el token del bot en bot.log" % ruidoso)


def test_los_tests_no_escriben_en_el_bot_log_de_produccion():
    """⚠️ Importar main engancha un FileHandler al bot.log REAL.

    Medido el 12-sep-2026: correr la suite dejó 18 líneas "Acceso denegado:
    X (user_id 6934038077)" en el log de producción — mis propios fakes. Leídas
    después parecen intentos de intrusión de verdad. Es la misma clase de señal
    falsa que el "bot apagado 62h" de agosto: un log que miente cuesta más que
    un log que falta.
    """
    import logging
    import os

    import main            # noqa: F401
    archivos = [h for h in logging.getLogger().handlers
                if isinstance(h, logging.FileHandler)
                and os.path.basename(getattr(h, "baseFilename", "")) == "bot.log"]
    assert archivos == [], (
        "la suite está escribiendo en bot.log: %s" % archivos)
