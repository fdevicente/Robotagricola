# -*- coding: utf-8 -*-
"""El reporte mensual en PDF: fuera del loop y con navegador de repuesto.

Dos cosas lo tienen roto o en riesgo, las dos vistas en el log de producción:

  · **Roto desde el 1-jul-2026.** `job_reporte_mensual` es una corrutina y
    llamaba a `generar_reporte_pdf()` en el hilo del loop. La API *sync* de
    Playwright se niega a correr ahí: "It looks like you are using Playwright
    Sync API inside the asyncio loop". Esa fue la única corrida que alcanzó a
    intentarlo: el 1-ago y el 1-sep el bot estaba abajo.
  · **Sin navegador de repuesto.** El 25-sep a las 08:00 Playwright no encontró
    su propio Chromium (`Executable doesn't exist ...chromium-1223`, dos veces
    en todo el log). El scraper del banco siguió andando porque cae al Chrome
    del sistema; el reporte no tiene esa red y se habría caído.

La escalera de navegadores vive en `modules/navegador.py` y la usan los dos:
importar el scraper para conseguirla traería consigo las claves del banco, que
las lee al importarse.
"""
import asyncio
import types

import pytest


class _Bot:
    def __init__(self):
        self.documentos = []
        self.mensajes = []

    async def send_document(self, chat_id=None, document=None, filename=None,
                            caption=None, **kw):
        self.documentos.append((chat_id, filename, document.read()[:4]))

    async def send_message(self, chat_id=None, text=None, **kw):
        self.mensajes.append(text)


def _contexto(bot=None):
    return types.SimpleNamespace(bot=bot or _Bot(),
                                 bot_data={"owner_chat_id": 42})


# ── El job no puede bloquear el loop ───────────────────────────────────────

def test_el_reporte_se_genera_fuera_del_loop(tmp_path, monkeypatch):
    """Playwright sync dentro del loop revienta: tiene que ir en un hilo."""
    from handlers import cash_flow_jobs
    from modules.cash_flow import reporte_pdf

    pdf = tmp_path / "reporte.pdf"
    pdf.write_bytes(b"%PDF-1.4 falso")
    en_el_loop = []

    def falso_generar(year, month, *a, **k):
        try:
            asyncio.get_running_loop()
            en_el_loop.append(True)      # está en el hilo del loop: mal
        except RuntimeError:
            en_el_loop.append(False)     # hilo aparte: bien
        return str(pdf)

    monkeypatch.setattr(reporte_pdf, "generar_reporte_pdf", falso_generar)
    ctx = _contexto()

    asyncio.run(cash_flow_jobs.job_reporte_mensual(ctx))

    assert en_el_loop == [False], \
        "el PDF se genera en el hilo del loop: Playwright sync se niega a correr ahí"
    assert ctx.bot.documentos, "no mandó el PDF"


def test_si_el_pdf_falla_el_dueno_se_entera(tmp_path, monkeypatch):
    """Lo que ya hacía bien: avisar en vez de callarse."""
    from handlers import cash_flow_jobs
    from modules.cash_flow import reporte_pdf

    def explota(year, month, *a, **k):
        raise RuntimeError("no hay navegador")

    monkeypatch.setattr(reporte_pdf, "generar_reporte_pdf", explota)
    ctx = _contexto()

    asyncio.run(cash_flow_jobs.job_reporte_mensual(ctx))

    assert ctx.bot.mensajes, "se cayó sin avisar"
    assert "no hay navegador" in " ".join(ctx.bot.mensajes)


# ── La escalera de navegadores ─────────────────────────────────────────────

class _Chromium:
    """Playwright de mentira: solo abre con el canal que se le diga."""

    def __init__(self, canal_bueno):
        self.canal_bueno = canal_bueno
        self.intentos = []

    def launch(self, **kw):
        self.intentos.append(kw.get("channel"))
        if kw.get("channel") != self.canal_bueno:
            raise Exception("Executable doesn't exist at "
                            r"C:\ms-playwright\chromium-1223\chrome.exe")
        return f"navegador-{self.canal_bueno}"


def _playwright(canal_bueno):
    chromium = _Chromium(canal_bueno)
    return types.SimpleNamespace(chromium=chromium), chromium


def test_si_falta_el_chromium_de_playwright_usa_el_chrome_del_sistema():
    """Lo que salvó al scraper el 25-sep."""
    from modules.navegador import lanzar_chromium
    p, chromium = _playwright("chrome")

    assert lanzar_chromium(p) == "navegador-chrome"
    assert chromium.intentos == [None, "chromium", "chrome"], \
        "no probó la escalera completa antes de rendirse"


def test_con_todo_instalado_usa_el_navegador_de_playwright():
    from modules.navegador import lanzar_chromium
    p, chromium = _playwright(None)

    assert lanzar_chromium(p) == "navegador-None"
    assert chromium.intentos == [None], "probó fallbacks sin necesidad"


def test_si_no_hay_ningun_navegador_lo_dice_con_la_solucion():
    from modules.navegador import lanzar_chromium
    p, _ = _playwright("no-existe")

    with pytest.raises(RuntimeError) as e:
        lanzar_chromium(p, "para el reporte mensual")

    assert "playwright install" in str(e.value)
    assert "para el reporte mensual" in str(e.value)


def test_el_scraper_del_banco_usa_la_misma_escalera():
    """Una sola escalera: si se arregla acá, se arregla para los dos."""
    import scotiabank_scraper
    p, chromium = _playwright("chrome")

    assert scotiabank_scraper._lanzar_chromium(p) == "navegador-chrome"
    assert chromium.intentos == [None, "chromium", "chrome"]


def test_el_reporte_no_lanza_el_navegador_por_su_cuenta():
    """Si vuelve a llamar a p.chromium.launch() directo, se queda sin red."""
    import inspect

    from modules.cash_flow import reporte_pdf
    fuente = inspect.getsource(reporte_pdf)

    assert "lanzar_chromium" in fuente, "el reporte no usa la escalera compartida"
    assert "chromium.launch(" not in fuente, "lanza el navegador por su cuenta"
