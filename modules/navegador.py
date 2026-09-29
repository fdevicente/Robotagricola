# -*- coding: utf-8 -*-
"""Abrir Chromium sin morir por una instalación incompleta de Playwright.

El 25-sep-2026 a las 08:00 Playwright no encontró su propio navegador
("Executable doesn't exist at ...chromium-1223\\chrome-win64\\chrome.exe") y el
scraper del banco siguió andando solo porque probó otros: terminó usando el
Chrome del sistema. Fue la única vez en todo el log, y el archivo estaba en su
sitio antes y después, así que lo más probable es que algo lo retuviera un
instante. Una red que se usa una vez al año sigue valiendo la pena.

Esta es esa escalera, en un lugar donde la usen los dos que la necesitan: el
banco y el reporte mensual en PDF. Vive aparte de `scotiabank_scraper` a
propósito, porque ese módulo lee las claves del banco al importarse y el
generador del reporte no tiene nada que hacer con ellas.
"""
import logging

logger = logging.getLogger(__name__)

ARGS = ["--no-sandbox", "--disable-blink-features=AutomationControlled"]

# De lo más específico a lo más general: el binario que Playwright usa por
# defecto, el Chromium completo que instala él mismo, y el Chrome que ya tiene
# el PC. El último no depende de Playwright para nada.
_INTENTOS = [
    ("headless shell (default)", {}),
    ("chromium completo",        {"channel": "chromium"}),
    ("chrome del sistema",       {"channel": "chrome"}),
]


def lanzar_chromium(p, para: str = ""):
    """Devuelve un navegador abierto, o explica cómo arreglarlo.

    `p` es lo que entrega `sync_playwright()`. `para` va en el mensaje de
    error: quien lo lee necesita saber qué se quedó sin hacer.
    """
    errores = []
    for nombre, extra in _INTENTOS:
        try:
            navegador = p.chromium.launch(headless=True, args=ARGS, **extra)
            if errores:          # solo avisar si hubo que recurrir a un fallback
                logger.warning("Chromium lanzado con fallback: %s", nombre)
            return navegador
        except Exception as e:   # noqa: BLE001 — cada intento falla a su manera
            errores.append(f"{nombre}: {str(e)[:120]}")
            logger.warning("Launch falló con %s: %s", nombre, str(e)[:150])
    raise RuntimeError(
        f"No pude abrir ningún navegador {para}".strip() + ".\n"
        "Solución: ejecuta en una terminal:\n"
        "  python -m playwright install chromium\n\n"
        "Detalle: " + " | ".join(errores))
