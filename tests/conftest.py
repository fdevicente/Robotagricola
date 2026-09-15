# -*- coding: utf-8 -*-
"""Fixtures de toda la suite.

🔴 La suite ESCRIBÍA EN PRODUCCIÓN. Medido el 15-sep-2026 comparando las
carpetas reales antes y después de una corrida completa: respaldos falsos en
Dropbox\\Backups que además se subían a Drive, y 497 correcciones falsas en el
correcciones_log.json real. Ver tests/test_suite_no_toca_produccion.py.

Aquí se desvían, para CADA test, las carpetas de producción que el código lee
de `config` al momento de usarlas. Un módulo que las congela al importarse
(`from config import X` arriba del archivo) NO queda cubierto: por eso
_registrar_correccion pasó a leer `config.DOWNLOAD_DIR` al momento.
"""
import pytest


@pytest.fixture(autouse=True)
def _produccion_desviada(tmp_path_factory, monkeypatch):
    import config
    base = tmp_path_factory.mktemp("produccion")
    facturas = base / "Facturas Recibidas por Telegram"
    facturas.mkdir()
    monkeypatch.setattr(config, "DROPBOX_BACKUP_PATH", str(base / "Backups"))
    monkeypatch.setattr(config, "DRIVE_COLA_PATH", str(base / "drive_cola.jsonl"))
    monkeypatch.setattr(config, "DOWNLOAD_DIR", str(facturas))
