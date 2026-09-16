# -*- coding: utf-8 -*-
"""La suite no puede escribir en las carpetas de producción.

🔴 Medido el 15-sep-2026 corriendo la suite completa y comparando las carpetas
reales antes y después:
  · tests/test_batch_categorize.py → batch_categorize_history → backup_master
    sin carpeta propia: dejaba un snapshot en Dropbox\\Backups, pisaba
    current.xlsx y lo encolaba para subirlo a Drive. Y como la rotación borra
    por orden alfabético, cada respaldo falso empujaba fuera uno real.
  · un test de facturas → _registrar_correccion → escribía en el
    correcciones_log.json REAL: +2.456 bytes por corrida. 497 de sus 616
    entradas eran de la suite.

Mismo patrón que el FileHandler de bot.log (d83337d) y la cola de Drive: un
default de producción escondido dentro de algo que el test creía aislado.
"""
import os


def _dentro_de(ruta, base):
    ruta, base = os.path.abspath(str(ruta)), os.path.abspath(str(base))
    return os.path.commonpath([ruta, base]) == base


def test_los_respaldos_de_la_suite_van_a_una_carpeta_temporal(tmp_path_factory):
    import config
    assert _dentro_de(config.DROPBOX_BACKUP_PATH, tmp_path_factory.getbasetemp()), \
        config.DROPBOX_BACKUP_PATH


def test_la_cola_de_drive_de_la_suite_es_temporal(tmp_path_factory):
    import config
    assert _dentro_de(config.DRIVE_COLA_PATH, tmp_path_factory.getbasetemp()), \
        config.DRIVE_COLA_PATH


def test_la_carpeta_de_facturas_de_la_suite_es_temporal(tmp_path_factory):
    import config
    assert _dentro_de(config.DOWNLOAD_DIR, tmp_path_factory.getbasetemp()), \
        config.DOWNLOAD_DIR


def test_el_lock_del_master_de_la_suite_es_temporal(tmp_path_factory):
    """El lock vive junto al bot. Sin desviarlo, la suite crearía archivos en
    producción y le disputaría el Master al bot corriendo."""
    import config
    assert _dentro_de(config.LOCK_DIR, tmp_path_factory.getbasetemp()), config.LOCK_DIR


def test_la_correccion_se_guarda_donde_dice_config_al_momento(tmp_path, monkeypatch):
    """_registrar_correccion usaba DOWNLOAD_DIR congelado al importar el módulo,
    así que redirigir config no la alcanzaba y escribía en el log real."""
    import config
    from handlers import facturas
    destino = tmp_path / "segun_config"
    trampa = tmp_path / "congelado_al_importar"
    destino.mkdir()
    trampa.mkdir()
    monkeypatch.setattr(config, "DOWNLOAD_DIR", str(destino))
    # La trampa es para que el ROJO tampoco toque producción
    monkeypatch.setattr(facturas, "DOWNLOAD_DIR", str(trampa))

    facturas._registrar_correccion({"Rut": "76.000.000-0"}, "Total Factura", 100, 200)

    assert (destino / "correcciones_log.json").exists()
    assert not (trampa / "correcciones_log.json").exists()
