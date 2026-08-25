# tests/test_config.py
from config import (EXCEL_PATH, DROPBOX_BACKUP_PATH, GUIAS_DIR,
                     DOCUMENTOS_DIR, FXP_PATH)

def test_paths_exist():
    assert EXCEL_PATH is not None
    assert DROPBOX_BACKUP_PATH is not None

def test_cash_flow_config_defaults():
    from config import CASH_FLOW_CONFIG
    assert CASH_FLOW_CONFIG['saldo_minimo_pct'] == 0.10
    assert CASH_FLOW_CONFIG['umbral_alerta_cat_pct'] == 0.90
    assert CASH_FLOW_CONFIG['umbral_confianza'] == 0.85
    assert CASH_FLOW_CONFIG['ventana_match_dias'] == 15
    assert CASH_FLOW_CONFIG['fecha_limite_cerezas'] == '12-15'
    assert CASH_FLOW_CONFIG['fecha_limite_nueces'] == '05-30'
    assert CASH_FLOW_CONFIG['dias_sin_guia_cierre'] == 7
    # 2026-08-17: 1.000 -> 910. Promedio de las dos unicas conversiones reales
    # del banco: 908,10 (16-mar-2026) y 932,00 (26-mar-2026).
    assert CASH_FLOW_CONFIG['usd_clp_estimado'] == 910
