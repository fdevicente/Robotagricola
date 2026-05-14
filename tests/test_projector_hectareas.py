import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import ensure_cash_flow_sheets, AJUSTES_SHEET


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_cash_flow_sheets(str(dst))
    return str(dst)


def test_load_hectareas_returns_dict(m):
    from modules.cash_flow.projector import load_hectareas
    hc = load_hectareas(excel_path=m)
    assert 2024 in hc and 2025 in hc and 2026 in hc
    assert hc[2024]["NOGALES"] == 65
    assert hc[2026]["AVELLANOS"] == 26.5


def test_load_ajustes_empty_if_no_data(m):
    from modules.cash_flow.projector import load_ajustes_manuales
    ajustes = load_ajustes_manuales(excel_path=m)
    assert ajustes == []


def test_load_ajustes_filters_inactive(m):
    from modules.cash_flow.projector import load_ajustes_manuales
    wb = load_workbook(m)
    ws = wb[AJUSTES_SHEET]
    ws.append(["2026-05-01", "2026-07", "Riego", "GENERAL", 5000000, "Bomba nueva", True])
    ws.append(["2026-05-01", "2026-08", "Fertilizantes", "NOGALES", 2000000, "Test", False])
    wb.save(m); wb.close()

    ajustes = load_ajustes_manuales(excel_path=m)
    assert len(ajustes) == 1
    assert ajustes[0]["categoria"] == "Riego"
    assert ajustes[0]["monto"] == 5000000
