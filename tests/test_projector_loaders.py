import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import (
    SHEET_NAME, ensure_new_columns,
    COL_CATEGORIA, COL_CULTIVO,
)


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def test_load_egresos_returns_aggregated(m):
    from modules.cash_flow.projector import load_historical_egresos
    egresos = load_historical_egresos(excel_path=m)
    assert isinstance(egresos, dict)
    for k, v in egresos.items():
        assert len(k) == 4
        assert isinstance(v, (int, float))


def test_load_egresos_skips_uncategorized(m):
    from modules.cash_flow.projector import load_historical_egresos
    egresos = load_historical_egresos(excel_path=m)
    cats = {k[2] for k in egresos.keys()}
    assert "REVISAR" not in cats


def test_load_egresos_by_year(m):
    from modules.cash_flow.projector import load_historical_egresos
    egresos = load_historical_egresos(excel_path=m, year=2025)
    years = {k[0] for k in egresos.keys()}
    assert years == {2025} or years == set()
