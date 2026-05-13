# tests/test_excel_sheets.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import ensure_cash_flow_sheets, COSECHAS_SHEET, GUIAS_SHEET
from excel_manager import FLUJO_CAJA_SHEET, AJUSTES_SHEET, CONFIG_SHEET, HECTAREAS_SHEET


@pytest.fixture
def test_excel(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..", "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test_master.xlsx"
    shutil.copy2(src, dst)
    return str(dst)


def test_creates_missing_sheets(test_excel):
    ensure_cash_flow_sheets(test_excel)
    wb = load_workbook(test_excel, read_only=True)
    names = wb.sheetnames
    assert COSECHAS_SHEET in names
    assert GUIAS_SHEET in names
    assert CONFIG_SHEET in names
    assert HECTAREAS_SHEET in names
    assert AJUSTES_SHEET in names
    wb.close()


def test_idempotent(test_excel):
    ensure_cash_flow_sheets(test_excel)
    ensure_cash_flow_sheets(test_excel)  # no error on second run
    wb = load_workbook(test_excel, read_only=True)
    count = wb.sheetnames.count(CONFIG_SHEET)
    assert count == 1
    wb.close()


def test_config_has_defaults(test_excel):
    ensure_cash_flow_sheets(test_excel)
    wb = load_workbook(test_excel, read_only=True)
    ws = wb[CONFIG_SHEET]
    params = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(2, ws.max_row + 1)}
    assert params['saldo_minimo_pct'] == 0.10
    assert params['usd_clp_estimado'] == 1000
    wb.close()


def test_hectareas_has_data(test_excel):
    ensure_cash_flow_sheets(test_excel)
    wb = load_workbook(test_excel, read_only=True)
    ws = wb[HECTAREAS_SHEET]
    assert ws.cell(2, 1).value == 2024
    assert ws.cell(2, 2).value == 65  # nogales 2024
    assert ws.cell(4, 4).value == 26.5  # avellanos 2026
    wb.close()
