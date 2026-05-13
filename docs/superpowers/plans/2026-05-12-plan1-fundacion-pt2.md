# Plan 1: Fundación — Parte 2/3

### Task 3: Función para crear hojas nuevas en Master

**Files:**
- Modify: `Robot/excel_manager.py`
- Test: `Robot/tests/test_excel_sheets.py`

- [ ] **Step 1: Write test**

```python
# tests/test_excel_sheets.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import ensure_cash_flow_sheets, COSECHAS_SHEET, GUIAS_SHEET
from excel_manager import FLUJO_CAJA_SHEET, AJUSTES_SHEET, CONFIG_SHEET, HECTAREAS_SHEET

TEST_EXCEL = "tests/fixtures/test_master.xlsx"

@pytest.fixture
def test_excel(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "MASTER Agricola Santa Elisa.xlsx")
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
    params = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(2, ws.max_row+1)}
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
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_excel_sheets.py -v`
Expected: ImportError — ensure_cash_flow_sheets not defined

- [ ] **Step 3: Implement ensure_cash_flow_sheets**

```python
# Agregar a excel_manager.py
from config import CASH_FLOW_CONFIG

def ensure_cash_flow_sheets(path=None):
    """Crea hojas nuevas en Master si no existen. Idempotente."""
    from config import EXCEL_PATH
    path = path or EXCEL_PATH
    wb = load_workbook(path)
    
    _ensure_sheet_with_headers(wb, COSECHAS_SHEET, [
        "Año", "Cultivo", "Kg total", "Exportadora", "Kg asignados",
        "Precio USD/kg", "N° cuotas", "Cuota #", "Fecha estimada",
        "Monto USD estimado", "Tipo cuota", "Estado",
        "Fecha real recibido", "Monto real recibido", "Moneda recibida", "Notas"
    ])
    
    _ensure_sheet_with_headers(wb, GUIAS_SHEET, [
        "Fecha", "N° Guía", "Cultivo", "Kg", "Exportadora destino",
        "Camión / Conductor", "Sector / Equipo", "Año cosecha",
        "Origen", "PDF_path", "Notas"
    ])
    
    _ensure_sheet_with_headers(wb, AJUSTES_SHEET, [
        "Fecha agregado", "Mes proyectado", "Categoria",
        "Cultivo", "Monto", "Razón", "Activo"
    ])
    
    if CONFIG_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(CONFIG_SHEET)
        ws.append(["Parámetro", "Valor"])
        for k, v in CASH_FLOW_CONFIG.items():
            ws.append([k, v])
    
    if HECTAREAS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(HECTAREAS_SHEET)
        ws.append(["Año", "Nogales", "Cerezos", "Avellanos", "Notas"])
        ws.append([2024, 65, 1.8, 0, "Sin avellanos"])
        ws.append([2025, 54, 3.8, 11.5, "Inicio replante avellanos"])
        ws.append([2026, 43, 3.8, 26.5, "+15 hc avellanos"])
    
    if FLUJO_CAJA_SHEET not in wb.sheetnames:
        wb.create_sheet(FLUJO_CAJA_SHEET)
    
    _save_wb(wb, path)
    wb.close()


def _ensure_sheet_with_headers(wb, name, headers):
    """Crea hoja con headers si no existe."""
    if name not in wb.sheetnames:
        ws = wb.create_sheet(name)
        ws.append(headers)
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_excel_sheets.py -v`

- [ ] **Step 5: Commit**

```bash
git add excel_manager.py tests/test_excel_sheets.py
git commit -m "feat: ensure_cash_flow_sheets creates 6 new sheets in Master"
```

---

### Task 4: Agregar columnas nuevas a Facturas y Cuenta Banco

**Files:**
- Modify: `Robot/excel_manager.py`
- Test: `Robot/tests/test_excel_columns.py`

- [ ] **Step 1: Write test**

```python
# tests/test_excel_columns.py
import shutil, pytest
from openpyxl import load_workbook
from excel_manager import (ensure_new_columns, COL_CATEGORIA, COL_CULTIVO,
    COL_CONFIANZA, COL_CATEGORIZADO_POR, COL_BANCO_TIPO,
    SHEET_NAME, CUENTA_BANCO_SHEET)

@pytest.fixture
def test_excel(tmp_path):
    import os
    src = os.path.join(os.path.dirname(__file__), "..", "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test_master.xlsx"
    shutil.copy2(src, dst)
    return str(dst)

def test_facturas_new_headers(test_excel):
    ensure_new_columns(test_excel)
    wb = load_workbook(test_excel, read_only=True)
    ws = wb[SHEET_NAME]
    assert ws.cell(1, COL_CATEGORIA).value == "Categoria"
    assert ws.cell(1, COL_CULTIVO).value == "Cultivo"
    assert ws.cell(1, COL_CONFIANZA).value == "Confianza"
    assert ws.cell(1, COL_CATEGORIZADO_POR).value == "Categorizado_por"
    wb.close()

def test_banco_new_headers(test_excel):
    ensure_new_columns(test_excel)
    wb = load_workbook(test_excel, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(1, COL_BANCO_TIPO).value == "Tipo"
    assert ws.cell(1, 8).value == "Categoria"
    assert ws.cell(1, 9).value == "Cultivo"
    assert ws.cell(1, 10).value == "Factura_linkeada"
    wb.close()

def test_idempotent(test_excel):
    ensure_new_columns(test_excel)
    ensure_new_columns(test_excel)
    wb = load_workbook(test_excel, read_only=True)
    ws = wb[SHEET_NAME]
    assert ws.cell(1, COL_CATEGORIA).value == "Categoria"
    wb.close()
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement ensure_new_columns**

```python
# Agregar a excel_manager.py

def ensure_new_columns(path=None):
    """Agrega headers de columnas nuevas si no existen. Idempotente."""
    from config import EXCEL_PATH
    path = path or EXCEL_PATH
    wb = load_workbook(path)
    
    # Facturas: cols Q-T
    ws = wb[SHEET_NAME]
    new_fact = {COL_CATEGORIA: "Categoria", COL_CULTIVO: "Cultivo",
                COL_CONFIANZA: "Confianza", COL_CATEGORIZADO_POR: "Categorizado_por"}
    for col, header in new_fact.items():
        if ws.cell(1, col).value != header:
            ws.cell(1, col, header)
    
    # Cuenta Banco: cols G-J
    ws2 = wb[CUENTA_BANCO_SHEET]
    new_banco = {COL_BANCO_TIPO: "Tipo", COL_BANCO_CATEGORIA: "Categoria",
                 COL_BANCO_CULTIVO: "Cultivo", COL_BANCO_FACTURA_LINK: "Factura_linkeada"}
    for col, header in new_banco.items():
        if ws2.cell(1, col).value != header:
            ws2.cell(1, col, header)
    
    _save_wb(wb, path)
    wb.close()
```

- [ ] **Step 4: Run test — expect PASS**

- [ ] **Step 5: Commit**

```bash
git add excel_manager.py tests/test_excel_columns.py
git commit -m "feat: add Categoria/Cultivo columns to Facturas and Cuenta Banco"
```
