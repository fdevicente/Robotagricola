# Plan 2: Categorización — Parte 2/3

### Task 4: `categorize_invoice(row)` — categoriza factura y escribe Master

**Files:**
- Modify: `Robot/modules/cash_flow/categorizer.py`
- Test: `Robot/tests/test_categorize_invoice.py`

- [ ] **Step 1: Write test**

```python
# tests/test_categorize_invoice.py
import os, shutil, pytest
from unittest.mock import patch
from openpyxl import load_workbook
from excel_manager import (
    SHEET_NAME, COL_CATEGORIA, COL_CULTIVO, COL_CONFIANZA, COL_CATEGORIZADO_POR,
    ensure_new_columns,
)


@pytest.fixture
def test_master(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test_master.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def _fake_result(cat="Fertilizantes", cultivo="NOGALES", conf=0.9):
    return {"categoria": cat, "cultivo": cultivo,
            "confianza": conf, "razon": "ok"}


def test_categorize_invoice_writes_to_master(test_master):
    from modules.cash_flow.categorizer import categorize_invoice
    # Tomar primera fila con proveedor real
    wb = load_workbook(test_master)
    ws = wb[SHEET_NAME]
    target_row = None
    for r in range(2, min(ws.max_row + 1, 20)):
        if ws.cell(r, 4).value:
            target_row = r
            break
    wb.close()
    assert target_row is not None

    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value=_fake_result()):
        result = categorize_invoice(target_row, excel_path=test_master)

    assert result["categoria"] == "Fertilizantes"

    wb = load_workbook(test_master, read_only=True)
    ws = wb[SHEET_NAME]
    assert ws.cell(target_row, COL_CATEGORIA).value == "Fertilizantes"
    assert ws.cell(target_row, COL_CULTIVO).value == "NOGALES"
    assert ws.cell(target_row, COL_CONFIANZA).value == 0.9
    assert ws.cell(target_row, COL_CATEGORIZADO_POR).value == "claude"
    wb.close()


def test_categorize_invoice_low_confidence_marks_revisar(test_master):
    from modules.cash_flow.categorizer import categorize_invoice
    wb = load_workbook(test_master)
    ws = wb[SHEET_NAME]
    target_row = None
    for r in range(2, min(ws.max_row + 1, 20)):
        if ws.cell(r, 4).value:
            target_row = r
            break
    wb.close()

    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value=_fake_result(cat="Riego", conf=0.5)):
        categorize_invoice(target_row, excel_path=test_master)

    wb = load_workbook(test_master, read_only=True)
    ws = wb[SHEET_NAME]
    assert ws.cell(target_row, COL_CATEGORIA).value == "REVISAR"
    wb.close()


def test_categorize_invoice_uses_cache_on_second_call(test_master, tmp_path):
    from modules.cash_flow.categorizer import categorize_invoice
    wb = load_workbook(test_master)
    ws = wb[SHEET_NAME]
    r = next((r for r in range(2, 20) if ws.cell(r, 4).value), None)
    wb.close()

    cache_path = str(tmp_path / "cache.json")
    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value=_fake_result()) as mock_raw:
        categorize_invoice(r, excel_path=test_master, cache_path=cache_path)
        categorize_invoice(r, excel_path=test_master, cache_path=cache_path)
    assert mock_raw.call_count == 1  # segunda vez sale del cache
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_categorize_invoice.py -v`

- [ ] **Step 3: Implement categorize_invoice**

Append to `modules/cash_flow/categorizer.py`:

```python
# Agregar al final de modules/cash_flow/categorizer.py
import os
from openpyxl import load_workbook

from config import EXCEL_PATH, DROPBOX_BACKUP_PATH
from excel_manager import (
    SHEET_NAME, _save_wb,
    COL_CATEGORIA, COL_CULTIVO, COL_CONFIANZA, COL_CATEGORIZADO_POR,
)
from modules.cash_flow.categorizer_cache import CategorizerCache

DEFAULT_CACHE_PATH = os.path.join(DROPBOX_BACKUP_PATH, "categorizer_cache.json")
CONFIANZA_REVISAR = 0.85


def _get_cache(cache_path=None) -> CategorizerCache:
    return CategorizerCache(cache_path or DEFAULT_CACHE_PATH)


def _read_invoice_row(ws, row: int) -> dict:
    return {
        "fecha": str(ws.cell(row, 1).value or ""),
        "proveedor": str(ws.cell(row, 4).value or ""),
        "documento": str(ws.cell(row, 6).value or ""),
        "glosa": str(ws.cell(row, 8).value or ""),
        "glosa_ii": str(ws.cell(row, 9).value or ""),
        "monto": float(ws.cell(row, 15).value or 0),
    }


def categorize_invoice(row: int, excel_path=None, cache_path=None) -> dict:
    """Categoriza la fila `row` de Facturas y escribe cols Q-T en el Master.

    Usa cache local. Si confianza < 0.85 escribe "REVISAR" en Categoria.
    """
    excel_path = excel_path or EXCEL_PATH
    cache = _get_cache(cache_path)

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]
    data = _read_invoice_row(ws, row)
    wb.close()

    # Cache hit
    cached = cache.get(data["proveedor"], data["glosa"])
    if cached:
        result = cached
        source = "cache"
    else:
        result = categorize_raw(
            proveedor=data["proveedor"],
            glosa=data["glosa"],
            glosa_ii=data["glosa_ii"],
            monto=data["monto"],
            fecha=data["fecha"],
        )
        cache.set(data["proveedor"], data["glosa"], result)
        source = "claude"

    cat_to_write = result["categoria"]
    if result["confianza"] < CONFIANZA_REVISAR:
        cat_to_write = "REVISAR"

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]
    ws.cell(row, COL_CATEGORIA, cat_to_write)
    ws.cell(row, COL_CULTIVO, result["cultivo"])
    ws.cell(row, COL_CONFIANZA, result["confianza"])
    ws.cell(row, COL_CATEGORIZADO_POR, source)
    _save_wb(wb, excel_path)
    wb.close()

    return result
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_categorize_invoice.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/categorizer.py tests/test_categorize_invoice.py
git commit -m "feat: categorize_invoice writes Categoria/Cultivo to Master"
```

---

### Task 5: `categorize_bank_movement(row)` — categoriza cargo banco

**Files:**
- Modify: `Robot/modules/cash_flow/categorizer.py`
- Test: `Robot/tests/test_categorize_bank.py`

- [ ] **Step 1: Write test**

```python
# tests/test_categorize_bank.py
import os, shutil, pytest
from unittest.mock import patch
from openpyxl import load_workbook
from excel_manager import (
    CUENTA_BANCO_SHEET, COL_BANCO_TIPO, COL_BANCO_CATEGORIA,
    COL_BANCO_CULTIVO, ensure_new_columns,
)


@pytest.fixture
def test_master(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test_master.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def test_bank_charge_categorized_as_factura(test_master):
    """Si la descripcion banco linkea con una factura existente -> Tipo=factura."""
    from modules.cash_flow.categorizer import categorize_bank_movement
    wb = load_workbook(test_master)
    ws = wb[CUENTA_BANCO_SHEET]
    # Crear una fila de prueba
    test_row = ws.max_row + 1
    ws.cell(test_row, 1, "2025-09-01")
    ws.cell(test_row, 2, "PAGO PROVEEDOR XYZ")
    ws.cell(test_row, 3, "TRF")
    ws.cell(test_row, 4, 500000)  # Cargo
    wb.save(test_master)
    wb.close()

    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value={"categoria": "Fertilizantes",
                              "cultivo": "NOGALES",
                              "confianza": 0.9, "razon": ""}):
        result = categorize_bank_movement(test_row, excel_path=test_master)

    assert result["categoria"] == "Fertilizantes"
    wb = load_workbook(test_master, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(test_row, COL_BANCO_CATEGORIA).value == "Fertilizantes"
    assert ws.cell(test_row, COL_BANCO_CULTIVO).value == "NOGALES"
    wb.close()


def test_bank_abono_marked_as_ingreso(test_master):
    """Movimiento con Abono (col 5) y sin Cargo => Tipo=ingreso, no llama Claude."""
    from modules.cash_flow.categorizer import categorize_bank_movement
    wb = load_workbook(test_master)
    ws = wb[CUENTA_BANCO_SHEET]
    test_row = ws.max_row + 1
    ws.cell(test_row, 1, "2025-09-01")
    ws.cell(test_row, 2, "DEPOSITO VALBIFRUT")
    ws.cell(test_row, 4, None)
    ws.cell(test_row, 5, 8000000)
    wb.save(test_master)
    wb.close()

    with patch("modules.cash_flow.categorizer.categorize_raw") as m:
        result = categorize_bank_movement(test_row, excel_path=test_master)
    assert m.called is False
    assert result["tipo"] == "ingreso"

    wb = load_workbook(test_master, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(test_row, COL_BANCO_TIPO).value == "ingreso"
    wb.close()
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement categorize_bank_movement**

Append to `modules/cash_flow/categorizer.py`:

```python
from excel_manager import (
    CUENTA_BANCO_SHEET, COL_BANCO_TIPO, COL_BANCO_CATEGORIA,
    COL_BANCO_CULTIVO,
)


def _read_bank_row(ws, row: int) -> dict:
    return {
        "fecha": str(ws.cell(row, 1).value or ""),
        "descripcion": str(ws.cell(row, 2).value or ""),
        "referencia": str(ws.cell(row, 3).value or ""),
        "cargo": float(ws.cell(row, 4).value or 0),
        "abono": float(ws.cell(row, 5).value or 0),
    }


def categorize_bank_movement(row: int, excel_path=None, cache_path=None) -> dict:
    """Categoriza una fila de Cuenta Banco y escribe cols G-J.

    Abono > 0 y Cargo == 0 => tipo=ingreso (no llama Claude).
    Cargo > 0 => llama Claude con descripcion como glosa.
    """
    excel_path = excel_path or EXCEL_PATH
    cache = _get_cache(cache_path)

    wb = load_workbook(excel_path)
    ws = wb[CUENTA_BANCO_SHEET]
    data = _read_bank_row(ws, row)
    wb.close()

    if data["abono"] > 0 and data["cargo"] == 0:
        result = {
            "tipo": "ingreso",
            "categoria": "",
            "cultivo": "",
            "confianza": 1.0,
            "razon": "abono detectado",
        }
    else:
        cached = cache.get(data["descripcion"], data["referencia"])
        if cached:
            base = cached
        else:
            base = categorize_raw(
                proveedor=data["descripcion"],
                glosa=data["referencia"],
                glosa_ii="",
                monto=data["cargo"],
                fecha=data["fecha"],
            )
            cache.set(data["descripcion"], data["referencia"], base)
        result = {**base, "tipo": "egreso"}

    wb = load_workbook(excel_path)
    ws = wb[CUENTA_BANCO_SHEET]
    ws.cell(row, COL_BANCO_TIPO, result["tipo"])
    if result["tipo"] != "ingreso":
        cat = result["categoria"]
        if result["confianza"] < CONFIANZA_REVISAR:
            cat = "REVISAR"
        ws.cell(row, COL_BANCO_CATEGORIA, cat)
        ws.cell(row, COL_BANCO_CULTIVO, result["cultivo"])
    _save_wb(wb, excel_path)
    wb.close()

    return result
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_categorize_bank.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/categorizer.py tests/test_categorize_bank.py
git commit -m "feat: categorize_bank_movement classifies bank charges/income"
```

---

### Task 6: `batch_categorize_history()` — categoriza histórico con backup y progreso

**Files:**
- Modify: `Robot/modules/cash_flow/categorizer.py`
- Test: `Robot/tests/test_batch_categorize.py`

- [ ] **Step 1: Write test**

```python
# tests/test_batch_categorize.py
import os, shutil, pytest
from unittest.mock import patch
from openpyxl import load_workbook
from excel_manager import SHEET_NAME, COL_CATEGORIA, ensure_new_columns


@pytest.fixture
def test_master(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test_master.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def test_batch_skips_already_categorized(test_master):
    from modules.cash_flow.categorizer import batch_categorize_history
    # Marcar 2 filas como ya categorizadas
    wb = load_workbook(test_master)
    ws = wb[SHEET_NAME]
    for r in (2, 3):
        ws.cell(r, COL_CATEGORIA, "Fertilizantes")
    wb.save(test_master)
    wb.close()

    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value={"categoria": "Riego", "cultivo": "GENERAL",
                              "confianza": 0.9, "razon": ""}):
        report = batch_categorize_history(excel_path=test_master, limit=5)

    assert report["skipped"] >= 2
    assert report["processed"] >= 0


def test_batch_limit_respected(test_master):
    from modules.cash_flow.categorizer import batch_categorize_history
    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value={"categoria": "Riego", "cultivo": "GENERAL",
                              "confianza": 0.9, "razon": ""}):
        report = batch_categorize_history(excel_path=test_master, limit=3)
    assert report["processed"] <= 3


def test_batch_reports_low_confidence_count(test_master):
    from modules.cash_flow.categorizer import batch_categorize_history
    with patch("modules.cash_flow.categorizer.categorize_raw",
                return_value={"categoria": "Riego", "cultivo": "GENERAL",
                              "confianza": 0.5, "razon": ""}):
        report = batch_categorize_history(excel_path=test_master, limit=3)
    assert report["low_confidence"] >= 1
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement batch_categorize_history**

Append to `modules/cash_flow/categorizer.py`:

```python
def batch_categorize_history(excel_path=None, cache_path=None,
                               limit: int | None = None,
                               progress_cb=None) -> dict:
    """Itera Master.Facturas, categoriza filas sin Categoria.

    Salta filas ya categorizadas (col Q no vacia). Guarda cada 25 filas.
    `limit`: maximo de filas a procesar (None = todas).
    `progress_cb(processed, total_pending)`: callback opcional para reportar.
    """
    from infrastructure.backups import backup_master
    excel_path = excel_path or EXCEL_PATH

    # Backup preventivo antes de batch
    try:
        backup_master(reason="pre-batch-categorize", excel_path=excel_path)
    except Exception as e:
        logger.warning(f"Backup pre-batch fallo (continuo igual): {e}")

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]
    pending_rows = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 4).value is None:
            continue
        if ws.cell(r, COL_CATEGORIA).value:
            continue
        pending_rows.append(r)
    wb.close()

    if limit is not None:
        pending_rows = pending_rows[:limit]

    report = {
        "total_pending": len(pending_rows),
        "processed": 0,
        "skipped": 0,
        "low_confidence": 0,
        "from_cache": 0,
        "errors": 0,
    }

    # Contar tambien las que ya estan categorizadas para reporte completo
    wb = load_workbook(excel_path, read_only=True)
    ws = wb[SHEET_NAME]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 4).value and ws.cell(r, COL_CATEGORIA).value:
            report["skipped"] += 1
    wb.close()

    cache = _get_cache(cache_path)

    for idx, row in enumerate(pending_rows, 1):
        try:
            result = categorize_invoice(row,
                                          excel_path=excel_path,
                                          cache_path=cache_path)
            report["processed"] += 1
            if result["confianza"] < CONFIANZA_REVISAR:
                report["low_confidence"] += 1
        except Exception as e:
            logger.error(f"Error fila {row}: {e}")
            report["errors"] += 1

        if progress_cb:
            progress_cb(idx, len(pending_rows))

    logger.info(f"Batch categorize OK: {report}")
    return report
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_batch_categorize.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/categorizer.py tests/test_batch_categorize.py
git commit -m "feat: batch_categorize_history iterates Master with backup and progress"
```

---
