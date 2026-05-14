# Plan 3: Banco + Matching — Parte 2/3

### Task 4: Lectores de facturas pendientes y movimientos nuevos

**Files:**
- Modify: `Robot/excel_manager.py` (agregar helpers)
- Test: `Robot/tests/test_matching_readers.py`

- [ ] **Step 1: Write test**

```python
# tests/test_matching_readers.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import (
    SHEET_NAME, CUENTA_BANCO_SHEET, ensure_new_columns,
    COL_BANCO_TIPO, COL_BANCO_FACTURA_LINK,
    read_facturas_pendientes, read_bank_movements_unlinked,
)


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def test_pendientes_excludes_paid(m):
    wb = load_workbook(m)
    ws = wb[SHEET_NAME]
    # Marcar fila 2 con Fecha Pago
    ws.cell(2, 3, "2025-09-01 (Banco)")
    wb.save(m); wb.close()
    pendientes = read_facturas_pendientes(excel_path=m)
    assert all(p["fila"] != 2 for p in pendientes)


def test_pendientes_has_required_fields(m):
    pendientes = read_facturas_pendientes(excel_path=m)
    if pendientes:
        p = pendientes[0]
        assert "fila" in p and "total" in p and "proveedor" in p
        assert "fecha_emision" in p and "nro_factura" in p


def test_bank_unlinked_excludes_linked(m):
    wb = load_workbook(m)
    ws = wb[CUENTA_BANCO_SHEET]
    r = ws.max_row + 1
    ws.cell(r, 1, "2025-09-01"); ws.cell(r, 2, "X"); ws.cell(r, 4, 1000)
    ws.cell(r, COL_BANCO_FACTURA_LINK, "FAC-123")  # ya linkeado
    wb.save(m); wb.close()
    movs = read_bank_movements_unlinked(excel_path=m)
    assert all(mv["fila"] != r for mv in movs)


def test_bank_unlinked_only_cargos(m):
    wb = load_workbook(m)
    ws = wb[CUENTA_BANCO_SHEET]
    r = ws.max_row + 1
    ws.cell(r, 1, "2025-09-01"); ws.cell(r, 2, "DEP"); ws.cell(r, 5, 5000)  # abono solo
    wb.save(m); wb.close()
    movs = read_bank_movements_unlinked(excel_path=m)
    assert all(mv["fila"] != r for mv in movs)
```

- [ ] **Step 2: Run — expect FAIL (ImportError)**

- [ ] **Step 3: Implement (append a excel_manager.py)**

```python
def read_facturas_pendientes(excel_path=None) -> list[dict]:
    """Lee Master.Facturas, devuelve facturas con Fecha Pago vacia."""
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    pendientes = []
    for r in range(2, ws.max_row + 1):
        proveedor = ws.cell(r, 4).value
        if not proveedor:
            continue
        if ws.cell(r, 3).value:  # Fecha Pago
            continue
        pendientes.append({
            "fila": r,
            "fecha_emision": ws.cell(r, 1).value,
            "fecha_vencimiento": ws.cell(r, 2).value,
            "proveedor": proveedor,
            "rut": ws.cell(r, 5).value,
            "nro_factura": ws.cell(r, 7).value,
            "glosa": ws.cell(r, 8).value,
            "total": ws.cell(r, 15).value,
        })
    wb.close()
    return pendientes


def read_bank_movements_unlinked(excel_path=None) -> list[dict]:
    """Lee Cuenta Banco, devuelve cargos sin Factura_linkeada."""
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    movs = []
    for r in range(2, ws.max_row + 1):
        fecha = ws.cell(r, 1).value
        if not fecha:
            continue
        cargo = float(ws.cell(r, 4).value or 0)
        if cargo <= 0:
            continue  # solo cargos
        if ws.cell(r, COL_BANCO_FACTURA_LINK).value:
            continue  # ya linkeado
        movs.append({
            "fila": r,
            "fecha": fecha,
            "descripcion": ws.cell(r, 2).value or "",
            "referencia": ws.cell(r, 3).value or "",
            "cargo": cargo,
            "abono": float(ws.cell(r, 5).value or 0),
        })
    wb.close()
    return movs
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_matching_readers.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
git add excel_manager.py tests/test_matching_readers.py
git commit -m "feat: readers for facturas pendientes and unlinked bank movs"
```

---

### Task 5: Linking — escribe Fecha Pago + Factura_linkeada

**Files:**
- Modify: `Robot/excel_manager.py`
- Test: `Robot/tests/test_matching_link.py`

- [ ] **Step 1: Write test**

```python
# tests/test_matching_link.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import (
    SHEET_NAME, CUENTA_BANCO_SHEET, ensure_new_columns,
    COL_BANCO_TIPO, COL_BANCO_FACTURA_LINK,
    apply_bank_factura_link,
)


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    # Agregar un cargo de prueba
    wb = load_workbook(str(dst))
    ws = wb[CUENTA_BANCO_SHEET]
    r = ws.max_row + 1
    ws.cell(r, 1, "2025-09-05"); ws.cell(r, 2, "PAGO X"); ws.cell(r, 4, 999000)
    wb.save(str(dst)); wb.close()
    return str(dst), r


def test_link_writes_both_sides(m):
    path, bank_row = m
    apply_bank_factura_link(
        bank_row=bank_row, factura_row=2,
        nro_factura="FAC-555", fecha_pago="2025-09-05",
        excel_path=path,
    )
    wb = load_workbook(path, read_only=True)
    ws_f = wb[SHEET_NAME]
    ws_b = wb[CUENTA_BANCO_SHEET]
    # Factura: Fecha Pago col 3
    assert "2025-09-05" in str(ws_f.cell(2, 3).value)
    # Banco: Tipo=factura, Factura_linkeada=nro
    assert ws_b.cell(bank_row, COL_BANCO_TIPO).value == "factura"
    assert ws_b.cell(bank_row, COL_BANCO_FACTURA_LINK).value == "FAC-555"
    wb.close()


def test_link_preserves_existing_fecha_pago(m):
    """Si la factura ya tiene Fecha Pago, no la sobrescribe."""
    path, bank_row = m
    wb = load_workbook(path)
    ws = wb[SHEET_NAME]
    ws.cell(2, 3, "2025-08-01 (Manual)")
    wb.save(path); wb.close()

    result = apply_bank_factura_link(
        bank_row=bank_row, factura_row=2,
        nro_factura="FAC-555", fecha_pago="2025-09-05",
        excel_path=path,
    )
    assert result["fecha_pago_skipped"] is True
    wb = load_workbook(path, read_only=True)
    ws = wb[SHEET_NAME]
    assert "Manual" in str(ws.cell(2, 3).value)
    wb.close()
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append a excel_manager.py)**

```python
def apply_bank_factura_link(bank_row: int, factura_row: int,
                              nro_factura: str, fecha_pago: str,
                              excel_path=None) -> dict:
    """Escribe Fecha Pago en Facturas y Tipo+Factura_linkeada en Cuenta Banco.

    Si Fecha Pago ya tiene valor, NO sobrescribe (devuelve fecha_pago_skipped=True).
    Idempotente para el lado banco: vuelve a setear los mismos valores sin problema.
    """
    excel_path = excel_path or EXCEL_PATH
    result = {"fecha_pago_skipped": False, "linked": False}
    wb = load_workbook(excel_path)

    # Lado factura
    ws_f = wb[SHEET_NAME]
    if ws_f.cell(factura_row, 3).value:
        result["fecha_pago_skipped"] = True
    else:
        ws_f.cell(factura_row, 3, f"{fecha_pago} (Banco)")

    # Lado banco
    ws_b = wb[CUENTA_BANCO_SHEET]
    ws_b.cell(bank_row, COL_BANCO_TIPO, "factura")
    ws_b.cell(bank_row, COL_BANCO_FACTURA_LINK, str(nro_factura))
    result["linked"] = True

    _save_wb(wb, excel_path)
    wb.close()
    return result
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_matching_link.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add excel_manager.py tests/test_matching_link.py
git commit -m "feat: apply_bank_factura_link writes Fecha Pago + Factura_linkeada"
```

---

### Task 6: Orquestador `match_new_bank_movements()`

**Files:**
- Modify: `Robot/modules/cash_flow/matcher.py`
- Test: `Robot/tests/test_match_orchestrator.py`

- [ ] **Step 1: Write test**

```python
# tests/test_match_orchestrator.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import (
    SHEET_NAME, CUENTA_BANCO_SHEET, ensure_new_columns,
    COL_BANCO_TIPO, COL_BANCO_FACTURA_LINK,
)


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def test_orchestrator_auto_matches_clear_case(m):
    from modules.cash_flow.matcher import match_new_bank_movements
    # Construir caso: leer una factura real, agregar cargo igual al total
    wb = load_workbook(m)
    ws_f = wb[SHEET_NAME]
    # Buscar primera factura sin Fecha Pago con total > 0
    target = None
    for r in range(2, min(ws_f.max_row + 1, 50)):
        if ws_f.cell(r, 3).value:
            continue
        total = ws_f.cell(r, 15).value
        if total and float(total) > 0:
            target = r
            break
    assert target is not None
    prov = str(ws_f.cell(target, 4).value)
    total = float(ws_f.cell(target, 15).value)
    fecha = ws_f.cell(target, 1).value

    ws_b = wb[CUENTA_BANCO_SHEET]
    br = ws_b.max_row + 1
    ws_b.cell(br, 1, fecha); ws_b.cell(br, 2, f"PAGO {prov}")
    ws_b.cell(br, 3, ""); ws_b.cell(br, 4, total)
    wb.save(m); wb.close()

    report = match_new_bank_movements(excel_path=m)
    assert report["auto_matched"] >= 1


def test_orchestrator_reports_categories(m):
    from modules.cash_flow.matcher import match_new_bank_movements
    report = match_new_bank_movements(excel_path=m)
    for k in ("scanned", "auto_matched", "ambiguous", "no_match"):
        assert k in report
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append matcher.py)**

```python
def match_new_bank_movements(excel_path=None) -> dict:
    """Lee movimientos unlinked, busca matches, aplica auto-matches.

    Devuelve: {scanned, auto_matched, ambiguous, no_match, errors}
    Los ambiguos quedan sin linkear (resolveran por Telegram en Plan 5).
    """
    from excel_manager import (
        read_facturas_pendientes, read_bank_movements_unlinked,
        apply_bank_factura_link,
    )
    movs = read_bank_movements_unlinked(excel_path)
    pendientes = read_facturas_pendientes(excel_path)

    report = {
        "scanned": len(movs), "auto_matched": 0,
        "ambiguous": 0, "no_match": 0, "errors": 0,
    }

    for mov in movs:
        try:
            candidates = find_matches(mov, pendientes)
            decision = classify_match(candidates)
            if decision["status"] == "auto":
                # Encontrar nro_factura
                fila = decision["fila"]
                factura = next((f for f in pendientes if f["fila"] == fila), None)
                if not factura:
                    report["errors"] += 1
                    continue
                fecha = mov["fecha"]
                fecha_str = (fecha.strftime("%Y-%m-%d")
                              if hasattr(fecha, "strftime") else str(fecha)[:10])
                apply_bank_factura_link(
                    bank_row=mov["fila"], factura_row=fila,
                    nro_factura=str(factura.get("nro_factura") or ""),
                    fecha_pago=fecha_str, excel_path=excel_path,
                )
                # Sacar de pendientes para no doble-matchear
                pendientes = [f for f in pendientes if f["fila"] != fila]
                report["auto_matched"] += 1
            elif decision["status"] == "ambiguous" or decision["status"] == "ambiguo":
                report["ambiguous"] += 1
            else:
                report["no_match"] += 1
        except Exception:
            report["errors"] += 1

    return report
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_match_orchestrator.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/matcher.py tests/test_match_orchestrator.py
git commit -m "feat: match_new_bank_movements orchestrator"
```

---
