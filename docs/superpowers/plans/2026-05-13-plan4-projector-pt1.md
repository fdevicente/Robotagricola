# Plan 4: Motor de proyección — Parte 1/3

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans

**Goal:** Construir el projector que calcula proyección mes × categoría × cultivo y escribe la hoja Flujo Caja.

**Architecture:** Módulo `modules/cash_flow/projector.py`. Funciones puras de cálculo + loaders de Master + writer de Flujo Caja. Reusa iter_rows() optimizado de Plan 3.

**Tech Stack:** Python 3.11, openpyxl, dataclasses para tipos

---

### Task 1: Loader de egresos históricos

Lee Master.Facturas (categorizadas) y agrupa por mes × categoría × cultivo.

**Files:**
- Create: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_loaders.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_loaders.py
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
    # Estructura: {(year, month, categoria, cultivo): total}
    for k, v in egresos.items():
        assert len(k) == 4
        assert isinstance(v, (int, float))


def test_load_egresos_skips_uncategorized(m):
    """Filas sin Categoria o con REVISAR se ignoran."""
    from modules.cash_flow.projector import load_historical_egresos
    egresos = load_historical_egresos(excel_path=m)
    # REVISAR no debe aparecer como categoria
    cats = {k[2] for k in egresos.keys()}
    assert "REVISAR" not in cats


def test_load_egresos_by_year(m):
    from modules.cash_flow.projector import load_historical_egresos
    egresos = load_historical_egresos(excel_path=m, year=2025)
    years = {k[0] for k in egresos.keys()}
    assert years == {2025} or years == set()
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_projector_loaders.py -v`
Expected: ImportError

- [ ] **Step 3: Implement load_historical_egresos**

```python
# modules/cash_flow/projector.py
"""Motor de proyección de flujo de caja.

Calcula proyeccion mes x categoria x cultivo escalando un ano base
por el factor de hectareas + aplica ajustes manuales del usuario.
"""
from collections import defaultdict
from datetime import date, datetime
from openpyxl import load_workbook

from config import EXCEL_PATH
from excel_manager import (
    SHEET_NAME, CUENTA_BANCO_SHEET,
    COSECHAS_SHEET, HECTAREAS_SHEET, AJUSTES_SHEET, FLUJO_CAJA_SHEET,
    CATEGORIAS, CULTIVOS,
)


def _to_year_month(v) -> tuple[int, int] | None:
    """Acepta date, datetime o string YYYY-MM-DD."""
    if isinstance(v, datetime):
        return (v.year, v.month)
    if isinstance(v, date):
        return (v.year, v.month)
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                d = datetime.strptime(v[:10], fmt).date()
                return (d.year, d.month)
            except ValueError:
                pass
    return None


def load_historical_egresos(excel_path: str | None = None,
                              year: int | None = None) -> dict:
    """Agrupa Facturas por (year, month, categoria, cultivo) -> total.

    Salta filas sin Categoria o con Categoria=REVISAR.
    Si `year` esta seteado, filtra solo ese ano.
    """
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    agg: dict = defaultdict(float)
    # Columns (1-indexed): 1=fecha emision, 4=proveedor, 15=total,
    # 17=Categoria (Q), 18=Cultivo (R)
    for row in ws.iter_rows(min_row=2, max_col=18, values_only=True):
        proveedor = row[3]
        if not proveedor:
            continue
        categoria = row[16]
        if not categoria or categoria == "REVISAR":
            continue
        cultivo = row[17] or "GENERAL"
        total = row[14]
        if not total:
            continue
        ym = _to_year_month(row[0])
        if not ym:
            continue
        if year is not None and ym[0] != year:
            continue
        agg[(ym[0], ym[1], categoria, cultivo)] += float(total)
    wb.close()
    return dict(agg)
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_loaders.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_loaders.py
git commit -m "feat: load_historical_egresos aggregates Facturas by (year,mes,cat,cultivo)"
```

---

### Task 2: Loaders de hectáreas y ajustes manuales

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_hectareas.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_hectareas.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import (
    ensure_cash_flow_sheets, AJUSTES_SHEET,
)


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
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement (append a projector.py)**

```python
def load_hectareas(excel_path: str | None = None) -> dict:
    """Devuelve {year: {cultivo: hc}}. Cultivos en uppercase."""
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[HECTAREAS_SHEET]
    hc: dict = {}
    # Headers row 1: Año, Nogales, Cerezos, Avellanos, Notas
    for row in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        year = row[0]
        if not isinstance(year, int):
            continue
        hc[year] = {
            "NOGALES": float(row[1] or 0),
            "CEREZOS": float(row[2] or 0),
            "AVELLANOS": float(row[3] or 0),
        }
    wb.close()
    return hc


def load_ajustes_manuales(excel_path: str | None = None) -> list:
    """Devuelve lista de ajustes activos.

    Cada item: {mes_proyectado: (year, month), categoria, cultivo, monto, razon}
    """
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[AJUSTES_SHEET]
    ajustes = []
    # Cols: fecha_agregado, mes_proyectado, categoria, cultivo, monto, razon, activo
    for row in ws.iter_rows(min_row=2, max_col=7, values_only=True):
        if not row[1]:  # mes_proyectado
            continue
        activo = row[6]
        if activo is False:
            continue
        ym = _parse_mes_str(row[1])
        if not ym:
            continue
        try:
            monto = float(row[4] or 0)
        except (TypeError, ValueError):
            continue
        ajustes.append({
            "mes_proyectado": ym,
            "categoria": row[2],
            "cultivo": row[3] or "GENERAL",
            "monto": monto,
            "razon": row[5] or "",
        })
    wb.close()
    return ajustes


def _parse_mes_str(v) -> tuple[int, int] | None:
    """Parsea '2026-07' o date o datetime a (year, month)."""
    if isinstance(v, datetime):
        return (v.year, v.month)
    if isinstance(v, date):
        return (v.year, v.month)
    if isinstance(v, str):
        try:
            parts = v.split("-")
            return (int(parts[0]), int(parts[1]))
        except (IndexError, ValueError):
            return None
    return None
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_hectareas.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_hectareas.py
git commit -m "feat: load_hectareas and load_ajustes_manuales"
```

---

### Task 3: Loader de ingresos esperados (Cosechas)

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_ingresos.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_ingresos.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import ensure_cash_flow_sheets, COSECHAS_SHEET


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_cash_flow_sheets(str(dst))
    return str(dst)


def test_load_ingresos_uses_real_if_present(m):
    """Si Estado=recibido y Monto real recibido > 0, usa real, no estimado."""
    from modules.cash_flow.projector import load_expected_ingresos
    wb = load_workbook(m)
    ws = wb[COSECHAS_SHEET]
    ws.append([2026, "NOGALES", 240000, "Valbifrut", 140000, 1.8, 2, 1,
                "2026-06-15", 126000, "adelanto", "recibido",
                "2026-06-20", 120000000, "CLP", ""])
    wb.save(m); wb.close()

    ingresos = load_expected_ingresos(excel_path=m)
    assert len(ingresos) == 1
    assert ingresos[0]["monto_clp"] == 120000000


def test_load_ingresos_estimates_usd_when_pending(m):
    """Estado=esperado con USD: convierte a CLP usando usd_clp_estimado."""
    from modules.cash_flow.projector import load_expected_ingresos
    wb = load_workbook(m)
    ws = wb[COSECHAS_SHEET]
    ws.append([2026, "NOGALES", 240000, "Valbifrut", 140000, 1.8, 2, 1,
                "2026-06-15", 126000, "adelanto", "esperado",
                None, None, "", ""])
    wb.save(m); wb.close()

    ingresos = load_expected_ingresos(excel_path=m)
    # 126000 USD * 1000 CLP = 126M
    assert ingresos[0]["monto_clp"] == 126_000_000


def test_load_ingresos_empty_returns_empty(m):
    from modules.cash_flow.projector import load_expected_ingresos
    ingresos = load_expected_ingresos(excel_path=m)
    assert ingresos == []
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement (append a projector.py)**

```python
def load_expected_ingresos(excel_path: str | None = None) -> list:
    """Lee Cosechas, devuelve ingresos proyectados convertidos a CLP.

    Si Estado=recibido y Monto real recibido > 0, usa el real.
    Si esperado y Moneda recibida=CLP, usa Monto USD * usd_clp_estimado
    como aproximacion (porque la columna USD esta).
    """
    from config import CASH_FLOW_CONFIG
    usd_clp = CASH_FLOW_CONFIG.get("usd_clp_estimado", 1000)

    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[COSECHAS_SHEET]
    ingresos = []
    # Cols (1-indexed): A=year, B=cultivo, D=exportadora, I=fecha estimada,
    # J=monto USD estimado, K=tipo cuota, L=estado,
    # M=fecha real recibido, N=monto real recibido, O=moneda recibida
    for row in ws.iter_rows(min_row=2, max_col=16, values_only=True):
        if not row[0]:
            continue
        estado = row[11]
        if estado == "recibido":
            monto_real = row[13]
            fecha = row[12]
            moneda = (row[14] or "CLP").upper()
            try:
                m_val = float(monto_real or 0)
            except (TypeError, ValueError):
                m_val = 0
            if m_val <= 0:
                continue
            monto_clp = m_val if moneda == "CLP" else m_val * usd_clp
            ym = _to_year_month(fecha)
        else:
            monto_usd = row[9]
            try:
                m_val = float(monto_usd or 0)
            except (TypeError, ValueError):
                m_val = 0
            if m_val <= 0:
                continue
            monto_clp = m_val * usd_clp
            ym = _to_year_month(row[8])
        if not ym:
            continue
        ingresos.append({
            "year": ym[0], "month": ym[1],
            "cultivo": row[1] or "GENERAL",
            "exportadora": row[3] or "",
            "tipo_cuota": row[10] or "",
            "estado": estado or "esperado",
            "monto_clp": float(monto_clp),
        })
    wb.close()
    return ingresos
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_ingresos.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_ingresos.py
git commit -m "feat: load_expected_ingresos converts Cosechas to CLP per month"
```

---
