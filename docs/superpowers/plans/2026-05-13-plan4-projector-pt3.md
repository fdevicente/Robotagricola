# Plan 4: Motor de proyección — Parte 3/3

### Task 7: Escribir hoja Flujo Caja con la proyección

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_writer.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_writer.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import ensure_cash_flow_sheets, FLUJO_CAJA_SHEET


@pytest.fixture
def m(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test.xlsx"
    shutil.copy2(src, dst)
    ensure_cash_flow_sheets(str(dst))
    return str(dst)


def test_write_flujo_caja_creates_header(m):
    from modules.cash_flow.projector import write_flujo_caja
    saldo_data = {
        (2026, 5): {"saldo_inicio": 100, "ingresos": 200, "egresos": 50, "saldo_cierre": 250},
        (2026, 6): {"saldo_inicio": 250, "ingresos": 0, "egresos": 30, "saldo_cierre": 220},
    }
    egresos = {
        (2026, 5, "Fertilizantes", "NOGALES"): 30,
        (2026, 5, "Riego", "GENERAL"): 20,
        (2026, 6, "Combustible", "NOGALES"): 30,
    }
    ingresos = [
        {"year": 2026, "month": 5, "monto_clp": 200, "exportadora": "Valbifrut",
         "estado": "recibido"},
    ]
    write_flujo_caja(saldo_data, egresos, ingresos,
                       months=[(2026, 5), (2026, 6)], excel_path=m)

    wb = load_workbook(m, read_only=True)
    ws = wb[FLUJO_CAJA_SHEET]
    # Header row 1: SECCION, mes-may, mes-jun
    assert ws.cell(1, 1).value in ("SECCION", "Seccion", "Sección")
    assert ws.cell(1, 2).value is not None  # mayo
    assert ws.cell(1, 3).value is not None  # junio
    wb.close()


def test_write_flujo_caja_has_saldo_rows(m):
    from modules.cash_flow.projector import write_flujo_caja
    saldo_data = {
        (2026, 5): {"saldo_inicio": 100, "ingresos": 200, "egresos": 50, "saldo_cierre": 250},
    }
    write_flujo_caja(saldo_data, {}, [], months=[(2026, 5)], excel_path=m)

    wb = load_workbook(m, read_only=True)
    ws = wb[FLUJO_CAJA_SHEET]
    # Buscar fila "SALDO INICIAL" y "SALDO CIERRE"
    labels = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
    assert any("SALDO INICIAL" in str(l or "") for l in labels)
    assert any("SALDO CIERRE" in str(l or "") for l in labels)
    assert any("TOTAL EGRESOS" in str(l or "") for l in labels)
    wb.close()
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_projector_writer.py -v`

- [ ] **Step 3: Implement (append projector.py)**

```python
def write_flujo_caja(saldo_data: dict, egresos: dict, ingresos: list,
                       months: list, excel_path: str | None = None):
    """Regenera la hoja Flujo Caja con la proyeccion.

    Estructura: filas = secciones, columnas = meses.
    """
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path)
    if FLUJO_CAJA_SHEET in wb.sheetnames:
        del wb[FLUJO_CAJA_SHEET]
    ws = wb.create_sheet(FLUJO_CAJA_SHEET)

    # Header
    header = ["SECCION"] + [_month_label(y, m) for (y, m) in months]
    ws.append(header)

    # === SALDO INICIAL ===
    ws.append(["SALDO INICIAL"] + [saldo_data.get(ym, {}).get("saldo_inicio", 0)
                                    for ym in months])

    # === INGRESOS (por exportadora) ===
    ws.append(["INGRESOS"])
    exportadoras = sorted({i.get("exportadora", "") for i in ingresos})
    for exp in exportadoras:
        if not exp:
            continue
        row = [f"  {exp}"]
        for ym in months:
            total = sum(i["monto_clp"] for i in ingresos
                          if i.get("exportadora") == exp
                          and i["year"] == ym[0] and i["month"] == ym[1])
            row.append(total)
        ws.append(row)
    ws.append(["  TOTAL INGRESOS"] + [saldo_data.get(ym, {}).get("ingresos", 0)
                                       for ym in months])

    # === EGRESOS (por categoria) ===
    ws.append(["EGRESOS"])
    cats_set = sorted({k[2] for k in egresos.keys()})
    for cat in cats_set:
        row = [f"  {cat}"]
        for ym in months:
            total = sum(monto for (y, m, c, _cu), monto in egresos.items()
                          if y == ym[0] and m == ym[1] and c == cat)
            row.append(total)
        ws.append(row)
    ws.append(["  TOTAL EGRESOS"] + [saldo_data.get(ym, {}).get("egresos", 0)
                                      for ym in months])

    # === SALDO CIERRE ===
    ws.append(["SALDO CIERRE MES"] + [saldo_data.get(ym, {}).get("saldo_cierre", 0)
                                       for ym in months])

    # Anchos
    ws.column_dimensions["A"].width = 28
    for i in range(2, len(months) + 2):
        ws.column_dimensions[chr(64 + i)].width = 14

    _save_wb_local(wb, excel_path)
    wb.close()


_MESES_ES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
              "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _month_label(year: int, month: int) -> str:
    return f"{_MESES_ES[month]}-{str(year)[-2:]}"


def _save_wb_local(wb, excel_path):
    """Wrapper para _save_wb evitando import circular."""
    from excel_manager import _save_wb
    _save_wb(wb, excel_path)
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_writer.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_writer.py
git commit -m "feat: write_flujo_caja regenerates Flujo Caja sheet"
```

---

### Task 8: API `get_cash_flow()` + runner smoke test

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Create: `Robot/recalc_flujo_caja.py`

- [ ] **Step 1: Add API `get_cash_flow`**

Append a `projector.py`:

```python
def get_cash_flow(start: tuple, end: tuple,
                    saldo_inicial: float,
                    base_year: int = 2025,
                    excel_path: str | None = None) -> dict:
    """API principal del projector.

    Args:
        start: (year, month) inicio del periodo
        end: (year, month) fin del periodo (inclusive)
        saldo_inicial: saldo banco al inicio
        base_year: ano historico a usar como base (default 2025)

    Returns:
        {
          "months": [(y, m), ...],
          "saldo": {(y, m): {...}},
          "egresos": {(y, m, cat, cul): monto},
          "ingresos": [...],
        }
    """
    historicos = load_historical_egresos(excel_path, year=base_year)
    hc = load_hectareas(excel_path)
    ajustes = load_ajustes_manuales(excel_path)
    ingresos = load_expected_ingresos(excel_path)

    # Iterar meses del periodo
    months = []
    y, mo = start
    while (y, mo) <= end:
        months.append((y, mo))
        mo += 1
        if mo > 12:
            mo = 1
            y += 1

    # Proyectar egresos para cada year en el periodo
    egresos_proj = {}
    years_target = {ym[0] for ym in months}
    for ty in years_target:
        e = compute_egresos_proyectados(
            historicos=historicos, ajustes=ajustes, hc=hc,
            base_year=base_year, target_year=ty,
        )
        egresos_proj.update(e)

    # Saldo running
    saldo = compute_saldo_mensual(saldo_inicial, ingresos, egresos_proj, months)

    return {
        "months": months, "saldo": saldo,
        "egresos": egresos_proj, "ingresos": ingresos,
    }
```

- [ ] **Step 2: Create runner**

```python
# recalc_flujo_caja.py
"""Recalcula la hoja Flujo Caja con la proyeccion actual.

Uso: py -3.11 recalc_flujo_caja.py
     py -3.11 recalc_flujo_caja.py --saldo 130600000
"""
import argparse
import logging

from infrastructure.backups import backup_master
from modules.cash_flow.projector import get_cash_flow, write_flujo_caja

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logger = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--saldo", type=float, default=130_600_000,
                         help="Saldo banco actual CLP (default: 130.6M)")
    parser.add_argument("--start", default="2026-05",
                         help="Mes inicial YYYY-MM")
    parser.add_argument("--end", default="2027-04",
                         help="Mes final YYYY-MM (inclusive)")
    parser.add_argument("--base-year", type=int, default=2025,
                         help="Ano historico base para escalamiento")
    args = parser.parse_args()

    sy, sm = map(int, args.start.split("-"))
    ey, em = map(int, args.end.split("-"))

    logger.info("Backup pre-recalc...")
    backup_master(reason="pre-recalc-flujo")

    logger.info(f"Calculando proyeccion {args.start} -> {args.end}...")
    result = get_cash_flow(
        start=(sy, sm), end=(ey, em),
        saldo_inicial=args.saldo, base_year=args.base_year,
    )

    logger.info(f"Escribiendo {len(result['months'])} meses en Flujo Caja...")
    write_flujo_caja(
        saldo_data=result["saldo"], egresos=result["egresos"],
        ingresos=result["ingresos"], months=result["months"],
    )

    logger.info("Backup post-recalc...")
    backup_master(reason="post-recalc-flujo")

    # Resumen
    total_ing = sum(s["ingresos"] for s in result["saldo"].values())
    total_eg = sum(s["egresos"] for s in result["saldo"].values())
    saldo_final = list(result["saldo"].values())[-1]["saldo_cierre"]
    logger.info("=== Proyeccion ===")
    logger.info(f"Ingresos totales: ${total_ing:>15,.0f}")
    logger.info(f"Egresos totales:  ${total_eg:>15,.0f}")
    logger.info(f"Saldo final:      ${saldo_final:>15,.0f}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Smoke test con datos reales**

Run: `py -3.11 recalc_flujo_caja.py`
Expected: corre OK, escribe Flujo Caja, muestra resumen con ingresos/egresos/saldo final.

- [ ] **Step 4: Verificar Flujo Caja**

Abrir Master en Excel → hoja Flujo Caja → debe tener 12 columnas (may-26 a abr-27),
filas SALDO INICIAL / INGRESOS / EGRESOS por categoria / SALDO CIERRE.

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py recalc_flujo_caja.py
git commit -m "feat: get_cash_flow API + recalc_flujo_caja runner"
```

---

## Resumen Plan 4

| Task | Descripcion | Archivos |
|---|---|---|
| 1 | load_historical_egresos | projector.py + tests |
| 2 | load_hectareas + load_ajustes | projector.py + tests |
| 3 | load_expected_ingresos | projector.py + tests |
| 4 | compute_factor_hc | projector.py + tests |
| 5 | compute_egresos_proyectados | projector.py + tests |
| 6 | compute_saldo_mensual | projector.py + tests |
| 7 | write_flujo_caja | projector.py + tests |
| 8 | get_cash_flow + runner | projector.py + recalc_flujo_caja.py |

**Resultado:**
- Master.Flujo Caja con proyeccion 12 meses
- Egresos historicos 2025 escalados por hc 2026
- Ingresos esperados (Cosechas) sumados por mes
- Ajustes manuales aplicados encima
- Saldo running balance mes a mes
