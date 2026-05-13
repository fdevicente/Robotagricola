# Plan 2: Categorización — Parte 3/3

### Task 7: Detector de patrones de ingreso (banco)

**Files:**
- Create: `Robot/modules/cash_flow/historical_importer.py`
- Test: `Robot/tests/test_income_detector.py`

Marca movimientos del banco con `Tipo` correcto: `venta_dolares`, `ingreso_clp`, `sueldo`, `honorario`, `factura` o `otro`. Esto permite que el projector separe ingresos de egresos sin llamar a Claude.

- [ ] **Step 1: Write test**

```python
# tests/test_income_detector.py
import os, shutil, pytest
from openpyxl import load_workbook
from excel_manager import (
    CUENTA_BANCO_SHEET, COL_BANCO_TIPO, ensure_new_columns,
)


@pytest.fixture
def test_master(tmp_path):
    src = os.path.join(os.path.dirname(__file__), "..", "..",
                        "MASTER Agricola Santa Elisa.xlsx")
    dst = tmp_path / "test_master.xlsx"
    shutil.copy2(src, dst)
    ensure_new_columns(str(dst))
    return str(dst)


def _add_bank_row(path, fecha, descripcion, cargo=None, abono=None):
    wb = load_workbook(path)
    ws = wb[CUENTA_BANCO_SHEET]
    r = ws.max_row + 1
    ws.cell(r, 1, fecha)
    ws.cell(r, 2, descripcion)
    ws.cell(r, 3, "REF")
    ws.cell(r, 4, cargo)
    ws.cell(r, 5, abono)
    wb.save(path)
    wb.close()
    return r


def test_detects_venta_dolares(test_master):
    from modules.cash_flow.historical_importer import detect_income_patterns
    r = _add_bank_row(test_master, "2025-09-01",
                       "VENTA DOLARES MISMA EMPRESA", abono=15000000)
    report = detect_income_patterns(excel_path=test_master)
    wb = load_workbook(test_master, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(r, COL_BANCO_TIPO).value == "venta_dolares"
    wb.close()
    assert report["venta_dolares"] >= 1


def test_detects_vitakai_as_ingreso_clp(test_master):
    from modules.cash_flow.historical_importer import detect_income_patterns
    r = _add_bank_row(test_master, "2025-09-02",
                       "DEPOSITO VITAKAI SPA", abono=5000000)
    detect_income_patterns(excel_path=test_master)
    wb = load_workbook(test_master, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(r, COL_BANCO_TIPO).value == "ingreso_clp"
    wb.close()


def test_detects_sueldo_egreso(test_master):
    from modules.cash_flow.historical_importer import detect_income_patterns
    r = _add_bank_row(test_master, "2025-09-03",
                       "PAGO REMUNERACIONES PERSONAL", cargo=8000000)
    detect_income_patterns(excel_path=test_master)
    wb = load_workbook(test_master, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(r, COL_BANCO_TIPO).value == "sueldo"
    wb.close()


def test_idempotent_skips_already_tagged(test_master):
    from modules.cash_flow.historical_importer import detect_income_patterns
    r = _add_bank_row(test_master, "2025-09-04",
                       "VENTA DOLARES", abono=1000000)
    detect_income_patterns(excel_path=test_master)
    # Cambiar manualmente y verificar que no lo sobrescribe
    wb = load_workbook(test_master)
    ws = wb[CUENTA_BANCO_SHEET]
    ws.cell(r, COL_BANCO_TIPO, "manual_override")
    wb.save(test_master)
    wb.close()
    detect_income_patterns(excel_path=test_master)
    wb = load_workbook(test_master, read_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    assert ws.cell(r, COL_BANCO_TIPO).value == "manual_override"
    wb.close()
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_income_detector.py -v`

- [ ] **Step 3: Implement historical_importer.py**

```python
# modules/cash_flow/historical_importer.py
"""Onboarding: detecta patrones en Cuenta Banco para clasificar Tipo
sin llamar a Claude (mas barato y deterministico).

Patrones detectados:
- "venta de dolares" / "venta dolares" -> Tipo=venta_dolares
- Exportadora conocida (Valbifrut, Pacific Nuts, Vitakai) + abono -> ingreso_clp
- Sueldo / remuneracion + cargo -> sueldo
- Honorario + cargo -> honorario
- Resto: deja vacio (lo categoriza Claude despues)
"""
import logging
import re
from openpyxl import load_workbook

from config import EXCEL_PATH
from excel_manager import (
    CUENTA_BANCO_SHEET, COL_BANCO_TIPO, _save_wb,
)

logger = logging.getLogger(__name__)


PATTERNS = [
    # (regex sobre descripcion lower, tipo, requiere_abono, requiere_cargo)
    (re.compile(r"venta.*dolares|venta.*dolar"), "venta_dolares", False, False),
    (re.compile(r"valbifrut|pacific\s*nuts|vitakai"), "ingreso_clp", True, False),
    (re.compile(r"remuneracion|sueldo|liquidacion|finiquito"), "sueldo", False, True),
    (re.compile(r"honorario|boleta\s+honorario"), "honorario", False, True),
]


def detect_income_patterns(excel_path: str | None = None) -> dict:
    """Clasifica Tipo en Cuenta Banco usando regex. No toca filas ya tagueadas."""
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path)
    ws = wb[CUENTA_BANCO_SHEET]

    counts = {
        "venta_dolares": 0,
        "ingreso_clp": 0,
        "sueldo": 0,
        "honorario": 0,
        "skipped_tagged": 0,
        "no_match": 0,
    }

    for r in range(2, ws.max_row + 1):
        descripcion = str(ws.cell(r, 2).value or "")
        if not descripcion:
            continue
        if ws.cell(r, COL_BANCO_TIPO).value:
            counts["skipped_tagged"] += 1
            continue

        cargo = float(ws.cell(r, 4).value or 0)
        abono = float(ws.cell(r, 5).value or 0)
        desc_low = descripcion.lower()

        matched = False
        for rx, tipo, req_abono, req_cargo in PATTERNS:
            if rx.search(desc_low):
                if req_abono and abono <= 0:
                    continue
                if req_cargo and cargo <= 0:
                    continue
                ws.cell(r, COL_BANCO_TIPO, tipo)
                counts[tipo] += 1
                matched = True
                break

        if not matched:
            counts["no_match"] += 1

    _save_wb(wb, excel_path)
    wb.close()
    logger.info(f"detect_income_patterns: {counts}")
    return counts
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_income_detector.py -v`
Expected: 4 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/historical_importer.py tests/test_income_detector.py
git commit -m "feat: detect_income_patterns tags bank rows by description regex"
```

---

### Task 8: Script de onboarding histórico (runner Fase 2)

**Files:**
- Create: `Robot/onboarding_cash_flow.py`

- [ ] **Step 1: Implement onboarding script**

```python
# onboarding_cash_flow.py
"""
Script de onboarding histórico — Fase 2 Cash Flow.
Ejecutar UNA VEZ después de setup_cash_flow.py.

Acciones:
1. Backup preventivo
2. Detectar patrones de ingreso en Cuenta Banco (venta dolares, exportadoras, sueldos)
3. Batch categorize Facturas con Claude (costo ~USD $5 para ~1300 facturas)
4. Backup post-import
5. Reportar # de filas a revisar manualmente (confianza < 0.85)

Uso:
  py -3.11 onboarding_cash_flow.py            # corre todo
  py -3.11 onboarding_cash_flow.py --limit 50 # solo 50 facturas (prueba)
  py -3.11 onboarding_cash_flow.py --skip-claude  # solo patrones banco
"""
import argparse
import logging
import sys

from infrastructure.backups import backup_master
from modules.cash_flow.historical_importer import detect_income_patterns
from modules.cash_flow.categorizer import batch_categorize_history
from config import EXCEL_PATH

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logger = logging.getLogger(__name__)


def _progress(done: int, total: int):
    if done % 25 == 0 or done == total:
        pct = 100 * done / total if total else 100
        logger.info(f"   ... {done}/{total} ({pct:.0f}%)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=None,
                         help="Maximo de facturas a categorizar (None=todas)")
    parser.add_argument("--skip-claude", action="store_true",
                         help="Solo correr patrones de banco, no llamar Claude")
    args = parser.parse_args()

    logger.info("=== Onboarding Cash Flow - Fase 2 ===")
    logger.info(f"Master: {EXCEL_PATH}")

    logger.info("1/4 Backup preventivo...")
    backup_master(reason="pre-fase-2-onboarding")

    logger.info("2/4 Detectando patrones de ingreso en banco...")
    bank_report = detect_income_patterns()
    logger.info(f"   Banco: {bank_report}")

    if args.skip_claude:
        logger.info("3/4 SKIP — flag --skip-claude")
        cat_report = {"processed": 0, "low_confidence": 0}
    else:
        logger.info(f"3/4 Categorizando facturas con Claude (limit={args.limit})...")
        logger.info("   Costo estimado: ~USD $5 para 1300 facturas")
        cat_report = batch_categorize_history(
            limit=args.limit, progress_cb=_progress,
        )
        logger.info(f"   Categorizacion: {cat_report}")

    logger.info("4/4 Backup post-onboarding...")
    backup_master(reason="post-fase-2-onboarding")

    logger.info("=== Onboarding completo ===")
    logger.info(f"Facturas categorizadas: {cat_report.get('processed', 0)}")
    logger.info(f"REVISAR (baja confianza): {cat_report.get('low_confidence', 0)}")
    logger.info(f"Banco - venta_dolares: {bank_report.get('venta_dolares', 0)}")
    logger.info(f"Banco - ingreso_clp:   {bank_report.get('ingreso_clp', 0)}")
    logger.info(f"Banco - sueldo:        {bank_report.get('sueldo', 0)}")
    logger.info(f"Banco - sin clasificar: {bank_report.get('no_match', 0)}")
    logger.info("")
    logger.info("Siguiente paso: revisar filas con Categoria=REVISAR en Master.Facturas")
    logger.info("Despues: ejecutar Plan 3 (banco + matching)")


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Smoke test sin Claude**

Run: `py -3.11 onboarding_cash_flow.py --skip-claude`
Expected: corre sin error, reporta # de filas banco tagueadas.

- [ ] **Step 3: Smoke test con Claude (3 facturas)**

Run: `py -3.11 onboarding_cash_flow.py --limit 3`
Expected: 3 facturas categorizadas en `cat_report["processed"]`.

- [ ] **Step 4: Verificar Master**

Run: `py -3.11 -c "from openpyxl import load_workbook; from excel_manager import SHEET_NAME, COL_CATEGORIA; wb=load_workbook(r'C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx', read_only=True); ws=wb[SHEET_NAME]; print('Categorizadas:', sum(1 for r in range(2, ws.max_row+1) if ws.cell(r, COL_CATEGORIA).value))"`
Expected: 3 categorizadas (después del smoke test).

- [ ] **Step 5: Commit**

```bash
git add onboarding_cash_flow.py
git commit -m "feat: add onboarding_cash_flow.py runner for Fase 2"
```

---

## Resumen Plan 2

| Task | Descripción | Archivos |
|---|---|---|
| 1 | Prompt builder + JSON parser | modules/cash_flow/prompt.py + tests |
| 2 | Cliente HTTP Claude | modules/cash_flow/categorizer.py + tests |
| 3 | Cache JSON local | modules/cash_flow/categorizer_cache.py + tests |
| 4 | categorize_invoice(row) | categorizer.py + tests |
| 5 | categorize_bank_movement(row) | categorizer.py + tests |
| 6 | batch_categorize_history() | categorizer.py + tests |
| 7 | detect_income_patterns() | modules/cash_flow/historical_importer.py + tests |
| 8 | onboarding_cash_flow.py | runner |

**Resultado:**
- Facturas históricas categorizadas (categoría × cultivo × confianza)
- Cargos banco categorizados
- Ingresos banco tagueados por patrón (venta_dolares, ingreso_clp, sueldo)
- Cache JSON para no re-llamar Claude
- Filas con confianza < 0.85 marcadas REVISAR
- Base lista para Plan 3 (matcher banco↔factura) y Plan 4 (projector)
