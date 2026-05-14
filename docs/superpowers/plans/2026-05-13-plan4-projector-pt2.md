# Plan 4: Motor de proyección — Parte 2/3

### Task 4: Factor de escalamiento por hectáreas

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_escalamiento.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_escalamiento.py
from modules.cash_flow.projector import compute_factor_hc


HC = {
    2024: {"NOGALES": 65, "CEREZOS": 1.8, "AVELLANOS": 0},
    2025: {"NOGALES": 54, "CEREZOS": 3.8, "AVELLANOS": 11.5},
    2026: {"NOGALES": 43, "CEREZOS": 3.8, "AVELLANOS": 26.5},
}


def test_factor_same_year_is_one():
    assert compute_factor_hc(HC, "NOGALES", base_year=2025, target_year=2025) == 1.0


def test_factor_nogales_2025_to_2026_smaller():
    f = compute_factor_hc(HC, "NOGALES", base_year=2025, target_year=2026)
    assert abs(f - (43 / 54)) < 0.001


def test_factor_avellanos_2025_to_2026_growth():
    f = compute_factor_hc(HC, "AVELLANOS", base_year=2025, target_year=2026)
    assert abs(f - (26.5 / 11.5)) < 0.001


def test_factor_general_uses_total():
    """Para cultivo=GENERAL usa suma total."""
    f = compute_factor_hc(HC, "GENERAL", base_year=2025, target_year=2026)
    total_2025 = 54 + 3.8 + 11.5
    total_2026 = 43 + 3.8 + 26.5
    assert abs(f - (total_2026 / total_2025)) < 0.001


def test_factor_base_zero_returns_one():
    """Si el base year tiene 0 hectareas, no escalar (evitar div/0)."""
    hc = {2024: {"AVELLANOS": 0}, 2025: {"AVELLANOS": 10}}
    assert compute_factor_hc(hc, "AVELLANOS", base_year=2024, target_year=2025) == 1.0
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_projector_escalamiento.py -v`

- [ ] **Step 3: Implement (append projector.py)**

```python
def compute_factor_hc(hc: dict, cultivo: str, base_year: int, target_year: int) -> float:
    """Factor de escalamiento por hectareas.

    Para cultivo=GENERAL usa la suma de todos los cultivos.
    Si base=0, devuelve 1.0 (no escalar).
    """
    if base_year == target_year:
        return 1.0
    if base_year not in hc or target_year not in hc:
        return 1.0

    if cultivo.upper() == "GENERAL":
        base = sum(hc[base_year].values())
        target = sum(hc[target_year].values())
    else:
        base = hc[base_year].get(cultivo.upper(), 0)
        target = hc[target_year].get(cultivo.upper(), 0)

    if base <= 0:
        return 1.0
    return target / base
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_escalamiento.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_escalamiento.py
git commit -m "feat: compute_factor_hc scales by hectareas year over year"
```

---

### Task 5: Egresos proyectados (base × factor + ajustes)

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_egresos_proyectados.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_egresos_proyectados.py
from modules.cash_flow.projector import compute_egresos_proyectados


def test_egresos_scaled_by_hc():
    historicos = {(2025, 5, "Fertilizantes", "NOGALES"): 10_000_000}
    hc = {2025: {"NOGALES": 54}, 2026: {"NOGALES": 43}}
    proj = compute_egresos_proyectados(
        historicos=historicos, ajustes=[], hc=hc,
        base_year=2025, target_year=2026,
    )
    # 10M × 43/54 ≈ 7.96M
    val = proj[(2026, 5, "Fertilizantes", "NOGALES")]
    assert abs(val - 10_000_000 * 43 / 54) < 100


def test_ajustes_added_to_projection():
    """Ajuste manual se suma al gasto proyectado."""
    historicos = {}
    ajustes = [{
        "mes_proyectado": (2026, 7), "categoria": "Riego",
        "cultivo": "GENERAL", "monto": 5_000_000, "razon": "Bomba",
    }]
    hc = {2025: {"NOGALES": 54}, 2026: {"NOGALES": 43}}
    proj = compute_egresos_proyectados(
        historicos=historicos, ajustes=ajustes, hc=hc,
        base_year=2025, target_year=2026,
    )
    assert proj[(2026, 7, "Riego", "GENERAL")] == 5_000_000


def test_ajustes_can_be_negative():
    """Ajuste negativo reduce el proyectado."""
    historicos = {(2025, 5, "Fertilizantes", "NOGALES"): 10_000_000}
    ajustes = [{
        "mes_proyectado": (2026, 5), "categoria": "Fertilizantes",
        "cultivo": "NOGALES", "monto": -2_000_000, "razon": "Sin compra",
    }]
    hc = {2025: {"NOGALES": 54}, 2026: {"NOGALES": 54}}  # factor=1
    proj = compute_egresos_proyectados(
        historicos=historicos, ajustes=ajustes, hc=hc,
        base_year=2025, target_year=2026,
    )
    assert proj[(2026, 5, "Fertilizantes", "NOGALES")] == 8_000_000
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement (append projector.py)**

```python
def compute_egresos_proyectados(historicos: dict, ajustes: list,
                                  hc: dict, base_year: int,
                                  target_year: int) -> dict:
    """Proyecta egresos del target_year escalando base_year + sumando ajustes.

    Output: {(target_year, month, categoria, cultivo): monto_proyectado}
    """
    from collections import defaultdict
    proj: dict = defaultdict(float)

    # 1. Escalar historicos del base_year
    for (y, m, cat, cul), monto in historicos.items():
        if y != base_year:
            continue
        factor = compute_factor_hc(hc, cul, base_year, target_year)
        proj[(target_year, m, cat, cul)] += monto * factor

    # 2. Sumar ajustes manuales del target_year
    for a in ajustes:
        ym = a["mes_proyectado"]
        if ym[0] != target_year:
            continue
        key = (target_year, ym[1], a["categoria"], a["cultivo"])
        proj[key] += a["monto"]

    return dict(proj)
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_egresos_proyectados.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_egresos_proyectados.py
git commit -m "feat: compute_egresos_proyectados scales base_year by hc + adds ajustes"
```

---

### Task 6: Saldo proyectado (running balance mes a mes)

**Files:**
- Modify: `Robot/modules/cash_flow/projector.py`
- Test: `Robot/tests/test_projector_saldo.py`

- [ ] **Step 1: Write test**

```python
# tests/test_projector_saldo.py
from modules.cash_flow.projector import compute_saldo_mensual


def test_saldo_running_balance():
    egresos = {
        (2026, 5, "Fertilizantes", "NOGALES"): 5_000_000,
        (2026, 5, "Riego", "GENERAL"): 1_000_000,
        (2026, 6, "Combustible", "NOGALES"): 2_000_000,
    }
    ingresos = [
        {"year": 2026, "month": 5, "monto_clp": 200_000_000, "estado": "recibido"},
        {"year": 2026, "month": 6, "monto_clp": 0, "estado": "esperado"},
    ]
    result = compute_saldo_mensual(
        saldo_inicial=100_000_000,
        ingresos=ingresos, egresos=egresos,
        months=[(2026, 5), (2026, 6)],
    )
    # Mayo: 100M + 200M - 5M - 1M = 294M
    assert result[(2026, 5)]["saldo_cierre"] == 294_000_000
    assert result[(2026, 5)]["ingresos"] == 200_000_000
    assert result[(2026, 5)]["egresos"] == 6_000_000
    # Junio: 294M + 0 - 2M = 292M
    assert result[(2026, 6)]["saldo_cierre"] == 292_000_000


def test_saldo_zero_ingresos():
    egresos = {(2026, 5, "Riego", "GENERAL"): 10_000_000}
    result = compute_saldo_mensual(
        saldo_inicial=50_000_000, ingresos=[], egresos=egresos,
        months=[(2026, 5)],
    )
    assert result[(2026, 5)]["saldo_cierre"] == 40_000_000


def test_saldo_negativo_se_propaga():
    """Saldo puede ir negativo y propagarse."""
    egresos = {
        (2026, 5, "X", "GENERAL"): 100_000_000,
        (2026, 6, "Y", "GENERAL"): 50_000_000,
    }
    result = compute_saldo_mensual(
        saldo_inicial=10_000_000, ingresos=[], egresos=egresos,
        months=[(2026, 5), (2026, 6)],
    )
    assert result[(2026, 5)]["saldo_cierre"] == -90_000_000
    assert result[(2026, 6)]["saldo_cierre"] == -140_000_000
```

- [ ] **Step 2: Run test — expect FAIL**

- [ ] **Step 3: Implement (append projector.py)**

```python
def compute_saldo_mensual(saldo_inicial: float, ingresos: list,
                            egresos: dict, months: list) -> dict:
    """Running balance mes a mes.

    Args:
        saldo_inicial: saldo banco al inicio del primer mes
        ingresos: lista de {year, month, monto_clp}
        egresos: dict {(year, month, cat, cultivo): monto}
        months: lista de (year, month) en orden cronologico

    Returns:
        {(year, month): {saldo_inicio, ingresos, egresos, saldo_cierre}}
    """
    from collections import defaultdict
    # Agregar ingresos por mes
    ing_mes: dict = defaultdict(float)
    for i in ingresos:
        ing_mes[(i["year"], i["month"])] += i["monto_clp"]

    # Agregar egresos por mes
    eg_mes: dict = defaultdict(float)
    for (y, m, _cat, _cul), monto in egresos.items():
        eg_mes[(y, m)] += monto

    result = {}
    saldo = saldo_inicial
    for ym in months:
        ing = ing_mes.get(ym, 0)
        eg = eg_mes.get(ym, 0)
        saldo_cierre = saldo + ing - eg
        result[ym] = {
            "saldo_inicio": saldo,
            "ingresos": ing,
            "egresos": eg,
            "saldo_cierre": saldo_cierre,
        }
        saldo = saldo_cierre
    return result
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_projector_saldo.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/projector.py tests/test_projector_saldo.py
git commit -m "feat: compute_saldo_mensual running balance"
```

---
