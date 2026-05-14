# Plan 3: Banco + Matching — Parte 1/3

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans

**Goal:** Conectar movimientos bancarios con facturas pendientes y escribir Fecha Pago bidireccional.

**Architecture:** Módulo `modules/cash_flow/matcher.py` puro (sin I/O). Linking se hace en `excel_manager`. Runner `daily_banco_18h.py` orquesta scraper → matcher → categorizer → backup.

**Tech Stack:** Python 3.11, openpyxl, scotiabank_scraper existente

---

### Task 1: Scoring puro (factura ↔ bank movement)

**Files:**
- Create: `Robot/modules/cash_flow/matcher.py`
- Test: `Robot/tests/test_matcher_score.py`

- [ ] **Step 1: Write test**

```python
# tests/test_matcher_score.py
from datetime import date
from modules.cash_flow.matcher import match_score


def _factura(total=1000000, fecha=date(2025, 9, 1), proveedor="COPEVAL", nro="123"):
    return {"fila": 10, "total": total, "fecha_emision": fecha,
            "proveedor": proveedor, "nro_factura": nro}


def _mov(cargo=1000000, fecha=date(2025, 9, 5), descripcion="PAGO COPEVAL", ref=""):
    return {"fila": 5, "cargo": cargo, "abono": 0,
            "fecha": fecha, "descripcion": descripcion, "referencia": ref}


def test_score_exact_amount_and_close_date():
    s = match_score(_factura(), _mov())
    assert s >= 150


def test_score_amount_diff_lowers_score():
    s = match_score(_factura(), _mov(cargo=900000))
    assert s < 100


def test_score_provider_in_description_bonus():
    base = match_score(_factura(proveedor="ZZZZ"), _mov(descripcion="PAGO ZZZZ"))
    miss = match_score(_factura(proveedor="ZZZZ"), _mov(descripcion="PAGO XYZ"))
    assert base > miss


def test_score_nro_factura_in_reference_strong_match():
    s = match_score(_factura(nro="555"), _mov(cargo=1234567, ref="N FAC 555"))
    assert s >= 100


def test_score_abono_only_no_match():
    f = _factura()
    m = {"fila": 5, "cargo": 0, "abono": 1000000,
         "fecha": date(2025, 9, 5), "descripcion": "X", "referencia": ""}
    assert match_score(f, m) == 0
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_matcher_score.py -v`

- [ ] **Step 3: Implement**

```python
# modules/cash_flow/matcher.py
"""Match facturas pendientes con cargos bancarios."""
from datetime import date, datetime, timedelta


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(v[:10], fmt).date()
            except ValueError:
                pass
    return None


def _provider_match(provider: str, description: str) -> float:
    p = (provider or "").lower()
    d = (description or "").lower()
    words = [w for w in p.split() if len(w) > 3]
    if not words:
        return 0.0
    hits = sum(1 for w in words if w in d)
    return hits / len(words)


def match_score(factura: dict, bank_mov: dict) -> float:
    """Devuelve score 0-200. >=100 candidato fuerte. >=30 ambiguo."""
    cargo = float(bank_mov.get("cargo") or 0)
    if cargo <= 0:
        return 0.0

    total = float(factura.get("total") or 0)
    if total <= 0:
        return 0.0

    score = 0.0
    diff_pct = abs(cargo - total) / total
    if diff_pct < 0.001:
        score += 100
    elif diff_pct < 0.05:
        score += 30

    f_factura = _to_date(factura.get("fecha_emision"))
    f_banco = _to_date(bank_mov.get("fecha"))
    if f_factura and f_banco:
        diff_dias = abs((f_banco - f_factura).days)
        if diff_dias <= 5:
            score += 50
        elif diff_dias <= 15:
            score += 20

    score += 40 * _provider_match(factura.get("proveedor", ""),
                                   bank_mov.get("descripcion", ""))

    nro = str(factura.get("nro_factura") or "").strip()
    ref = str(bank_mov.get("referencia") or "")
    if nro and nro in ref:
        score += 50

    return score
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_matcher_score.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/matcher.py tests/test_matcher_score.py
git commit -m "feat: matcher.match_score scores factura<->bank_mov pairs"
```

---

### Task 2: `find_matches(bank_mov, facturas_pendientes)` — buscar candidatos

**Files:**
- Modify: `Robot/modules/cash_flow/matcher.py`
- Test: `Robot/tests/test_matcher_find.py`

- [ ] **Step 1: Write test**

```python
# tests/test_matcher_find.py
from datetime import date
from modules.cash_flow.matcher import find_matches


def _f(fila, total=1000000, prov="X", fecha=date(2025, 9, 1), nro="1"):
    return {"fila": fila, "total": total, "proveedor": prov,
            "fecha_emision": fecha, "nro_factura": nro}


def _m(cargo=1000000, prov="X", fecha=date(2025, 9, 3), ref=""):
    return {"fila": 7, "cargo": cargo, "abono": 0,
            "descripcion": f"PAGO {prov}", "fecha": fecha, "referencia": ref}


def test_find_single_match():
    facturas = [_f(10), _f(11, total=999999), _f(12, total=1, prov="Z")]
    candidates = find_matches(_m(), facturas)
    assert candidates[0]["fila"] == 10


def test_find_returns_empty_below_threshold():
    facturas = [_f(10, total=500000, prov="Z", fecha=date(2024, 1, 1))]
    assert find_matches(_m(), facturas) == []


def test_find_returns_sorted_desc():
    facturas = [_f(10, prov="ZZZ"), _f(11, prov="X")]  # 11 matches better
    cands = find_matches(_m(prov="X"), facturas)
    assert cands[0]["fila"] == 11
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append a matcher.py)**

```python
MATCH_THRESHOLD = 30


def find_matches(bank_mov: dict, facturas_pendientes: list[dict]) -> list[dict]:
    """Devuelve candidatos sobre threshold, ordenados por score desc."""
    scored = []
    for f in facturas_pendientes:
        s = match_score(f, bank_mov)
        if s >= MATCH_THRESHOLD:
            scored.append({**f, "score": s})
    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_matcher_find.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/matcher.py tests/test_matcher_find.py
git commit -m "feat: matcher.find_matches returns sorted candidates"
```

---

### Task 3: `classify_match(candidates)` — auto/ambiguo/none

**Files:**
- Modify: `Robot/modules/cash_flow/matcher.py`
- Test: `Robot/tests/test_matcher_classify.py`

- [ ] **Step 1: Write test**

```python
# tests/test_matcher_classify.py
from modules.cash_flow.matcher import classify_match


def test_no_candidates_returns_none():
    r = classify_match([])
    assert r["status"] == "no_match"


def test_single_strong_candidate_auto():
    r = classify_match([{"fila": 10, "score": 150}])
    assert r["status"] == "auto"
    assert r["fila"] == 10


def test_two_strong_candidates_ambiguo():
    r = classify_match([{"fila": 10, "score": 150}, {"fila": 11, "score": 140}])
    assert r["status"] == "ambiguo"


def test_strong_plus_weak_picks_strong():
    r = classify_match([{"fila": 10, "score": 150}, {"fila": 11, "score": 35}])
    assert r["status"] == "auto"
    assert r["fila"] == 10


def test_only_weak_candidate_ambiguo():
    r = classify_match([{"fila": 10, "score": 35}])
    assert r["status"] == "ambiguo"
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append)**

```python
STRONG_THRESHOLD = 100
GAP_MIN = 30  # gap entre top y siguiente para considerar auto


def classify_match(candidates: list[dict]) -> dict:
    """Decide auto / ambiguo / no_match.

    - auto: top score >= STRONG_THRESHOLD y gap >= GAP_MIN con el 2do
    - ambiguo: top score >= MATCH_THRESHOLD pero no llega a auto
    - no_match: lista vacia
    """
    if not candidates:
        return {"status": "no_match"}

    top = candidates[0]
    if top["score"] < STRONG_THRESHOLD:
        return {"status": "ambiguo", "candidates": candidates}

    if len(candidates) == 1:
        return {"status": "auto", "fila": top["fila"], "score": top["score"]}

    gap = top["score"] - candidates[1]["score"]
    if gap < GAP_MIN:
        return {"status": "ambiguo", "candidates": candidates}

    return {"status": "auto", "fila": top["fila"], "score": top["score"]}
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_matcher_classify.py -v`
Expected: 5 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/matcher.py tests/test_matcher_classify.py
git commit -m "feat: matcher.classify_match auto/ambiguo/no_match"
```

---
