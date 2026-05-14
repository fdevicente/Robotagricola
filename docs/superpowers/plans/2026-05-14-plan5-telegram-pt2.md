# Plan 5: Telegram — Parte 2/3 (Alertas + jobs)

### Task 5: Detectores de alertas (puras)

**Files:**
- Create: `Robot/modules/cash_flow/alerts.py`
- Test: `Robot/tests/test_alerts.py`

- [ ] **Step 1: Write test**

```python
# tests/test_alerts.py
from datetime import date
from modules.cash_flow.alerts import (
    detect_saldo_bajo, detect_factura_por_vencer,
)


def test_saldo_bajo_below_minimo():
    a = detect_saldo_bajo(saldo_actual=20_000_000, saldo_minimo=36_000_000)
    assert a is not None
    assert "20" in a["mensaje"]


def test_saldo_ok_no_alerta():
    assert detect_saldo_bajo(saldo_actual=100_000_000,
                              saldo_minimo=36_000_000) is None


def test_factura_vence_en_3_dias():
    hoy = date(2026, 5, 14)
    facts = [{
        "fila": 10, "proveedor": "COPEVAL",
        "fecha_vencimiento": date(2026, 5, 17),
        "total": 5_000_000, "nro_factura": "F123",
    }]
    alertas = detect_factura_por_vencer(facts, hoy=hoy, dias=3)
    assert len(alertas) == 1
    assert alertas[0]["fila"] == 10


def test_factura_vencida_no_alerta():
    hoy = date(2026, 5, 14)
    facts = [{"fila": 10, "proveedor": "X",
              "fecha_vencimiento": date(2026, 5, 10),
              "total": 100, "nro_factura": "X"}]
    alertas = detect_factura_por_vencer(facts, hoy=hoy, dias=3)
    assert alertas == []
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement**

```python
# modules/cash_flow/alerts.py
"""Detectores de alertas (puras, sin I/O Telegram)."""
from datetime import date, datetime, timedelta


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        try:
            return datetime.strptime(v[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def detect_saldo_bajo(saldo_actual: float, saldo_minimo: float) -> dict | None:
    """Alerta 🔴 si saldo < minimo."""
    if saldo_actual >= saldo_minimo:
        return None
    diff = saldo_minimo - saldo_actual
    return {
        "tipo": "saldo_bajo",
        "nivel": "🔴",
        "mensaje": (f"🔴 Saldo bajo: ${saldo_actual:,.0f} CLP "
                     f"(falta ${diff:,.0f} para el minimo de ${saldo_minimo:,.0f})"),
    }


def detect_factura_por_vencer(facturas: list, hoy: date | None = None,
                                dias: int = 3) -> list:
    """Alerta 🟡 por cada factura que vence en `dias` o menos (no vencidas)."""
    hoy = hoy or date.today()
    alertas = []
    for f in facturas:
        venc = _to_date(f.get("fecha_vencimiento"))
        if not venc:
            continue
        delta = (venc - hoy).days
        if 0 <= delta <= dias:
            alertas.append({
                "tipo": "factura_por_vencer",
                "nivel": "🟡",
                "fila": f["fila"],
                "mensaje": (
                    f"🟡 Vence en {delta}d: {f.get('proveedor', '')} "
                    f"factura {f.get('nro_factura', '')} ${f.get('total', 0):,.0f}"
                ),
            })
    return alertas
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_alerts.py -v`

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/alerts.py tests/test_alerts.py
git commit -m "feat: alert detectors saldo bajo + factura por vencer"
```

---

### Task 6: Dedupe de alertas (no enviar 2 veces el mismo mes)

**Files:**
- Modify: `Robot/modules/cash_flow/alerts.py`
- Test: `Robot/tests/test_alerts_dedupe.py`

- [ ] **Step 1: Write test**

```python
# tests/test_alerts_dedupe.py
from modules.cash_flow.alerts import AlertDedupe


def test_first_fire_returns_true(tmp_path):
    d = AlertDedupe(path=str(tmp_path / "d.json"))
    assert d.should_fire("saldo_bajo", "2026-05") is True


def test_second_same_month_blocks(tmp_path):
    p = str(tmp_path / "d.json")
    d = AlertDedupe(path=p)
    d.should_fire("saldo_bajo", "2026-05")
    d.mark_fired("saldo_bajo", "2026-05")
    d2 = AlertDedupe(path=p)
    assert d2.should_fire("saldo_bajo", "2026-05") is False


def test_different_month_allowed(tmp_path):
    p = str(tmp_path / "d.json")
    d = AlertDedupe(path=p)
    d.mark_fired("saldo_bajo", "2026-05")
    assert d.should_fire("saldo_bajo", "2026-06") is True
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append alerts.py)**

```python
import json
import os


class AlertDedupe:
    """Recuerda que alertas ya se enviaron por (tipo, periodo)."""

    def __init__(self, path: str):
        self.path = path
        self._data = self._load()

    def _load(self) -> dict:
        if not os.path.exists(self.path):
            return {}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}

    def _save(self):
        os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self._data, f)

    def should_fire(self, tipo: str, periodo: str) -> bool:
        return self._data.get(f"{tipo}|{periodo}") is None

    def mark_fired(self, tipo: str, periodo: str):
        self._data[f"{tipo}|{periodo}"] = True
        self._save()
```

- [ ] **Step 4: Run — expect PASS**

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/alerts.py tests/test_alerts_dedupe.py
git commit -m "feat: AlertDedupe persists fired alerts per (tipo,periodo)"
```

---

### Task 7: Job semanal lunes 8am (resumen)

**Files:**
- Create: `Robot/handlers/cash_flow_jobs.py`
- Test: `Robot/tests/test_cash_flow_jobs.py`

- [ ] **Step 1: Write test**

```python
# tests/test_cash_flow_jobs.py
from unittest.mock import MagicMock
from handlers.cash_flow_jobs import format_resumen_semanal


def test_resumen_includes_saldo_y_alertas():
    cf = {
        "months": [(2026, 5)],
        "saldo": {(2026, 5): {"saldo_inicio": 100, "ingresos": 200,
                                "egresos": 50, "saldo_cierre": 250}},
        "egresos": {}, "ingresos": [],
    }
    alertas = [{"mensaje": "🟡 Vence en 2d: COPEVAL F123"}]
    text = format_resumen_semanal(cf, alertas)
    assert "250" in text or "$250" in text
    assert "COPEVAL" in text


def test_resumen_sin_alertas_dice_ok():
    cf = {
        "months": [(2026, 5)],
        "saldo": {(2026, 5): {"saldo_inicio": 100, "ingresos": 0,
                                "egresos": 50, "saldo_cierre": 50}},
        "egresos": {}, "ingresos": [],
    }
    text = format_resumen_semanal(cf, [])
    assert "sin alertas" in text.lower() or "ok" in text.lower()
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement**

```python
# handlers/cash_flow_jobs.py
"""Jobs programados de cash flow (resumen semanal, etc.)."""
import logging
from datetime import date

logger = logging.getLogger(__name__)

_MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
          "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _label(y, m):
    return f"{_MESES[m]}-{str(y)[-2:]}"


def format_resumen_semanal(cf: dict, alertas: list) -> str:
    lines = ["📅 *Resumen semanal*", ""]
    ym = cf["months"][0] if cf["months"] else None
    if ym:
        s = cf["saldo"][ym]
        lines.append(f"Mes {_label(*ym)}: saldo cierre ${s['saldo_cierre']:,.0f}")
        lines.append(f"  ingresos ${s['ingresos']:,.0f}, "
                      f"egresos ${s['egresos']:,.0f}")
    lines.append("")
    if alertas:
        lines.append("*Alertas:*")
        for a in alertas:
            lines.append(f"  {a['mensaje']}")
    else:
        lines.append("✅ Sin alertas activas")
    return "\n".join(lines)


async def job_resumen_semanal(context):
    """Envia resumen al chat configurado. Trigger: cron Lunes 8am."""
    from config import TELEGRAM_CHAT_ID
    from modules.cash_flow.projector import get_cash_flow
    from modules.cash_flow.alerts import detect_factura_por_vencer
    from excel_manager import read_facturas_pendientes

    if not TELEGRAM_CHAT_ID:
        logger.warning("TELEGRAM_CHAT_ID no configurado, skip resumen semanal")
        return

    today = date.today()
    cf = get_cash_flow(start=(today.year, today.month),
                       end=(today.year, today.month),
                       saldo_inicial=130_600_000)
    facturas = read_facturas_pendientes()
    alertas = detect_factura_por_vencer(facturas, hoy=today, dias=7)
    text = format_resumen_semanal(cf, alertas)
    await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID,
                                    text=text, parse_mode="Markdown")
```

- [ ] **Step 4: Run — expect PASS**

- [ ] **Step 5: Commit**

```bash
git add handlers/cash_flow_jobs.py tests/test_cash_flow_jobs.py
git commit -m "feat: job_resumen_semanal sends weekly cash flow summary"
```

---

### Task 8: Registrar job semanal en main.py

**Files:**
- Modify: `Robot/main.py`

- [ ] **Step 1: Registrar JobQueue**

Buscar en main.py el bloque que registra otros jobs (`job_sync_banco`, `job_vacaciones_mensuales`). Agregar:

```python
from handlers.cash_flow_jobs import job_resumen_semanal
from datetime import time

# Job semanal: lunes 8:00 AM
application.job_queue.run_daily(
    job_resumen_semanal,
    time=time(hour=8, minute=0),
    days=(0,),  # lunes (0 = monday en JobQueue)
    name="resumen_semanal",
)
```

- [ ] **Step 2: Commit**

```bash
git add main.py
git commit -m "feat: schedule job_resumen_semanal every Monday 8am"
```

---
