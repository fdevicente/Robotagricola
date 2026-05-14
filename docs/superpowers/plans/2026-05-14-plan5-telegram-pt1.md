# Plan 5: Telegram — Parte 1/3 (Comandos lectura)

> Sub-skill: superpowers:executing-plans

**Goal:** Comandos Telegram para consultar proyeccion, categorias, replante.

**Tech:** python-telegram-bot, reusa modules/cash_flow/.

---

### Task 1: `/proyeccion` muestra saldo proyectado N meses

**Files:**
- Create: `Robot/handlers/cash_flow_cmds.py`
- Test: `Robot/tests/test_cash_flow_cmds.py`

- [ ] **Step 1: Write test**

```python
# tests/test_cash_flow_cmds.py
from unittest.mock import patch, MagicMock
from handlers.cash_flow_cmds import format_proyeccion


def _fake_cf():
    return {
        "months": [(2026, 5), (2026, 6)],
        "saldo": {
            (2026, 5): {"saldo_inicio": 100, "ingresos": 200, "egresos": 50, "saldo_cierre": 250},
            (2026, 6): {"saldo_inicio": 250, "ingresos": 0, "egresos": 30, "saldo_cierre": 220},
        },
        "egresos": {}, "ingresos": [],
    }


def test_format_proyeccion_has_each_month():
    text = format_proyeccion(_fake_cf())
    assert "May-26" in text and "Jun-26" in text
    assert "220" in text or "$220" in text


def test_format_proyeccion_marks_negative_red():
    cf = _fake_cf()
    cf["saldo"][(2026, 6)]["saldo_cierre"] = -1000
    text = format_proyeccion(cf)
    assert "🔴" in text or "-1" in text
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement**

```python
# handlers/cash_flow_cmds.py
"""Comandos Telegram para cash flow."""
import logging
from modules.cash_flow.projector import get_cash_flow

logger = logging.getLogger(__name__)

_MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
          "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _label(y, m):
    return f"{_MESES[m]}-{str(y)[-2:]}"


def _fmt_money(v):
    return f"${v:,.0f}"


def format_proyeccion(cf: dict) -> str:
    """Formato texto para Telegram."""
    lines = ["📊 *Proyeccion flujo de caja*", ""]
    lines.append(f"`{'Mes':<8} {'Ingr':>12} {'Egr':>12} {'Saldo':>14}`")
    for ym in cf["months"]:
        s = cf["saldo"][ym]
        emoji = "🔴" if s["saldo_cierre"] < 0 else ""
        lines.append(
            f"`{_label(*ym):<8} {_fmt_money(s['ingresos']):>12} "
            f"{_fmt_money(s['egresos']):>12} {_fmt_money(s['saldo_cierre']):>14}` {emoji}"
        )
    return "\n".join(lines)


async def cmd_proyeccion(update, context):
    """`/proyeccion [meses]` (default 6)."""
    args = context.args or []
    n_meses = int(args[0]) if args and args[0].isdigit() else 6
    saldo_actual = 130_600_000  # TODO leer del banco

    from datetime import date
    today = date.today()
    sy, sm = today.year, today.month
    ey, em = sy, sm
    for _ in range(n_meses - 1):
        em += 1
        if em > 12:
            em = 1
            ey += 1

    cf = get_cash_flow(start=(sy, sm), end=(ey, em),
                       saldo_inicial=saldo_actual)
    text = format_proyeccion(cf)
    await update.message.reply_text(text, parse_mode="Markdown")
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_cash_flow_cmds.py -v`

- [ ] **Step 5: Commit**

```bash
git add handlers/cash_flow_cmds.py tests/test_cash_flow_cmds.py
git commit -m "feat: cmd_proyeccion shows N months projection in Telegram"
```

---

### Task 2: `/categoria <nombre>` muestra detalle mensual

**Files:**
- Modify: `Robot/handlers/cash_flow_cmds.py`
- Test: `Robot/tests/test_cmd_categoria.py`

- [ ] **Step 1: Write test**

```python
# tests/test_cmd_categoria.py
from handlers.cash_flow_cmds import format_categoria


def test_format_categoria_groups_by_month():
    egresos = {
        (2026, 5, "Fertilizantes", "NOGALES"): 5_000_000,
        (2026, 5, "Fertilizantes", "GENERAL"): 1_000_000,
        (2026, 6, "Fertilizantes", "NOGALES"): 3_000_000,
        (2026, 5, "Riego", "GENERAL"): 999,
    }
    text = format_categoria("Fertilizantes", egresos,
                              months=[(2026, 5), (2026, 6)])
    assert "Fertilizantes" in text
    assert "6,000,000" in text or "6.000.000" in text  # May total
    assert "3,000,000" in text or "3.000.000" in text  # Jun
    assert "Riego" not in text
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append handlers/cash_flow_cmds.py)**

```python
def format_categoria(cat_name: str, egresos: dict, months: list) -> str:
    """Texto Telegram con detalle de una categoria."""
    lines = [f"📋 *Categoria: {cat_name}*", ""]
    total = 0
    for ym in months:
        m_total = sum(v for (y, mo, c, _cul), v in egresos.items()
                       if y == ym[0] and mo == ym[1] and c == cat_name)
        total += m_total
        lines.append(f"`{_label(*ym):<8} {_fmt_money(m_total):>14}`")
    lines.append("")
    lines.append(f"`Total      {_fmt_money(total):>14}`")
    return "\n".join(lines)


async def cmd_categoria(update, context):
    """`/categoria <nombre>` muestra gasto mensual."""
    args = context.args or []
    if not args:
        await update.message.reply_text(
            "Uso: /categoria <nombre>. Ej: /categoria Fertilizantes")
        return
    cat = " ".join(args)
    from datetime import date
    today = date.today()
    months = [(today.year, m) for m in range(1, today.month + 1)]
    cf = get_cash_flow(start=months[0], end=months[-1],
                       saldo_inicial=0)
    text = format_categoria(cat, cf["egresos"], months)
    await update.message.reply_text(text, parse_mode="Markdown")
```

- [ ] **Step 4: Run — expect PASS**

- [ ] **Step 5: Commit**

```bash
git add handlers/cash_flow_cmds.py tests/test_cmd_categoria.py
git commit -m "feat: cmd_categoria shows monthly spend per category"
```

---

### Task 3: `/replante <cultivo> <hc>` affordability check

**Files:**
- Create: `Robot/modules/cash_flow/replante.py`
- Test: `Robot/tests/test_replante.py`

- [ ] **Step 1: Write test**

```python
# tests/test_replante.py
from modules.cash_flow.replante import afford_check


def test_afford_check_yes_if_enough():
    r = afford_check(cultivo="AVELLANOS", hc=2,
                       saldo_proyectado=50_000_000,
                       saldo_minimo=10_000_000,
                       costo_por_hc=5_000_000)
    assert r["alcanza"] is True
    assert r["disponible"] == 40_000_000
    assert r["costo_total"] == 10_000_000


def test_afford_check_no_if_short():
    r = afford_check(cultivo="AVELLANOS", hc=10,
                       saldo_proyectado=20_000_000,
                       saldo_minimo=5_000_000,
                       costo_por_hc=5_000_000)
    assert r["alcanza"] is False
    assert r["deficit"] == 35_000_000


def test_afford_check_zero_hc():
    r = afford_check(cultivo="AVELLANOS", hc=0,
                       saldo_proyectado=10_000_000,
                       saldo_minimo=5_000_000,
                       costo_por_hc=1_000_000)
    assert r["alcanza"] is True
    assert r["costo_total"] == 0
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement**

```python
# modules/cash_flow/replante.py
"""Affordability check para replante."""


def afford_check(cultivo: str, hc: float,
                   saldo_proyectado: float,
                   saldo_minimo: float,
                   costo_por_hc: float) -> dict:
    """Verifica si alcanza la caja para replantar X hectareas."""
    costo_total = hc * costo_por_hc
    disponible = saldo_proyectado - saldo_minimo
    alcanza = costo_total <= disponible
    return {
        "cultivo": cultivo,
        "hc": hc,
        "costo_por_hc": costo_por_hc,
        "costo_total": costo_total,
        "saldo_proyectado": saldo_proyectado,
        "saldo_minimo": saldo_minimo,
        "disponible": disponible,
        "alcanza": alcanza,
        "deficit": 0 if alcanza else (costo_total - disponible),
    }
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_replante.py -v`

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/replante.py tests/test_replante.py
git commit -m "feat: replante.afford_check pure affordability calc"
```

---

### Task 4: Wire commands in main.py

**Files:**
- Modify: `Robot/main.py`

- [ ] **Step 1: Importar y registrar**

Buscar `application.add_handler(CommandHandler(` en main.py y agregar:

```python
from handlers.cash_flow_cmds import cmd_proyeccion, cmd_categoria

application.add_handler(CommandHandler("proyeccion", cmd_proyeccion))
application.add_handler(CommandHandler("categoria", cmd_categoria))
```

- [ ] **Step 2: Smoke test manual**

Start bot: `py -3.11 main.py`
En Telegram: `/proyeccion 3` → debe responder con tabla 3 meses.

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: register /proyeccion and /categoria in main bot"
```

---
