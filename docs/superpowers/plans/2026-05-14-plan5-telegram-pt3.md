# Plan 5: Telegram — Parte 3/3 (Wizard cosecha FSM)

### Task 9: Estado del wizard (FSM puro)

**Files:**
- Create: `Robot/modules/cash_flow/cosecha_wizard.py`
- Test: `Robot/tests/test_cosecha_wizard.py`

- [ ] **Step 1: Write test**

```python
# tests/test_cosecha_wizard.py
from modules.cash_flow.cosecha_wizard import CosechaWizard


def test_inicio_pide_kg_totales():
    w = CosechaWizard(cultivo="NOGALES")
    assert w.estado == "esperando_kg_totales"
    assert "kg" in w.prompt.lower()


def test_kg_totales_avanza_a_exportadoras():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("240000")
    assert w.estado == "esperando_exportadoras"
    assert w.data["kg_total"] == 240000


def test_exportadoras_parsea_lista():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("240000")
    w.responder("Valbifrut 140000, Pacific Nuts 100000")
    assert len(w.data["exportadoras"]) == 2
    assert w.data["exportadoras"][0]["nombre"] == "Valbifrut"
    assert w.data["exportadoras"][0]["kg"] == 140000


def test_kg_no_numerico_se_queda():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("abc")
    assert w.estado == "esperando_kg_totales"
    assert "error" in w.prompt.lower() or "numero" in w.prompt.lower()


def test_resumen_al_terminar():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("240000")
    w.responder("Valbifrut 140000")
    w.responder("1.8")  # precio
    w.responder("1")    # cuotas
    w.responder("2026-06-15 252000")  # fecha + monto USD
    w.responder("no")  # liquidacion
    assert w.estado == "resumen"
    assert w.data["exportadoras"][0]["precio_usd_kg"] == 1.8
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (FSM lineal simple)**

```python
# modules/cash_flow/cosecha_wizard.py
"""FSM para wizard de cosecha (cierre por cultivo)."""
import re


PROMPTS = {
    "esperando_kg_totales": "Kg totales cosechados? (numero)",
    "esperando_exportadoras": ("Exportadoras y kg. Formato: "
                                "'Nombre1 kg1, Nombre2 kg2'"),
    "esperando_precio": "Precio USD/kg de {exp}?",
    "esperando_cuotas": "Cuantas cuotas? (1-5)",
    "esperando_cuota_data": "Cuota {n}: fecha YYYY-MM-DD y monto USD",
    "esperando_liquidacion": "Liquidacion final? (si/no)",
    "esperando_liquidacion_data": "Fecha y monto USD estimado",
    "resumen": "Resumen listo. /guardar para confirmar, /cancelar para descartar.",
}


class CosechaWizard:
    def __init__(self, cultivo: str):
        self.cultivo = cultivo
        self.data = {"cultivo": cultivo, "exportadoras": []}
        self.estado = "esperando_kg_totales"
        self._exp_idx = 0
        self._cuota_idx = 0
        self.prompt = PROMPTS["esperando_kg_totales"]
        self._error = ""

    def _set_error(self, msg):
        self._error = msg
        self.prompt = f"Error ({msg}). " + PROMPTS[self.estado]

    def _advance(self, nuevo_estado, **fmt):
        self.estado = nuevo_estado
        self._error = ""
        self.prompt = PROMPTS[nuevo_estado].format(**fmt)

    def responder(self, texto: str):
        texto = (texto or "").strip()

        if self.estado == "esperando_kg_totales":
            try:
                self.data["kg_total"] = int(float(texto))
                self._advance("esperando_exportadoras")
            except ValueError:
                self._set_error("numero invalido")

        elif self.estado == "esperando_exportadoras":
            exps = []
            for chunk in texto.split(","):
                m = re.match(r"(.+?)\s+(\d+)\s*$", chunk.strip())
                if m:
                    exps.append({"nombre": m.group(1).strip(),
                                 "kg": int(m.group(2))})
            if not exps:
                self._set_error("formato invalido")
                return
            self.data["exportadoras"] = exps
            self._exp_idx = 0
            self._advance("esperando_precio",
                          exp=exps[0]["nombre"])

        elif self.estado == "esperando_precio":
            try:
                price = float(texto)
            except ValueError:
                self._set_error("precio invalido")
                return
            self.data["exportadoras"][self._exp_idx]["precio_usd_kg"] = price
            self.data["exportadoras"][self._exp_idx]["cuotas"] = []
            self._advance("esperando_cuotas")

        elif self.estado == "esperando_cuotas":
            try:
                n = int(texto)
                if n < 1 or n > 5:
                    raise ValueError
            except ValueError:
                self._set_error("entre 1 y 5")
                return
            self.data["exportadoras"][self._exp_idx]["n_cuotas"] = n
            self._cuota_idx = 0
            self._advance("esperando_cuota_data", n=1)

        elif self.estado == "esperando_cuota_data":
            m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d+(?:\.\d+)?)\s*$", texto)
            if not m:
                self._set_error("formato fecha + monto")
                return
            self.data["exportadoras"][self._exp_idx]["cuotas"].append({
                "fecha": m.group(1), "usd": float(m.group(2)),
            })
            self._cuota_idx += 1
            total = self.data["exportadoras"][self._exp_idx]["n_cuotas"]
            if self._cuota_idx < total:
                self._advance("esperando_cuota_data", n=self._cuota_idx + 1)
            else:
                # Siguiente exportadora o liquidacion
                self._exp_idx += 1
                if self._exp_idx < len(self.data["exportadoras"]):
                    self._cuota_idx = 0
                    self._advance("esperando_precio",
                                  exp=self.data["exportadoras"][self._exp_idx]["nombre"])
                else:
                    self._advance("esperando_liquidacion")

        elif self.estado == "esperando_liquidacion":
            if texto.lower() in ("si", "sí", "s", "y"):
                self._advance("esperando_liquidacion_data")
            else:
                self.data["liquidacion"] = None
                self._advance("resumen")

        elif self.estado == "esperando_liquidacion_data":
            m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d+(?:\.\d+)?)\s*$", texto)
            if not m:
                self._set_error("formato fecha + monto USD")
                return
            self.data["liquidacion"] = {"fecha": m.group(1),
                                         "usd": float(m.group(2))}
            self._advance("resumen")
```

- [ ] **Step 4: Run — expect PASS**

Run: `py -3.11 -m pytest tests/test_cosecha_wizard.py -v`

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/cosecha_wizard.py tests/test_cosecha_wizard.py
git commit -m "feat: CosechaWizard FSM for harvest income data entry"
```

---

### Task 10: Guardar wizard data en Master.Cosechas

**Files:**
- Modify: `Robot/modules/cash_flow/cosecha_wizard.py`
- Test: `Robot/tests/test_cosecha_wizard_save.py`

- [ ] **Step 1: Write test**

```python
# tests/test_cosecha_wizard_save.py
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


def test_save_wizard_data_writes_rows(m):
    from modules.cash_flow.cosecha_wizard import save_to_cosechas
    data = {
        "cultivo": "NOGALES",
        "kg_total": 240000,
        "exportadoras": [{
            "nombre": "Valbifrut", "kg": 140000,
            "precio_usd_kg": 1.8, "n_cuotas": 1,
            "cuotas": [{"fecha": "2026-06-15", "usd": 252000}],
        }],
        "liquidacion": {"fecha": "2026-12-15", "usd": 50000},
    }
    rows_added = save_to_cosechas(data, year=2026, excel_path=m)
    assert rows_added == 2  # 1 cuota + 1 liquidacion

    wb = load_workbook(m, read_only=True)
    ws = wb[COSECHAS_SHEET]
    # Buscar fila con Valbifrut + cuota 1
    found = False
    for r in range(2, ws.max_row + 1):
        if (ws.cell(r, 4).value == "Valbifrut"
            and ws.cell(r, 8).value == 1):
            assert ws.cell(r, 11).value == "adelanto"
            assert ws.cell(r, 10).value == 252000
            found = True
            break
    wb.close()
    assert found
```

- [ ] **Step 2: Run — expect FAIL**

- [ ] **Step 3: Implement (append cosecha_wizard.py)**

```python
def save_to_cosechas(data: dict, year: int,
                       excel_path: str | None = None) -> int:
    """Escribe filas de wizard en Master.Cosechas. Devuelve # filas agregadas."""
    from openpyxl import load_workbook
    from config import EXCEL_PATH
    from excel_manager import COSECHAS_SHEET, _save_wb
    excel_path = excel_path or EXCEL_PATH

    wb = load_workbook(excel_path)
    ws = wb[COSECHAS_SHEET]
    added = 0
    for exp in data["exportadoras"]:
        n_cuotas = exp.get("n_cuotas", len(exp.get("cuotas", [])))
        for i, cuota in enumerate(exp.get("cuotas", []), start=1):
            ws.append([
                year, data["cultivo"], data["kg_total"], exp["nombre"],
                exp["kg"], exp["precio_usd_kg"],
                n_cuotas + (1 if data.get("liquidacion") else 0),
                i, cuota["fecha"], cuota["usd"],
                "adelanto", "esperado",
                None, None, "", "",
            ])
            added += 1

        # Liquidacion (1 sola por wizard, no por exportadora)
        # se agrega despues del loop principal
        pass

    if data.get("liquidacion"):
        liq = data["liquidacion"]
        # Asignar a la primera exportadora (placeholder)
        exp = data["exportadoras"][0]
        n_cuotas = exp.get("n_cuotas", 0) + 1
        ws.append([
            year, data["cultivo"], data["kg_total"], exp["nombre"],
            exp["kg"], exp["precio_usd_kg"], n_cuotas, n_cuotas,
            liq["fecha"], liq["usd"], "liquidacion final", "esperado",
            None, None, "", "",
        ])
        added += 1

    _save_wb(wb, excel_path)
    wb.close()
    return added
```

- [ ] **Step 4: Run — expect PASS**

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/cosecha_wizard.py tests/test_cosecha_wizard_save.py
git commit -m "feat: save_to_cosechas persists wizard data to Master.Cosechas"
```

---

### Task 11: Handler `/cosecha <cultivo>` con ConversationHandler

**Files:**
- Modify: `Robot/handlers/cash_flow_cmds.py`

- [ ] **Step 1: Implement (no test unit — integration manual)**

Append a `handlers/cash_flow_cmds.py`:

```python
from telegram.ext import (
    ConversationHandler, CommandHandler, MessageHandler, filters,
)
from modules.cash_flow.cosecha_wizard import CosechaWizard, save_to_cosechas

WIZARD_STATE = 1


async def cmd_cosecha_start(update, context):
    args = context.args or []
    if not args:
        await update.message.reply_text("Uso: /cosecha <NOGALES|CEREZOS|AVELLANOS>")
        return ConversationHandler.END
    cultivo = args[0].upper()
    if cultivo not in ("NOGALES", "CEREZOS", "AVELLANOS"):
        await update.message.reply_text("Cultivo invalido")
        return ConversationHandler.END
    w = CosechaWizard(cultivo=cultivo)
    context.user_data["cosecha_wizard"] = w
    await update.message.reply_text(w.prompt)
    return WIZARD_STATE


async def cb_cosecha_resp(update, context):
    w = context.user_data.get("cosecha_wizard")
    if not w:
        return ConversationHandler.END
    w.responder(update.message.text)
    if w.estado == "resumen":
        from datetime import date
        year = date.today().year
        added = save_to_cosechas(w.data, year=year)
        await update.message.reply_text(
            f"Listo. {added} filas guardadas en Cosechas.")
        context.user_data.pop("cosecha_wizard", None)
        return ConversationHandler.END
    await update.message.reply_text(w.prompt)
    return WIZARD_STATE


async def cmd_cosecha_cancel(update, context):
    context.user_data.pop("cosecha_wizard", None)
    await update.message.reply_text("Wizard cancelado.")
    return ConversationHandler.END


cosecha_conv = ConversationHandler(
    entry_points=[CommandHandler("cosecha", cmd_cosecha_start)],
    states={WIZARD_STATE: [MessageHandler(filters.TEXT & ~filters.COMMAND,
                                            cb_cosecha_resp)]},
    fallbacks=[CommandHandler("cancelar", cmd_cosecha_cancel)],
)
```

- [ ] **Step 2: Registrar en main.py**

```python
from handlers.cash_flow_cmds import cosecha_conv
application.add_handler(cosecha_conv)
```

- [ ] **Step 3: Smoke test manual**

`/cosecha NOGALES` → bot responde con prompt → ir respondiendo hasta el resumen.

- [ ] **Step 4: Commit**

```bash
git add handlers/cash_flow_cmds.py main.py
git commit -m "feat: /cosecha ConversationHandler wraps CosechaWizard"
```

---

## Resumen Plan 5

| Task | Descripcion |
|---|---|
| 1 | cmd_proyeccion + format |
| 2 | cmd_categoria |
| 3 | replante.afford_check |
| 4 | Wire main.py |
| 5 | detect_saldo_bajo + factura_por_vencer |
| 6 | AlertDedupe |
| 7 | job_resumen_semanal |
| 8 | Schedule job lunes 8am |
| 9 | CosechaWizard FSM |
| 10 | save_to_cosechas |
| 11 | /cosecha ConversationHandler |

**Resultado:** Bot Telegram con comandos lectura, alertas auto, wizard cosecha completo.
