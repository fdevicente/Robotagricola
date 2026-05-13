# Plan 1: Fundación — Parte 1/3

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans

**Goal:** Preparar Master Excel con nuevas columnas/hojas, configuración del sistema y backup automático a Dropbox.

**Architecture:** Extender excel_manager.py con funciones para las 6 hojas nuevas. Crear módulo infrastructure/ para backups. Agregar constantes a config.py.

**Tech Stack:** Python 3.11, openpyxl, shutil (backups), config existente

---

### Task 1: Agregar constantes de configuración a config.py

**Files:**
- Modify: `Robot/config.py`
- Test: `Robot/tests/test_config.py`

- [ ] **Step 1: Write test**

```python
# tests/test_config.py
from config import (EXCEL_PATH, DROPBOX_BACKUP_PATH, GUIAS_DIR,
                     DOCUMENTOS_DIR, FXP_PATH)

def test_paths_exist():
    assert EXCEL_PATH is not None
    assert DROPBOX_BACKUP_PATH is not None

def test_cash_flow_config_defaults():
    from config import CASH_FLOW_CONFIG
    assert CASH_FLOW_CONFIG['saldo_minimo_pct'] == 0.10
    assert CASH_FLOW_CONFIG['umbral_alerta_cat_pct'] == 0.90
    assert CASH_FLOW_CONFIG['umbral_confianza'] == 0.85
    assert CASH_FLOW_CONFIG['ventana_match_dias'] == 15
    assert CASH_FLOW_CONFIG['fecha_limite_cerezas'] == '12-15'
    assert CASH_FLOW_CONFIG['fecha_limite_nueces'] == '05-30'
    assert CASH_FLOW_CONFIG['dias_sin_guia_cierre'] == 7
    assert CASH_FLOW_CONFIG['usd_clp_estimado'] == 1000
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_config.py -v`
Expected: ImportError — CASH_FLOW_CONFIG not defined

- [ ] **Step 3: Implement — agregar a config.py**

```python
# Agregar al final de config.py

# --- Cash Flow (Fase 1) ---
DROPBOX_BASE = os.getenv("DROPBOX_BASE",
    r"C:\Users\Windows\Dropbox\Agricola Santa Elisa")
DROPBOX_BACKUP_PATH = os.path.join(DROPBOX_BASE, "Backups")
FXP_PATH = os.path.join(DROPBOX_BASE, "FXP.xlsx")
GUIAS_DIR = os.getenv("GUIAS_DIR",
    os.path.join(os.path.dirname(EXCEL_PATH), "Guias Recibidas por Telegram"))
DOCUMENTOS_DIR = os.path.join(DROPBOX_BASE, "Documentos")
REPORTES_DIR = os.path.join(os.path.dirname(EXCEL_PATH), "Reportes")

for _d in [GUIAS_DIR, DOCUMENTOS_DIR, REPORTES_DIR,
           os.path.join(DOCUMENTOS_DIR, "Guias Despacho"),
           os.path.join(DROPBOX_BACKUP_PATH, "Master", "snapshots"),
           os.path.join(DROPBOX_BACKUP_PATH, "Robot")]:
    os.makedirs(_d, exist_ok=True)

CASH_FLOW_CONFIG = {
    'saldo_minimo_pct': 0.10,
    'umbral_alerta_cat_pct': 0.90,
    'umbral_confianza': 0.85,
    'ventana_match_dias': 15,
    'fecha_limite_cerezas': '12-15',
    'fecha_limite_nueces': '05-30',
    'dias_sin_guia_cierre': 7,
    'usd_clp_estimado': 1000,
}
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_config.py -v`

- [ ] **Step 5: Commit**

```bash
git add config.py tests/test_config.py
git commit -m "feat: add cash flow config constants and paths"
```

---

### Task 2: Constantes de hojas y columnas en excel_manager.py

**Files:**
- Modify: `Robot/excel_manager.py`

- [ ] **Step 1: Agregar constantes al inicio de excel_manager.py**

```python
# Agregar después de las constantes existentes (SHEET_NAME, BOLETAS_SHEET, etc.)

# --- Hojas nuevas (Fase 1 Cash Flow) ---
COSECHAS_SHEET = "Cosechas"
GUIAS_SHEET = "Guias Despacho"
FLUJO_CAJA_SHEET = "Flujo Caja"
AJUSTES_SHEET = "Ajustes Manuales"
CONFIG_SHEET = "Config"
HECTAREAS_SHEET = "Hectareas"
CUENTA_BANCO_SHEET = "Cuenta Banco"

# Columnas nuevas en Facturas (después de col 16)
COL_CATEGORIA = 17       # Q
COL_CULTIVO = 18          # R
COL_CONFIANZA = 19        # S
COL_CATEGORIZADO_POR = 20 # T

# Columnas nuevas en Cuenta Banco (después de col 6)
COL_BANCO_TIPO = 7             # G
COL_BANCO_CATEGORIA = 8        # H
COL_BANCO_CULTIVO = 9           # I
COL_BANCO_FACTURA_LINK = 10     # J

# Categorías válidas
CATEGORIAS = [
    "Mano de obra planta",
    "Mano de obra temporal",
    "Fertilizantes",
    "Fitosanitarios",
    "Combustible",
    "Maquinaria - mantención",
    "Riego",
    "Servicios profesionales",
    "Arriendos / Patentes / Seguros",
    "Inversión / Replante",
    "Caja chica / Imprevistos",
]

CULTIVOS = ["NOGALES", "CEREZOS", "AVELLANOS", "GENERAL"]
```

- [ ] **Step 2: Commit**

```bash
git add excel_manager.py
git commit -m "feat: add cash flow sheet/column constants to excel_manager"
```
