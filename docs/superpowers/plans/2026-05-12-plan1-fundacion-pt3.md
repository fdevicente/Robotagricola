# Plan 1: Fundación — Parte 3/3

### Task 5: Módulo de backups automáticos

**Files:**
- Create: `Robot/infrastructure/__init__.py`
- Create: `Robot/infrastructure/backups.py`
- Test: `Robot/tests/test_backups.py`

- [ ] **Step 1: Create __init__.py**

```python
# infrastructure/__init__.py
```

- [ ] **Step 2: Write test**

```python
# tests/test_backups.py
import os, shutil, pytest
from unittest.mock import patch

@pytest.fixture
def setup_dirs(tmp_path):
    master = tmp_path / "master.xlsx"
    master.write_text("fake excel")
    backup_dir = tmp_path / "Backups"
    return str(master), str(backup_dir)

def test_backup_master_creates_current(setup_dirs):
    master, backup_dir = setup_dirs
    from infrastructure.backups import backup_master
    backup_master(reason="test", excel_path=master, backup_base=backup_dir)
    current = os.path.join(backup_dir, "Master", "current.xlsx")
    assert os.path.exists(current)

def test_backup_master_creates_snapshot(setup_dirs):
    master, backup_dir = setup_dirs
    from infrastructure.backups import backup_master
    backup_master(reason="test", excel_path=master, backup_base=backup_dir)
    snapshots = os.path.join(backup_dir, "Master", "snapshots")
    files = os.listdir(snapshots)
    assert len(files) == 1
    assert files[0].endswith(".xlsx")

def test_daily_snapshot_rotates_30_days(setup_dirs):
    master, backup_dir = setup_dirs
    from infrastructure.backups import backup_master
    snapshots_dir = os.path.join(backup_dir, "Master", "snapshots")
    os.makedirs(snapshots_dir, exist_ok=True)
    # Create 35 fake old snapshots
    for i in range(35):
        f = os.path.join(snapshots_dir, f"2026-04-{i:02d}_18-00.xlsx")
        open(f, 'w').close()
    backup_master(reason="rotation", excel_path=master, backup_base=backup_dir)
    files = os.listdir(snapshots_dir)
    assert len(files) <= 31  # 30 old + 1 new

def test_backup_codebase(setup_dirs, tmp_path):
    _, backup_dir = setup_dirs
    robot_dir = tmp_path / "Robot"
    robot_dir.mkdir()
    (robot_dir / "main.py").write_text("print('hi')")
    (robot_dir / "config.py").write_text("X=1")
    from infrastructure.backups import backup_codebase
    backup_codebase(robot_dir=str(robot_dir), backup_base=backup_dir)
    dest = os.path.join(backup_dir, "Robot", "current", "main.py")
    assert os.path.exists(dest)
```

- [ ] **Step 3: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_backups.py -v`

- [ ] **Step 4: Implement backups.py**

```python
# infrastructure/backups.py
"""Backup automático de Master.xlsx y código del Robot a Dropbox."""
import os, shutil, logging
from datetime import datetime

logger = logging.getLogger(__name__)

def backup_master(reason: str, excel_path=None, backup_base=None):
    """Copia Master a Dropbox: current.xlsx + snapshot con timestamp."""
    from config import EXCEL_PATH, DROPBOX_BACKUP_PATH
    excel_path = excel_path or EXCEL_PATH
    backup_base = backup_base or DROPBOX_BACKUP_PATH

    master_dir = os.path.join(backup_base, "Master")
    snap_dir = os.path.join(master_dir, "snapshots")
    os.makedirs(snap_dir, exist_ok=True)

    # current.xlsx (sobrescribe)
    shutil.copy2(excel_path, os.path.join(master_dir, "current.xlsx"))

    # snapshot con timestamp
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    shutil.copy2(excel_path, os.path.join(snap_dir, f"{ts}.xlsx"))

    # rotación: mantener últimos 30
    _rotate_snapshots(snap_dir, keep=30)

    logger.info(f"Backup Master ({reason}): {ts}")

def backup_codebase(robot_dir=None, backup_base=None):
    """Copia código del Robot a Dropbox (semanal, sobrescribe)."""
    from config import DROPBOX_BACKUP_PATH
    backup_base = backup_base or DROPBOX_BACKUP_PATH

    if robot_dir is None:
        robot_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    dest = os.path.join(backup_base, "Robot", "current")
    if os.path.exists(dest):
        shutil.rmtree(dest)

    skip = {'venv', '__pycache__', '.git', 'node_modules'}
    shutil.copytree(robot_dir, dest,
                    ignore=shutil.ignore_patterns(*skip))
    logger.info("Backup codebase completado")

def _rotate_snapshots(snap_dir, keep=30):
    """Mantiene solo los últimos N snapshots."""
    files = sorted(os.listdir(snap_dir))
    while len(files) > keep:
        os.remove(os.path.join(snap_dir, files.pop(0)))
```

- [ ] **Step 5: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_backups.py -v`

- [ ] **Step 6: Commit**

```bash
git add infrastructure/ tests/test_backups.py
git commit -m "feat: add backup module for Master and codebase to Dropbox"
```

---

### Task 6: Script de setup inicial (onboarding runner)

**Files:**
- Create: `Robot/setup_cash_flow.py`

- [ ] **Step 1: Implement setup script**

```python
# setup_cash_flow.py
"""
Script de setup inicial para Fase 1 — Cash Flow.
Ejecutar UNA VEZ antes de usar el módulo cash_flow.

Uso: py -3.11 setup_cash_flow.py
"""
import logging
from infrastructure.backups import backup_master
from excel_manager import ensure_cash_flow_sheets, ensure_new_columns
from config import EXCEL_PATH

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logger = logging.getLogger(__name__)

def main():
    logger.info("=== Setup Cash Flow — Fase 1 ===")

    # 1. Backup preventivo
    logger.info("1/3 Backup preventivo...")
    backup_master(reason="pre-fase-1-setup")

    # 2. Crear hojas nuevas
    logger.info("2/3 Creando hojas nuevas en Master...")
    ensure_cash_flow_sheets()
    logger.info("   Hojas creadas: Cosechas, Guias Despacho, Flujo Caja, "
                "Ajustes Manuales, Config, Hectareas")

    # 3. Agregar columnas nuevas
    logger.info("3/3 Agregando columnas nuevas...")
    ensure_new_columns()
    logger.info("   Facturas: +Categoria, +Cultivo, +Confianza, +Categorizado_por")
    logger.info("   Cuenta Banco: +Tipo, +Categoria, +Cultivo, +Factura_linkeada")

    # 4. Backup post-setup
    backup_master(reason="post-fase-1-setup")

    logger.info("=== Setup completo ===")
    logger.info(f"Master: {EXCEL_PATH}")
    logger.info("Siguiente paso: ejecutar Plan 2 (categorización del histórico)")

if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Test manually**

Run: `py -3.11 setup_cash_flow.py`
Expected output:
```
=== Setup Cash Flow — Fase 1 ===
1/3 Backup preventivo...
2/3 Creando hojas nuevas en Master...
3/3 Agregando columnas nuevas...
=== Setup completo ===
```

- [ ] **Step 3: Verificar Master**

Run: `py -3.11 -c "from openpyxl import load_workbook; wb=load_workbook(r'...\MASTER Agricola Santa Elisa.xlsx', read_only=True); print(wb.sheetnames)"`
Expected: lista incluye Cosechas, Guias Despacho, Config, Hectareas, etc.

- [ ] **Step 4: Verificar backup en Dropbox**

Run: `ls "C:\Users\Windows\Dropbox\Agricola Santa Elisa\Backups\Master\"`
Expected: `current.xlsx` + `snapshots/` con 1 archivo

- [ ] **Step 5: Commit**

```bash
git add setup_cash_flow.py
git commit -m "feat: add setup_cash_flow.py for one-time Fase 1 initialization"
```

---

## Resumen Plan 1

| Task | Descripción | Archivos |
|---|---|---|
| 1 | Config constants + paths | config.py |
| 2 | Sheet/column constants | excel_manager.py |
| 3 | ensure_cash_flow_sheets() | excel_manager.py + tests |
| 4 | ensure_new_columns() | excel_manager.py + tests |
| 5 | Backup module | infrastructure/backups.py + tests |
| 6 | Setup script (runner) | setup_cash_flow.py |

**Resultado:** Master con 6 hojas nuevas + 8 columnas nuevas + backup funcionando + config lista. Base para Plan 2 (categorización).
