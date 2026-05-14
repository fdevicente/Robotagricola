# Plan 3: Banco + Matching — Parte 3/3

### Task 7: Runner diario 18:00

**Files:**
- Create: `Robot/daily_banco_18h.py`

Orquesta: scraper → guardar mov → matcher → categorize_bank_movement (no-factura) → backup.

- [ ] **Step 1: Implement runner**

```python
# daily_banco_18h.py
"""
Cron 18:00 diario: trae movimientos del banco, matchea con facturas pendientes,
categoriza los no-factura via Claude, hace backup.

Uso: py -3.11 daily_banco_18h.py
"""
import argparse
import logging
import sys

from openpyxl import load_workbook

from config import EXCEL_PATH
from excel_manager import (
    CUENTA_BANCO_SHEET, guardar_movimientos_banco,
    COL_BANCO_TIPO,
)
from infrastructure.backups import backup_master
from modules.cash_flow.matcher import match_new_bank_movements
from modules.cash_flow.categorizer import categorize_bank_movement

logging.basicConfig(level=logging.INFO,
                     format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)


def _categorize_uncategorized_cargos(excel_path: str) -> dict:
    """Categoriza cargos sin Tipo (los nuevos que matcher no resolvio)."""
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[CUENTA_BANCO_SHEET]
    target_rows = []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        cargo = float(ws.cell(r, 4).value or 0)
        if cargo <= 0:
            continue
        if ws.cell(r, COL_BANCO_TIPO).value:
            continue
        target_rows.append(r)
    wb.close()

    counts = {"categorizados": 0, "errors": 0}
    for r in target_rows:
        try:
            categorize_bank_movement(r, excel_path=excel_path)
            counts["categorizados"] += 1
        except Exception as e:
            logger.error(f"Cat banco fila {r}: {e}")
            counts["errors"] += 1
    return counts


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-scrape", action="store_true",
                         help="No correr scraper, solo matcher + categorizer")
    args = parser.parse_args()

    logger.info("=== Daily Banco 18:00 ===")

    logger.info("1/5 Backup preventivo")
    backup_master(reason="pre-daily-18h")

    if args.skip_scrape:
        logger.info("2/5 SKIP scraper")
        save_report = {"nuevos": 0, "duplicados": 0}
    else:
        logger.info("2/5 Scraping Scotiabank...")
        from scotiabank_scraper import sync_scotiabank_movements
        try:
            movs = sync_scotiabank_movements()
            save_report = guardar_movimientos_banco(movs)
            logger.info(f"   Banco: {save_report}")
        except Exception as e:
            logger.error(f"Scraper fallo: {e}")
            save_report = {"nuevos": 0, "duplicados": 0, "error": str(e)}

    logger.info("3/5 Matcheando facturas vs banco...")
    match_report = match_new_bank_movements()
    logger.info(f"   Match: {match_report}")

    logger.info("4/5 Categorizando cargos no-factura con Claude...")
    cat_report = _categorize_uncategorized_cargos(EXCEL_PATH)
    logger.info(f"   Categorizacion: {cat_report}")

    logger.info("5/5 Backup post-18h")
    backup_master(reason="post-daily-18h")

    logger.info("=== Daily completo ===")
    logger.info(f"Mov nuevos: {save_report.get('nuevos', 0)}")
    logger.info(f"Auto-matched: {match_report.get('auto_matched', 0)}")
    logger.info(f"Ambiguos:     {match_report.get('ambiguous', 0)}")
    logger.info(f"Cargos cat:   {cat_report.get('categorizados', 0)}")


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Smoke test sin scrape (solo matcher + categorizer)**

Run: `py -3.11 daily_banco_18h.py --skip-scrape`
Expected: corre OK, reporta matches encontrados en banco actual.

- [ ] **Step 3: Verificar Master**

Crear script auxiliar y correr:

```python
# _verify_matched.py
from openpyxl import load_workbook
from config import EXCEL_PATH
from excel_manager import SHEET_NAME, CUENTA_BANCO_SHEET, COL_BANCO_FACTURA_LINK

wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
ws_f = wb[SHEET_NAME]
ws_b = wb[CUENTA_BANCO_SHEET]
pagadas = sum(1 for r in range(2, ws_f.max_row + 1)
              if ws_f.cell(r, 3).value and "Banco" in str(ws_f.cell(r, 3).value))
linkeadas = sum(1 for r in range(2, ws_b.max_row + 1)
                if ws_b.cell(r, COL_BANCO_FACTURA_LINK).value)
print(f"Facturas con Fecha Pago (Banco): {pagadas}")
print(f"Movimientos linkeados: {linkeadas}")
wb.close()
```

Run: `py -3.11 _verify_matched.py`
Expected: cuentas consistentes.

- [ ] **Step 4: Commit**

```bash
git add daily_banco_18h.py
git commit -m "feat: add daily_banco_18h.py runner orchestrating scrape+match+categorize"
```

---

### Task 8: Schedule via Windows Task Scheduler (documentacion + script)

**Files:**
- Create: `Robot/scripts/register_daily_banco_task.ps1`
- Create: `Robot/scripts/README_scheduler.md`

- [ ] **Step 1: Implement PowerShell script**

```powershell
# scripts/register_daily_banco_task.ps1
# Registra Task Scheduler: corre daily_banco_18h.py todos los dias a las 18:00.
# Uso: ejecutar como Administrador.

$TaskName = "AgricolaSantaElisa-DailyBanco-18h"
$Python = "py.exe"
$Args = "-3.11 $PSScriptRoot\..\daily_banco_18h.py"
$WorkingDir = "$PSScriptRoot\.."

$Action = New-ScheduledTaskAction -Execute $Python -Argument $Args -WorkingDirectory $WorkingDir
$Trigger = New-ScheduledTaskTrigger -Daily -At 6:00pm
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERNAME" -LogonType Interactive
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 60)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
    -Principal $Principal -Settings $Settings -Force

Write-Host "Tarea registrada: $TaskName (corre cada dia a las 18:00)"
```

- [ ] **Step 2: Implement README**

```markdown
# scripts/README_scheduler.md

## Registrar cron 18:00 (Windows Task Scheduler)

1. Abrir PowerShell como **Administrador**
2. Cambiar a la carpeta del proyecto:
   ```
   cd "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\Robot"
   ```
3. Permitir scripts (una sola vez):
   ```
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
4. Registrar tarea:
   ```
   .\scripts\register_daily_banco_task.ps1
   ```
5. Verificar en Task Scheduler:
   - Abrir `taskschd.msc`
   - Buscar `AgricolaSantaElisa-DailyBanco-18h`

## Reintentos

Si el scraper falla, el Settings tiene `-RestartCount 3 -RestartInterval 60 min`,
asi que reintenta 19h y 20h automaticamente.

## Logs

El runner escribe a stdout. Para capturar log a archivo, editar la accion
en Task Scheduler para redirigir: `> daily_banco.log 2>&1`.

## Desactivar

```
Unregister-ScheduledTask -TaskName "AgricolaSantaElisa-DailyBanco-18h" -Confirm:$false
```
```

- [ ] **Step 3: Commit**

```bash
git add scripts/
git commit -m "feat: Windows Task Scheduler script + docs for 18:00 cron"
```

---

## Resumen Plan 3

| Task | Descripcion | Archivos |
|---|---|---|
| 1 | match_score puro | matcher.py + tests |
| 2 | find_matches | matcher.py + tests |
| 3 | classify_match auto/ambiguo | matcher.py + tests |
| 4 | Readers facturas/banco | excel_manager.py + tests |
| 5 | apply_bank_factura_link | excel_manager.py + tests |
| 6 | match_new_bank_movements | matcher.py + tests |
| 7 | daily_banco_18h.py runner | runner |
| 8 | PowerShell scheduler | scripts/ |

**Resultado:**
- Cada movimiento banco con cargo busca factura pendiente
- Match auto solo si gap claro (score >=100 y diff >=30 con 2do)
- Ambiguos se marcan para resolucion manual (Telegram en Plan 5)
- Cargos sin match → categorizer Claude
- Cron 18:00 ejecuta scraper→matcher→categorizer→backup
