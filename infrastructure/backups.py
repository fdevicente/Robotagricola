"""Backup automatico de Master.xlsx y codigo del Robot a Dropbox."""
import os, shutil, logging
from datetime import datetime

logger = logging.getLogger(__name__)


def backup_master(reason: str, excel_path=None, backup_base=None):
    """Copia Master a Dropbox: current.xlsx + snapshot con timestamp."""
    if excel_path is None or backup_base is None:
        from config import EXCEL_PATH, DROPBOX_BACKUP_PATH
        excel_path = excel_path or EXCEL_PATH
        backup_base = backup_base or DROPBOX_BACKUP_PATH

    master_dir = os.path.join(backup_base, "Master")
    snap_dir = os.path.join(master_dir, "snapshots")
    os.makedirs(snap_dir, exist_ok=True)

    shutil.copy2(excel_path, os.path.join(master_dir, "current.xlsx"))

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    shutil.copy2(excel_path, os.path.join(snap_dir, f"{ts}.xlsx"))

    _rotate_snapshots(snap_dir, keep=30)

    logger.info(f"Backup Master ({reason}): {ts}")


def backup_codebase(robot_dir=None, backup_base=None):
    """Copia codigo del Robot a Dropbox (semanal, sobrescribe)."""
    if backup_base is None:
        from config import DROPBOX_BACKUP_PATH
        backup_base = DROPBOX_BACKUP_PATH

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
    """Mantiene solo los ultimos N snapshots."""
    files = sorted(os.listdir(snap_dir))
    while len(files) > keep:
        os.remove(os.path.join(snap_dir, files.pop(0)))
