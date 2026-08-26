"""Backup automatico de Master.xlsx y codigo del Robot a Dropbox."""
import os, shutil, logging
from datetime import datetime

logger = logging.getLogger(__name__)


def backup_master(reason: str, excel_path=None, backup_base=None,
                   cola_path=None):
    """Copia Master a Dropbox: current.xlsx + snapshot con timestamp.

    `cola_path` existe para las PRUEBAS. Sin él, un test que pasa `excel_path` y
    `backup_base` propios igual encolaba en la cola de PRODUCCIÓN, porque el
    encolado leía la ruta de config. Cada corrida de la suite dejaba basura ahí
    apuntando a carpetas temporales de pytest ya borradas. Es el mismo patrón
    que una vez destruyó el Master real: confiar en un default dentro de algo
    que el test creía haber aislado.
    """
    if excel_path is None or backup_base is None:
        from config import EXCEL_PATH, DROPBOX_BACKUP_PATH
        excel_path = excel_path or EXCEL_PATH
        backup_base = backup_base or DROPBOX_BACKUP_PATH

    master_dir = os.path.join(backup_base, "Master")
    snap_dir = os.path.join(master_dir, "snapshots")
    os.makedirs(snap_dir, exist_ok=True)

    shutil.copy2(excel_path, os.path.join(master_dir, "current.xlsx"))

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    snap_path = os.path.join(snap_dir, f"{ts}.xlsx")
    shutil.copy2(excel_path, snap_path)

    try:
        from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS
        from modules.drive.cola import Cola
        Cola(cola_path or DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS).encolar(
            snap_path, "Respaldos/Master", os.path.basename(snap_path))
    except Exception as e:
        logger.warning("No pude encolar el respaldo para Drive: %s", e)

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


def cuales_borrar(snapshots: list[dict], hoy=None) -> list[dict]:
    """Cuáles respaldos sobran.

    Regla: todos los de los últimos 30 días · uno por mes del año en curso ·
    uno por año hacia atrás. Se conserva siempre el más reciente de cada grupo.
    """
    from datetime import date, timedelta
    hoy = hoy or date.today()
    limite_diario = hoy - timedelta(days=30)

    recientes, por_mes, por_anio = [], {}, {}
    for s in snapshots:
        f = s["fecha"]
        if f >= limite_diario:
            recientes.append(s)
        elif f.year == hoy.year:
            por_mes.setdefault((f.year, f.month), []).append(s)
        else:
            por_anio.setdefault(f.year, []).append(s)

    borrar = []
    for grupo in list(por_mes.values()) + list(por_anio.values()):
        grupo.sort(key=lambda s: s["fecha"])
        borrar.extend(grupo[:-1])          # se conserva el más reciente
    return borrar
