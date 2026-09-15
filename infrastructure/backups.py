"""Backup automatico de Master.xlsx y codigo del Robot a Dropbox."""
import hashlib
import logging
import os
import re
import shutil
import time
from datetime import date, datetime

logger = logging.getLogger(__name__)

# Nombre de los respaldos automáticos: 2026-09-15_09-24.xlsx. Solo a estos se les
# aplica la retención; los manuales (MASTER_pre_*, de los scripts de carga) no se
# borran nunca solos.
_AUTOMATICO = re.compile(r"^(\d{4}-\d{2}-\d{2})_\d{2}-\d{2}\.xlsx$")


def backup_master(reason: str, excel_path=None, backup_base=None,
                   cola_path=None):
    """Copia Master a Dropbox: current.xlsx + snapshot con timestamp.

    `cola_path` existe para las PRUEBAS. Sin él, un test que pasa `excel_path` y
    `backup_base` propios igual encolaba en la cola de PRODUCCIÓN, porque el
    encolado leía la ruta de config. Cada corrida de la suite dejaba basura ahí
    apuntando a carpetas temporales de pytest ya borradas. Es el mismo patrón
    que una vez destruyó el Master real: confiar en un default dentro de algo
    que el test creía haber aislado.

    Devuelve la ruta del snapshot. Si el Master cambia mientras se copia y no se
    logra una copia fiel, levanta RuntimeError en vez de dejar un respaldo a
    medias.
    """
    if excel_path is None or backup_base is None:
        from config import EXCEL_PATH, DROPBOX_BACKUP_PATH
        excel_path = excel_path or EXCEL_PATH
        backup_base = backup_base or DROPBOX_BACKUP_PATH

    master_dir = os.path.join(backup_base, "Master")
    snap_dir = os.path.join(master_dir, "snapshots")
    os.makedirs(snap_dir, exist_ok=True)

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    snap_path = os.path.join(snap_dir, f"{ts}.xlsx")
    _copiar_verificado(excel_path, snap_path)
    _copiar_verificado(snap_path, os.path.join(master_dir, "current.xlsx"))

    try:
        from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS
        from modules.drive.cola import Cola
        Cola(cola_path or DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS).encolar(
            snap_path, "Respaldos/Master", os.path.basename(snap_path))
    except Exception as e:
        logger.warning("No pude encolar el respaldo para Drive: %s", e)

    _aplicar_retencion(snap_dir)

    logger.info(f"Backup Master ({reason}): {ts}")
    return snap_path


def respaldar_si_cambio(reason: str, excel_path=None, backup_base=None,
                        cola_path=None):
    """Respalda el Master solo si cambió desde el último respaldo automático.

    Es lo que corre el job del bot (al arrancar y cada 6 h): sin cambios no se
    llenan Dropbox ni Drive de copias idénticas. Devuelve la ruta del snapshot,
    o None si no hacía falta.
    """
    if excel_path is None or backup_base is None:
        from config import EXCEL_PATH, DROPBOX_BACKUP_PATH
        excel_path = excel_path or EXCEL_PATH
        backup_base = backup_base or DROPBOX_BACKUP_PATH

    snap_dir = os.path.join(backup_base, "Master", "snapshots")
    if os.path.isdir(snap_dir):
        automaticos = sorted(n for n in os.listdir(snap_dir) if _AUTOMATICO.match(n))
        if automaticos and (_huella(os.path.join(snap_dir, automaticos[-1]))
                            == _huella(excel_path)):
            return None
    return backup_master(reason, excel_path=excel_path, backup_base=backup_base,
                         cola_path=cola_path)


def _huella(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _copiar_verificado(origen, destino, intentos=3, espera=2):
    """Copia `origen` a `destino` solo si la copia es fiel.

    Se copia a un temporal y se compara su huella con la del origen antes y
    después: si el Master se guardó mientras se copiaba, la copia puede quedar a
    medias. Solo cuando calzan las tres se reemplaza el destino.
    """
    parcial = destino + ".parcial"
    for intento in range(1, intentos + 1):
        antes = _huella(origen)
        shutil.copy2(origen, parcial)
        if _huella(parcial) == antes == _huella(origen):
            os.replace(parcial, destino)
            return
        os.remove(parcial)
        if intento < intentos:
            time.sleep(espera)
    raise RuntimeError(f"{origen} cambió mientras se copiaba ({intentos} intentos): "
                       "no dejé un respaldo a medias")


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


def _aplicar_retencion(snap_dir, hoy=None):
    """Borra los respaldos AUTOMÁTICOS que sobran según `cuales_borrar`.

    Reemplaza a la rotación que dejaba "los últimos 30" por orden ALFABÉTICO:
    los 'MASTER_pre_*' ordenan después de '2026-…', se quedaban con los cupos y
    lo que se borraba eran los respaldos nuevos (de 47 quedaban 8). Lo que no
    tiene nombre de respaldo automático no se toca.
    """
    snaps = []
    for nombre in sorted(os.listdir(snap_dir)):     # por nombre: en empate de día gana el último
        m = _AUTOMATICO.match(nombre)
        if not m:
            continue
        try:
            snaps.append({"nombre": nombre, "fecha": date.fromisoformat(m.group(1))})
        except ValueError:
            continue
    for s in cuales_borrar(snaps, hoy=hoy):
        try:
            os.remove(os.path.join(snap_dir, s["nombre"]))
        except OSError as e:
            logger.warning("No pude borrar el respaldo viejo %s: %s", s["nombre"], e)


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
