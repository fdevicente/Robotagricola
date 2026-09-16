"""
vacaciones_manager.py — Gestión de Vacaciones y Personal
Hojas: Personal, Vacaciones
"""
import logging
from collections import defaultdict
from datetime import date, datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from config import EXCEL_PATH
from excel_manager import _save_wb  # guardado con reintentos si Excel está abierto
from infrastructure.escritura_master import escribe_master

logger = logging.getLogger(__name__)

PERSONAL_SHEET = "Personal"
VACACIONES_SHEET = "Vacaciones"
BASE_SHEET = "Vacaciones Pendientes"

PERSONAL_HEADERS = [
    "Nombre", "RUT", "Cargo", "Fecha Ingreso",
    "Días Pendientes", "Días Tomados Total", "Última Vacación"
]
VACACIONES_HEADERS = [
    "Trabajador", "Fecha Inicio", "Fecha Fin", "Días",
    "Estado", "Observaciones"
]

# Chile: 15 días hábiles = ~21 corridos por año
DIAS_ANUALES = 15
DIAS_POR_MES = DIAS_ANUALES / 12        # 1,25


def _open_wb(path=None):
    return load_workbook(path or EXCEL_PATH)


def _create_sheet(wb, name, headers, widths, color):
    if name not in wb.sheetnames:
        ws = wb.create_sheet(name)
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        font = Font(bold=True, color="FFFFFF", size=11)
        thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill = fill
            c.font = font
            c.alignment = Alignment(horizontal="center")
            c.border = thin
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i) if i <= 26 else 'A'].width = w
    return wb[name]


def _ensure_personal(wb):
    return _create_sheet(wb, PERSONAL_SHEET, PERSONAL_HEADERS,
                         [28, 14, 22, 12, 16, 16, 12], "4A148C")


def _ensure_vacaciones(wb):
    return _create_sheet(wb, VACACIONES_SHEET, VACACIONES_HEADERS,
                         [28, 12, 12, 8, 12, 35], "6A1B9A")


@escribe_master
def crear_hojas_vacaciones():
    """Crea las hojas Personal y Vacaciones."""
    wb = _open_wb()
    _ensure_personal(wb)
    _ensure_vacaciones(wb)
    _save_wb(wb)


@escribe_master
def agregar_trabajador(nombre: str, rut: str = "", cargo: str = "",
                       fecha_ingreso: str = "") -> bool:
    """Agrega un trabajador a la hoja Personal."""
    wb = _open_wb()
    ws = _ensure_personal(wb)

    # Verificar que no exista
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] and str(row[0]).strip().lower() == nombre.strip().lower():
            wb.close()
            return False

    if not fecha_ingreso:
        fecha_ingreso = date.today().strftime("%Y-%m-%d")

    # Calcular días pendientes según antigüedad
    try:
        fi = datetime.strptime(fecha_ingreso[:10], "%Y-%m-%d").date()
        meses = (date.today().year - fi.year) * 12 + date.today().month - fi.month
        dias_acumulados = round(DIAS_ANUALES * meses / 12)
    except Exception:
        dias_acumulados = 0

    ws.append([nombre, rut, cargo, fecha_ingreso, dias_acumulados, 0, ""])
    _save_wb(wb)
    logger.info(f"Trabajador agregado: {nombre}, {dias_acumulados} días pendientes")
    return True


def _fecha(v):
    """date desde lo que venga en la celda: date, datetime o texto AAAA-MM-DD."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except (TypeError, ValueError):
        return None


def dias_habiles(inicio: date, fin: date) -> int:
    """Días hábiles entre dos fechas, ambas incluidas: lunes a viernes sin feriados.

    Antes se contaba con toordinal() % 7, que salta viernes y sábado y cuenta el
    domingo: una semana de lunes a viernes daba 4.
    """
    from modules.feriados import es_habil
    return sum(1 for i in range((fin - inicio).days + 1)
               if es_habil(inicio + timedelta(days=i)))


def saldos_calculados(wb, hasta: date) -> dict:
    """Saldo de vacaciones de cada trabajador de Personal, hasta el mes de `hasta`.

    saldo base (hoja Vacaciones Pendientes) + 1,25 por mes desde la fecha de ese
    saldo − días de las vacaciones aprobadas. Quien no tiene saldo base parte en
    0 desde su fecha de ingreso. Es la misma cuenta de
    src/dashboard_data.get_vacaciones_pendientes y de la carga del 4-ago-2026.
    """
    tomados, ultima = defaultdict(float), {}
    if VACACIONES_SHEET in wb.sheetnames:
        for row in wb[VACACIONES_SHEET].iter_rows(min_row=2, values_only=True):
            if not row or not row[0] or "aprobad" not in str(row[4] or "").lower():
                continue
            clave = str(row[0]).strip().upper()
            try:
                tomados[clave] += float(row[3] or 0)
            except (TypeError, ValueError):
                continue
            fin = _fecha(row[2])
            if fin and (clave not in ultima or fin > ultima[clave]):
                ultima[clave] = fin

    base = {}
    if BASE_SHEET in wb.sheetnames:
        for row in wb[BASE_SHEET].iter_rows(min_row=2, values_only=True):
            desde = _fecha(row[4]) if row and row[0] else None
            if not desde:
                continue
            try:
                base[str(row[0]).strip().upper()] = (float(row[3] or 0), desde)
            except (TypeError, ValueError):
                continue

    saldos = {}
    for row in _ensure_personal(wb).iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        nombre = str(row[0]).strip()
        clave = nombre.upper()
        saldo, desde = base.get(clave, (0.0, _fecha(row[3])))
        if not desde:
            continue
        meses = max(0, (hasta.year - desde.year) * 12 + hasta.month - desde.month)
        saldos[nombre] = {
            "pendientes": round(saldo + meses * DIAS_POR_MES - tomados[clave], 2),
            "tomados": tomados[clave],
            "ultima": ultima.get(clave),
        }
    return saldos


def _escribir_saldo(ws, row_idx, saldo) -> bool:
    """Deja en la fila de Personal el saldo calculado. True si cambió algo."""
    cambio = False
    for col, valor in ((5, saldo["pendientes"]), (6, saldo["tomados"])):
        try:
            igual = abs(float(ws.cell(row=row_idx, column=col).value or 0) - valor) < 0.005
        except (TypeError, ValueError):
            igual = False
        if not igual:
            ws.cell(row=row_idx, column=col).value = (
                int(valor) if float(valor).is_integer() else valor)
            cambio = True
    if saldo["ultima"] and _fecha(ws.cell(row=row_idx, column=7).value) != saldo["ultima"]:
        ws.cell(row=row_idx, column=7).value = saldo["ultima"]
        cambio = True
    return cambio


@escribe_master
def registrar_vacacion(nombre: str, fecha_inicio: str, fecha_fin: str,
                       observaciones: str = "", path=None, hasta: date = None) -> dict:
    """Registra vacaciones para un trabajador y le deja el saldo recalculado."""
    wb = _open_wb(path)
    ws_per = _ensure_personal(wb)
    ws_vac = _ensure_vacaciones(wb)

    fi, ff = _fecha(fecha_inicio), _fecha(fecha_fin)
    if fi and ff and ff >= fi:
        dias = (ff - fi).days + 1
        habiles = dias_habiles(fi, ff)
    else:
        dias = habiles = 0

    ws_vac.append([nombre, fi or fecha_inicio, ff or fecha_fin, habiles,
                   "Aprobado", observaciones])

    saldos = saldos_calculados(wb, hasta or date.today())
    encontrado = False
    for row_idx in range(2, ws_per.max_row + 1):
        actual = str(ws_per.cell(row=row_idx, column=1).value or "").strip()
        if actual.lower() == nombre.strip().lower():
            encontrado = True
            if actual in saldos:
                _escribir_saldo(ws_per, row_idx, saldos[actual])
            break

    _save_wb(wb, path)
    logger.info(f"Vacación registrada: {nombre}, {habiles} días hábiles")
    return {"nombre": nombre, "inicio": fecha_inicio, "fin": fecha_fin,
            "dias_habiles": habiles, "dias_corridos": dias, "encontrado": encontrado}


def listar_personal() -> list[dict]:
    """Lista todo el personal con días pendientes."""
    wb = _open_wb()
    ws = _ensure_personal(wb)
    personal = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        personal.append({
            "nombre": str(row[0]),
            "rut": str(row[1] or ""),
            "cargo": str(row[2] or ""),
            "fecha_ingreso": str(row[3] or ""),
            "dias_pendientes": float(row[4] or 0),
            "dias_tomados": float(row[5] or 0),
            "ultima_vacacion": str(row[6] or ""),
        })
    wb.close()
    return personal


def vacaciones_pendientes() -> list[dict]:
    """Retorna trabajadores con vacaciones pendientes."""
    return [p for p in listar_personal() if p["dias_pendientes"] > 0]


@escribe_master
def actualizar_dias_mensuales(hasta: date = None, path=None) -> dict:
    """Deja el saldo de vacaciones de cada trabajador al día hasta el mes de `hasta`.

    RECALCULA, no suma. Antes sumaba +1,25 cada vez que corría sin anotar el mes:
    un mes que no corría se perdía y uno que corría dos veces se duplicaba. En
    2026 no acumuló con éxito ni una vez. Solo guarda el Master si algo cambió.
    """
    hasta = hasta or date.today()
    wb = _open_wb(path)
    ws = _ensure_personal(wb)
    saldos = saldos_calculados(wb, hasta)
    actualizados = 0
    for row_idx in range(2, ws.max_row + 1):
        nombre = str(ws.cell(row=row_idx, column=1).value or "").strip()
        if nombre in saldos and _escribir_saldo(ws, row_idx, saldos[nombre]):
            actualizados += 1

    if actualizados:
        _save_wb(wb, path)
        logger.info(f"Vacaciones al día hasta {hasta:%m-%Y}: "
                    f"{actualizados} trabajador(es) actualizados")
    wb.close()
    return {"actualizados": actualizados, "incremento": round(DIAS_POR_MES, 2),
            "hasta": f"{hasta.year}-{hasta.month:02d}"}


def ultimas_vacaciones(n: int = 10) -> list[dict]:
    """Retorna las últimas N vacaciones registradas."""
    wb = _open_wb()
    ws = _ensure_vacaciones(wb)
    vacaciones = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        vacaciones.append({
            "trabajador": str(row[0]),
            "inicio": str(row[1] or ""),
            "fin": str(row[2] or ""),
            "dias": int(row[3] or 0),
            "estado": str(row[4] or ""),
            "observaciones": str(row[5] or ""),
        })
    wb.close()
    return vacaciones[-n:]
