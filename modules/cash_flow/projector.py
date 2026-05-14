"""Motor de proyeccion de flujo de caja.

Calcula proyeccion mes x categoria x cultivo escalando un ano base
por el factor de hectareas + aplica ajustes manuales del usuario.
"""
from collections import defaultdict
from datetime import date, datetime
from openpyxl import load_workbook

from config import EXCEL_PATH
from excel_manager import (
    SHEET_NAME, CUENTA_BANCO_SHEET,
    COSECHAS_SHEET, HECTAREAS_SHEET, AJUSTES_SHEET, FLUJO_CAJA_SHEET,
    CATEGORIAS, CULTIVOS,
)


def _to_year_month(v) -> tuple[int, int] | None:
    if isinstance(v, datetime):
        return (v.year, v.month)
    if isinstance(v, date):
        return (v.year, v.month)
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                d = datetime.strptime(v[:10], fmt).date()
                return (d.year, d.month)
            except ValueError:
                pass
    return None


def load_historical_egresos(excel_path: str | None = None,
                              year: int | None = None) -> dict:
    """Agrupa Facturas por (year, month, categoria, cultivo) -> total.

    Salta filas sin Categoria o con Categoria=REVISAR.
    """
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    agg: dict = defaultdict(float)
    for row in ws.iter_rows(min_row=2, max_col=18, values_only=True):
        proveedor = row[3]
        if not proveedor:
            continue
        categoria = row[16]
        if not categoria or categoria == "REVISAR":
            continue
        cultivo = row[17] or "GENERAL"
        total = row[14]
        if not total:
            continue
        ym = _to_year_month(row[0])
        if not ym:
            continue
        if year is not None and ym[0] != year:
            continue
        agg[(ym[0], ym[1], categoria, cultivo)] += float(total)
    wb.close()
    return dict(agg)


def load_hectareas(excel_path: str | None = None) -> dict:
    """Devuelve {year: {cultivo: hc}}. Cultivos en uppercase."""
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[HECTAREAS_SHEET]
    hc: dict = {}
    for row in ws.iter_rows(min_row=2, max_col=5, values_only=True):
        year = row[0]
        if not isinstance(year, int):
            continue
        hc[year] = {
            "NOGALES": float(row[1] or 0),
            "CEREZOS": float(row[2] or 0),
            "AVELLANOS": float(row[3] or 0),
        }
    wb.close()
    return hc


def _parse_mes_str(v):
    if isinstance(v, datetime):
        return (v.year, v.month)
    if isinstance(v, date):
        return (v.year, v.month)
    if isinstance(v, str):
        try:
            parts = v.split("-")
            return (int(parts[0]), int(parts[1]))
        except (IndexError, ValueError):
            return None
    return None


def load_ajustes_manuales(excel_path: str | None = None) -> list:
    """Devuelve lista de ajustes activos."""
    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[AJUSTES_SHEET]
    ajustes = []
    for row in ws.iter_rows(min_row=2, max_col=7, values_only=True):
        if not row[1]:
            continue
        if row[6] is False:
            continue
        ym = _parse_mes_str(row[1])
        if not ym:
            continue
        try:
            monto = float(row[4] or 0)
        except (TypeError, ValueError):
            continue
        ajustes.append({
            "mes_proyectado": ym,
            "categoria": row[2],
            "cultivo": row[3] or "GENERAL",
            "monto": monto,
            "razon": row[5] or "",
        })
    wb.close()
    return ajustes


def load_expected_ingresos(excel_path: str | None = None) -> list:
    """Lee Cosechas, devuelve ingresos proyectados convertidos a CLP."""
    from config import CASH_FLOW_CONFIG
    usd_clp = CASH_FLOW_CONFIG.get("usd_clp_estimado", 1000)

    excel_path = excel_path or EXCEL_PATH
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[COSECHAS_SHEET]
    ingresos = []
    for row in ws.iter_rows(min_row=2, max_col=16, values_only=True):
        if not row[0]:
            continue
        estado = row[11]
        if estado == "recibido":
            monto_real = row[13]
            fecha = row[12]
            moneda = (row[14] or "CLP").upper()
            try:
                m_val = float(monto_real or 0)
            except (TypeError, ValueError):
                m_val = 0
            if m_val <= 0:
                continue
            monto_clp = m_val if moneda == "CLP" else m_val * usd_clp
            ym = _to_year_month(fecha)
        else:
            monto_usd = row[9]
            try:
                m_val = float(monto_usd or 0)
            except (TypeError, ValueError):
                m_val = 0
            if m_val <= 0:
                continue
            monto_clp = m_val * usd_clp
            ym = _to_year_month(row[8])
        if not ym:
            continue
        ingresos.append({
            "year": ym[0], "month": ym[1],
            "cultivo": row[1] or "GENERAL",
            "exportadora": row[3] or "",
            "tipo_cuota": row[10] or "",
            "estado": estado or "esperado",
            "monto_clp": float(monto_clp),
        })
    wb.close()
    return ingresos


def compute_factor_hc(hc: dict, cultivo: str, base_year: int, target_year: int) -> float:
    """Factor de escalamiento por hectareas."""
    if base_year == target_year:
        return 1.0
    if base_year not in hc or target_year not in hc:
        return 1.0

    if cultivo.upper() == "GENERAL":
        base = sum(hc[base_year].values())
        target = sum(hc[target_year].values())
    else:
        base = hc[base_year].get(cultivo.upper(), 0)
        target = hc[target_year].get(cultivo.upper(), 0)

    if base <= 0:
        return 1.0
    return target / base


def compute_egresos_proyectados(historicos: dict, ajustes: list,
                                  hc: dict, base_year: int,
                                  target_year: int) -> dict:
    """Proyecta egresos del target_year escalando base_year + sumando ajustes."""
    proj: dict = defaultdict(float)

    for (y, m, cat, cul), monto in historicos.items():
        if y != base_year:
            continue
        factor = compute_factor_hc(hc, cul, base_year, target_year)
        proj[(target_year, m, cat, cul)] += monto * factor

    for a in ajustes:
        ym = a["mes_proyectado"]
        if ym[0] != target_year:
            continue
        key = (target_year, ym[1], a["categoria"], a["cultivo"])
        proj[key] += a["monto"]

    return dict(proj)


def compute_saldo_mensual(saldo_inicial: float, ingresos: list,
                            egresos: dict, months: list) -> dict:
    """Running balance mes a mes."""
    ing_mes: dict = defaultdict(float)
    for i in ingresos:
        ing_mes[(i["year"], i["month"])] += i["monto_clp"]

    eg_mes: dict = defaultdict(float)
    for (y, m, _cat, _cul), monto in egresos.items():
        eg_mes[(y, m)] += monto

    result = {}
    saldo = saldo_inicial
    for ym in months:
        ing = ing_mes.get(ym, 0)
        eg = eg_mes.get(ym, 0)
        saldo_cierre = saldo + ing - eg
        result[ym] = {
            "saldo_inicio": saldo,
            "ingresos": ing,
            "egresos": eg,
            "saldo_cierre": saldo_cierre,
        }
        saldo = saldo_cierre
    return result
