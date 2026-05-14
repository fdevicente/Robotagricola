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
