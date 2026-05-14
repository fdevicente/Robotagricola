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
