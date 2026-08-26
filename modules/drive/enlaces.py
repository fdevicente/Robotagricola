# -*- coding: utf-8 -*-
"""Escribe en la hoja Facturas el enlace del documento en Drive."""
import logging

logger = logging.getLogger(__name__)

COL_NUMERO, COL_DRIVE = 7, 22
URL = "https://drive.google.com/file/d/%s/view"


def guardar_enlace(excel_path: str, numero_factura: str, file_id: str) -> bool:
    """Pone el enlace en TODAS las líneas de esa factura. No pisa lo existente."""
    from openpyxl import load_workbook
    wb = load_workbook(excel_path)
    try:
        ws = wb["Facturas"]
        objetivo = _normalizar(numero_factura)
        tocadas = 0
        for f in range(2, ws.max_row + 1):
            if _normalizar(ws.cell(f, COL_NUMERO).value) != objetivo:
                continue
            if ws.cell(f, COL_DRIVE).value:
                continue
            ws.cell(f, COL_DRIVE).value = URL % file_id
            tocadas += 1
        if tocadas:
            wb.save(excel_path)          # ruta EXPLÍCITA siempre
        return tocadas > 0
    finally:
        wb.close()


def _normalizar(valor) -> str:
    """'2777', 2777 y 2777.0 son el mismo número de factura."""
    s = str(valor if valor is not None else "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s
