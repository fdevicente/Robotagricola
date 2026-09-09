# -*- coding: utf-8 -*-
"""El resumen de facturas del dashboard cuenta FACTURAS, no lineas, y no
muestra como deuda lo que el dueño decidio no pagar.

Medido contra el Master real el 9-sep-2026: la tarjeta decia "19 vencidas + 14
por pagar" cuando habia 18 facturas por pagar. Dos errores sumados:

  1. Contaba LINEAS. Una factura con tres items se contaba tres veces.
  2. No descontaba las marcadas `NN-no-pagar`: 13 de esas 33 lineas eran 10
     facturas viejas (2023-2025, $3.425.211) que el dueño decidio no pagar y
     el dashboard las seguia mostrando como deuda.
"""
import os
import sys

from openpyxl import Workbook

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "src"))

from dashboard_data import get_facturas_summary   # noqa: E402

ENC = ["Fecha Emision / Fecha", "Fecha Vencimiento", "Fecha Pago",
       "Nombre Factura / Proveedor", "Rut", "Documento",
       "Numero Factura / Nro Documento", "Detalle", "Glosa II",
       "Valor unitario", "Cantidad", "TOTAL NETO", "IVA", "ESPECIFICO",
       "Total por Item", "TOTAL FACTURA", "Categoria", "Cultivo",
       "Confianza", "Categorizado_por", "N° Archivo"]


def _fila(emision, venc, pago, prov, nro, item, nota=""):
    r = [None] * len(ENC)
    r[0], r[1], r[2], r[3] = emision, venc, pago, prov
    r[5], r[6] = "Factura", nro
    r[14], r[15] = item, item
    r[19] = nota
    return r


def _master(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append(ENC)
    # Una factura PAGADA de tres lineas: es UNA factura, no tres.
    for parte in (1000, 2000, 3000):
        ws.append(_fila("2026-05-01", "2026-06-01", "2026-06-05",
                        "COPEVAL", "111", parte))
    # Una VENCIDA de dos lineas.
    ws.append(_fila("2026-06-01", "2026-07-01", None, "LIPIGAS", "222", 500))
    ws.append(_fila("2026-06-01", "2026-07-01", None, "LIPIGAS", "222", 700))
    # Una POR PAGAR, todavia no vence.
    ws.append(_fila("2026-09-01", "2099-01-01", None, "NOGALTEC", "333", 900))
    # Dos lineas que el dueño marco NN: no son deuda.
    ws.append(_fila("2023-03-09", "2023-04-09", None, "CONTRERAS", "94",
                    404600, "NN-no-pagar"))
    ws.append(_fila("2023-03-09", "2023-04-09", None, "CONTRERAS", "94",
                    238000, "NN-no-pagar"))
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_una_factura_de_tres_lineas_se_cuenta_una_vez(tmp_path):
    r = get_facturas_summary(_master(tmp_path))
    assert r["pagadas"] == 1
    assert r["total_facturas"] == 4        # 111, 222, 333 y la NN


def test_las_vencidas_son_facturas_no_lineas(tmp_path):
    r = get_facturas_summary(_master(tmp_path))
    assert r["vencidas"] == 1              # la 222 tiene dos lineas


def test_lo_marcado_NN_no_se_muestra_como_deuda(tmp_path):
    """10 facturas viejas por $3.425.211 figuraban como pendientes."""
    r = get_facturas_summary(_master(tmp_path))
    assert r["vencidas"] == 1
    assert r["por_pagar"] == 1
    assert r["no_se_pagan"] == 1           # se informa aparte, no se esconde


def test_los_estados_suman_el_total(tmp_path):
    """Si no suman, la pantalla se contradice sola."""
    r = get_facturas_summary(_master(tmp_path))
    assert (r["pagadas"] + r["vencidas"] + r["por_pagar"]
            + r["no_se_pagan"]) == r["total_facturas"]


def test_el_monto_total_sigue_sumando_los_items(tmp_path):
    """El total en plata se suma por ITEM: cambiar el conteo no puede moverlo."""
    r = get_facturas_summary(_master(tmp_path))
    assert r["total_monto"] == 1000 + 2000 + 3000 + 500 + 700 + 900 + 404600 + 238000
