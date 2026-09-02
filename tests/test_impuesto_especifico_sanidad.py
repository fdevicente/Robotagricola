# -*- coding: utf-8 -*-
"""Un impuesto específico que rompe la aritmética de la factura no es real.

BUG REPORTADO el 1-sep-2026: *"las últimas 3 facturas las leyó mal, y dio
entre 2000 o 3000 pesos menos de neto y le sumó ese monto al impuesto
específico"*.

El origen: cuando el impuesto viene en 0, el extractor barre el OCR con
`IMP_PATTERNS` (`ief`, `iev`, `fepp`, `imp.esp`, `i.e.c.`) y SUMA todas las
coincidencias. Los patrones son laxos y pescan falsos positivos. Nada valida
después si el número resultante cuadra con el resto de la factura.

Evidencia medida sobre el Master (109 filas tienen impuesto específico):

    fila 2185  Admin. Ventas 24248754   neto 28.084  imp 11.619  TOTAL 33.420
               -> 28.084 × 1,19 = 33.420 EXACTO. El impuesto es inventado.
    fila 2187  Admin. Ventas 24140536   neto 594.456 imp 1       TOTAL 707.400
               -> un "impuesto" de UN PESO. Ruido de OCR puro.

La regla: si el neto por el IVA ya da el total, el impuesto tiene que ser 0.
El impuesto específico es un cargo REAL de combustible, no el lugar donde se
esconde lo que no cuadra.
"""
import pytest

from processors.extractor import sanear_impuesto_especifico


def _item(neto, imp, total, doc="Factura Electronica", cant=1):
    return {"Documento": doc, "Valor unitario": neto / cant, "Cantidad": cant,
            "Impuesto Especifico": imp, "Total Factura": total,
            "Monto / TOTAL": total}


# ── Los casos reales que lo motivaron ──────────────────────────────────────

def test_el_caso_2185_el_total_ya_cuadra_sin_impuesto():
    """28.084 x 1,19 = 33.420 exacto -> los 11.619 sobran."""
    items = [_item(28084, 11619, 33420)]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 0


def test_el_caso_2187_un_impuesto_de_un_peso_es_ruido():
    items = [_item(594456, 1, 707400)]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 0


# ── Lo que NO se puede romper ──────────────────────────────────────────────

def test_un_impuesto_de_combustible_DE_VERDAD_se_respeta():
    """Copec: neto 100.000, IEF 7.660, total = 100.000*1,19 + 7.660."""
    total = round(100000 * 1.19) + 7660
    items = [_item(100000, 7660, total)]
    items[0]["Detalle / Glosa"] = "petroleo diesel"
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 7660


def test_no_toca_la_factura_sin_impuesto():
    items = [_item(10000, 0, 11900)]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 0


def test_una_diferencia_de_redondeo_no_dispara_la_limpieza():
    """29.277 x 1,19 = 34.839,6 -> 34.840. Un peso de diferencia es redondeo."""
    items = [_item(29277, 0, 34840)]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 0


def test_la_factura_exenta_no_lleva_iva_y_se_evalua_igual():
    """Exenta: total = neto, sin 1,19. Un impuesto que sobra igual se saca."""
    items = [_item(50000, 4000, 50000, doc="Factura Exenta Electronica")]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 0


def test_si_el_impuesto_es_lo_que_hace_cuadrar_se_deja():
    """Albino Fuentealba 607: 300.000 + 53.982 = 353.982, sin IVA."""
    items = [_item(300000, 53982, 353982, doc="Factura Exenta Electronica")]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 53982


def test_sin_total_factura_no_puede_juzgar_y_no_toca_nada():
    """Sin el total no hay contra qué comparar: no inventar."""
    items = [_item(10000, 5000, 0)]
    sanear_impuesto_especifico(items)
    assert items[0]["Impuesto Especifico"] == 5000


def test_varios_items_se_evaluan_como_una_sola_factura():
    """El total es de la factura, no de cada linea."""
    items = [_item(10000, 3000, 23800), _item(10000, 0, 23800)]
    sanear_impuesto_especifico(items)
    assert sum(float(i["Impuesto Especifico"]) for i in items) == 0


def test_el_pipeline_lo_llama_de_verdad():
    """Sin esto seria codigo muerto: el barrido de OCR seguiria inventando."""
    import inspect

    from processors import extractor
    fuente = inspect.getsource(extractor.process_file) if hasattr(
        extractor, "process_file") else inspect.getsource(extractor)
    assert "sanear_impuesto_especifico(items)" in fuente, \
        "el saneo no esta enganchado al pipeline de extraccion"
