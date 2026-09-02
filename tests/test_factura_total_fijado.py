# -*- coding: utf-8 -*-
"""Si el dueño fija el Total Factura a mano, ese valor NO se vuelve a mover.

BUG REPORTADO el 1-sep-2026, con una factura de Administradora de Ventas al
Detalle cuyo impuesto específico estaba mal:

    "fijo el total de la factura, después cambio el monto del total neto y al
     hacer el 2do cambio, de nuevo me cambia el total de la factura"

La causa está al final de `handle_text_edit_factura`: cada vez que se edita un
campo numérico se recalcula

    Total Factura = suma de los Monto / TOTAL

y como `Total Factura` no está en `NUMERICOS`, lo que el dueño fijó a mano se
pisa en la edición siguiente.

Lo más llamativo es que la marca YA EXISTÍA: la rama `_total_factura` escribe
`context.user_data["total_override"]`... y nadie la leía nunca. Un grep sobre
todo el repo devolvía esa única línea.

Lo que se espera ahora: con el total fijado, editar el neto reparte los
valores HACIA ADENTRO (sumatoria inversa) para cuadrar contra el total, en vez
de recalcular el total. Solo volver a tocar el botón de total lo suelta.
"""
import asyncio
import types

import pytest

from handlers import facturas


class _Msg:
    def __init__(self, texto):
        self.text = texto
        self.respuestas = []

    async def reply_text(self, texto, **kw):
        self.respuestas.append(texto)


def _update(texto):
    return types.SimpleNamespace(message=_Msg(texto))


def _ctx(items, campo, **extra):
    ud = {"pending_items": items, "editing_field": campo,
          "editing_item_idx": None}
    ud.update(extra)
    return types.SimpleNamespace(user_data=ud)


def _editar(items, campo, texto, **extra):
    ctx = _ctx(items, campo, **extra)
    asyncio.run(facturas.handle_text_edit_factura(_update(texto), ctx))
    return ctx


def _factura():
    """Un ítem simple: neto 10.000, IVA 19% -> total 11.900."""
    return [{"Documento": "Factura Electronica", "Valor unitario": 10000,
             "Cantidad": 1, "Monto / TOTAL": 11900, "Total Factura": 11900,
             "Impuesto Especifico": 0}]


# ── El caso reportado ──────────────────────────────────────────────────────

def test_fijar_el_total_y_despues_editar_el_neto_NO_mueve_el_total():
    """El bug tal cual lo reportó el dueño."""
    items = _factura()
    ctx = _editar(items, "_total_factura", "20000")
    assert items[0]["Total Factura"] == 20000

    # segunda edición: cambia el neto
    _editar(items, "TOTAL NETO", "15000",
            total_override=ctx.user_data.get("total_override"))
    assert items[0]["Total Factura"] == 20000, \
        "el total fijado a mano se volvió a recalcular"


def test_con_el_total_fijo_editar_el_neto_hace_la_sumatoria_INVERSA():
    """Los otros valores se acomodan al total, no al revés."""
    items = _factura()
    _editar(items, "TOTAL NETO", "15000", total_override=20000)
    # el Monto/TOTAL del ítem tiene que sumar el total fijado
    assert round(sum(float(i["Monto / TOTAL"]) for i in items)) == 20000


def test_editar_el_valor_unitario_tampoco_mueve_el_total_fijo():
    items = _factura()
    _editar(items, "Valor unitario", "7000", total_override=20000)
    assert items[0]["Total Factura"] == 20000


def test_editar_la_cantidad_tampoco_lo_mueve():
    items = _factura()
    _editar(items, "Cantidad", "3", total_override=20000)
    assert items[0]["Total Factura"] == 20000


# ── Sin fijar, el comportamiento de siempre ────────────────────────────────

def test_sin_fijar_el_total_se_sigue_recalculando():
    """No romper el flujo normal: si nadie lo fijó, se calcula solo."""
    items = _factura()
    _editar(items, "TOTAL NETO", "20000")          # sin total_override
    assert items[0]["Total Factura"] == round(20000 * 1.19)


def test_volver_a_fijar_el_total_lo_suelta_y_lo_reemplaza():
    """El botón de total es la única forma de cambiarlo."""
    items = _factura()
    ctx = _editar(items, "_total_factura", "20000")
    ctx2 = _editar(items, "_total_factura", "30000",
                   total_override=ctx.user_data.get("total_override"))
    assert items[0]["Total Factura"] == 30000
    assert ctx2.user_data["total_override"] == 30000


def test_al_fijar_el_total_queda_la_marca():
    items = _factura()
    ctx = _editar(items, "_total_factura", "20000")
    assert ctx.user_data.get("total_override") == 20000


# ── Varios ítems ───────────────────────────────────────────────────────────

def test_con_varios_items_el_total_fijo_manda_igual():
    items = [
        {"Documento": "Factura Electronica", "Valor unitario": 10000,
         "Cantidad": 1, "Monto / TOTAL": 11900, "Total Factura": 23800,
         "Impuesto Especifico": 0},
        {"Documento": "Factura Electronica", "Valor unitario": 10000,
         "Cantidad": 1, "Monto / TOTAL": 11900, "Total Factura": 23800,
         "Impuesto Especifico": 0},
    ]
    _editar(items, "TOTAL NETO", "5000", total_override=23800)
    assert all(i["Total Factura"] == 23800 for i in items)
