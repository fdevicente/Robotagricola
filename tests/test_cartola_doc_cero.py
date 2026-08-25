"""El N° de documento "0" es un relleno del banco, no un identificador.

CASO REAL 2026-08-25: la cartola traía dos compras del mismo día

    25-08-2026,REDCOMPRA FERRETERIA M Y G TAL,...,0,-23300,,67647904
    25-08-2026,REDCOMPRA FERRETERIA M Y G TAL,...,0,-29700,,67671204

y el importador las descartó como duplicadas. El Master tenía UNA fila vieja
con documento "0" (Facaz, 27-mar-2026), y la dedup por documento hizo match
contra ella: `if m["doc"] and m["doc"] in docs`. Como "0" es una cadena
truthy, los dos movimientos se perdieron en silencio.

El banco pone 0 en las compras con tarjeta (REDCOMPRA). Cualquier movimiento
así queda tragado apenas exista una sola fila con documento 0 en el Master.
"""
from datetime import date

import pytest

from modules.banco_import import _doc_identificador


# ── El documento de relleno no identifica nada ─────────────────────────────

@pytest.mark.parametrize("doc", ["0", "00", "0000", "", "   ", "-", "--", None])
def test_documentos_de_relleno_no_cuentan_como_identificador(doc):
    assert _doc_identificador(doc) == ""


@pytest.mark.parametrize("doc", [
    "5865450093",           # TEF
    "11864885",             # REDCOMPRA con documento real
    "50765765",             # COMEX
    "0053",                 # ceros a la izquierda pero con dígito
    "F195279",
])
def test_documentos_reales_si_identifican(doc):
    assert _doc_identificador(doc) == str(doc).strip()


def test_el_cero_no_se_confunde_con_un_documento_que_lo_contiene():
    assert _doc_identificador("1000") == "1000"
    assert _doc_identificador("0100") == "0100"


# ── El caso real, de punta a punta ─────────────────────────────────────────

def _fila(desc, doc, cargo, saldo, dia):
    return {"fecha": date(2026, 8, dia), "desc": desc, "doc": doc,
            "cargo": float(cargo), "abono": 0.0, "saldo": float(saldo)}


def test_dos_compras_con_documento_cero_el_mismo_dia_son_dos_movimientos():
    """No se pueden colapsar entre sí: montos distintos, mismo día."""
    from modules.banco_import import _clave
    a = _fila("REDCOMPRA FERRETERIA M Y G TAL", "0", 23300, 67647904, 25)
    b = _fila("REDCOMPRA FERRETERIA M Y G TAL", "0", 29700, 67671204, 25)
    assert _clave(a["fecha"], a["cargo"], a["abono"]) != \
           _clave(b["fecha"], b["cargo"], b["abono"])


def test_una_fila_vieja_con_doc_cero_no_puede_tragarse_las_nuevas():
    """El Master tenía Facaz con doc 0 el 27-mar; no debe deduplicar nada de agosto."""
    vieja = _doc_identificador("0")
    nueva = _doc_identificador("0")
    # ambas quedan sin identificador -> la dedup cae al multiconjunto
    # (fecha, cargo, abono), que SÍ las distingue
    assert vieja == "" and nueva == ""
