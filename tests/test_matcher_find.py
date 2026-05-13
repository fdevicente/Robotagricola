from datetime import date
from modules.cash_flow.matcher import find_matches


def _f(fila, total=1000000, prov="X", fecha=date(2025, 9, 1), nro="1"):
    return {"fila": fila, "total": total, "proveedor": prov,
            "fecha_emision": fecha, "nro_factura": nro}


def _m(cargo=1000000, prov="X", fecha=date(2025, 9, 3), ref=""):
    return {"fila": 7, "cargo": cargo, "abono": 0,
            "descripcion": f"PAGO {prov}", "fecha": fecha, "referencia": ref}


def test_find_single_match():
    facturas = [_f(10), _f(11, total=999999), _f(12, total=1, prov="Z")]
    candidates = find_matches(_m(), facturas)
    assert candidates[0]["fila"] == 10


def test_find_returns_empty_below_threshold():
    facturas = [_f(10, total=500000, prov="Z", fecha=date(2024, 1, 1))]
    assert find_matches(_m(), facturas) == []


def test_find_returns_sorted_desc():
    facturas = [_f(10, prov="OTROS"), _f(11, prov="COPEVAL")]
    cands = find_matches(_m(prov="COPEVAL"), facturas)
    assert cands[0]["fila"] == 11
