from datetime import date
from modules.cash_flow.matcher import match_score


def _factura(total=1000000, fecha=date(2025, 9, 1), proveedor="COPEVAL", nro="123"):
    return {"fila": 10, "total": total, "fecha_emision": fecha,
            "proveedor": proveedor, "nro_factura": nro}


def _mov(cargo=1000000, fecha=date(2025, 9, 5), descripcion="PAGO COPEVAL", ref=""):
    return {"fila": 5, "cargo": cargo, "abono": 0,
            "fecha": fecha, "descripcion": descripcion, "referencia": ref}


def test_score_exact_amount_and_close_date():
    s = match_score(_factura(), _mov())
    assert s >= 150


def test_score_amount_diff_lowers_score():
    s = match_score(_factura(), _mov(cargo=900000))
    assert s < 100


def test_score_provider_in_description_bonus():
    base = match_score(_factura(proveedor="ZZZZ"), _mov(descripcion="PAGO ZZZZ"))
    miss = match_score(_factura(proveedor="ZZZZ"), _mov(descripcion="PAGO XYZ"))
    assert base > miss


def test_score_nro_factura_in_reference_strong_match():
    s = match_score(_factura(nro="555"), _mov(cargo=1234567, ref="N FAC 555"))
    assert s >= 100


def test_score_abono_only_no_match():
    f = _factura()
    m = {"fila": 5, "cargo": 0, "abono": 1000000,
         "fecha": date(2025, 9, 5), "descripcion": "X", "referencia": ""}
    assert match_score(f, m) == 0
