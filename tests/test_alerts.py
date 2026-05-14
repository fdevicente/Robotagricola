from datetime import date
from modules.cash_flow.alerts import (
    detect_saldo_bajo, detect_factura_por_vencer,
)


def test_saldo_bajo_below_minimo():
    a = detect_saldo_bajo(saldo_actual=20_000_000, saldo_minimo=36_000_000)
    assert a is not None
    assert "20" in a["mensaje"]


def test_saldo_ok_no_alerta():
    assert detect_saldo_bajo(saldo_actual=100_000_000,
                              saldo_minimo=36_000_000) is None


def test_factura_vence_en_3_dias():
    hoy = date(2026, 5, 14)
    facts = [{
        "fila": 10, "proveedor": "COPEVAL",
        "fecha_vencimiento": date(2026, 5, 17),
        "total": 5_000_000, "nro_factura": "F123",
    }]
    alertas = detect_factura_por_vencer(facts, hoy=hoy, dias=3)
    assert len(alertas) == 1
    assert alertas[0]["fila"] == 10


def test_factura_vencida_no_alerta():
    hoy = date(2026, 5, 14)
    facts = [{"fila": 10, "proveedor": "X",
              "fecha_vencimiento": date(2026, 5, 10),
              "total": 100, "nro_factura": "X"}]
    alertas = detect_factura_por_vencer(facts, hoy=hoy, dias=3)
    assert alertas == []
