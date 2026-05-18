import json
import pytest
from unittest.mock import patch


@pytest.fixture
def client():
    import sys, os
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
    from dashboard import app
    return app.test_client()


def test_cash_flow_endpoint_returns_json(client):
    fake_cf = {
        "months": [(2026, 5), (2026, 6)],
        "saldo": {
            (2026, 5): {"saldo_inicio": 100, "ingresos": 200,
                          "egresos": 50, "saldo_cierre": 250},
            (2026, 6): {"saldo_inicio": 250, "ingresos": 0,
                          "egresos": 30, "saldo_cierre": 220},
        },
        "egresos": {(2026, 5, "Riego", "GENERAL"): 50},
        "ingresos": [{"year": 2026, "month": 5, "monto_clp": 200,
                       "exportadora": "Valbifrut", "estado": "recibido"}],
    }
    with patch("dashboard.get_cash_flow", return_value=fake_cf):
        resp = client.get("/api/cash-flow")
    assert resp.status_code == 200
    data = resp.get_json()
    assert "labels" in data
    assert len(data["labels"]) == 2
    assert "saldo_cierre" in data
    assert data["saldo_cierre"][0] == 250


def test_replante_endpoint_post(client):
    resp = client.post("/api/cash-flow/replante",
                        json={"cultivo": "AVELLANOS", "hc": 2,
                              "saldo_proyectado": 50_000_000,
                              "saldo_minimo": 10_000_000,
                              "costo_por_hc": 5_000_000})
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["alcanza"] is True
    assert data["costo_total"] == 10_000_000
