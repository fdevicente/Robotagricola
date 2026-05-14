from modules.cash_flow.projector import compute_saldo_mensual


def test_saldo_running_balance():
    egresos = {
        (2026, 5, "Fertilizantes", "NOGALES"): 5_000_000,
        (2026, 5, "Riego", "GENERAL"): 1_000_000,
        (2026, 6, "Combustible", "NOGALES"): 2_000_000,
    }
    ingresos = [
        {"year": 2026, "month": 5, "monto_clp": 200_000_000, "estado": "recibido"},
        {"year": 2026, "month": 6, "monto_clp": 0, "estado": "esperado"},
    ]
    result = compute_saldo_mensual(
        saldo_inicial=100_000_000,
        ingresos=ingresos, egresos=egresos,
        months=[(2026, 5), (2026, 6)],
    )
    assert result[(2026, 5)]["saldo_cierre"] == 294_000_000
    assert result[(2026, 5)]["ingresos"] == 200_000_000
    assert result[(2026, 5)]["egresos"] == 6_000_000
    assert result[(2026, 6)]["saldo_cierre"] == 292_000_000


def test_saldo_zero_ingresos():
    egresos = {(2026, 5, "Riego", "GENERAL"): 10_000_000}
    result = compute_saldo_mensual(
        saldo_inicial=50_000_000, ingresos=[], egresos=egresos,
        months=[(2026, 5)],
    )
    assert result[(2026, 5)]["saldo_cierre"] == 40_000_000


def test_saldo_negativo_se_propaga():
    egresos = {
        (2026, 5, "X", "GENERAL"): 100_000_000,
        (2026, 6, "Y", "GENERAL"): 50_000_000,
    }
    result = compute_saldo_mensual(
        saldo_inicial=10_000_000, ingresos=[], egresos=egresos,
        months=[(2026, 5), (2026, 6)],
    )
    assert result[(2026, 5)]["saldo_cierre"] == -90_000_000
    assert result[(2026, 6)]["saldo_cierre"] == -140_000_000
