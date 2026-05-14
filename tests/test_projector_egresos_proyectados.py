from modules.cash_flow.projector import compute_egresos_proyectados


def test_egresos_scaled_by_hc():
    historicos = {(2025, 5, "Fertilizantes", "NOGALES"): 10_000_000}
    hc = {2025: {"NOGALES": 54}, 2026: {"NOGALES": 43}}
    proj = compute_egresos_proyectados(
        historicos=historicos, ajustes=[], hc=hc,
        base_year=2025, target_year=2026,
    )
    val = proj[(2026, 5, "Fertilizantes", "NOGALES")]
    assert abs(val - 10_000_000 * 43 / 54) < 100


def test_ajustes_added_to_projection():
    historicos = {}
    ajustes = [{
        "mes_proyectado": (2026, 7), "categoria": "Riego",
        "cultivo": "GENERAL", "monto": 5_000_000, "razon": "Bomba",
    }]
    hc = {2025: {"NOGALES": 54}, 2026: {"NOGALES": 43}}
    proj = compute_egresos_proyectados(
        historicos=historicos, ajustes=ajustes, hc=hc,
        base_year=2025, target_year=2026,
    )
    assert proj[(2026, 7, "Riego", "GENERAL")] == 5_000_000


def test_ajustes_can_be_negative():
    historicos = {(2025, 5, "Fertilizantes", "NOGALES"): 10_000_000}
    ajustes = [{
        "mes_proyectado": (2026, 5), "categoria": "Fertilizantes",
        "cultivo": "NOGALES", "monto": -2_000_000, "razon": "Sin compra",
    }]
    hc = {2025: {"NOGALES": 54}, 2026: {"NOGALES": 54}}
    proj = compute_egresos_proyectados(
        historicos=historicos, ajustes=ajustes, hc=hc,
        base_year=2025, target_year=2026,
    )
    assert proj[(2026, 5, "Fertilizantes", "NOGALES")] == 8_000_000
