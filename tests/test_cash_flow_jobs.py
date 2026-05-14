from handlers.cash_flow_jobs import format_resumen_semanal


def test_resumen_includes_saldo_y_alertas():
    cf = {
        "months": [(2026, 5)],
        "saldo": {(2026, 5): {"saldo_inicio": 100, "ingresos": 200,
                                "egresos": 50, "saldo_cierre": 250}},
        "egresos": {}, "ingresos": [],
    }
    alertas = [{"mensaje": "🟡 Vence en 2d: COPEVAL F123"}]
    text = format_resumen_semanal(cf, alertas)
    assert "250" in text or "$250" in text
    assert "COPEVAL" in text


def test_resumen_sin_alertas_dice_ok():
    cf = {
        "months": [(2026, 5)],
        "saldo": {(2026, 5): {"saldo_inicio": 100, "ingresos": 0,
                                "egresos": 50, "saldo_cierre": 50}},
        "egresos": {}, "ingresos": [],
    }
    text = format_resumen_semanal(cf, [])
    assert "sin alertas" in text.lower() or "ok" in text.lower()
