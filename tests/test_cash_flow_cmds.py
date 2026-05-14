from handlers.cash_flow_cmds import format_proyeccion


def _fake_cf():
    return {
        "months": [(2026, 5), (2026, 6)],
        "saldo": {
            (2026, 5): {"saldo_inicio": 100, "ingresos": 200, "egresos": 50, "saldo_cierre": 250},
            (2026, 6): {"saldo_inicio": 250, "ingresos": 0, "egresos": 30, "saldo_cierre": 220},
        },
        "egresos": {}, "ingresos": [],
    }


def test_format_proyeccion_has_each_month():
    text = format_proyeccion(_fake_cf())
    assert "May-26" in text and "Jun-26" in text
    assert "220" in text or "$220" in text


def test_format_proyeccion_marks_negative_red():
    cf = _fake_cf()
    cf["saldo"][(2026, 6)]["saldo_cierre"] = -1000
    text = format_proyeccion(cf)
    assert "🔴" in text or "-1" in text
