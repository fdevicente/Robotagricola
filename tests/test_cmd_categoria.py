from handlers.cash_flow_cmds import format_categoria


def test_format_categoria_groups_by_month():
    egresos = {
        (2026, 5, "Fertilizantes", "NOGALES"): 5_000_000,
        (2026, 5, "Fertilizantes", "GENERAL"): 1_000_000,
        (2026, 6, "Fertilizantes", "NOGALES"): 3_000_000,
        (2026, 5, "Riego", "GENERAL"): 999,
    }
    text = format_categoria("Fertilizantes", egresos,
                              months=[(2026, 5), (2026, 6)])
    assert "Fertilizantes" in text
    assert "6,000,000" in text or "6.000.000" in text
    assert "3,000,000" in text or "3.000.000" in text
    assert "Riego" not in text
