from modules.cash_flow.alerts import AlertDedupe


def test_first_fire_returns_true(tmp_path):
    d = AlertDedupe(path=str(tmp_path / "d.json"))
    assert d.should_fire("saldo_bajo", "2026-05") is True


def test_second_same_month_blocks(tmp_path):
    p = str(tmp_path / "d.json")
    d = AlertDedupe(path=p)
    d.should_fire("saldo_bajo", "2026-05")
    d.mark_fired("saldo_bajo", "2026-05")
    d2 = AlertDedupe(path=p)
    assert d2.should_fire("saldo_bajo", "2026-05") is False


def test_different_month_allowed(tmp_path):
    p = str(tmp_path / "d.json")
    d = AlertDedupe(path=p)
    d.mark_fired("saldo_bajo", "2026-05")
    assert d.should_fire("saldo_bajo", "2026-06") is True
