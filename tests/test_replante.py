from modules.cash_flow.replante import afford_check


def test_afford_check_yes_if_enough():
    r = afford_check(cultivo="AVELLANOS", hc=2,
                       saldo_proyectado=50_000_000,
                       saldo_minimo=10_000_000,
                       costo_por_hc=5_000_000)
    assert r["alcanza"] is True
    assert r["disponible"] == 40_000_000
    assert r["costo_total"] == 10_000_000


def test_afford_check_no_if_short():
    r = afford_check(cultivo="AVELLANOS", hc=10,
                       saldo_proyectado=20_000_000,
                       saldo_minimo=5_000_000,
                       costo_por_hc=5_000_000)
    assert r["alcanza"] is False
    assert r["deficit"] == 35_000_000


def test_afford_check_zero_hc():
    r = afford_check(cultivo="AVELLANOS", hc=0,
                       saldo_proyectado=10_000_000,
                       saldo_minimo=5_000_000,
                       costo_por_hc=1_000_000)
    assert r["alcanza"] is True
    assert r["costo_total"] == 0
