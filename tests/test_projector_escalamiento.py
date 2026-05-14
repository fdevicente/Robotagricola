from modules.cash_flow.projector import compute_factor_hc


HC = {
    2024: {"NOGALES": 65, "CEREZOS": 1.8, "AVELLANOS": 0},
    2025: {"NOGALES": 54, "CEREZOS": 3.8, "AVELLANOS": 11.5},
    2026: {"NOGALES": 43, "CEREZOS": 3.8, "AVELLANOS": 26.5},
}


def test_factor_same_year_is_one():
    assert compute_factor_hc(HC, "NOGALES", base_year=2025, target_year=2025) == 1.0


def test_factor_nogales_2025_to_2026_smaller():
    f = compute_factor_hc(HC, "NOGALES", base_year=2025, target_year=2026)
    assert abs(f - (43 / 54)) < 0.001


def test_factor_avellanos_2025_to_2026_growth():
    f = compute_factor_hc(HC, "AVELLANOS", base_year=2025, target_year=2026)
    assert abs(f - (26.5 / 11.5)) < 0.001


def test_factor_general_uses_total():
    f = compute_factor_hc(HC, "GENERAL", base_year=2025, target_year=2026)
    total_2025 = 54 + 3.8 + 11.5
    total_2026 = 43 + 3.8 + 26.5
    assert abs(f - (total_2026 / total_2025)) < 0.001


def test_factor_base_zero_returns_one():
    hc = {2024: {"AVELLANOS": 0}, 2025: {"AVELLANOS": 10}}
    assert compute_factor_hc(hc, "AVELLANOS", base_year=2024, target_year=2025) == 1.0
