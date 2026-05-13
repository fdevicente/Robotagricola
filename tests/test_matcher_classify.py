from modules.cash_flow.matcher import classify_match


def test_no_candidates_returns_none():
    r = classify_match([])
    assert r["status"] == "no_match"


def test_single_strong_candidate_auto():
    r = classify_match([{"fila": 10, "score": 150}])
    assert r["status"] == "auto"
    assert r["fila"] == 10


def test_two_strong_candidates_ambiguo():
    r = classify_match([{"fila": 10, "score": 150}, {"fila": 11, "score": 140}])
    assert r["status"] == "ambiguo"


def test_strong_plus_weak_picks_strong():
    r = classify_match([{"fila": 10, "score": 150}, {"fila": 11, "score": 35}])
    assert r["status"] == "auto"
    assert r["fila"] == 10


def test_only_weak_candidate_ambiguo():
    r = classify_match([{"fila": 10, "score": 35}])
    assert r["status"] == "ambiguo"
