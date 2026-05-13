from modules.cash_flow.categorizer_cache import (
    CategorizerCache,
    make_cache_key,
)


def test_make_cache_key_normalizes_whitespace():
    k1 = make_cache_key("Agrosuper  S.A.", "Urea N46")
    k2 = make_cache_key("agrosuper s.a.", "urea n46")
    assert k1 == k2


def test_make_cache_key_strips_punctuation():
    k1 = make_cache_key("Agrosuper, S.A.!", "Urea")
    k2 = make_cache_key("Agrosuper SA", "Urea")
    assert k1 == k2


def test_cache_get_miss_returns_none(tmp_path):
    cache = CategorizerCache(str(tmp_path / "cache.json"))
    assert cache.get("Agrosuper", "Urea") is None


def test_cache_set_then_get(tmp_path):
    cache = CategorizerCache(str(tmp_path / "cache.json"))
    cache.set("Agrosuper", "Urea", {
        "categoria": "Fertilizantes", "cultivo": "NOGALES",
        "confianza": 0.95, "razon": "test"})
    hit = cache.get("Agrosuper", "Urea")
    assert hit["categoria"] == "Fertilizantes"
    assert hit["confianza"] == 0.95


def test_cache_persists_across_instances(tmp_path):
    path = str(tmp_path / "cache.json")
    c1 = CategorizerCache(path)
    c1.set("X", "Y", {"categoria": "Riego", "cultivo": "GENERAL",
                       "confianza": 0.8, "razon": ""})
    c2 = CategorizerCache(path)
    assert c2.get("X", "Y")["categoria"] == "Riego"


def test_cache_low_confidence_not_stored(tmp_path):
    cache = CategorizerCache(str(tmp_path / "cache.json"))
    cache.set("X", "Y", {"categoria": "REVISAR", "cultivo": "GENERAL",
                          "confianza": 0.4, "razon": ""})
    assert cache.get("X", "Y") is None
