from modules.cash_flow.cosecha_wizard import CosechaWizard


def test_inicio_pide_kg_totales():
    w = CosechaWizard(cultivo="NOGALES")
    assert w.estado == "esperando_kg_totales"
    assert "kg" in w.prompt.lower()


def test_kg_totales_avanza_a_exportadoras():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("240000")
    assert w.estado == "esperando_exportadoras"
    assert w.data["kg_total"] == 240000


def test_exportadoras_parsea_lista():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("240000")
    w.responder("Valbifrut 140000, Pacific Nuts 100000")
    assert len(w.data["exportadoras"]) == 2
    assert w.data["exportadoras"][0]["nombre"] == "Valbifrut"
    assert w.data["exportadoras"][0]["kg"] == 140000


def test_kg_no_numerico_se_queda():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("abc")
    assert w.estado == "esperando_kg_totales"
    assert "error" in w.prompt.lower() or "numero" in w.prompt.lower()


def test_resumen_al_terminar():
    w = CosechaWizard(cultivo="NOGALES")
    w.responder("240000")
    w.responder("Valbifrut 140000")
    w.responder("1.8")
    w.responder("1")
    w.responder("2026-06-15 252000")
    w.responder("no")
    assert w.estado == "resumen"
    assert w.data["exportadoras"][0]["precio_usd_kg"] == 1.8
