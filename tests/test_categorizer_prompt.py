from modules.cash_flow.prompt import build_categorization_prompt, parse_categorization_response


def test_prompt_includes_categorias():
    prompt = build_categorization_prompt(
        proveedor="AGROSUPER S.A.",
        glosa="Fertilizante NPK",
        glosa_ii="",
        monto=1500000,
        fecha="2025-09-15",
    )
    assert "Fertilizantes" in prompt
    assert "Mano de obra planta" in prompt
    assert "NOGALES" in prompt
    assert "GENERAL" in prompt
    assert "AGROSUPER" in prompt
    assert "1500000" in prompt or "1,500,000" in prompt


def test_prompt_includes_all_11_categorias():
    prompt = build_categorization_prompt("X", "Y", "", 1, "2025-01-01")
    for cat in [
        "Mano de obra planta", "Mano de obra temporal", "Fertilizantes",
        "Fitosanitarios", "Combustible", "Maquinaria",
        "Riego", "Servicios profesionales", "Arriendos",
        "Inversion", "Caja chica",
    ]:
        assert cat in prompt


def test_parse_valid_json_response():
    raw = '{"categoria": "Fertilizantes", "cultivo": "NOGALES", "confianza": 0.92, "razon": "NPK para nogales"}'
    result = parse_categorization_response(raw)
    assert result["categoria"] == "Fertilizantes"
    assert result["cultivo"] == "NOGALES"
    assert result["confianza"] == 0.92


def test_parse_response_with_markdown_fence():
    raw = '```json\n{"categoria": "Riego", "cultivo": "GENERAL", "confianza": 0.7}\n```'
    result = parse_categorization_response(raw)
    assert result["categoria"] == "Riego"


def test_parse_invalid_returns_low_confidence():
    result = parse_categorization_response("no soy json")
    assert result["confianza"] == 0.0
    assert result["categoria"] == "REVISAR"
