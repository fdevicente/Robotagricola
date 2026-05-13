from unittest.mock import patch, MagicMock
from modules.cash_flow.categorizer import categorize_raw


def _mock_claude_response(text):
    resp = MagicMock()
    resp.status_code = 200
    resp.json.return_value = {
        "content": [{"text": text}],
        "stop_reason": "end_turn",
        "usage": {"output_tokens": 30},
    }
    return resp


def test_categorize_raw_returns_parsed_dict():
    fake_text = '{"categoria": "Fertilizantes", "cultivo": "NOGALES", "confianza": 0.9, "razon": "NPK"}'
    with patch("modules.cash_flow.categorizer.requests.post",
               return_value=_mock_claude_response(fake_text)) as mock_post:
        result = categorize_raw(
            proveedor="Agrosuper", glosa="NPK 15-15-15",
            glosa_ii="", monto=500000, fecha="2025-09-01",
        )
    assert result["categoria"] == "Fertilizantes"
    assert result["cultivo"] == "NOGALES"
    assert result["confianza"] == 0.9
    assert mock_post.called


def test_categorize_raw_handles_http_error():
    fake = MagicMock()
    fake.status_code = 500
    fake.text = "server error"
    with patch("modules.cash_flow.categorizer.requests.post", return_value=fake):
        result = categorize_raw("X", "Y", "", 0, "2025-01-01")
    assert result["categoria"] == "REVISAR"
    assert result["confianza"] == 0.0


def test_categorize_raw_handles_invalid_json():
    with patch("modules.cash_flow.categorizer.requests.post",
               return_value=_mock_claude_response("no json aqui")):
        result = categorize_raw("X", "Y", "", 0, "2025-01-01")
    assert result["categoria"] == "REVISAR"
