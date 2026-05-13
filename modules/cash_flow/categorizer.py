"""Cliente Claude para categorizar facturas y cargos bancarios.

Usa requests directo (mismo patron que processors/extractor.py).
"""
import logging
import requests

from config import ANTHROPIC_API_KEY
from modules.cash_flow.prompt import (
    build_categorization_prompt,
    parse_categorization_response,
)

logger = logging.getLogger(__name__)

CLAUDE_URL = "https://api.anthropic.com/v1/messages"
CLAUDE_MODELS = ["claude-haiku-4-5-20251001", "claude-sonnet-4-6"]
MAX_TOKENS = 200
TIMEOUT_SEC = 30


def categorize_raw(proveedor: str, glosa: str, glosa_ii: str,
                    monto: float, fecha: str) -> dict:
    """Llama a Claude directo, sin cache. Devuelve dict con categoria/cultivo/confianza/razon."""
    prompt = build_categorization_prompt(proveedor, glosa, glosa_ii, monto, fecha)

    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload_base = {
        "max_tokens": MAX_TOKENS,
        "messages": [{"role": "user", "content": prompt}],
    }

    for model in CLAUDE_MODELS:
        payload = {**payload_base, "model": model}
        try:
            resp = requests.post(CLAUDE_URL, headers=headers,
                                  json=payload, timeout=TIMEOUT_SEC)
        except requests.RequestException as e:
            logger.warning(f"Claude {model} excepcion: {e}")
            continue

        if resp.status_code != 200:
            logger.warning(f"Claude {model} HTTP {resp.status_code}: {resp.text[:200]}")
            continue

        try:
            data = resp.json()
            raw_text = data["content"][0]["text"]
        except (KeyError, IndexError, ValueError) as e:
            logger.warning(f"Claude {model} respuesta inesperada: {e}")
            continue

        return parse_categorization_response(raw_text)

    logger.error(f"Todos los modelos Claude fallaron para {proveedor[:40]}")
    return {
        "categoria": "REVISAR",
        "cultivo": "GENERAL",
        "confianza": 0.0,
        "razon": "Claude API fallo",
    }
