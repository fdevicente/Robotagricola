"""Cliente Claude para categorizar facturas y cargos bancarios.

Usa requests directo (mismo patron que processors/extractor.py).
"""
import logging
import os
import requests
from openpyxl import load_workbook

from config import ANTHROPIC_API_KEY, EXCEL_PATH, DROPBOX_BACKUP_PATH
from excel_manager import (
    SHEET_NAME, _save_wb,
    COL_CATEGORIA, COL_CULTIVO, COL_CONFIANZA, COL_CATEGORIZADO_POR,
)
from modules.cash_flow.prompt import (
    build_categorization_prompt,
    parse_categorization_response,
)
from modules.cash_flow.categorizer_cache import CategorizerCache

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


DEFAULT_CACHE_PATH = os.path.join(DROPBOX_BACKUP_PATH, "categorizer_cache.json")
CONFIANZA_REVISAR = 0.85


def _get_cache(cache_path=None) -> CategorizerCache:
    return CategorizerCache(cache_path or DEFAULT_CACHE_PATH)


def _read_invoice_row(ws, row: int) -> dict:
    return {
        "fecha": str(ws.cell(row, 1).value or ""),
        "proveedor": str(ws.cell(row, 4).value or ""),
        "documento": str(ws.cell(row, 6).value or ""),
        "glosa": str(ws.cell(row, 8).value or ""),
        "glosa_ii": str(ws.cell(row, 9).value or ""),
        "monto": float(ws.cell(row, 15).value or 0),
    }


def categorize_invoice(row: int, excel_path=None, cache_path=None) -> dict:
    """Categoriza la fila `row` de Facturas y escribe cols Q-T en el Master."""
    excel_path = excel_path or EXCEL_PATH
    cache = _get_cache(cache_path)

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]
    data = _read_invoice_row(ws, row)
    wb.close()

    cached = cache.get(data["proveedor"], data["glosa"])
    if cached:
        result = cached
        source = "cache"
    else:
        result = categorize_raw(
            proveedor=data["proveedor"], glosa=data["glosa"],
            glosa_ii=data["glosa_ii"], monto=data["monto"],
            fecha=data["fecha"],
        )
        cache.set(data["proveedor"], data["glosa"], result)
        source = "claude"

    cat_to_write = result["categoria"]
    if result["confianza"] < CONFIANZA_REVISAR:
        cat_to_write = "REVISAR"

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]
    ws.cell(row, COL_CATEGORIA, cat_to_write)
    ws.cell(row, COL_CULTIVO, result["cultivo"])
    ws.cell(row, COL_CONFIANZA, result["confianza"])
    ws.cell(row, COL_CATEGORIZADO_POR, source)
    _save_wb(wb, excel_path)
    wb.close()

    return result
