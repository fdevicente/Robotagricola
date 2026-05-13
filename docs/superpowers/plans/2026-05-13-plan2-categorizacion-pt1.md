# Plan 2: Categorización — Parte 1/3

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans

**Goal:** Construir el sistema de categorización automática con Claude AI para clasificar facturas y movimientos bancarios en (categoría × cultivo × confianza), con cache local para minimizar costos.

**Architecture:** Módulo nuevo `modules/cash_flow/` con `categorizer.py` (cliente Claude + prompt) y `categorizer_cache.py` (cache JSON proveedor+glosa→categoría). Reutiliza el patrón HTTP del extractor existente (`processors/extractor.py`) en vez de instalar SDK nuevo.

**Tech Stack:** Python 3.11, requests (Claude API), openpyxl, json (cache), config existente (ANTHROPIC_API_KEY)

---

### Task 1: Estructura del paquete cash_flow + prompt template

**Files:**
- Create: `Robot/modules/__init__.py`
- Create: `Robot/modules/cash_flow/__init__.py`
- Create: `Robot/modules/cash_flow/prompt.py`
- Test: `Robot/tests/test_categorizer_prompt.py`

- [ ] **Step 1: Create package __init__.py files**

```python
# modules/__init__.py
```

```python
# modules/cash_flow/__init__.py
```

- [ ] **Step 2: Write test for prompt builder**

```python
# tests/test_categorizer_prompt.py
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
```

- [ ] **Step 3: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_categorizer_prompt.py -v`
Expected: ImportError — module not found

- [ ] **Step 4: Implement prompt.py**

```python
# modules/cash_flow/prompt.py
"""Prompt builder y parser para categorizacion via Claude."""
import json
import re
from excel_manager import CATEGORIAS, CULTIVOS


SYSTEM_INSTRUCTIONS = """Eres un asistente que categoriza facturas y cargos bancarios
de una explotacion agricola en Chile (nogales, cerezos, avellanos).

Devuelves SIEMPRE un JSON valido con las claves:
- categoria: una de la lista
- cultivo: NOGALES / CEREZOS / AVELLANOS / GENERAL
- confianza: float 0.0-1.0
- razon: explicacion breve (max 80 chars)

Si no sabes, usa confianza < 0.6. Nunca inventes categorias fuera de la lista."""


def build_categorization_prompt(proveedor: str, glosa: str, glosa_ii: str,
                                  monto: float, fecha: str) -> str:
    """Arma el prompt para clasificar una factura/cargo."""
    cats_list = "\n".join(f"- {c}" for c in CATEGORIAS)
    cultivos_list = " / ".join(CULTIVOS)
    return f"""{SYSTEM_INSTRUCTIONS}

Categorias validas:
{cats_list}

Cultivos validos: {cultivos_list}

Datos del documento:
- Proveedor: {proveedor or "(sin nombre)"}
- Glosa: {glosa or "(sin glosa)"}
- Glosa II: {glosa_ii or ""}
- Monto: {monto}
- Fecha: {fecha}

Responde SOLO con el JSON, sin texto adicional."""


def parse_categorization_response(raw: str) -> dict:
    """Extrae el JSON de la respuesta de Claude. Devuelve dict con keys estandar."""
    if not raw:
        return _low_confidence("respuesta vacia")

    # Remover markdown fences si los hay
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    # Extraer primer bloque {...}
    match = re.search(r"\{.*\}", cleaned, re.DOTALL)
    if not match:
        return _low_confidence("sin JSON")

    try:
        data = json.loads(match.group(0))
    except json.JSONDecodeError:
        return _low_confidence("JSON invalido")

    categoria = data.get("categoria") or "REVISAR"
    cultivo = data.get("cultivo") or "GENERAL"
    confianza = float(data.get("confianza") or 0.0)
    razon = (data.get("razon") or "")[:80]

    # Validar categoria contra lista
    if categoria != "REVISAR" and categoria not in CATEGORIAS:
        return _low_confidence(f"categoria desconocida: {categoria}")

    if cultivo not in CULTIVOS:
        cultivo = "GENERAL"

    return {
        "categoria": categoria,
        "cultivo": cultivo,
        "confianza": max(0.0, min(1.0, confianza)),
        "razon": razon,
    }


def _low_confidence(reason: str) -> dict:
    return {
        "categoria": "REVISAR",
        "cultivo": "GENERAL",
        "confianza": 0.0,
        "razon": reason,
    }
```

- [ ] **Step 5: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_categorizer_prompt.py -v`
Expected: 5 passed

- [ ] **Step 6: Commit**

```bash
git add modules/ tests/test_categorizer_prompt.py
git commit -m "feat: add categorization prompt builder and JSON parser"
```

---

### Task 2: Cliente HTTP Claude para categorización

**Files:**
- Create: `Robot/modules/cash_flow/categorizer.py`
- Test: `Robot/tests/test_categorizer_client.py`

- [ ] **Step 1: Write test with mocked HTTP**

```python
# tests/test_categorizer_client.py
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
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_categorizer_client.py -v`
Expected: ImportError

- [ ] **Step 3: Implement categorizer.py**

```python
# modules/cash_flow/categorizer.py
"""Cliente Claude para categorizar facturas y cargos bancarios.

Usa requests directo (mismo patron que processors/extractor.py).
"""
import logging
import time
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
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_categorizer_client.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/categorizer.py tests/test_categorizer_client.py
git commit -m "feat: add Claude HTTP client for categorization"
```

---

### Task 3: Cache JSON local para evitar re-categorizar

**Files:**
- Create: `Robot/modules/cash_flow/categorizer_cache.py`
- Test: `Robot/tests/test_categorizer_cache.py`

- [ ] **Step 1: Write test**

```python
# tests/test_categorizer_cache.py
import os
import pytest
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
```

- [ ] **Step 2: Run test — expect FAIL**

Run: `py -3.11 -m pytest tests/test_categorizer_cache.py -v`

- [ ] **Step 3: Implement categorizer_cache.py**

```python
# modules/cash_flow/categorizer_cache.py
"""Cache JSON simple para resultados de categorizacion.

Clave: (proveedor_normalizado, glosa_normalizada). Solo guarda hits
de alta confianza (>=0.7) para no envenenar el cache.
"""
import json
import logging
import os
import re
import threading

logger = logging.getLogger(__name__)

MIN_CONFIDENCE_TO_CACHE = 0.7


def make_cache_key(proveedor: str, glosa: str) -> str:
    """Normaliza proveedor + glosa en clave estable."""
    return f"{_normalize(proveedor)}||{_normalize(glosa)}"


def _normalize(text: str) -> str:
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


class CategorizerCache:
    def __init__(self, path: str):
        self.path = path
        self._lock = threading.Lock()
        self._data = self._load()

    def _load(self) -> dict:
        if not os.path.exists(self.path):
            return {}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.warning(f"Cache corrupto, reseteando: {e}")
            return {}

    def _save(self):
        os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self.path)

    def get(self, proveedor: str, glosa: str) -> dict | None:
        key = make_cache_key(proveedor, glosa)
        with self._lock:
            return self._data.get(key)

    def set(self, proveedor: str, glosa: str, result: dict):
        if result.get("confianza", 0) < MIN_CONFIDENCE_TO_CACHE:
            return
        if result.get("categoria") == "REVISAR":
            return
        key = make_cache_key(proveedor, glosa)
        with self._lock:
            self._data[key] = result
            self._save()

    def size(self) -> int:
        return len(self._data)
```

- [ ] **Step 4: Run test — expect PASS**

Run: `py -3.11 -m pytest tests/test_categorizer_cache.py -v`
Expected: 6 passed

- [ ] **Step 5: Commit**

```bash
git add modules/cash_flow/categorizer_cache.py tests/test_categorizer_cache.py
git commit -m "feat: add JSON cache for categorization results"
```

---
