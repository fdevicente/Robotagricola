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
    # Eliminar puntuacion intra-token (".SA" -> "SA", "S.A." -> "SA")
    text = re.sub(r"[.\'`]", "", text)
    # El resto de no-alfanum lo convierto en espacio
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
