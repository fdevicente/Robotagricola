# -*- coding: utf-8 -*-
"""Cola de subidas a Drive, persistida en disco.

Se persiste a propósito: el bot corre bajo watchdog y se reinicia seguido, así
que una cola en memoria perdería lo pendiente en cada reinicio. El formato es
append-only (una línea JSON por evento) para que una escritura interrumpida no
corrompa lo anterior.
"""
import json
import logging
import os
import uuid
from datetime import datetime, timezone

logger = logging.getLogger(__name__)


class Cola:
    def __init__(self, ruta: str, max_intentos: int = 5):
        self.ruta = ruta
        self.max_intentos = max_intentos

    # ── lectura ────────────────────────────────────────────────────────────
    def _eventos(self) -> list[dict]:
        if not os.path.exists(self.ruta):
            return []
        out = []
        with open(self.ruta, encoding="utf-8") as fh:
            for ln in fh:
                ln = ln.strip()
                if not ln:
                    continue
                try:
                    out.append(json.loads(ln))
                except json.JSONDecodeError:
                    # una línea corrupta no puede inutilizar el resto
                    logger.warning("Línea ilegible en la cola de Drive")
        return out

    def _estado(self) -> dict:
        """Reconstruye el estado actual aplicando los eventos en orden."""
        items: dict = {}
        for e in self._eventos():
            tipo, iid = e.get("evento"), e.get("id")
            if not iid:
                continue
            if tipo == "encolado":
                items.setdefault(iid, {
                    "id": iid, "ruta_local": e["ruta_local"],
                    "carpeta": e["carpeta"], "nombre": e["nombre"],
                    "intentos": 0, "ultimo_error": "", "file_id": None,
                    "listo": False})
            elif iid in items:
                if tipo == "ok":
                    items[iid]["listo"] = True
                    items[iid]["file_id"] = e.get("file_id")
                elif tipo == "error":
                    items[iid]["intentos"] += 1
                    items[iid]["ultimo_error"] = e.get("motivo", "")
        return items

    def pendientes(self) -> list[dict]:
        return [i for i in self._estado().values()
                if not i["listo"] and i["intentos"] < self.max_intentos]

    def rendidos(self) -> list[dict]:
        """Los que agotaron los reintentos. El archivo sigue en disco."""
        return [i for i in self._estado().values()
                if not i["listo"] and i["intentos"] >= self.max_intentos]

    # ── escritura ──────────────────────────────────────────────────────────
    def _append(self, fila: dict) -> None:
        os.makedirs(os.path.dirname(self.ruta) or ".", exist_ok=True)
        fila["cuando"] = datetime.now(timezone.utc).isoformat()
        with open(self.ruta, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(fila, ensure_ascii=False) + "\n")

    def encolar(self, ruta_local: str, carpeta: str, nombre: str) -> str:
        """Encola una subida. Si ese archivo ya está pendiente, no duplica."""
        for i in self.pendientes():
            if i["ruta_local"] == ruta_local and i["carpeta"] == carpeta:
                return i["id"]
        iid = uuid.uuid4().hex
        self._append({"evento": "encolado", "id": iid, "ruta_local": ruta_local,
                      "carpeta": carpeta, "nombre": nombre})
        return iid

    def marcar_ok(self, iid: str, file_id: str) -> None:
        self._append({"evento": "ok", "id": iid, "file_id": file_id})

    def marcar_error(self, iid: str, motivo: str) -> None:
        self._append({"evento": "error", "id": iid, "motivo": str(motivo)[:200]})
