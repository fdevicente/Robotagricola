"""Detectores de alertas (puras, sin I/O Telegram)."""
from datetime import date, datetime


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        try:
            return datetime.strptime(v[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def detect_saldo_bajo(saldo_actual: float, saldo_minimo: float):
    """Alerta 🔴 si saldo < minimo."""
    if saldo_actual >= saldo_minimo:
        return None
    diff = saldo_minimo - saldo_actual
    return {
        "tipo": "saldo_bajo",
        "nivel": "🔴",
        "mensaje": (f"🔴 Saldo bajo: ${saldo_actual:,.0f} CLP "
                     f"(falta ${diff:,.0f} para el minimo de ${saldo_minimo:,.0f})"),
    }


def detect_factura_por_vencer(facturas: list, hoy=None, dias: int = 3) -> list:
    """Alerta 🟡 por cada factura que vence en `dias` o menos."""
    hoy = hoy or date.today()
    alertas = []
    for f in facturas:
        venc = _to_date(f.get("fecha_vencimiento"))
        if not venc:
            continue
        delta = (venc - hoy).days
        if 0 <= delta <= dias:
            alertas.append({
                "tipo": "factura_por_vencer",
                "nivel": "🟡",
                "fila": f["fila"],
                "mensaje": (
                    f"🟡 Vence en {delta}d: {f.get('proveedor', '')} "
                    f"factura {f.get('nro_factura', '')} ${f.get('total', 0):,.0f}"
                ),
            })
    return alertas


import json
import os


class AlertDedupe:
    """Recuerda que alertas ya se enviaron por (tipo, periodo)."""

    def __init__(self, path: str):
        self.path = path
        self._data = self._load()

    def _load(self) -> dict:
        if not os.path.exists(self.path):
            return {}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}

    def _save(self):
        os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(self._data, f)

    def should_fire(self, tipo: str, periodo: str) -> bool:
        return self._data.get(f"{tipo}|{periodo}") is None

    def mark_fired(self, tipo: str, periodo: str):
        self._data[f"{tipo}|{periodo}"] = True
        self._save()
