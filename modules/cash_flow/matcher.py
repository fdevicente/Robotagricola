"""Match facturas pendientes con cargos bancarios."""
from datetime import date, datetime


def _to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(v[:10], fmt).date()
            except ValueError:
                pass
    return None


def _provider_match(provider: str, description: str) -> float:
    p = (provider or "").lower()
    d = (description or "").lower()
    words = [w for w in p.split() if len(w) > 3]
    if not words:
        return 0.0
    hits = sum(1 for w in words if w in d)
    return hits / len(words)


def match_score(factura: dict, bank_mov: dict) -> float:
    """Devuelve score 0-200. >=100 candidato fuerte. >=30 ambiguo."""
    cargo = float(bank_mov.get("cargo") or 0)
    if cargo <= 0:
        return 0.0

    total = float(factura.get("total") or 0)
    if total <= 0:
        return 0.0

    score = 0.0
    diff_pct = abs(cargo - total) / total
    if diff_pct < 0.001:
        score += 100
    elif diff_pct < 0.05:
        score += 30

    f_factura = _to_date(factura.get("fecha_emision"))
    f_banco = _to_date(bank_mov.get("fecha"))
    if f_factura and f_banco:
        diff_dias = abs((f_banco - f_factura).days)
        if diff_dias <= 5:
            score += 50
        elif diff_dias <= 15:
            score += 20

    score += 40 * _provider_match(factura.get("proveedor", ""),
                                   bank_mov.get("descripcion", ""))

    nro = str(factura.get("nro_factura") or "").strip()
    ref = str(bank_mov.get("referencia") or "")
    if nro and nro in ref:
        score += 50

    return score


MATCH_THRESHOLD = 30


def find_matches(bank_mov: dict, facturas_pendientes: list[dict]) -> list[dict]:
    """Devuelve candidatos sobre threshold, ordenados por score desc."""
    scored = []
    for f in facturas_pendientes:
        s = match_score(f, bank_mov)
        if s >= MATCH_THRESHOLD:
            scored.append({**f, "score": s})
    scored.sort(key=lambda x: x["score"], reverse=True)
    return scored
