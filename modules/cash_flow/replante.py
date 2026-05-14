"""Affordability check para replante."""


def afford_check(cultivo: str, hc: float,
                   saldo_proyectado: float,
                   saldo_minimo: float,
                   costo_por_hc: float) -> dict:
    """Verifica si alcanza la caja para replantar X hectareas."""
    costo_total = hc * costo_por_hc
    disponible = saldo_proyectado - saldo_minimo
    alcanza = costo_total <= disponible
    return {
        "cultivo": cultivo,
        "hc": hc,
        "costo_por_hc": costo_por_hc,
        "costo_total": costo_total,
        "saldo_proyectado": saldo_proyectado,
        "saldo_minimo": saldo_minimo,
        "disponible": disponible,
        "alcanza": alcanza,
        "deficit": 0 if alcanza else (costo_total - disponible),
    }
