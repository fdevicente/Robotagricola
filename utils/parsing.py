"""
utils/parsing.py - Funciones de parseo reutilizadas en multiples handlers.
- parsear_fecha: texto libre (DD/MM/YYYY, YYYY-MM-DD, "hoy") -> "YYYY-MM-DD"
- parsear_monto: texto con formato chileno ($, puntos miles, coma decimal) -> float
"""
from datetime import datetime


def parsear_fecha(texto: str) -> str | None:
    """Convierte texto a fecha YYYY-MM-DD. Retorna None si no es valido.

    Formatos aceptados:
      - "hoy"           -> fecha actual
      - "15/03/2026"    -> DD/MM/YYYY
      - "2026-03-15"    -> ISO
      - Cualquier otro  -> intenta dateutil.parser
    """
    texto = texto.strip()
    if texto.lower() == "hoy":
        return datetime.now().strftime("%Y-%m-%d")
    try:
        if "/" in texto:
            parts = texto.split("/")
            if len(parts) == 3 and len(parts[0]) <= 2:
                return datetime.strptime(texto, "%d/%m/%Y").strftime("%Y-%m-%d")
        from dateutil import parser as date_parser
        return date_parser.parse(texto).strftime("%Y-%m-%d")
    except Exception:
        return None


def parsear_monto(texto: str) -> float | None:
    """Convierte texto con formato chileno a float. Retorna None si no es valido.

    Formatos aceptados:
      - "$1.234.567"  -> 1234567.0  (puntos = miles)
      - "7.164,32"    -> 7164.32    (coma = decimal)
      - "500000"      -> 500000.0
    """
    try:
        n = texto.replace("$", "").replace(" ", "").strip()
        if "," in n:
            n = n.replace(".", "").replace(",", ".")
        elif n.count(".") > 1:
            n = n.replace(".", "")
        return float(n)
    except (ValueError, AttributeError):
        return None
