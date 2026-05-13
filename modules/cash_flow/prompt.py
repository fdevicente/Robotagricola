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

    cleaned = raw.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

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
