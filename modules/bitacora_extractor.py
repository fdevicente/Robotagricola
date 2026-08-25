"""Extrae campos estructurados de un registro de bitácora en lenguaje natural.

Usa Claude (mismo patrón que el categorizer de facturas). El capataz escribe
algo como "hoy poda en avellanos, estuvieron todos: Felicito, Patricio,
Ramiro, Richard y Jorge" y devuelve los campos estructurados.
"""
import json
import logging
import re
import requests

from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)

CLAUDE_URL = "https://api.anthropic.com/v1/messages"
CLAUDE_MODELS = ["claude-haiku-4-5-20251001", "claude-sonnet-4-6"]
MAX_TOKENS = 600
TIMEOUT_SEC = 30

# Trabajadores conocidos (la IA normaliza nombres a estos)
TRABAJADORES_CONOCIDOS = [
    "Felix De Vicente",
    "Juan Parada",
    "Felicito Amigo",
    "Agustin Mora",
    "Patricio Mora",
    "Ramiro Amigo",
    "Jorge",
    "Richard Padilla",
    # El HIJO, trabajando con ellos desde julio-2026. Se llama igual que el
    # padre y se distingue por el "Crespo": son DOS personas, no un duplicado.
    "Richard Padilla Crespo",
]
# Alias frecuentes → nombre canónico
ALIAS = {
    "felix": "Felix De Vicente", "felipe": "Felix De Vicente",
    "juan": "Juan Parada", "parada": "Juan Parada",
    "felicito": "Felicito Amigo", "felicto": "Felicito Amigo",
    "agustin": "Agustin Mora", "agustín": "Agustin Mora",
    "patricio": "Patricio Mora", "pato": "Patricio Mora",
    "ramiro": "Ramiro Amigo",
    "jorge": "Jorge",
    # "richard" a secas es el PADRE: lleva años y es el default histórico.
    # El hijo solo se reconoce si aparece el "Crespo".
    "richard": "Richard Padilla", "ricardo": "Richard Padilla",
    "crespo": "Richard Padilla Crespo",
}

CULTIVOS = ["NOGALES", "CEREZOS", "AVELLANOS", "GENERAL"]
TIPOS = ["LABOR", "APLICACION", "RIEGO", "EVENTO", "MAQUINARIA", "OTRO"]

# Máquinas frecuentes (la IA normaliza a estas)
MAQUINAS = ["EXCAVADORA", "TRACTOR", "CAMION", "CAMIONETA", "RETROEXCAVADORA",
            "BARREDORA", "SOPLADOR", "PULVERIZADORA", "SORTER", "GRUA"]


def _build_prompt(texto: str, fecha_hoy: str) -> str:
    trabajadores = ", ".join(TRABAJADORES_CONOCIDOS)
    return f"""Eres un asistente que estructura registros de bitácora de una agrícola en Chile.
El capataz (Juan Parada) describe el trabajo del día. Extrae los campos en JSON.

Fecha de hoy: {fecha_hoy}

TRABAJADORES CONOCIDOS (normaliza los nombres a estos exactos):
{trabajadores}

CULTIVOS posibles: NOGALES, CEREZOS, AVELLANOS, GENERAL (si no aplica a un cultivo).

TIPOS de registro:
- LABOR: trabajo manual (poda, raleo, cosecha, limpieza, plantación, deshierbe, etc.)
- APLICACION: aplicación de pesticida/fungicida/herbicida/fertilizante (lleva insumo + cantidad)
- RIEGO: riego de cultivos
- MAQUINARIA: trabajo con máquina (excavadora, tractor, etc.) — limpieza de camellones,
  sacar tocones, arar, transporte. Lleva máquina + odómetro/horómetro si lo menciona.
- EVENTO: clima (lluvia, helada, viento) o incidente (rotura, falla de maquinaria, corte de luz)
- OTRO: cualquier otra cosa

MAQUINA: si es tipo MAQUINARIA, identifica la máquina (EXCAVADORA, TRACTOR, CAMION, etc.).
ODOMETRO: si menciona horómetro/odómetro/horas de máquina, extrae el NÚMERO (ej: 1250).

JORNADAS HOMBRE (JH): cantidad de trabajadores que participaron ese día (1 trabajador = 1 JH).
Si dice "estuvieron todos" sin listar, deja la lista vacía pero NO inventes JH; pon jornadas_hombre=null.
Si lista nombres, jornadas_hombre = cantidad de nombres.

SECTOR: si menciona un sector/parcela/cuartel del campo, ponlo; sino deja "".

FECHA DEL TRABAJO: el capataz suele reportar días después. Si el texto menciona
una fecha o día (ej: "23 de julio", "lunes 15 de junio 2026", "asistencia 8 julio"),
devuélvela como "YYYY-MM-DD" usando el año de la fecha de hoy si no lo dice.
Si NO menciona ninguna fecha, devuelve null (se asumirá hoy).

Devuelve SOLO este JSON (sin texto extra, sin markdown):
{{
  "fecha": "<YYYY-MM-DD o null>",
  "tipo": "<LABOR|APLICACION|RIEGO|MAQUINARIA|EVENTO|OTRO>",
  "actividad": "<descripción corta, ej: Poda, Aplicación foliar, Limpieza camellones, Lluvia>",
  "cultivo": "<NOGALES|CEREZOS|AVELLANOS|GENERAL>",
  "sector": "<sector o vacío>",
  "jornadas_hombre": <número o null>,
  "trabajadores": ["<nombre canónico>", ...],
  "insumo": "<nombre del insumo o vacío>",
  "cantidad": <número o null>,
  "unidad": "<L|kg|cc|unidad o vacío>",
  "maquina": "<EXCAVADORA|TRACTOR|... o vacío>",
  "odometro": <número del horómetro/odómetro o null>,
  "superficie_ha": <hectáreas trabajadas o null>,
  "confianza": <0.0 a 1.0>,
  "resumen": "<frase corta de lo que se hizo>"
}}

Registro del capataz:
\"\"\"{texto}\"\"\""""


def _normalizar(data) -> dict:
    """Limpia y valida los campos extraídos.

    Si el texto trae varios días, la IA a veces devuelve una LISTA de objetos;
    en ese caso se toma el primero (el desglose por día lo hace
    `bitacora_asistencia.parsear_asistencia_multi`, que es determinista).
    """
    if isinstance(data, list):
        data = data[0] if data and isinstance(data[0], dict) else {}
    if not isinstance(data, dict):
        data = {}
    out = {
        "fecha": str(data.get("fecha") or "").strip()[:10],
        "tipo": str(data.get("tipo") or "OTRO").upper(),
        "actividad": str(data.get("actividad") or "").strip(),
        "cultivo": str(data.get("cultivo") or "GENERAL").upper(),
        "sector": str(data.get("sector") or "").strip(),
        "jornadas_hombre": data.get("jornadas_hombre"),
        "trabajadores": data.get("trabajadores") or [],
        "insumo": str(data.get("insumo") or "").strip(),
        "cantidad": data.get("cantidad"),
        "unidad": str(data.get("unidad") or "").strip(),
        "maquina": str(data.get("maquina") or "").strip().upper(),
        "odometro": data.get("odometro"),
        "superficie_ha": data.get("superficie_ha"),
        "confianza": float(data.get("confianza") or 0.5),
        "resumen": str(data.get("resumen") or "").strip(),
    }
    if out["tipo"] not in TIPOS:
        out["tipo"] = "OTRO"
    # Validar la fecha extraída (formato YYYY-MM-DD y no futura)
    if out["fecha"]:
        try:
            from datetime import datetime as _dt, date as _d
            f = _dt.strptime(out["fecha"], "%Y-%m-%d").date()
            if f > _d.today():
                out["fecha"] = ""
        except (ValueError, TypeError):
            out["fecha"] = ""
    if out["cultivo"] not in CULTIVOS:
        out["cultivo"] = "GENERAL"
    # Normalizar trabajadores con alias
    norm = []
    for t in out["trabajadores"]:
        key = str(t).strip().lower()
        norm.append(ALIAS.get(key, str(t).strip()))
    # quitar duplicados preservando orden
    seen = set()
    out["trabajadores"] = [x for x in norm if not (x in seen or seen.add(x))]
    # JH: si no vino pero hay lista, usar largo de lista
    if out["jornadas_hombre"] in (None, "") and out["trabajadores"]:
        out["jornadas_hombre"] = len(out["trabajadores"])
    try:
        out["jornadas_hombre"] = (int(out["jornadas_hombre"])
                                   if out["jornadas_hombre"] not in (None, "") else None)
    except (TypeError, ValueError):
        out["jornadas_hombre"] = None
    try:
        out["cantidad"] = (float(out["cantidad"])
                            if out["cantidad"] not in (None, "") else None)
    except (TypeError, ValueError):
        out["cantidad"] = None
    # Odómetro y superficie numéricos
    try:
        out["odometro"] = (float(out["odometro"])
                            if out["odometro"] not in (None, "") else None)
    except (TypeError, ValueError):
        out["odometro"] = None
    try:
        out["superficie_ha"] = (float(out["superficie_ha"])
                                 if out["superficie_ha"] not in (None, "") else None)
    except (TypeError, ValueError):
        out["superficie_ha"] = None
    return out


def extraer_bitacora(texto: str, fecha_hoy: str) -> dict:
    """Extrae campos estructurados del texto. Devuelve dict normalizado."""
    prompt = _build_prompt(texto, fecha_hoy)
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload_base = {"max_tokens": MAX_TOKENS,
                    "messages": [{"role": "user", "content": prompt}]}

    for model in CLAUDE_MODELS:
        try:
            resp = requests.post(CLAUDE_URL, headers=headers,
                                  json={**payload_base, "model": model},
                                  timeout=TIMEOUT_SEC)
        except requests.RequestException as e:
            logger.warning(f"Claude {model} excepcion bitacora: {e}")
            continue
        if resp.status_code != 200:
            logger.warning(f"Claude {model} HTTP {resp.status_code}: {resp.text[:150]}")
            continue
        try:
            raw = resp.json()["content"][0]["text"].strip()
            # Quitar fences si vinieron
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
            data = json.loads(raw)
            return {**_normalizar(data), "texto_original": texto}
        except (KeyError, IndexError, ValueError, json.JSONDecodeError) as e:
            logger.warning(f"Claude {model} respuesta no parseable: {e}")
            continue

    logger.error("Todos los modelos Claude fallaron para bitácora")
    return {
        "fecha": "",
        "tipo": "OTRO", "actividad": "", "cultivo": "GENERAL", "sector": "",
        "jornadas_hombre": None, "trabajadores": [], "insumo": "",
        "cantidad": None, "unidad": "", "maquina": "", "odometro": None,
        "superficie_ha": None, "confianza": 0.0,
        "resumen": texto[:60], "texto_original": texto,
        "error": "IA no disponible",
    }


# ── Mensajes que NO son un registro ─────────────────────────────────────────
# Juan escribe la sección sola antes de mandar el parte ("Bitácora", "Personal",
# "Asistencia /"). El guard viejo era `len(texto) < 3`, así que todas pasaban y
# se guardaban como OTRO con la propia palabra de actividad. Se limpiaron a mano
# el 10-ago y el 18-ago de 2026 y volvieron al día siguiente.
_PALABRAS_COMANDO = {
    "bitacora", "bitácora", "personal", "maquinaria", "maquina", "máquina",
    "asistencia", "inventario", "tarea", "tareas", "hecho", "vacaciones",
    "uso", "saldo", "reporte", "banco", "ayuda", "deposito", "depósito",
    "pagado", "vencimientos", "bodega", "conciliar", "correlativos",
    "basedatos", "estado", "cancelar", "deshacer", "dashboard", "proyeccion",
    "proyección", "categoria", "categoría", "start", "soydueno",
}


def es_mensaje_sin_contenido(texto: str | None) -> bool:
    """True si el mensaje es solo un comando o puntuación, sin nada que anotar.

    Nunca descarta un mensaje que traiga información: "Bitácora" sola es basura,
    pero "Bitácora: hoy se podó el sector 3" es un registro bueno.
    """
    if not texto:
        return True
    limpio = re.sub(r"[^\w\sáéíóúñü]", " ", texto, flags=re.IGNORECASE)
    palabras = [p for p in limpio.lower().split() if p]
    if not palabras:
        return True
    return all(p in _PALABRAS_COMANDO for p in palabras)
