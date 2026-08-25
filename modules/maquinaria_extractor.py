"""Extrae mantenciones y fichas de maquinaria desde texto libre.

Juan escribe como habla ("al 5085 le cambiaron aceite y filtros el 20 de julio
a las 3100 horas, lo hizo Álamos"). Acá se convierte en registros.
"""
import json
import logging

import requests

from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)

URL = "https://api.anthropic.com/v1/messages"
MODELOS = ["claude-haiku-4-5-20251001", "claude-sonnet-4-6"]
TIMEOUT = 40


def _prompt(texto: str, maquinas: list, hoy: str) -> str:
    lista = "\n".join(f"- {m}" for m in maquinas) or "(ninguna registrada aún)"
    return f"""Estructura lo que cuenta el capataz de una agrícola chilena sobre su maquinaria.

Fecha de hoy: {hoy}

MÁQUINAS YA CONOCIDAS (usa EXACTAMENTE estos nombres si se refiere a una de ellas):
{lista}

Si menciona una máquina que no está en la lista, usa el nombre completo que él dé.
Ojo con las abreviaciones: "JD"=John Deere, "MF"=Massey Ferguson.
Suele nombrarlas solo por el modelo ("el 5085", "la 6711").

Puede venir MANTENCIONES (algo que se le hizo) y/o FICHAS (datos de la máquina).
Un mismo mensaje puede traer varias de cada una.

MANTENCIÓN: algo que se le hizo, o algo que HAY QUE hacerle.
  ⚠️ DISTINGUE:
    - estado "HECHA": ya se hizo ("le cambiaron el aceite", "le hicieron engrase").
    - estado "PENDIENTE": todavía no ("necesita neumáticos", "hay que cambiarle
      el filtro", "está fallando", "falta hacerle"). NO le pongas fecha pasada
      a algo pendiente.
  tipo: ACEITE MOTOR | FILTROS | HIDRAULICO | TRANSMISION | ENGRASE |
        NEUMATICOS | FRENOS | REPARACION | REVISION | OTRO
  fecha: "YYYY-MM-DD" de cuando se hizo. Si dice "el 20 de julio" usa el año de
        hoy; "la semana pasada" cuéntala desde hoy. Si es PENDIENTE o no lo
        dice, null.
  odometro: el número de horas/km al momento de la mantención (null si no dice).
  proxima_h: cada cuántas horas toca repetirla, si lo menciona (null si no).

FICHA: marca, modelo, año, patente, n° de serie, si es propia o arrendada,
  tipo (TRACTOR/CAMIONETA/EXCAVADORA/OTRO) y estado.

Devuelve SOLO este JSON, sin texto alrededor:
{{
  "mantenciones": [
    {{"maquina": "<nombre>", "estado": "<HECHA|PENDIENTE>", "tipo": "<TIPO>",
      "descripcion": "<qué se hizo o hay que hacer>",
      "fecha": "<YYYY-MM-DD o null>", "odometro": <número o null>,
      "proveedor": "<quién o vacío>", "costo": <número o null>,
      "proxima_h": <número o null>, "notas": "<algo relevante o vacío>"}}
  ],
  "fichas": [
    {{"maquina": "<nombre>", "tipo": "<TRACTOR|CAMIONETA|EXCAVADORA|OTRO o vacío>",
      "marca": "", "modelo": "", "anio": <número o null>, "patente": "",
      "serie": "", "propiedad": "<PROPIA|ARRENDADA o vacío>",
      "estado": "", "notas": ""}}
  ]
}}

Deja vacía la lista que no aplique. No inventes datos que no estén en el texto.

Mensaje del capataz:
\"\"\"{texto}\"\"\""""


def _limpiar(dato) -> dict:
    if isinstance(dato, list):
        dato = dato[0] if dato and isinstance(dato[0], dict) else {}
    if not isinstance(dato, dict):
        return {"mantenciones": [], "fichas": []}
    out = {"mantenciones": [], "fichas": []}
    for m in dato.get("mantenciones") or []:
        if isinstance(m, dict) and m.get("maquina"):
            out["mantenciones"].append(m)
    for f in dato.get("fichas") or []:
        if isinstance(f, dict) and f.get("maquina"):
            # No guardar fichas que vienen sin ningún dato útil
            if any(f.get(k) for k in ("marca", "modelo", "anio", "patente",
                                       "serie", "propiedad", "estado", "tipo")):
                out["fichas"].append(f)
    return out


def extraer(texto: str, maquinas: list, hoy: str) -> dict:
    """Devuelve {'mantenciones': [...], 'fichas': [...]}."""
    if not ANTHROPIC_API_KEY:
        raise RuntimeError("falta ANTHROPIC_API_KEY")
    cuerpo = {"max_tokens": 1200,
              "messages": [{"role": "user",
                            "content": _prompt(texto, maquinas, hoy)}]}
    headers = {"x-api-key": ANTHROPIC_API_KEY,
               "anthropic-version": "2023-06-01",
               "content-type": "application/json"}
    ultimo = None
    for modelo in MODELOS:
        try:
            r = requests.post(URL, headers=headers,
                              json={**cuerpo, "model": modelo}, timeout=TIMEOUT)
            if r.status_code != 200:
                ultimo = f"HTTP {r.status_code}"
                continue
            crudo = r.json()["content"][0]["text"].strip()
            if crudo.startswith("```"):
                crudo = crudo.split("```")[1]
                if crudo.startswith("json"):
                    crudo = crudo[4:]
            return _limpiar(json.loads(crudo))
        except Exception as e:
            ultimo = str(e)
            logger.warning(f"maquinaria_extractor {modelo}: {e}")
            continue
    raise RuntimeError(ultimo or "todos los modelos fallaron")
