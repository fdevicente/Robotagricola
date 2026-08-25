"""Registro persistente del estado del bot.

Se guardan DOS marcas distintas, y confundirlas da avisos falsos:

  · `ultima_actividad_utc` — cuándo llegó el último mensaje ENTRANTE. Sirve
    para saber si pudieron perderse mensajes (Telegram solo los retiene 24 h
    con polling).
  · `ultimo_latido_utc` — cuándo estuvo vivo el PROCESO por última vez. Lo
    escribe un job periódico.

Antes solo existía la primera, así que un fin de semana sin que nadie le
escribiera se reportaba como "62 h apagado" aunque el bot estuviera corriendo
y mandando su heartbeat.
"""
import os
import json
import logging
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STATE_FILE = os.path.join(_DIR, ".bot_state.json")


def cargar_estado() -> dict:
    """Lee el estado persistido. Devuelve {} si no existe."""
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.warning(f"No pude leer bot_state: {e}")
        return {}


def guardar_actividad(update_id: int | None = None,
                       chat_id=None,
                       resumen: str = "") -> None:
    """Registra la última actividad procesada (timestamp + update_id)."""
    estado = cargar_estado()
    estado["ultimo_update_id"] = update_id
    estado["ultimo_chat_id"] = chat_id
    estado["ultima_actividad_utc"] = datetime.now(timezone.utc).isoformat()
    estado["ultimo_resumen"] = resumen[:200]
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"No pude guardar bot_state: {e}")


def guardar_latido() -> None:
    """Marca que el proceso sigue vivo. Lo llama un job periódico."""
    estado = cargar_estado()
    estado["ultimo_latido_utc"] = datetime.now(timezone.utc).isoformat()
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f"No pude guardar el latido: {e}")


def _horas_desde(clave: str) -> float | None:
    ts = cargar_estado().get(clave)
    if not ts:
        return None
    try:
        return (datetime.now(timezone.utc)
                - datetime.fromisoformat(ts)).total_seconds() / 3600
    except Exception:
        return None


def horas_desde_ultima_actividad() -> float | None:
    """Horas desde el último mensaje ENTRANTE. None si no hay registro."""
    return _horas_desde("ultima_actividad_utc")


def horas_apagado() -> float | None:
    """Horas que el proceso estuvo realmente caído.

    Se mide contra el latido, no contra los mensajes: que nadie escriba en el
    fin de semana no significa que el bot estuviera abajo.
    """
    return _horas_desde("ultimo_latido_utc")


def _local(ts: str) -> str:
    try:
        from zoneinfo import ZoneInfo
        return (datetime.fromisoformat(ts)
                .astimezone(ZoneInfo("America/Santiago"))
                .strftime("%d-%m-%Y %H:%M"))
    except Exception:
        return (ts or "")[:16]


# Margen sobre el intervalo del latido: por debajo de esto fue un reinicio
# normal (o el watchdog relanzando), no una caída que valga avisar.
UMBRAL_CAIDA_H = 0.5


def mensaje_reconexion() -> str | None:
    """Aviso de reconexión, o None si no hay nada que avisar.

    El corte se mide con el LATIDO del proceso. Si no hay latido (primer
    arranque tras la actualización) no se avisa nada: no se puede distinguir
    "estuvo caído" de "nadie escribió".
    """
    caido = horas_apagado()
    if caido is None or caido < UMBRAL_CAIDA_H:
        return None

    estado = cargar_estado()
    desde = _local(estado.get("ultimo_latido_utc", ""))

    if caido < 24:
        return (f"🔄 *Bot reconectado*\n"
                f"Estuvo caído {caido:.0f}h (desde {desde}).\n"
                f"Telegram me entrega los mensajes de las últimas 24h, "
                f"así que no debería faltar nada.")
    return (f"⚠️ *Bot reconectado tras {caido:.0f}h caído*\n"
            f"Se cayó el {desde}.\n"
            f"Telegram solo retiene mensajes 24h, así que pueden haberse "
            f"perdido mensajes/facturas enviados hace más de 1 día. "
            f"Por favor revisa y reenvía lo que falte.")
