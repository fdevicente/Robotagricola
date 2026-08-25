# -*- coding: utf-8 -*-
"""Respaldo crudo de todo lo que entra por Telegram.

POR QUÉ EXISTE
El 24-ago-2026 un parte de horómetro de Juan se perdió en silencio: el bot lo
recibió, respondió y no guardó nada. No se pudo saber qué pasó porque **el texto
original no quedaba en ninguna parte**: el mirror lo reenvía al chat del dueño
pero no lo persiste, y la persistencia de PTB solo guarda configuración.
Lo mismo con los 8 rechazos de odómetro: el log dejó el número mal extraído,
nunca el mensaje que lo produjo, así que no hay forma de recuperar la lectura.

QUÉ HACE
Una línea JSON por mensaje, en `files/telegram/YYYY-MM.jsonl`. Texto plano,
append-only, un archivo por mes. Guarda el mensaje TAL CUAL llegó, antes de que
nadie lo interprete.

REGLA DE ORO: esto NUNCA puede voltear el bot. Si falla al escribir, se traga el
error y sigue — perder el respaldo es malo, perder el mensaje es peor.
"""
import json
import logging
import os
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "files", "telegram")


def ruta_del_mes(cuando=None, base: str | None = None) -> str:
    """`files/telegram/YYYY-MM.jsonl` del mes de `cuando` (UTC)."""
    cuando = cuando or datetime.now(timezone.utc)
    return os.path.join(base or _DIR, f"{cuando.strftime('%Y-%m')}.jsonl")


def _resumir(msg) -> dict:
    """Saca del mensaje lo que sirve para reconstruirlo después."""
    d = {}
    if msg is None:
        return d
    d["message_id"] = getattr(msg, "message_id", None)
    for campo in ("text", "caption"):
        val = getattr(msg, campo, None)
        if val:
            d[campo] = val
    # Adjuntos: no se guarda el binario, sí con qué identificarlo
    doc = getattr(msg, "document", None)
    if doc is not None:
        d["documento"] = {"file_name": getattr(doc, "file_name", None),
                          "mime": getattr(doc, "mime_type", None),
                          "file_id": getattr(doc, "file_id", None),
                          "size": getattr(doc, "file_size", None)}
    fotos = getattr(msg, "photo", None)
    if fotos:
        try:
            d["foto"] = {"file_id": fotos[-1].file_id}
        except Exception:
            d["foto"] = {"file_id": None}
    if getattr(msg, "voice", None) is not None:
        d["voz"] = True
    return d


def guardar_update(update, base: str | None = None) -> str | None:
    """Escribe una línea con el update crudo. Devuelve la ruta, o None si no aplica.

    Nunca lanza: si no puede escribir, lo deja en el log y sigue.
    """
    try:
        msg = getattr(update, "effective_message", None)
        if msg is None:
            return None
        usuario = getattr(update, "effective_user", None)
        chat = getattr(update, "effective_chat", None)
        fecha = getattr(msg, "date", None) or datetime.now(timezone.utc)

        fila = {
            "recibido_utc": datetime.now(timezone.utc).isoformat(),
            "fecha_mensaje_utc": fecha.isoformat() if hasattr(fecha, "isoformat")
                                  else str(fecha),
            "update_id": getattr(update, "update_id", None),
            "chat_id": getattr(chat, "id", None),
            "user_id": getattr(usuario, "id", None),
            "nombre": getattr(usuario, "full_name", None),
        }
        fila.update(_resumir(msg))

        destino = ruta_del_mes(base=base)
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        with open(destino, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(fila, ensure_ascii=False) + "\n")
        return destino
    except Exception as e:                      # nunca voltear el bot por esto
        logger.warning(f"No pude respaldar el update de Telegram: {e}")
        return None


def leer_mes(anio_mes: str, base: str | None = None) -> list[dict]:
    """Lee un mes del respaldo. `anio_mes` en formato 'YYYY-MM'."""
    ruta = os.path.join(base or _DIR, f"{anio_mes}.jsonl")
    if not os.path.exists(ruta):
        return []
    filas = []
    with open(ruta, encoding="utf-8") as fh:
        for ln in fh:
            ln = ln.strip()
            if not ln:
                continue
            try:
                filas.append(json.loads(ln))
            except json.JSONDecodeError:
                # una línea corrupta no puede inutilizar el resto del respaldo
                logger.warning("Línea ilegible en el respaldo %s", ruta)
    return filas


def buscar(texto: str, base: str | None = None) -> list[dict]:
    """Busca en todo el respaldo los mensajes que contengan `texto`."""
    carpeta = base or _DIR
    if not os.path.isdir(carpeta):
        return []
    out = []
    for nombre in sorted(os.listdir(carpeta)):
        if not nombre.endswith(".jsonl"):
            continue
        for fila in leer_mes(nombre[:-6], base=carpeta):
            cuerpo = f"{fila.get('text') or ''} {fila.get('caption') or ''}"
            if texto.lower() in cuerpo.lower():
                out.append(fila)
    return out
