# -*- coding: utf-8 -*-
"""Vida util de los flujos conversacionales (deposito, tarea, bitacora, ...).

POR QUE EXISTE
El 28-ago-2026 Juan escribio /deposito y quiso salir con "/ cancelar" -- con un
espacio. Telegram solo marca como comando lo que va pegado, asi que ese mensaje
NO ejecuto cmd_cancelar: entro como texto normal y se lo comio el propio flujo
de deposito. `deposito_state` quedo abierto en `.bot_persistence.pickle`, que
sobrevive a los reinicios, y como es el primero de la fila en handlers/chat.py
cada parte de Juan se leyo como "el monto del deposito", fallo, y corto ahi.
Doce dias de partes de asistencia y horometros que nunca llegaron a la bitacora,
sin una sola linea en el log.

LA REGLA QUE IMPORTA
El reloj corre desde que el flujo se ABRIO, no desde el ultimo mensaje. Juan
mandaba un parte cada ~80 s: un timeout "desde el ultimo mensaje" se habria
refrescado con cada intento fallido y no habria vencido nunca.
"""
import logging

logger = logging.getLogger(__name__)

MINUTOS_VIDA = 30

# Claves que ABREN un flujo: mientras una tenga valor, su handler se queda con
# todo el texto que llegue. Son las que hay que vigilar.
CLAVES_ESTADO = (
    "deposito_state",
    "pagado_state",
    "tarea_state",
    "bitacora_state",
    "venc_state",          # faltaba en /cancelar: tambien se puede trabar
    "uso_state",
    "vacacion_state",
    "trabajador_state",
    "editing_field",
)

# Datos que acompanan a cada flujo y que hay que soltar junto con el estado.
CLAVES_DATOS = (
    "deposito_monto",
    "pagado_nro",
    "tarea_desc", "tarea_id_hecho",
    "bitacora_registrado_por", "bitacora_pending",
    "venc_pendientes", "venc_idx",
    "uso_data",
    "vacacion_data",
    "editing_item_idx", "editing_field_label",
)

_TS = "flujo_ts"


def flujos_abiertos(user_data) -> list[str]:
    """Nombres de los flujos con estado abierto ('deposito', 'tarea', ...)."""
    return [c[:-6] if c.endswith("_state") else c
            for c in CLAVES_ESTADO if user_data.get(c)]


def limpiar_flujos(user_data) -> None:
    """Cierra todos los flujos. No toca nada que no sea de un flujo."""
    for clave in CLAVES_ESTADO + CLAVES_DATOS:
        if clave in user_data:
            user_data[clave] = None
    user_data.pop(_TS, None)


def revisar_flujos(user_data, ahora: float | None = None) -> str | None:
    """Descarta los flujos vencidos. Devuelve cual se descarto, o None.

    - Sin flujo abierto: no hace nada (y suelta la marca de tiempo).
    - Flujo visto por primera vez: le pone fecha y lo deja seguir.
    - Flujo vivo: lo deja seguir SIN refrescarle la fecha.
    - Flujo pasado de MINUTOS_VIDA: lo cierra para que el mensaje siga su camino.
    """
    if ahora is None:
        import time
        ahora = time.time()

    abiertos = flujos_abiertos(user_data)
    if not abiertos:
        user_data.pop(_TS, None)
        return None

    desde = user_data.get(_TS)
    if desde is None:
        user_data[_TS] = ahora          # primera vez que lo vemos
        return None

    if ahora - desde <= MINUTOS_VIDA * 60:
        return None                     # sigue vivo; OJO: no se refresca

    limpiar_flujos(user_data)
    logger.info("Flujo(s) vencido(s) tras %d min, se descartan: %s",
                MINUTOS_VIDA, ", ".join(abiertos))
    return abiertos[0]
