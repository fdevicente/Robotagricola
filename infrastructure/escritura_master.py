# -*- coding: utf-8 -*-
"""El Master tiene un solo escritor a la vez.

El bot (en varios hilos, con sus tareas programadas), el dashboard (otro
proceso) y los scripts cargan el libro ENTERO, lo cambian y lo guardan. Si dos
se cruzan, los dos cargan la misma versión y el segundo en guardar borra lo
del primero sin avisar. Por eso el lock cubre la operación completa (cargar,
cambiar, guardar) y no solo el guardado: ver `escribe_master`.

El lock es un archivo en config.LOCK_DIR, leído al momento, así que la suite lo
desvía a un temporal. Lo toma el sistema operativo: si el proceso muere, se
suelta solo. Es reentrante dentro de un mismo hilo, así que una función con
lock puede llamar a otra que también lo pide. El guardado atómico está en
`excel_manager._save_wb`. Ver tests/test_escritura_master.py.
"""
import functools
import logging
import os

from filelock import FileLock, Timeout

logger = logging.getLogger(__name__)

# Lo máximo que un escritor espera su turno. Un guardado normal tarda unos 2 s
# (cargar 0,9 + guardar 1,0 en este PC). Pasado esto hay otro proceso colgado,
# o un lote largo corriendo con el bot vivo: mejor fallar con aviso que
# quedarse esperando para siempre.
ESPERA_MAXIMA = 120


def bloqueo() -> FileLock:
    """El lock del Master, para usar con `with bloqueo():`."""
    import config
    os.makedirs(config.LOCK_DIR, exist_ok=True)
    return FileLock(os.path.join(config.LOCK_DIR, "master.lock"),
                    timeout=ESPERA_MAXIMA, is_singleton=True)


def escribe_master(funcion):
    """Decorador: la función entera (cargar, cambiar, guardar) corre con el lock."""
    @functools.wraps(funcion)
    def con_lock(*args, **kwargs):
        try:
            with bloqueo():
                return funcion(*args, **kwargs)
        except Timeout:
            logger.error("El Master siguió ocupado %d s: %s no se ejecutó",
                         ESPERA_MAXIMA, funcion.__name__)
            raise
    return con_lock
