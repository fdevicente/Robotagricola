"""Feriados legales de Chile, para no confundirlos con días sin registrar.

Sin esto, el chequeo de huecos de la bitácora marca como "día faltante" fiestas
en que nadie trabajó (ej: 16-jul Virgen del Carmen).

Los años se cargan a mano porque varios feriados son MOVIBLES (San Pedro y San
Pablo, Encuentro de Dos Mundos, Iglesias Evangélicas se corren al lunes o viernes
según el día en que caen). `anio_cubierto()` avisa cuando falta cargar un año, en
vez de devolver una respuesta equivocada en silencio.
"""
from datetime import date

# Fuente: Ley 2.977 y modificaciones (19.668, 19.973, 20.148, 20.299).
FERIADOS = {
    2026: {
        date(2026, 1, 1):   "Año Nuevo",
        date(2026, 4, 3):   "Viernes Santo",
        date(2026, 4, 4):   "Sábado Santo",
        date(2026, 5, 1):   "Día del Trabajo",
        date(2026, 5, 21):  "Glorias Navales",
        date(2026, 6, 21):  "Día de los Pueblos Indígenas",
        date(2026, 6, 29):  "San Pedro y San Pablo",
        date(2026, 7, 16):  "Virgen del Carmen",
        date(2026, 8, 15):  "Asunción de la Virgen",
        date(2026, 9, 18):  "Independencia Nacional",
        date(2026, 9, 19):  "Glorias del Ejército",
        date(2026, 10, 12): "Encuentro de Dos Mundos",
        date(2026, 10, 31): "Iglesias Evangélicas",
        date(2026, 11, 1):  "Día de Todos los Santos",
        date(2026, 12, 8):  "Inmaculada Concepción",
        date(2026, 12, 25): "Navidad",
    },
}


def anio_cubierto(anio: int) -> bool:
    return anio in FERIADOS


def es_feriado(d: date) -> bool:
    return d in FERIADOS.get(d.year, {})


def nombre_feriado(d: date) -> str:
    return FERIADOS.get(d.year, {}).get(d, "")


def es_habil(d: date) -> bool:
    """Lunes a viernes que no sea feriado."""
    return d.weekday() < 5 and not es_feriado(d)
