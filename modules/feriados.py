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
    # 2027. Los movibles ya resueltos:
    #   San Pedro y San Pablo: el 29-jun cae MARTES -> lunes 28 (Ley 19.668).
    #   Encuentro de Dos Mundos: el 12-oct cae MARTES -> lunes 11 (Ley 19.668).
    #   Iglesias Evangelicas: el 31-oct cae DOMINGO -> se queda (la Ley 20.299
    #     solo lo mueve si cae martes o miercoles).
    #   Semana Santa: Pascua es el domingo 28-mar, asi que Viernes Santo es el
    #     26 y Sabado Santo el 27.
    #   El 18-sep cae SABADO, asi que no aplica el dia extra de la Ley 20.215.
    2027: {
        date(2027, 1, 1):   "Año Nuevo",
        date(2027, 3, 26):  "Viernes Santo",
        date(2027, 3, 27):  "Sábado Santo",
        date(2027, 5, 1):   "Día del Trabajo",
        date(2027, 5, 21):  "Glorias Navales",
        date(2027, 6, 21):  "Día de los Pueblos Indígenas",
        date(2027, 6, 28):  "San Pedro y San Pablo",
        date(2027, 7, 16):  "Virgen del Carmen",
        date(2027, 8, 15):  "Asunción de la Virgen",
        date(2027, 9, 18):  "Independencia Nacional",
        date(2027, 9, 19):  "Glorias del Ejército",
        date(2027, 10, 11): "Encuentro de Dos Mundos",
        date(2027, 10, 31): "Iglesias Evangélicas",
        date(2027, 11, 1):  "Día de Todos los Santos",
        date(2027, 12, 8):  "Inmaculada Concepción",
        date(2027, 12, 25): "Navidad",
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
