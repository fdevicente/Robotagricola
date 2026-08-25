"""La bitácora debe respetar la fecha del TRABAJO, no la del mensaje.

Juan reporta días atrasados ("hoy subo lo de ayer") y a veces manda varios
días juntos en un solo mensaje.
"""
from datetime import date

import pytest

from modules.bitacora_asistencia import fecha_de_linea, parsear_asistencia_multi

HOY = date(2026, 8, 4)   # martes


@pytest.mark.parametrize("linea, esperada", [
    ("Asistencia 3 de agosto 2026", date(2026, 8, 3)),
    ("Martes 28 julio",             date(2026, 7, 28)),
    ("lunes 15 jun 2026",           date(2026, 6, 15)),
    ("29/07/2026",                  date(2026, 7, 29)),
    ("30-07",                       date(2026, 7, 30)),
    # Sin año y en el futuro → es del año pasado, no una fecha por venir
    ("Asistencia 28 de diciembre",  date(2025, 12, 28)),
    # No son encabezados de fecha
    ("Felicito amigo : desaguar",   None),
    ("hoy llovio todo el dia",      None),
])
def test_fecha_de_linea(linea, esperada):
    assert fecha_de_linea(linea, hoy=HOY) == esperada


def test_dia_pasado_escrito_hoy():
    """Escrito el 4-ago pero el trabajo es del 3-ago."""
    dias = parsear_asistencia_multi(
        "Asistencia 3 de agosto 2026\n"
        "Felicito amigo : desaguar\n"
        "Ramiro amigo : desaguar\n"
        "Patricio Mora : mantencion maquinaria", hoy=HOY)
    assert len(dias) == 1
    assert dias[0]["fecha"] == date(2026, 8, 3)
    assert sum(g["jornadas_hombre"] for g in dias[0]["grupos"]) == 3


def test_sin_fecha_queda_none():
    """Sin encabezado no se inventa fecha: el handler usa hoy."""
    dias = parsear_asistencia_multi(
        "Felicito amigo : poda nogales\n"
        "Ramiro amigo : bajar ramas", hoy=HOY)
    assert len(dias) == 1
    assert dias[0]["fecha"] is None


def test_varios_dias_en_un_mensaje():
    """Cada bloque conserva SU fecha; no se fusionan ni se pierden."""
    dias = parsear_asistencia_multi(
        "Martes 28 julio\n"
        "Felicito amigo : riego\n"
        "Patricio Mora : riego\n"
        "29/07\n"
        "Felicito amigo : poda nogales\n"
        "Ramiro amigo : poda nogales\n"
        "Asistencia 30 de julio\n"
        "Felicito amigo : vacaciones\n"
        "Ramiro amigo : barvecho avellanos", hoy=HOY)

    assert [d["fecha"] for d in dias] == [
        date(2026, 7, 28), date(2026, 7, 29), date(2026, 7, 30)]
    # Las vacaciones no suman jornadas-hombre
    ultimo = {g["actividad"]: g["jornadas_hombre"] for g in dias[-1]["grupos"]}
    assert ultimo["Vacaciones"] is None
    assert ultimo["Barbecho avellanos"] == 1


def test_texto_libre_no_es_asistencia():
    assert parsear_asistencia_multi("hoy llovio y no se pudo trabajar",
                                     hoy=HOY) is None


def test_extractor_tolera_lista():
    """Con varios días la IA a veces devuelve una lista: no debe reventar."""
    from modules.bitacora_extractor import _normalizar
    assert _normalizar([{"tipo": "RIEGO", "actividad": "Desaguar"}])["tipo"] == "RIEGO"
    assert _normalizar([])["tipo"] == "OTRO"
    assert _normalizar(None)["tipo"] == "OTRO"
