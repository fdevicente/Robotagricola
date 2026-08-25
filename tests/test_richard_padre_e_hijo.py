# -*- coding: utf-8 -*-
"""Richard Padilla (padre) y Richard Padilla Crespo (hijo) son DOS personas.

CASO REAL 2026-08-18: el hijo trabaja con ellos desde julio-2026 y se llama igual
que el padre. En el parte de Juan del 18-ago aparecen los dos:

    Richard padilla : sacar restos poda nogales
    Richard padilla crespo : sacar restos poda nogales

y el bot guardo **JH=5** en vez de 6: `_canonico()` devolvia el PRIMER nombre de
pila reconocido, y "richard" mapeaba al padre en los dos casos, asi que el
`if trabajador not in g["trabajadores"]` se comia al hijo.

No es deduplicacion: es perder a un trabajador y subcontar su jornada.
"""
import pytest

from modules.bitacora_asistencia import _canonico, parsear_asistencia

PARTE_DEL_18 = """Martes 18 de agosto 2026
Felicito amigo : sacar restos poda nogales
Patricio Mora : vacaciones
Ramiro amigo : sacar restos poda nogales
Agustín mora : sacar restos poda nogales
Javier Gonzales : sacar restos poda nogales
Richard padilla : sacar restos poda nogales
Richard padilla crespo : sacar restos poda nogales"""


# ── Distinguir padre de hijo ────────────────────────────────────────────────

def test_padre_solo_con_apellido():
    assert _canonico("Richard padilla") == "Richard Padilla"


def test_hijo_lleva_el_segundo_apellido():
    assert _canonico("Richard padilla crespo") == "Richard Padilla Crespo"


def test_hijo_aunque_falte_el_primer_apellido():
    assert _canonico("Richard crespo") == "Richard Padilla Crespo"


def test_richard_a_secas_es_el_padre():
    """Ambiguo a proposito: el padre es el que lleva anios, es el default."""
    assert _canonico("richard") == "Richard Padilla"


def test_no_distingue_por_mayusculas_ni_tildes():
    assert _canonico("RICHARD PADILLA CRESPO") == "Richard Padilla Crespo"
    assert _canonico("richard padilla crespo") == "Richard Padilla Crespo"


# ── El parte real ───────────────────────────────────────────────────────────

def test_el_parte_del_18_cuenta_seis_jornadas():
    grupos = parsear_asistencia(PARTE_DEL_18)
    poda = [g for g in grupos if "poda" in g["actividad"].lower()][0]
    assert poda["jornadas_hombre"] == 6, poda["trabajadores"]


def test_el_parte_del_18_lista_a_los_dos_richard():
    grupos = parsear_asistencia(PARTE_DEL_18)
    poda = [g for g in grupos if "poda" in g["actividad"].lower()][0]
    assert "Richard Padilla" in poda["trabajadores"]
    assert "Richard Padilla Crespo" in poda["trabajadores"]


def test_patricio_de_vacaciones_no_suma_jornada():
    grupos = parsear_asistencia(PARTE_DEL_18)
    vac = [g for g in grupos if "vacacion" in g["actividad"].lower()][0]
    assert vac["jornadas_hombre"] is None
    assert vac["trabajadores"] == ["Patricio Mora"]


def test_si_el_mismo_nombre_se_repite_igual_sigue_siendo_uno():
    """La dedup real no se rompe: dos lineas identicas son una sola persona."""
    texto = ("Richard padilla : sacar restos\n"
             "Richard padilla : sacar restos\n"
             "Felicito amigo : sacar restos")
    poda = parsear_asistencia(texto)[0]
    assert poda["jornadas_hombre"] == 2
    assert poda["trabajadores"] == ["Felicito Amigo", "Richard Padilla"]


# ── Nadie mas puede verse afectado ──────────────────────────────────────────

@pytest.mark.parametrize("crudo,esperado", [
    ("Felicito amigo", "Felicito Amigo"),
    ("Ramiro amigo", "Ramiro Amigo"),
    ("Agustín mora", "Agustin Mora"),
    ("Patricio Mora", "Patricio Mora"),
    ("pato", "Patricio Mora"),
    ("Javier Gonzales", "Javier Gonzalez"),
    ("Javier Gonzalez", "Javier Gonzalez"),
    ("Juan Parada", "Juan Parada"),
])
def test_el_resto_del_equipo_no_cambia(crudo, esperado):
    assert _canonico(crudo) == esperado


def test_un_desconocido_sigue_devolviendo_none():
    assert _canonico("Pedro Perez") is None
