# -*- coding: utf-8 -*-
"""El parte tiene que quedar igual escriba /bitacora o no.

🔴 PASO EL 11-SEP-2026. Juan mando tres partes usando /bitacora y los tres
quedaron mal:

  2-sep : UNA fila "Sacar restos poda" con 11 jornadas   (debian ser 3 filas)
  10-sep: NADA                                            (debian ser 4 filas)
  11-sep: UNA fila tipo OTRO, "Asistencia y labores varias" (debian ser 4)

La causa: hay DOS caminos. El texto libre en modo capataz pasa por
`auto_guardar_bitacora`, que usa `parsear_asistencia_multi` y escribe UNA FILA
POR LABOR. Pero `/bitacora` abre un flujo y su texto va a
`handle_text_bitacora`, que lo manda a la IA, y la IA RESUME: junta todo en una
sola fila y pierde quien hizo que.

El parser determinista lee los tres partes perfecto. El problema no era leerlos:
era por donde entraban.
"""
import inspect

import handlers.bitacora as bita
from modules.bitacora_asistencia import parsear_asistencia_multi

PARTE_2SEP = """Asistencia miércoles 2 de septiembre 2026
Felicito amigo sacar restos poda nogales
Patricio Mora aplicación herbicida
Ramiro amigo aplicación herbicida
Agustín mora sacar restos poda nogales
Javier Gonzales sacar restos poda nogales
Richard padilla replante avellanos
Richard padilla crespo aplicación herbicida"""

PARTE_10SEP = """Asistencia jueves 10 de septiembre 2026
Felicito amigo poda cerezos
Patricio Mora aplicación herbicida
Ramiro amigo aplicación herbicida
Agustín mora pasar rastra
Javier Gonzales sacar amarras cerezos
Richard padilla poda cerezos
Richard padilla crespo aplicación herbicida"""

PARTE_11SEP = """Asistencia viernes 11 de septiembre 2026
Felicito amigo mantencion riego
Patricio Mora mantencion maquinaria
Ramiro amigo aseo general
Agustín mora mantencion maquinaria
Javier Gonzales aucente
Richard padilla mantencion riego
Richard padilla crespo aseo general"""


def _labores(texto):
    return [g for d in parsear_asistencia_multi(texto) for g in d["grupos"]]


def test_el_parte_del_2sep_son_tres_labores_no_una():
    """Lo que se escribió fue UNA fila de 11 jornadas."""
    labores = {g["actividad"] for g in _labores(PARTE_2SEP)}
    assert len(labores) == 3
    assert "Aplicación herbicida" in labores
    assert "Replante avellanos" in labores


def test_el_parte_del_10sep_son_cuatro_labores_no_cero():
    """Lo que se escribió fue NADA: el parte desapareció."""
    labores = {g["actividad"] for g in _labores(PARTE_10SEP)}
    assert len(labores) == 4
    assert {"Poda cerezos", "Pasar rastra", "Sacar amarras cerezos"} <= labores


def test_el_parte_del_11sep_separa_las_labores():
    """Lo que se escribió fue UNA fila tipo OTRO, 'labores varias'."""
    labores = {g["actividad"] for g in _labores(PARTE_11SEP)}
    assert {"Mantención riego", "Mantención maquinaria", "Aseo general"} <= labores


def test_una_ausencia_no_suma_jornadas():
    """'Javier Gonzales aucente' no trabajó: no puede contar como jornada."""
    por_labor = {g["actividad"]: g["jornadas_hombre"] for g in _labores(PARTE_11SEP)}
    assert por_labor.get("Aucente") in (None, 0)


def test_el_flujo_de_bitacora_usa_el_parser_determinista():
    """⚠️ ESTE ES EL TEST QUE FALTABA. Si /bitacora se salta el parser, el
    mismo parte da 3 filas por un camino y 1 por el otro."""
    fuente = inspect.getsource(bita.handle_text_bitacora)
    assert ("parsear_asistencia_multi" in fuente
            or "auto_guardar_bitacora" in fuente), (
        "handle_text_bitacora tiene que pasar por el parser determinista, "
        "igual que el texto libre: si no, /bitacora pierde labores")
