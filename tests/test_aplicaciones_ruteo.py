# -*- coding: utf-8 -*-
"""Los partes de aplicacion tienen que llegar a su handler.

De los 15 que mando Juan entre junio y agosto de 2026, solo 2 quedaron anotados.
Dos de las razones estan fijadas aca:

  - el porton de maquinaria se llevaba 2 por delante, porque traen numeros y
    palabras que le parecen maquina;
  - los que sobrevivian caian en la bitacora automatica, que no sabe que hacer
    con un parte sin personas.
"""
import inspect

import handlers.chat as chat
from handlers.maquinaria import parece_maquinaria
from modules.aplicaciones_parser import es_aplicacion

FUNGICIDA = """Aplicación fungicida
Miércoles 17 de junio 2026
Producto :nordox super 75 wp
Dosis :180 gramos por 100
Cerezos producción
Mojamiento 1500 litros por ha
Total ha : 2
Total Mojamiento : 3000 litros
Total producto : 5.4 kilos"""


def test_las_aplicaciones_se_miran_antes_que_el_porton_de_maquinaria():
    fuente = inspect.getsource(chat.handle_text)
    assert "procesar_aplicacion" in fuente
    assert fuente.index("procesar_aplicacion") < fuente.index("parece_maquinaria")


def test_las_aplicaciones_se_miran_antes_que_la_bitacora_automatica():
    fuente = inspect.getsource(chat.handle_text)
    assert fuente.index("procesar_aplicacion") < fuente.index("AUTO_SAVE_USERS")


def test_este_parte_real_lo_reclamaban_los_dos():
    """El caso exacto que motivó el orden: el portón SÍ lo agarra."""
    assert es_aplicacion(FUNGICIDA)
    assert parece_maquinaria(FUNGICIDA), (
        "si esto deja de ser verdad, el orden ya no importa tanto, "
        "pero el test de orden sigue siendo el que manda")
