# -*- coding: utf-8 -*-
"""Los partes de asistencia SIN dos puntos, con los 6 mensajes reales de Juan.

Juan dejo de escribir "Nombre : actividad" y paso a "Nombre actividad". El parser
descartaba esas lineas EN SILENCIO: del parte del 26-ago leia 2 jornadas de 7, y
de los otros cinco no leia ninguna. Son 67 jornadas-hombre.

TRES TRAMPAS QUE TIENEN QUE QUEDAR FIJADAS
1. Hay CUATRO Juanes distintos en la cuadrilla de temporada --Juan Rios, Juan
   condori, Juan jaque, Juan quiroz-- y ninguno es Juan Parada. El calce por
   nombre de pila los convertiria a todos en el jefe de campo.
2. Richard Padilla y Richard Padilla Crespo son padre e hijo, no un duplicado.
3. En el parte del 1-sep hay una linea con DOS personas pegadas, porque a Juan
   se le fue el salto de linea.
"""
from modules.bitacora_asistencia import parsear_asistencia

P_25 = """Asistencia Martes 25 de agosto 2026
Felicito amigo mantencion general
Patricio Mora mantencion general
Ramiro amigo mantencion general
Agustín mora mantencion maquinaria
Javier Gonzales mantencion maquinaria
Richard padilla mantencion general
Richard padilla crespo aseo general"""

P_26_MIXTO = """Asistencia miércoles 26 agosto 2026
Felicito amigo : mantencion planta
Patricio Mora mantencion maquinaria
Ramiro amigo : aseo general
Agustín mora mantencion planta
Javier Gonzales mantencion maquinaria
Richard padilla mantencion planta
Richard padilla crespo aseo general"""

P_31 = """Lunes 31 de agosto 2026
Felicito amigo sacar restos poda nogales
Patricio Mora aplicación herbicida nogales
Ramiro amigo sacar restos poda nogales
Agustín mora sacar restos poda nogales
Javier Gonzales sacar restos poda nogales
Richard padilla aplicación herbicida
Richard padilla crespo aplicación herbicida nogales
Juan Ríos sacar restos poda nogales
Josefina quiroga sacar restos poda nogales
Senobio Fernández sacar restos poda nogales
Juan cóndori sacar restos poda nogales
Juan jaque sacar restos poda nogales
Eusebio quiroz sacar restos poda nogales
Fabián Pacheco sacar restos poda nogales
Alejandro Cabrera sacar restos poda nogales
Mauricio González sacar restos poda nogales
Patricio abar a sacar restos poda nogales
Juan quiroz sacar restos poda nogales
Maribel abarza sacar restos poda nogales"""

P_1SEP = """Martes 1 de septiembre 2026
Felicito amigo sacar restos poda nogales
Patricio Mora aplicación herbicida
Ramiro amigo aplicación herbicida
Agustín mora sacar restos poda nogales
Javier Gonzales sacar restos poda nogales
Richard padilla replante avellanos
Richard padilla crespo aplicación herbicida
Juan Ríos sacar restos poda nogales
Josefina quiroga sacar restos poda nogales
Senobio Fernández sacar restos poda nogales
Juan cóndori sacar restos poda nogales
Juan jaque sacar restos poda nogales
Eusebio quiroz sacar restos poda nogales
Fabián Pacheco sacar restos poda nogales
Alejandro Cabrera sacar restos poda nogales Mauricio González sacar restos poda nogales
Patricio abarza sacar restos poda nogales
Abdias Vásquez replante avellanos
Maribel abarza replante avellanos"""


def gente(grupos):
    return [t for g in grupos for t in g["trabajadores"]]


def de(grupos, palabra):
    return {t for g in grupos if palabra in g["actividad"].lower()
            for t in g["trabajadores"]}


def test_el_parte_sin_ningun_dos_puntos_se_lee_entero():
    """Antes devolvia None: 7 jornadas perdidas."""
    g = parsear_asistencia(P_25)
    assert g is not None
    assert len(gente(g)) == 7, gente(g)


def test_el_parte_mixto_lee_las_siete_y_no_dos():
    """El caso que mas engana: parecia guardado con 2 de 7."""
    g = parsear_asistencia(P_26_MIXTO)
    assert len(gente(g)) == 7, gente(g)


def test_el_parte_de_temporada_lee_las_19_personas():
    g = parsear_asistencia(P_31)
    assert len(gente(g)) == 19, gente(g)


def test_no_mezcla_el_herbicida_con_la_poda():
    g = parsear_asistencia(P_31)
    poda, herb = de(g, "poda"), de(g, "herbicida")
    assert herb and not (poda & herb)
    assert "Patricio Mora" in herb


def test_los_cuatro_juanes_de_temporada_no_son_juan_parada():
    """El calce por nombre de pila los volvia a todos el jefe de campo."""
    nombres = gente(parsear_asistencia(P_31))
    juanes = [n for n in nombres if n.lower().startswith("juan")]
    assert len(juanes) == 4, juanes
    assert "Juan Parada" not in nombres
    assert len(set(juanes)) == 4, juanes


def test_el_padre_y_el_hijo_siguen_separados():
    nombres = gente(parsear_asistencia(P_25))
    assert "Richard Padilla" in nombres
    assert "Richard Padilla Crespo" in nombres


def test_dos_personas_en_una_linea_se_separan():
    """A Juan se le fue el salto de linea el 1-sep."""
    nombres = gente(parsear_asistencia(P_1SEP))
    assert len(nombres) == 19, nombres
    assert any("Alejandro" in n for n in nombres)
    assert any("Mauricio" in n for n in nombres)


def test_el_encabezado_no_es_un_trabajador():
    for texto in (P_25, P_31, P_1SEP):
        nombres = gente(parsear_asistencia(texto))
        assert not any("agosto" in n.lower() or "septiembre" in n.lower()
                       or "lunes" in n.lower() or "asistencia" in n.lower()
                       for n in nombres), nombres


def test_no_repite_a_nadie():
    for texto in (P_25, P_26_MIXTO, P_31, P_1SEP):
        nombres = gente(parsear_asistencia(texto))
        assert len(nombres) == len(set(nombres)), nombres


def test_las_jornadas_cuadran_con_la_gente():
    for texto in (P_25, P_26_MIXTO, P_31, P_1SEP):
        g = parsear_asistencia(texto)
        for grupo in g:
            assert grupo["jornadas_hombre"] == len(grupo["trabajadores"])


def test_un_texto_que_no_es_asistencia_sigue_devolviendo_None():
    assert parsear_asistencia("Hola, se acabó el petróleo") is None
    assert parsear_asistencia("") is None
