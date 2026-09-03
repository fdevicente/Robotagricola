# -*- coding: utf-8 -*-
"""Los botones salen del Master, no van escritos a mano.

Las labores cambian con la temporada: hoy poda, en marzo cosecha. Una lista fija
envejece sin que nadie lo note y Juan termina apretando "Otra..." siempre.
"""
from openpyxl import Workbook

from modules.opciones_capataz import labores_frecuentes, maquinas_recientes

CTX = {"maquinas": [
    {"maquina": "TRACTOR MASSEY FERGUSON 6711", "ultimo_odometro": 2057,
     "fecha": "2026-09-01", "unidad": "h"},
    {"maquina": "TRACTOR MASSEY FERGUSON 4292", "ultimo_odometro": 5239,
     "fecha": "2026-09-01", "unidad": "h"},
    {"maquina": "EXCAVADORA", "ultimo_odometro": 7240, "fecha": "2026-06-12",
     "unidad": "h"},
    {"maquina": "CAMIONETA RAM MODELO 1500", "ultimo_odometro": None,
     "fecha": None, "unidad": "km"},
]}


def test_las_maquinas_van_de_mas_reciente_a_mas_vieja():
    m = maquinas_recientes(CTX)
    assert m[0] in ("TRACTOR MASSEY FERGUSON 6711", "TRACTOR MASSEY FERGUSON 4292")
    assert m.index("EXCAVADORA") > m.index("TRACTOR MASSEY FERGUSON 6711")


def test_las_maquinas_sin_ninguna_lectura_van_al_final():
    """Nunca se usaron; no pueden ocupar los primeros botones."""
    m = maquinas_recientes(CTX)
    assert m[-1] == "CAMIONETA RAM MODELO 1500"


def test_no_devuelve_una_pared_de_botones():
    ctx = {"maquinas": [{"maquina": "M%d" % i, "ultimo_odometro": i,
                         "fecha": "2026-01-%02d" % (i + 1), "unidad": "h"}
                        for i in range(20)]}
    assert len(maquinas_recientes(ctx)) <= 6


def test_sin_maquinas_devuelve_lista_vacia():
    assert maquinas_recientes({"maquinas": []}) == []
    assert maquinas_recientes({}) == []


def _excel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(["Fecha", "Hora", "Tipo", "Actividad"])
    for _ in range(9):
        ws.append(["2026-08-20", "10:00", "LABOR", "Poda nogales"])
    for _ in range(4):
        ws.append(["2026-08-20", "10:00", "LABOR", "Aplicación herbicida"])
    for _ in range(7):
        ws.append(["2026-08-20", "10:00", "MAQUINARIA", "Lectura de horómetro"])
    ws.append(["2026-08-20", "10:00", "LABOR", "poda nogales"])   # misma, otra grafía
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_las_labores_van_de_mas_a_menos_usada(tmp_path):
    l = labores_frecuentes(_excel(tmp_path))
    assert l[0] == "Poda nogales"
    assert "Aplicación herbicida" in l


def test_lectura_de_horometro_no_es_una_labor(tmp_path):
    """La escribe el propio bot; ofrecersela a Juan como labor no tiene sentido."""
    assert "Lectura de horómetro" not in labores_frecuentes(_excel(tmp_path))


def test_la_misma_labor_con_otra_grafia_no_se_cuenta_dos_veces(tmp_path):
    l = labores_frecuentes(_excel(tmp_path))
    assert sum(1 for x in l if x.lower() == "poda nogales") == 1


def test_un_excel_ilegible_no_revienta(tmp_path):
    assert labores_frecuentes(str(tmp_path / "no_existe.xlsx")) == []
