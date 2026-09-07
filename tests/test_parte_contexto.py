# -*- coding: utf-8 -*-
"""El contexto es lo que el bot ya sabe: a quien conoce y que maquinas tiene.

OJO: los trabajadores NO salen solo de la hoja Personal. Medido el 2-sep-2026,
Personal tiene 6 filas con el nombre legal completo ("Felicito Amigo Soto") y no
incluye a Richard Padilla ni a su hijo, mientras la columna Trabajadores de la
bitacora usa los 8 nombres canonicos que el bot viene usando hace meses.
Armar el contexto solo con Personal dejaria a la IA peor informada que hoy.
"""
from openpyxl import Workbook

from modules.parte_contexto import construir


def _excel(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(["Fecha", "Hora", "Tipo", "Actividad", "Cultivo", "Sector",
               "Jornadas Hombre", "Trabajadores", "Insumo", "Cantidad",
               "Unidad", "Registro", "Registrado por", "Máquina", "Odómetro",
               "Horas Día", "Superficie ha", "Días Cubiertos"])
    ws.append(["2026-08-20", "14:09", "LABOR", "Poda", "NOGALES", "", 2,
               "Richard Padilla, Richard Padilla Crespo", "", None, "",
               "texto", "Juan Parada", "", None, None, None, None])
    per = wb.create_sheet("Personal")
    per.append(["Nombre", "RUT", "Cargo", "Fecha Ingreso"])
    per.append(["Felicito Amigo Soto", "9.850.887-2", None, None])
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_trae_los_nombres_de_la_bitacora(tmp_path):
    ctx = construir(_excel(tmp_path))
    assert "Richard Padilla" in ctx["trabajadores"]
    assert "Richard Padilla Crespo" in ctx["trabajadores"]


def test_trae_los_canonicos_de_siempre_aunque_no_esten_en_el_excel(tmp_path):
    """Sin esto se perderian los apodos y la regla del padre/hijo."""
    ctx = construir(_excel(tmp_path))
    assert "Patricio Mora" in ctx["trabajadores"]
    assert ctx["alias"]["pato"] == "Patricio Mora"
    assert ctx["alias"]["richard"] == "Richard Padilla"


def test_no_repite_nombres(tmp_path):
    from modules.parte_contexto import _clave
    nombres = construir(_excel(tmp_path))["trabajadores"]
    claves = [_clave(n) for n in nombres]
    assert len(claves) == len(set(claves))


def test_las_maquinas_traen_unidad_y_ultima_lectura(tmp_path):
    from openpyxl import load_workbook
    ruta = _excel(tmp_path)
    wb = load_workbook(ruta)
    wb["Bitácora"].append(["2026-08-21", "10:00", "MAQUINARIA",
                           "Lectura de horómetro", "GENERAL", "", None, "", "",
                           None, "", "t", "Juan Parada",
                           "TRACTOR MASSEY FERGUSON 4292", 5237, None, None,
                           None])
    wb.save(ruta)
    maquinas = construir(ruta)["maquinas"]
    assert maquinas, "no devolvió ninguna máquina"
    for m in maquinas:
        assert set(m) >= {"maquina", "ultimo_odometro", "fecha", "unidad"}
    mf = [m for m in maquinas if "4292" in m["maquina"]]
    assert mf and mf[0]["ultimo_odometro"] == 5237


def test_un_excel_sin_hojas_no_revienta(tmp_path):
    wb = Workbook()
    ruta = tmp_path / "vacio.xlsx"
    wb.save(ruta)
    ctx = construir(str(ruta))
    assert ctx["trabajadores"]          # quedan los canónicos de siempre
    assert ctx["maquinas"] == []


def test_los_canonicos_van_antes_que_los_nombres_nuevos_de_personal(tmp_path):
    """El orden es el que ve el modelo: primero los nombres que el bot escribe.

    Antes se comparaba contra "Felicito Amigo Soto", pero ese nombre ya no
    aparece en ctx["trabajadores"]: _canonico lo reconoce como "Felicito Amigo"
    y no lo duplica. La comparacion de orden hay que hacerla contra alguien
    genuinamente nuevo en Personal, que es el unico caso en que Personal aporta.
    """
    from openpyxl import load_workbook
    ruta = _excel(tmp_path)
    wb = load_workbook(ruta)
    wb["Personal"].append(["Josefina Quiroga", "", None, None])
    wb.save(ruta)
    nombres = construir(ruta)["trabajadores"]
    assert nombres.index("Patricio Mora") < nombres.index("Josefina Quiroga")


def test_personal_no_duplica_a_alguien_que_ya_conocemos(tmp_path):
    """Personal guarda el nombre LEGAL y la bitacora el canonico.

    Medido el 7-sep-2026 sobre el Master real: de las 6 filas de Personal, 5 son
    el nombre legal de alguien ya conocido. Meter los dos le da a la IA dos
    nombres para la misma persona.
    """
    ctx = construir(_excel(tmp_path))
    assert "Felicito Amigo" in ctx["trabajadores"]
    assert "Felicito Amigo Soto" not in ctx["trabajadores"]


def test_personal_agrega_a_alguien_nuevo_aunque_comparta_nombre_de_pila(tmp_path):
    """El caso que el filtro anterior perdia en silencio.

    "Juan Soto Rivera" no es Juan Parada. El calce por nombre de pila lo daba
    por conocido y no entraba nunca al vocabulario.
    """
    from openpyxl import load_workbook
    ruta = _excel(tmp_path)
    wb = load_workbook(ruta)
    wb["Personal"].append(["Juan Soto Rivera", "", None, None])
    wb["Personal"].append(["Josefina Quiroga", "", None, None])
    wb.save(ruta)
    nombres = construir(ruta)["trabajadores"]
    assert "Juan Soto Rivera" in nombres
    assert "Josefina Quiroga" in nombres
    assert "Juan Parada" in nombres          # el de siempre sigue estando


def test_las_variantes_de_un_mismo_nombre_colapsan(tmp_path):
    """La columna ahora la escribe la IA: una variante suya no puede quedarse."""
    from openpyxl import load_workbook
    ruta = _excel(tmp_path)
    wb = load_workbook(ruta)
    wb["Bitácora"].append(["2026-08-21", "10:00", "LABOR", "Poda", "NOGALES", "",
                           3, "ramiro amigo, RAMIRO AMIGO, Ramiro  Amigo", "",
                           None, "", "t", "Juan Parada", "", None, None, None,
                           None])
    wb.save(ruta)
    nombres = construir(ruta)["trabajadores"]
    cuantos = sum(1 for n in nombres if "amigo" in n.lower()
                  and "ramiro" in n.lower())
    assert cuantos == 1, [n for n in nombres if "ramiro" in n.lower()]
    assert "Ramiro Amigo" in nombres, "sobrevivio la variante y no el canonico"


def test_una_variante_de_la_ia_no_desplaza_al_nombre_canonico(tmp_path):
    """La columna ahora la escribe la IA. Si emite "ramiro amigo", esa grafia
    no puede ser la que el prompt le pide al modelo usar."""
    from openpyxl import load_workbook
    ruta = _excel(tmp_path)
    wb = load_workbook(ruta)
    wb["Bitácora"].append(["2026-08-21", "10:00", "LABOR", "Poda", "NOGALES", "",
                           1, "ramiro amigo", "", None, "", "t", "Juan Parada",
                           "", None, None, None, None])
    wb.save(ruta)
    nombres = construir(ruta)["trabajadores"]
    assert "Ramiro Amigo" in nombres
    assert "ramiro amigo" not in nombres


def test_una_bitacora_vacia_no_borra_las_maquinas(tmp_path):
    """StopIteration se comia la hoja Maquinaria entera, en silencio."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.title = "Bitácora"              # existe pero sin ninguna fila
    maq = wb.create_sheet("Maquinaria")
    maq.append(["Máquina", "Marca", "Modelo"])
    maq.append(["TRACTOR MASSEY FERGUSON 4292", "Massey", "4292"])
    ruta = tmp_path / "sin_filas.xlsx"
    wb.save(ruta)
    ctx = construir(str(ruta))
    assert len(ctx["maquinas"]) == 1, ctx["maquinas"]


def test_sin_columna_trabajadores_no_revienta(tmp_path):
    """Esquema viejo de la hoja: se cae con gracia a los canonicos."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(["Fecha", "Hora", "Tipo", "Actividad"])
    ws.append(["2026-08-20", "10:00", "LABOR", "Poda"])
    ruta = tmp_path / "viejo.xlsx"
    wb.save(ruta)
    assert "Patricio Mora" in construir(str(ruta))["trabajadores"]
