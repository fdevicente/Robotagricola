# -*- coding: utf-8 -*-
"""El acumulado de labores sale de la Bitácora del Master.

OJO CON LOS FALSOS HERMANOS: "Poda nogales", "Pintar poda nogales" y "Sacar
restos poda nogales" comparten la palabra pero son TRES trabajos distintos.
Agrupar por "poda" a secas los colapsa y borra justo lo que importa: medido
contra el Master real, sacar los restos costo mas del doble de jornadas que
podar.
"""
from openpyxl import Workbook

from modules.labores import resumen_labores

ENC = ["Fecha", "Hora", "Tipo", "Actividad", "Cultivo", "Sector",
       "Jornadas Hombre", "Trabajadores", "Insumo", "Cantidad", "Unidad",
       "Registro", "Registrado por", "Máquina", "Odómetro", "Horas Día",
       "Superficie ha", "Días Cubiertos"]


def _fila(fecha, actividad, jh, trabajadores, cultivo="NOGALES", sector="",
          tipo="LABOR"):
    r = [None] * len(ENC)
    r[0], r[2], r[3] = fecha, tipo, actividad
    r[4], r[5], r[6], r[7] = cultivo, sector, jh, trabajadores
    return r


def _libro(tmp_path, filas):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(ENC)
    for f in filas:
        ws.append(f)
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_la_misma_labor_escrita_distinto_queda_en_un_grupo(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "Felicito Amigo, Ramiro Amigo"),
        _fila("2026-06-11", "poda  nogales", 1, "Felicito Amigo"),
        _fila("2026-06-12", "Poda de nogales", 1, "Felicito Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert len(r) == 1
    assert r[0]["jornadas"] == 4
    assert r[0]["dias"] == 3
    assert set(r[0]["etiquetas"]) == {"Poda nogales", "poda  nogales",
                                      "Poda de nogales"}


def test_poda_pintar_y_sacar_restos_son_tres_labores(tmp_path):
    """Comparten la palabra 'poda' y son tres trabajos distintos."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-06-10", "Pintar poda nogales", 1, "B"),
        _fila("2026-06-10", "Sacar restos poda nogales", 3, "C"),
    ])
    nombres = {x["labor"] for x in resumen_labores(path=ruta)}
    assert len(nombres) == 3


def test_cuenta_personas_distintas_no_apariciones(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Desaguar", 2, "Felicito Amigo, Ramiro Amigo"),
        _fila("2026-06-11", "Desaguar", 1, "Felicito Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert r[0]["personas"] == 2
    assert r[0]["jornadas"] == 3


def test_una_fila_sin_jornadas_suma_dia_pero_no_jornadas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Aseo general", None, "Felicito Amigo"),
        _fila("2026-06-11", "Aseo general", 2, "Felicito Amigo, Ramiro Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert r[0]["jornadas"] == 2
    assert r[0]["dias"] == 2


def test_la_lectura_de_horometro_no_es_una_labor(tmp_path):
    """La escribe el propio bot al guardar una lectura."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Lectura de horómetro", None, ""),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda"}


def test_sin_hoja_bitacora_devuelve_vacio(tmp_path):
    wb = Workbook()
    ruta = tmp_path / "vacio.xlsx"
    wb.save(ruta)
    assert resumen_labores(path=str(ruta)) == []


def test_un_excel_ilegible_no_revienta(tmp_path):
    assert resumen_labores(path=str(tmp_path / "no_existe.xlsx")) == []


def test_filtra_por_rango_de_fechas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-08-10", "Poda nogales", 3, "A"),
    ])
    r = resumen_labores(desde="2026-08-01", path=ruta)
    assert r[0]["jornadas"] == 3


# ── Lo que NO es trabajo ───────────────────────────────────────────────────
# Medido contra el Master: la hoja mezcla labores con ausencias y con las filas
# que escribe el propio bot. Por Tipo: LABOR 313 jornadas, RIEGO 41,
# APLICACION 24, MAQUINARIA 4 y OTRO 0. Sin filtrar, "Vacaciones" y "Ausente"
# salian en la lista de labores como si fueran trabajo hecho.


def test_las_ausencias_no_son_una_labor(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Vacaciones", None, "Ramiro Amigo", tipo="OTRO"),
        _fila("2026-06-10", "Ausente", None, "Javier Gonzalez", tipo="OTRO"),
        _fila("2026-06-10", "Aucente", None, "Javier Gonzalez"),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda"}


def test_lo_que_escribe_el_bot_por_la_maquina_no_es_una_labor(tmp_path):
    """Carga de combustible y horómetro roto: filas de maquinaria, sin jornadas."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Carga de combustible", None, "", tipo="MAQUINARIA"),
        _fila("2026-06-10", "Horómetro en mal estado", None, "", tipo="MAQUINARIA"),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda"}


def test_una_fila_de_maquinaria_CON_jornadas_si_es_trabajo(tmp_path):
    """Destroncar nogales entra como MAQUINARIA y es trabajo de verdad."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Destroncar nogales", 2, "A", tipo="MAQUINARIA"),
    ])
    assert resumen_labores(path=ruta)[0]["jornadas"] == 2


def test_una_aplicacion_sin_jornadas_anotadas_igual_cuenta_como_dia(tmp_path):
    """Se hizo el trabajo aunque nadie anotó las jornadas: no se puede borrar."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Aplicación fungicida", None, "", tipo="APLICACION"),
    ])
    r = resumen_labores(path=ruta)
    assert len(r) == 1
    assert r[0]["dias"] == 1
    assert r[0]["jornadas"] == 0
