"""Maquinaria: fichas, lecturas de horómetro y mantenciones.

El riesgo mayor es que una foto de horómetro entre al lector de FACTURAS y
cree un registro basura. `parece_maquinaria` es lo que lo evita.
"""
import pytest
from openpyxl import Workbook

from handlers.maquinaria import parece_maquinaria
from modules.maquinaria import (campos_faltantes, detectar_maquina,
                                 extraer_odometro, guardar_ficha,
                                 listar_fichas, listar_mantenciones,
                                 norm_maquina, normalizar_tipo,
                                 registrar_mantencion, unidad_de)

CONOCIDAS = [
    {"maquina": "TRACTOR JOHN DEERE 5085"},
    {"maquina": "TRACTOR JOHN DEERE 5425"},
    {"maquina": "TRACTOR MASSEY FERGUSON 6711"},
    {"maquina": "TRACTOR MASSEY FERGUSON 4275"},
    {"maquina": "EXCAVADORA"},
    {"maquina": "SSANGYONG 1"},
]


@pytest.fixture
def master(tmp_path):
    wb = Workbook()
    wb.active.title = "Bitácora"
    wb.active.append(["Fecha", "Hora", "Tipo", "Actividad", "Cultivo", "Sector",
                      "Jornadas Hombre", "Trabajadores", "Insumo", "Cantidad",
                      "Unidad", "Registro", "Registrado por", "Máquina",
                      "Odómetro", "Horas Día", "Superficie ha"])
    p = tmp_path / "master.xlsx"
    wb.save(p); wb.close()
    return str(p)


# ── Nombres ──────────────────────────────────────────────────────────────

@pytest.mark.parametrize("escrito, esperado", [
    ("jd 5085", "JOHN DEERE 5085"),
    ("MF 6711", "MASSEY FERGUSON 6711"),
    ("Jhon Deere 5425", "JOHN DEERE 5425"),
    ("  excavadora  ", "EXCAVADORA"),
])
def test_normaliza_como_escribe_juan(escrito, esperado):
    assert norm_maquina(escrito) == esperado


@pytest.mark.parametrize("texto, esperada", [
    ("el 5085 quedó listo", "TRACTOR JOHN DEERE 5085"),
    ("MF 6711 horómetro 1980", "TRACTOR MASSEY FERGUSON 6711"),
    ("la excavadora trabajó todo el día", "EXCAVADORA"),
    ("john deere 5425 sigue con el horómetro malo", "TRACTOR JOHN DEERE 5425"),
])
def test_detecta_la_maquina_aunque_la_nombre_corto(texto, esperada):
    assert detectar_maquina(texto, CONOCIDAS) == esperada


def test_no_inventa_maquina_cuando_no_la_hay():
    assert detectar_maquina("hoy llovió todo el día", CONOCIDAS) is None


def test_las_camionetas_miden_kilometros():
    assert unidad_de("SSANGYONG 1") == "km"
    assert unidad_de("TRACTOR JOHN DEERE 5085") == "h"


# ── Lectura del número ───────────────────────────────────────────────────

@pytest.mark.parametrize("texto, valor", [
    ("horómetro 3200", 3200),
    ("horometro 3.166", 3166),          # punto de miles chileno
    ("odómetro 7240,7", 7240.7),        # coma decimal
    ("kilometraje 145000", 145000),
    ("JD 5085: 3200", 3200),
    ("marca 1964 horas", 1964),
])
def test_extrae_el_numero(texto, valor):
    assert extraer_odometro(texto) == pytest.approx(valor)


def test_sin_numero_devuelve_none():
    assert extraer_odometro("el horómetro está malo") is None
    assert extraer_odometro("") is None


# ── El guard de fotos ────────────────────────────────────────────────────

@pytest.mark.parametrize("caption", [
    "horómetro del 5085",
    "Odometro excavadora",
    "kilometraje camioneta",
    "mantención del tractor",
    "cambio de aceite JD 5085",
    "MF 6711 1980",
])
def test_reconoce_mensajes_de_maquinaria(caption):
    assert parece_maquinaria(caption)


@pytest.mark.parametrize("caption", [
    "",
    "Factura Copeval",
    "boleta de honorarios de Francisco",
    "foto de la poda de nogales",
])
def test_una_factura_no_se_confunde_con_maquinaria(caption):
    """Si esto falla, las facturas dejarían de procesarse."""
    assert not parece_maquinaria(caption)


# ── Fichas ───────────────────────────────────────────────────────────────

def test_guardar_y_completar_ficha_de_a_poco(master):
    guardar_ficha({"maquina": "JD 5085", "marca": "John Deere",
                   "anio": 2018}, excel_path=master)
    guardar_ficha({"maquina": "jd 5085", "patente": "ABCD12"},
                  excel_path=master)

    fichas = listar_fichas(master)
    assert len(fichas) == 1, "debe actualizar la misma ficha, no crear otra"
    f = fichas[0]
    assert f["maquina"] == "JOHN DEERE 5085"
    assert f["marca"] == "John Deere"
    assert f["anio"] == 2018
    assert f["patente"] == "ABCD12"


def test_no_borra_lo_que_ya_estaba(master):
    guardar_ficha({"maquina": "MF 6711", "marca": "Massey"}, excel_path=master)
    guardar_ficha({"maquina": "MF 6711", "marca": "", "modelo": "6711"},
                  excel_path=master)
    f = listar_fichas(master)[0]
    assert f["marca"] == "Massey"
    assert f["modelo"] == "6711"


def test_ficha_sin_nombre_se_rechaza(master):
    with pytest.raises(ValueError):
        guardar_ficha({"marca": "John Deere"}, excel_path=master)


def test_dice_que_le_falta_a_la_ficha():
    faltan = campos_faltantes({"maquina": "X", "marca": "John Deere"})
    assert "modelo" in faltan and "patente" in faltan
    assert "marca" not in faltan


# ── Mantenciones ─────────────────────────────────────────────────────────

@pytest.mark.parametrize("texto, tipo", [
    ("cambio de aceite motor", "ACEITE MOTOR"),
    ("le pusieron filtros nuevos", "FILTROS"),
    ("engrase general", "ENGRASE"),
    ("cambio de neumáticos", "NEUMATICOS"),
    ("repararon la bomba", "REPARACION"),
    ("cualquier otra cosa", "OTRO"),
])
def test_clasifica_el_tipo_de_mantencion(texto, tipo):
    assert normalizar_tipo(texto) == tipo


def test_registrar_y_listar_mantenciones(master):
    registrar_mantencion({"maquina": "JD 5085", "descripcion": "Cambio de aceite",
                          "fecha": "2026-07-20", "odometro": 3100,
                          "proveedor": "Comercial Álamos", "costo": 180_000},
                         registrado_por="juan", excel_path=master)
    registrar_mantencion({"maquina": "MF 6711", "descripcion": "Engrase"},
                         excel_path=master)

    todas = listar_mantenciones(excel_path=master)
    assert len(todas) == 2

    solo = listar_mantenciones("jd 5085", excel_path=master)
    assert len(solo) == 1
    assert solo[0]["tipo"] == "ACEITE MOTOR"
    assert solo[0]["odometro"] == 3100
    assert solo[0]["proveedor"] == "Comercial Álamos"


def test_lo_pendiente_no_se_guarda_como_hecho(master):
    """«necesita neumáticos» es algo POR hacer: no debe entrar al historial
    como realizado ni llevarse la fecha de hoy."""
    registrar_mantencion({"maquina": "MF 4275", "estado": "PENDIENTE",
                          "descripcion": "Cambiar neumáticos adelante"},
                         excel_path=master)
    registrar_mantencion({"maquina": "MF 4275", "descripcion": "Engrase"},
                         excel_path=master)

    pend = listar_mantenciones(estado="PENDIENTE", excel_path=master)
    hechas = listar_mantenciones(estado="HECHA", excel_path=master)
    assert len(pend) == 1 and len(hechas) == 1
    assert pend[0]["fecha"] is None, "lo pendiente no lleva fecha"
    assert hechas[0]["fecha"] is not None


def test_estado_invalido_cae_en_hecha(master):
    registrar_mantencion({"maquina": "X", "estado": "cualquier cosa",
                          "descripcion": "algo"}, excel_path=master)
    assert listar_mantenciones(excel_path=master)[0]["estado"] == "HECHA"


def test_mantencion_sin_maquina_se_rechaza(master):
    with pytest.raises(ValueError):
        registrar_mantencion({"descripcion": "algo"}, excel_path=master)


def test_sin_hojas_todavia_no_revienta(master):
    assert listar_fichas(master) == []
    assert listar_mantenciones(excel_path=master) == []
