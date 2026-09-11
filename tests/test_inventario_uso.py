# -*- coding: utf-8 -*-
"""Descontar un insumo del inventario sin inventar productos ni saldos.

🔴 QUE PASO EL 10-SEP-2026
Juan hizo /uso y dijo "Katana". El inventario tiene "KATANA 1 KG - herbicida",
pero el buscador exigia coincidencia EXACTA, asi que no lo encontro y
`registrar_uso` creo una fila NUEVA con stock -5, sin avisar a nadie. Tambien
quedo una fila con el producto en blanco. Y ya habia un "Agrocupper" en -2,1 de
agosto: el mismo agujero, dos veces.

LO QUE NO SE PUEDE HACER: adivinar. Hay TRES filas de Ripper Full en el
inventario ("Ripper Full", "RIPPER FULL SL - herbicida fit", "RIPPER FULL SL
20 LT - herbici"). Elegir una al azar es peor que preguntar.
"""
import pytest

from inventario_manager import buscar_producto

NOMBRES = [
    "Ripper Full",
    "Agrocupper",
    "KATANA 1 KG - herbicida",
    "RIPPER FULL SL - herbicida fitosanitario",
    "RIPPER FULL SL 20 LT - herbicida",
    "Nordox Super 75 Wp",
]


def test_el_nombre_exacto_gana_aunque_haya_parecidos():
    """'Ripper Full' existe tal cual: no puede quedar ambiguo con los otros dos."""
    idx, cand = buscar_producto(NOMBRES, "Ripper Full")
    assert NOMBRES[idx] == "Ripper Full"
    assert cand == []


def test_el_nombre_corto_calza_con_el_comercial_largo():
    """El caso Katana: Juan escribe el nombre corto, el inventario el largo."""
    idx, cand = buscar_producto(NOMBRES, "Katana")
    assert NOMBRES[idx] == "KATANA 1 KG - herbicida"


def test_no_importan_mayusculas_ni_espacios_de_mas():
    idx, _ = buscar_producto(NOMBRES, "  nordox   super 75 wp ")
    assert NOMBRES[idx] == "Nordox Super 75 Wp"


def test_un_nombre_ambiguo_NO_elige_uno_al_azar():
    """'Ripper full sl' calza con dos: hay que preguntar, no adivinar."""
    idx, cand = buscar_producto(NOMBRES, "Ripper full sl")
    assert idx is None
    assert len(cand) == 2
    assert all("RIPPER FULL SL" in c for c in cand)


def test_un_producto_que_no_existe_no_calza_con_nada():
    idx, cand = buscar_producto(NOMBRES, "Glifosato")
    assert idx is None
    assert cand == []


def test_un_nombre_vacio_no_calza_con_nada():
    """Asi nacio la fila con el producto en blanco."""
    assert buscar_producto(NOMBRES, "") == (None, [])
    assert buscar_producto(NOMBRES, "   ") == (None, [])
    assert buscar_producto(NOMBRES, None) == (None, [])


def test_una_sola_letra_no_alcanza_para_elegir():
    """'k' calzaria con Katana por prefijo, pero eso no es identificar nada."""
    idx, _ = buscar_producto(NOMBRES, "k")
    assert idx is None


# ── Registrar el uso sin inventar productos ni saldos ──────────────────────

from openpyxl import Workbook   # noqa: E402

from inventario_manager import (APLICACIONES_HEADERS, INVENTARIO_HEADERS,
                                registrar_uso)   # noqa: E402


def _libro(tmp_path, inventario):
    """`inventario`: (producto, unidad, stock)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventario"
    ws.append(INVENTARIO_HEADERS)
    for prod, uni, stock in inventario:
        ws.append([prod, "Herbicida", uni, stock, 0, "2026-08-01", None])
    wa = wb.create_sheet("Aplicaciones")
    wa.append(APLICACIONES_HEADERS)
    ruta = tmp_path / "inv.xlsx"
    wb.save(ruta)
    return str(ruta)


def _leer(ruta):
    from openpyxl import load_workbook
    wb = load_workbook(ruta, data_only=True)
    inv = [r for r in wb["Inventario"].iter_rows(min_row=2, values_only=True)
           if r and r[0]]
    app = [r for r in wb["Aplicaciones"].iter_rows(min_row=2, values_only=True)
           if r and any(r)]
    wb.close()
    return inv, app


def test_descuenta_del_producto_que_calza_aunque_el_nombre_sea_corto(tmp_path):
    ruta = _libro(tmp_path, [("KATANA 1 KG - herbicida", "Kg", 5)])
    r = registrar_uso("Katana", 2, "NOGALES", path=ruta)
    inv, app = _leer(ruta)
    assert len(inv) == 1                     # NO crea una fila nueva
    assert inv[0][3] == 3
    assert r["stock_restante"] == 3
    assert len(app) == 1


def test_un_producto_desconocido_NO_crea_una_fila_en_negativo(tmp_path):
    """Así nacieron "Katana -5" y "Agrocupper -2,1". El uso se anota igual
    --el trabajo se hizo-- pero el inventario no se inventa un producto."""
    ruta = _libro(tmp_path, [("KATANA 1 KG - herbicida", "Kg", 5)])
    r = registrar_uso("Glifosato", 3, "NOGALES", path=ruta)
    inv, app = _leer(ruta)
    assert len(inv) == 1                     # sigue habiendo UNA sola
    assert r["producto_desconocido"] is True
    assert len(app) == 1                     # pero el uso quedó anotado


def test_un_nombre_ambiguo_no_escribe_nada_y_devuelve_los_candidatos(tmp_path):
    """Tres Ripper Full: hay que preguntar, no elegir."""
    ruta = _libro(tmp_path, [("RIPPER FULL SL - herbicida", "L", 80),
                             ("RIPPER FULL SL 20 LT - herbicida", "L", 60)])
    r = registrar_uso("Ripper full sl", 60, "NOGALES", path=ruta)
    inv, app = _leer(ruta)
    assert r["error"] == "ambiguo"
    assert len(r["candidatos"]) == 2
    assert [x[3] for x in inv] == [80, 60]   # no se tocó el stock
    assert app == []                         # ni se anotó el uso


def test_un_producto_vacio_no_escribe_nada(tmp_path):
    """La fila con el producto en blanco del 10-sep nació así."""
    ruta = _libro(tmp_path, [("KATANA 1 KG - herbicida", "Kg", 5)])
    r = registrar_uso("", 5, "NOGALES", path=ruta)
    inv, app = _leer(ruta)
    assert r["error"] == "sin_producto"
    assert len(inv) == 1
    assert app == []


def test_si_el_stock_no_alcanza_se_anota_pero_se_avisa(tmp_path):
    """No se puede negar lo que se usó, pero tampoco callar que queda en rojo."""
    ruta = _libro(tmp_path, [("KATANA 1 KG - herbicida", "Kg", 5)])
    r = registrar_uso("Katana", 8, "NOGALES", path=ruta)
    inv, _ = _leer(ruta)
    assert r["stock_negativo"] is True
    assert inv[0][3] == -3


# ── Unificar productos repetidos ───────────────────────────────────────────
# El dueño, 11-sep-2026: "deberían tomar por el nombre, y unificarlos en uno
# solo". Medido contra el Master: "Ripper Full" está tres veces (20 + 80 + 60 L)
# y "Agrocupper" dos.
#
# ⚠️ EL CASO QUE ROMPE: agrupar por la primera palabra juntaría los CUATRO
# Defender --Zn, Calcio, K (Potasio) y Boro-- que son productos DISTINTOS. La
# regla tiene que ser que un nombre completo sea PREFIJO DE PALABRAS del otro.

from inventario_manager import agrupar_duplicados   # noqa: E402


def test_las_tres_filas_de_ripper_full_son_un_solo_producto():
    prods = [{"nombre": "Ripper Full", "unidad": "L", "categoria": "Herbicida", "stock": 20},
             {"nombre": "RIPPER FULL SL - herbicida fitosanitario", "unidad": "L",
              "categoria": "Herbicida", "stock": 80},
             {"nombre": "RIPPER FULL SL 20 LT - herbicida", "unidad": "L",
              "categoria": "Herbicida", "stock": 60}]
    grupos = agrupar_duplicados(prods)
    assert len(grupos) == 1
    assert len(grupos[0]) == 3


def test_los_cuatro_Defender_son_productos_DISTINTOS():
    """Comparten la primera palabra y no se pueden juntar: son Zn, Calcio, K y Boro."""
    prods = [{"nombre": n, "unidad": "L", "categoria": "Fertilizante foliar", "stock": 1}
             for n in ("Defender Zn", "Defender Calcio", "Defender K (Potasio)",
                       "Defender Boro")]
    assert agrupar_duplicados(prods) == []


def test_un_producto_solo_no_forma_grupo():
    prods = [{"nombre": "Nordox Super 75 Wp", "unidad": "Kg",
              "categoria": "Fungicida", "stock": 3}]
    assert agrupar_duplicados(prods) == []


def test_agrupa_aunque_cambien_mayusculas_y_espacios():
    prods = [{"nombre": "katana", "unidad": "Kg", "categoria": "Herbicida", "stock": 1},
             {"nombre": "KATANA  1 KG - herbicida", "unidad": "Kg",
              "categoria": "Herbicida", "stock": 4}]
    assert len(agrupar_duplicados(prods)) == 1
