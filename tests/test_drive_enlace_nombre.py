# -*- coding: utf-8 -*-
r"""De qué parte del nombre del archivo sale el número de factura.

BUG REAL, medido el 26-ago-2026: 124 de los 826 archivos a migrar se llaman
`PROVEEDOR_6945_20260826_130058.jpg`. El timestamp se lo pega
`_renombrar_archivo` cuando el nombre ya existía. El regex de `_enlazar` era
`_(\d+)\.ext$`, o sea sacaba **130058** — la hora, no el 6945 — y los 124
enlaces se perdían en silencio.

(Verificado: ninguno de esos 124 caía en una fila ajena, porque ningún HHMMSS
coincide con un número de factura de este Master. Se perdía el enlace, no se
corrompía nada — pero se perdían los 124.)
"""
from modules.drive.subidor import _partes_del_nombre


def test_saca_el_numero_aunque_el_nombre_traiga_timestamp():
    """El caso de los 124: el 6945 gana, no la hora."""
    p = _partes_del_nombre("FERRETERIA_M_Y_G_CAMARICO_SPA_6945_20260826_130058.jpg")
    assert p is not None
    assert p[1] == "6945"


def test_del_nombre_con_timestamp_tambien_sale_el_proveedor():
    p = _partes_del_nombre("FERRETERIA_M_Y_G_CAMARICO_SPA_6945_20260826_130058.jpg")
    assert p[0] == "FERRETERIA_M_Y_G_CAMARICO_SPA"


def test_el_nombre_normal_sigue_funcionando():
    p = _partes_del_nombre("Silpa_Sur_Spa_952.jpg")
    assert p == ("Silpa_Sur_Spa", "952")


def test_un_respaldo_del_master_no_se_intenta_enlazar():
    """`2026-08-26_16-40.xlsx` no es una factura y no tiene que colar."""
    assert _partes_del_nombre("2026-08-26_16-40.xlsx") is None


def test_un_nombre_sin_numero_no_cuela():
    assert _partes_del_nombre("cartola_agosto.pdf") is None


def test_un_numero_que_es_solo_fecha_y_hora_no_es_una_factura():
    """Sin proveedor delante, `20260826_130058.jpg` no identifica nada."""
    assert _partes_del_nombre("20260826_130058.jpg") is None
