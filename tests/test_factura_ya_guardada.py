# -*- coding: utf-8 -*-
"""Una factura que ya está en el Master no se vuelve a guardar.

El 2-sep-2026 aparecieron **7 facturas cargadas dos veces**, $1.061.875 de
montos de ítem fantasma y $34.840 contando como deuda viva que ya estaba
pagada. Salieron de reenviar por Telegram una foto ya procesada.

La deteccion EXISTÍA y funcionaba — el log muestra
`Factura duplicada detectada: 77416674-2 Nº6945` — pero solo pintaba un
⚠️ ADVERTENCIA arriba del preview y **dejaba el botón de guardar igual de
disponible**. Un aviso que no frena nada no sirve cuando el que aprieta es el
mismo que mandó la foto de nuevo.

Además miraba `facturas_log.json`, que es un registro paralelo. La verdad está
en el Master: si el log se pierde o se rota, el duplicado pasa igual. Ahora se
consulta el Master, que es lo que el dueño quiso decir con "si ya está
guardada".
"""
import openpyxl
import pytest

from handlers.facturas import buscar_en_master

COL_PROV, COL_RUT, COL_NUM, COL_TOTAL = 4, 5, 7, 16


def _master(tmp_path, facturas):
    """facturas = [(proveedor, rut, nro, total), ...]"""
    ruta = tmp_path / "master.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append([None] * 21 + ["Archivo Drive"])
    for prov, rut, nro, total in facturas:
        f = [None] * 22
        f[COL_PROV - 1], f[COL_RUT - 1] = prov, rut
        f[COL_NUM - 1], f[COL_TOTAL - 1] = nro, total
        ws.append(f)
    wb.save(ruta)
    wb.close()
    return str(ruta)


def _item(nro, rut="77.416.674-2", prov="FERRETERIA M Y G CAMARICO SPA"):
    return {"Numero Factura / Nro Documento": nro, "Rut": rut,
            "Nombre Factura / Proveedor": prov}


class TestLaEncuentra:
    def test_la_factura_ya_guardada_se_detecta(self, tmp_path):
        """El caso real: la 6950 de Ferreteria M y G, reenviada."""
        m = _master(tmp_path, [("FERRETERIA M Y G CAMARICO SPA",
                                "77416674-2", 6950, 23300)])
        assert buscar_en_master([_item(6950)], excel_path=m)

    def test_devuelve_los_datos_para_poder_avisar_bien(self, tmp_path):
        m = _master(tmp_path, [("FERRETERIA M Y G CAMARICO SPA",
                                "77416674-2", 6950, 23300)])
        hallada = buscar_en_master([_item(6950)], excel_path=m)
        assert hallada["nro"] == "6950"
        assert hallada["total"] == 23300
        assert hallada["filas"] == [2]

    def test_el_rut_calza_aunque_este_escrito_distinto(self, tmp_path):
        """En el Master conviven '77.416.674-2' y '77416674-2'."""
        m = _master(tmp_path, [("FERRETERIA M Y G", "77.416.674-2", 6950, 23300)])
        assert buscar_en_master([_item(6950, rut="77416674-2")], excel_path=m)

    def test_el_numero_calza_aunque_venga_como_texto(self, tmp_path):
        m = _master(tmp_path, [("FERRETERIA M Y G", "77416674-2", 6950, 23300)])
        assert buscar_en_master([_item("6950")], excel_path=m)

    def test_una_factura_de_varias_lineas_devuelve_todas(self, tmp_path):
        m = _master(tmp_path, [("EFRAIN MORALES", "06170168-0", 78322, 34840)] * 4)
        hallada = buscar_en_master(
            [_item(78322, rut="06170168-0", prov="EFRAIN MORALES")], excel_path=m)
        assert hallada["filas"] == [2, 3, 4, 5]


class TestNoSeConfunde:
    def test_una_factura_nueva_NO_es_duplicado(self, tmp_path):
        m = _master(tmp_path, [("FERRETERIA M Y G", "77416674-2", 6950, 23300)])
        assert buscar_en_master([_item(9999)], excel_path=m) is None

    def test_el_mismo_numero_de_OTRO_proveedor_no_es_duplicado(self, tmp_path):
        """Hay 19 numeros que usan mas de un proveedor: es el caso de siempre."""
        m = _master(tmp_path, [("ECOSMART", "76.111.111-1", 264, 50000)])
        assert buscar_en_master(
            [_item(264, rut="77.222.222-2", prov="INV. SANTA VICTORIA")],
            excel_path=m) is None

    def test_un_master_vacio_no_rompe(self, tmp_path):
        assert buscar_en_master([_item(6950)], excel_path=_master(tmp_path, [])) is None

    def test_sin_numero_no_puede_juzgar(self, tmp_path):
        m = _master(tmp_path, [("FERRETERIA M Y G", "77416674-2", 6950, 23300)])
        assert buscar_en_master([_item("")], excel_path=m) is None

    def test_sin_items_no_rompe(self, tmp_path):
        assert buscar_en_master([], excel_path=_master(tmp_path, [])) is None


def test_un_excel_ilegible_no_voltea_el_bot(tmp_path):
    """Perder la deteccion es malo; perder el mensaje del dueño es peor."""
    malo = tmp_path / "no_existe.xlsx"
    assert buscar_en_master([_item(6950)], excel_path=str(malo)) is None


# ── El bloqueo de verdad, en el flujo ──────────────────────────────────────

class TestNoLaGuarda:
    """Detectar no basta: antes se detectaba y se guardaba igual."""

    def test_el_flujo_corta_antes_de_dejar_la_factura_pendiente(self):
        """Sin `pending_items` no hay nada que el boton de guardar pueda subir."""
        import inspect

        from handlers import facturas
        fuente = inspect.getsource(facturas._process_and_reply)
        assert "buscar_en_master" in fuente, \
            "_process_and_reply no consulta el Master"
        i_check = fuente.index("buscar_en_master")
        i_pending = fuente.index('ud["pending_items"]')
        assert i_check < i_pending, \
            "consulta el Master DESPUES de dejar la factura lista para guardar"

    def test_el_modo_capataz_tambien_queda_cubierto(self):
        """Juan guarda directo, sin preview: ahi el bloqueo importa mas."""
        import inspect

        from handlers import facturas
        fuente = inspect.getsource(facturas._process_and_reply)
        i_check = fuente.index("buscar_en_master")
        i_auto = fuente.index('ud.get("auto_mode")')
        assert i_check < i_auto, "el modo capataz guarda sin pasar por el chequeo"
