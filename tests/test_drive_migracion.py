# -*- coding: utf-8 -*-
"""Dónde va cada archivo en la migración única a Drive.

Tres cosas que se descubrieron migrando de verdad, con las 7 carpetas medidas:

1. EL AÑO NO SALE DEL ARCHIVO. Los 827 archivos tienen fecha de modificación
   2026 porque Dropbox los resincronizó, pero las fechas de emisión reales van
   de 2021 a 2026 (393 son de 2025). Repartir por mtime mandaba 404 documentos
   al año equivocado. El año sale del Master.

2. LA ESTRUCTURA DE CARPETAS IMPORTA. `Legal` tiene CBR, Escrituras Públicas,
   Inscripciones y Publicaciones. Aplanarla con os.walk + basename dejó los 25
   archivos sueltos en Drive/Legal. Para `Facturas Recibidas` en cambio SÍ se
   aplana, porque ahí el orden es por año, no por la subcarpeta de Telegram.

3. HAY BASURA. Subieron `desktop.ini` y `.DS_Store` a Drive.
"""
import os

import pytest

from modules.drive.migracion import destino_de, es_basura


class TestBasura:
    @pytest.mark.parametrize("nombre", [
        ".DS_Store", ".ds_store", "desktop.ini", "Desktop.ini", "Thumbs.db",
    ])
    def test_los_archivos_de_sistema_no_se_suben(self, nombre):
        assert es_basura(nombre) is True

    @pytest.mark.parametrize("nombre", [
        "Constitución Santa Elisa.pdf", "GD198 VITAKAI.pdf", "F194.jpg",
    ])
    def test_un_documento_de_verdad_si_se_sube(self, nombre):
        assert es_basura(nombre) is False


class TestEstructura:
    """Las carpetas que no son Facturas Recibidas conservan sus subcarpetas."""

    def test_un_archivo_en_la_raiz_va_a_la_carpeta_de_arriba(self):
        d = destino_de("Legal", os.path.join("/base", "Legal", "x.pdf"),
                       os.path.join("/base", "Legal"))
        assert d == "Legal"

    def test_un_archivo_en_una_subcarpeta_conserva_la_subcarpeta(self):
        d = destino_de("Legal",
                       os.path.join("/base", "Legal", "CBR", "x.pdf"),
                       os.path.join("/base", "Legal"))
        assert d == "Legal/CBR"

    def test_la_subcarpeta_anidada_tambien(self):
        d = destino_de("Legal",
                       os.path.join("/base", "Legal", "a", "b", "x.pdf"),
                       os.path.join("/base", "Legal"))
        assert d == "Legal/a/b"

    def test_el_nombre_en_drive_puede_diferir_del_local(self):
        """'BH' en disco es 'Boletas Honorarios' en Drive."""
        d = destino_de("BH", os.path.join("/base", "BH", "x.pdf"),
                       os.path.join("/base", "BH"))
        assert d == "Boletas Honorarios"


class TestAnio:
    """Facturas Recibidas se reparte por AÑO DE EMISION, no por subcarpeta."""

    def test_va_al_anio_que_le_corresponde(self):
        d = destino_de("Facturas Recibidas",
                       os.path.join("/base", "Facturas Recibidas",
                                    "Facturas Recibidas por telegram", "x.jpg"),
                       os.path.join("/base", "Facturas Recibidas"),
                       anio="2025")
        assert d == "Facturas Recibidas/2025"

    def test_ignora_la_subcarpeta_de_telegram(self):
        """No queremos 'Facturas Recibidas/Facturas Recibidas por telegram'."""
        d = destino_de("Facturas Recibidas",
                       os.path.join("/base", "Facturas Recibidas",
                                    "Facturas Recibidas por telegram", "x.jpg"),
                       os.path.join("/base", "Facturas Recibidas"),
                       anio="2025")
        assert "telegram" not in d

    def test_sin_anio_va_a_una_carpeta_que_se_ve(self):
        """Las 227 que no calzan con el Master no se esconden dentro de 2026."""
        d = destino_de("Facturas Recibidas",
                       os.path.join("/base", "Facturas Recibidas", "x.jpg"),
                       os.path.join("/base", "Facturas Recibidas"),
                       anio=None)
        assert d == "Facturas Recibidas/Sin año"


class TestAnioDelMaster:
    """El año sale de la Fecha Emisión del Master, cruzando nº + proveedor."""

    @pytest.fixture
    def indice(self, tmp_path):
        import datetime
        import openpyxl
        from modules.drive.migracion import indice_de_anios
        ruta = tmp_path / "m.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Facturas"
        ws.append(["Fecha Emision"] + [None] * 21)
        for fecha, prov, num in [
            (datetime.date(2025, 3, 4), "ECOSMART", 264),
            (datetime.date(2026, 8, 25), "Silpa Sur Spa", 952),
            (datetime.date(2024, 1, 9), "INV. SANTA VICTORIA", 264),
        ]:
            f = [None] * 22
            f[0], f[3], f[6] = fecha, prov, num
            ws.append(f)
        wb.save(ruta)
        wb.close()
        return indice_de_anios(str(ruta))

    def test_encuentra_el_anio_por_numero_y_proveedor(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "Silpa_Sur_Spa_952.jpg") == "2026"

    def test_el_mismo_numero_de_otro_proveedor_da_otro_anio(self, indice):
        """El Nº264 lo usan ECOSMART (2025) e INV. SANTA VICTORIA (2024)."""
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "ECOSMART_264.pdf") == "2025"
        assert anio_de(indice, "INV._SANTA_VICTORIA_264.pdf") == "2024"

    def test_lo_que_no_esta_en_el_master_no_tiene_anio(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "PROVEEDOR_DESCONOCIDO_1.pdf") is None

    def test_un_nombre_que_no_es_factura_no_tiene_anio(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "cartola_agosto.pdf") is None

    def test_el_sello_de_hora_no_estorba(self, indice):
        """`Silpa_Sur_Spa_952_20260826_130058.jpg` sigue siendo la 952."""
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "Silpa_Sur_Spa_952_20260826_130058.jpg") == "2026"


class TestFechaComoTexto:
    """178 filas del Master guardan la Fecha Emisión como TEXTO, no como fecha.

    openpyxl devuelve `'2026-05-05'` (str) en vez de un datetime, y leer solo
    `.year` las dejaba a todas sin año: 97 documentos se iban a 'Sin año' sin
    motivo. El Master tiene las dos formas mezcladas y hay que aceptar ambas.
    """

    @pytest.fixture
    def indice(self, tmp_path):
        import datetime
        import openpyxl
        from modules.drive.migracion import indice_de_anios
        ruta = tmp_path / "m.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Facturas"
        ws.append(["Fecha Emision"] + [None] * 21)
        for fecha, prov, num in [
            ("2026-05-05", "COPEVAL", 100),               # texto
            (datetime.datetime(2025, 11, 24), "AGROCAMPO", 200),  # fecha
            ("", "SIN FECHA", 300),                        # vacío
            ("no es una fecha", "BASURA", 400),            # texto inservible
        ]:
            f = [None] * 22
            f[0], f[3], f[6] = fecha, prov, num
            ws.append(f)
        wb.save(ruta)
        wb.close()
        return indice_de_anios(str(ruta))

    def test_la_fecha_escrita_como_texto_igual_da_el_anio(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "COPEVAL_100.jpg") == "2026"

    def test_la_fecha_de_verdad_sigue_funcionando(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "AGROCAMPO_200.jpg") == "2025"

    def test_la_celda_vacia_no_inventa_anio(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "SIN_FECHA_300.jpg") is None

    def test_un_texto_que_no_es_fecha_no_inventa_anio(self, indice):
        from modules.drive.migracion import anio_de
        assert anio_de(indice, "BASURA_400.jpg") is None
