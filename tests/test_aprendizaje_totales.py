# -*- coding: utf-8 -*-
"""Un factor de corrección "aprendido" de UNA sola factura no es aprendizaje.

BUG ENCONTRADO en la revisión del 14-sep-2026. `_aplicar_correcciones_aprendidas`
multiplica el Total Factura por un factor cuando un RUT tiene >=2 correcciones
de total parecidas, pero nunca mira si son de la MISMA factura.

Medido sobre el correcciones_log.json real (121 entradas, 15-sep-2026): el
único RUT al que hoy se le aplicaría un factor es 76.029.046-7, Adrian Barrios
e Hijo, con 0,8398 sacado de sus dos correcciones:

    2026-04-23 15:13  Nº93885  claude 40.841 -> usuario 34.300  factor 0,839842
    2026-04-23 15:18  Nº93885  claude 40.841 -> usuario 34.300  factor 0,839842

Es UNA factura que llegó dos veces (facturas_log la tiene procesada a las 15:11
y a las 15:15) y se corrigió en cada pasada. Sus ítems suman 34.300: lo
que Claude leyó mal fue el total de cabecera de ESA factura, no un sesgo del
proveedor. Y es la única de ese RUT en el Master, así que la próxima que
llegara quedaría con un 16 % menos sin que nadie lo viera.

Lo que se espera ahora:
  · hacen falta correcciones de al menos DOS facturas distintas;
  · cuando el factor se aplica, se ve: en el preview, y en el modo capataz
    (que guarda sin preview) en un mensaje aparte.
"""
import asyncio
import json
import os
import types

import pytest

import config
from processors import extractor

RUT = "76.029.046-7"


def _correccion(nro, claude, usuario, rut=RUT):
    """Una entrada como las que escribe handlers.facturas._registrar_correccion."""
    return {"timestamp": "2026-04-23T15:13:31", "rut": rut,
            "proveedor": "ADRIAN BARRIOS E HIJO LIMITADA", "nro_factura": nro,
            "campo": "Total Factura", "valor_claude": claude,
            "valor_usuario": float(usuario),
            "factor": round(usuario / claude, 6)}


def _escribir_log(entradas):
    """En la carpeta de facturas que el conftest ya desvió a un temporal."""
    ruta = os.path.join(config.DOWNLOAD_DIR, "correcciones_log.json")
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(entradas, f)


def _item(total):
    return {"Rut": RUT, "Nombre Factura / Proveedor": "ADRIAN BARRIOS E HIJO LIMITADA",
            "Numero Factura / Nro Documento": "94500", "Documento": "Factura Electronica",
            "Total Factura": total}


def _dos_facturas_que_ensenan():
    """El mismo error en dos facturas distintas: 0,8398 y 0,8403."""
    _escribir_log([_correccion("93885", 40841, 34300),
                   _correccion("94120", 59500, 50000)])


# ── El caso real ───────────────────────────────────────────────────────────

def test_una_factura_corregida_dos_veces_no_ensena_un_factor():
    """La Nº93885 de Adrian Barrios, tal cual está en el log de producción."""
    _escribir_log([_correccion("93885", 40841, 34300),
                   _correccion("93885", 40841, 34300)])
    items = [_item(52000)]
    extractor._aplicar_correcciones_aprendidas(items)
    assert items[0]["Total Factura"] == 52000, \
        "aplicó un factor sacado de una sola factura"


# ── Lo que no hay que romper ───────────────────────────────────────────────

def test_dos_facturas_distintas_con_el_mismo_error_si_ensenan():
    """No se trata de apagar el aprendizaje: con dos facturas de verdad, se aplica."""
    _dos_facturas_que_ensenan()
    items = [_item(119000)]
    extractor._aplicar_correcciones_aprendidas(items)
    assert items[0]["Total Factura"] == pytest.approx(119000 * 0.84, rel=0.001)


def test_una_factura_corregida_con_valores_distintos_no_vota_dos_veces():
    """Administradora de Ventas al Detalle, 1-sep-2026, también del log real.

    La Nº24248754 se corrigió dos veces (51.389 -> 33.420 y 37.421 -> 33.420) y
    la Nº24140536 una (755.205 -> 707.400). Las tres vienen del impuesto
    específico inventado, que ya se arregló en el extractor. Si cada factura
    votara con su ÚLTIMA corrección quedarían 0,893 y 0,937, "consistentes", y
    se le aplicaría un 0,915 a todo lo que mande ese proveedor. Contar
    facturas distintas no puede significar tirar las correcciones que
    muestran que el error no es parejo.
    """
    rut = "77215640-5"
    _escribir_log([_correccion("24248754", 51389, 33420, rut=rut),
                   _correccion("24248754", 37421, 33420, rut=rut),
                   _correccion("24140536", 755205, 707400, rut=rut)])
    items = [dict(_item(100000), Rut=rut)]
    extractor._aplicar_correcciones_aprendidas(items)
    assert items[0]["Total Factura"] == 100000


# ── Si se aplica, se tiene que ver ─────────────────────────────────────────

def test_cuando_se_aplica_deja_un_aviso_con_el_total_de_antes_y_el_de_despues():
    _dos_facturas_que_ensenan()
    items = [_item(119000)]
    avisos = []
    extractor._aplicar_correcciones_aprendidas(items, avisos)
    assert len(avisos) == 1
    assert "119,000" in avisos[0]
    assert "{:,.0f}".format(items[0]["Total Factura"]) in avisos[0]
    assert "93885" in avisos[0] and "94120" in avisos[0], \
        "el aviso no dice de qué facturas salió el factor"


def test_el_aviso_sale_una_vez_aunque_la_factura_tenga_varios_items():
    _dos_facturas_que_ensenan()
    avisos = []
    extractor._aplicar_correcciones_aprendidas(
        [_item(119000), _item(119000), _item(119000)], avisos)
    assert len(avisos) == 1


def test_si_no_se_aplica_no_hay_aviso():
    _escribir_log([_correccion("93885", 40841, 34300),
                   _correccion("93885", 40841, 34300)])
    avisos = []
    extractor._aplicar_correcciones_aprendidas([_item(52000)], avisos)
    assert avisos == []


def test_process_file_devuelve_el_aviso(monkeypatch):
    """Un aviso que se queda dentro del extractor no lo ve nadie."""
    _dos_facturas_que_ensenan()
    monkeypatch.setattr(extractor, "_scan_document", lambda *a, **k: a[0])
    monkeypatch.setattr(extractor, "_resize_image", lambda *a, **k: a[0])
    monkeypatch.setattr(extractor, "_ocr_text", lambda *a, **k: "")
    monkeypatch.setattr(extractor, "_call_ia", lambda *a, **k: [_item(119000)])
    monkeypatch.setattr(extractor, "_load_proveedores_rows", lambda *a, **k: [])
    resultado = extractor.process_file("factura.jpg")
    assert resultado["status"] == "ok"
    assert len(resultado.get("avisos") or []) == 1


# ── En el flujo del bot ────────────────────────────────────────────────────

AVISO = "⚠️ AVISO DE PRUEBA: total ajustado"


class _Estado:
    """El mensaje "Leyendo documento…" que después se convierte en el preview."""
    chat_id = 1

    def __init__(self):
        self.textos = []

    async def edit_text(self, texto, **kw):
        self.textos.append(texto)


class _Bot:
    def __init__(self):
        self.enviados = []

    async def send_message(self, chat_id=None, text=None, **kw):
        self.enviados.append(text)


@pytest.fixture
def facturas(monkeypatch):
    """_process_and_reply sin IA, sin Drive y sin tocar el Master."""
    from handlers import facturas as modulo
    resultado = {"status": "ok", "items": [_item(99971)], "duplicado": False,
                 "avisos": [AVISO]}
    monkeypatch.setattr(modulo, "process_file", lambda *a, **k: resultado)
    monkeypatch.setattr(modulo, "_renombrar_archivo", lambda *a, **k: a[0])
    monkeypatch.setattr(modulo, "encolar_documento", lambda *a, **k: None)
    monkeypatch.setattr(modulo, "buscar_en_master", lambda *a, **k: None)
    monkeypatch.setattr(modulo, "main_keyboard", lambda *a, **k: None)
    return modulo


def test_el_dueno_ve_el_aviso_en_el_preview(facturas):
    estado = _Estado()
    ctx = types.SimpleNamespace(user_data={}, bot=_Bot())
    asyncio.run(facturas._process_and_reply(None, ctx, estado, "factura.jpg"))
    assert AVISO in estado.textos[-1]


def test_en_el_modo_capataz_el_aviso_sale_en_un_mensaje_aparte(facturas, monkeypatch):
    """Juan guarda directo, sin preview: ahí el ajuste volvía a ser silencioso."""
    guardadas = []

    async def guardar(query, context, items, file_path):
        guardadas.append(items)

    monkeypatch.setattr(facturas, "_guardar_excel", guardar)
    monkeypatch.setattr(facturas, "_rut_existe", lambda *a, **k: True)
    bot = _Bot()
    ctx = types.SimpleNamespace(user_data={"auto_mode": True}, bot=bot)
    asyncio.run(facturas._process_and_reply(None, ctx, _Estado(), "factura.jpg"))
    assert guardadas, "el modo capataz dejó de guardar"
    assert AVISO in bot.enviados
