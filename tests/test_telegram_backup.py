# -*- coding: utf-8 -*-
"""El respaldo crudo de Telegram guarda el mensaje TAL CUAL llegó.

Nace del 24-ago-2026: un parte de horómetro de Juan se perdió en silencio y no
se pudo diagnosticar porque el texto original no quedaba en ninguna parte.
"""
import json
import types

import pytest

from modules.telegram_backup import guardar_update, leer_mes, buscar, ruta_del_mes


def _update(texto=None, uid=8840816610, nombre="Juan Parada", chat=8840816610,
            update_id=1, doc=None, foto=None, caption=None):
    from datetime import datetime, timezone
    msg = types.SimpleNamespace(
        message_id=99, text=texto, caption=caption, document=doc, photo=foto,
        voice=None, date=datetime(2026, 8, 24, 14, 15, tzinfo=timezone.utc))
    return types.SimpleNamespace(
        update_id=update_id, effective_message=msg,
        effective_user=types.SimpleNamespace(id=uid, full_name=nombre),
        effective_chat=types.SimpleNamespace(id=chat))


PARTE = ("Tractor jhon deere 5085\nHorometro inicio 3263\n"
         "Horometro termino 3265\nTotal horas 2\nSector 1")


def test_guarda_el_texto_tal_cual(tmp_path):
    guardar_update(_update(PARTE), base=str(tmp_path))
    filas = leer_mes("2026-08", base=str(tmp_path))
    assert len(filas) == 1
    assert filas[0]["text"] == PARTE          # sin recortar ni interpretar


def test_guarda_quien_y_cuando(tmp_path):
    guardar_update(_update("hola"), base=str(tmp_path))
    f = leer_mes("2026-08", base=str(tmp_path))[0]
    assert f["user_id"] == 8840816610
    assert f["nombre"] == "Juan Parada"
    assert f["chat_id"] == 8840816610
    assert f["fecha_mensaje_utc"].startswith("2026-08-24T14:15")
    assert f["recibido_utc"]                   # cuándo lo vio el bot


def test_es_append_no_pisa(tmp_path):
    for i in range(3):
        guardar_update(_update("msg %d" % i, update_id=i), base=str(tmp_path))
    filas = leer_mes("2026-08", base=str(tmp_path))
    assert [f["text"] for f in filas] == ["msg 0", "msg 1", "msg 2"]


def test_una_linea_json_por_mensaje(tmp_path):
    guardar_update(_update("uno"), base=str(tmp_path))
    guardar_update(_update("dos"), base=str(tmp_path))
    contenido = open(ruta_del_mes(base=str(tmp_path)), encoding="utf-8").read()
    lineas = [l for l in contenido.splitlines() if l.strip()]
    assert len(lineas) == 2
    for l in lineas:
        json.loads(l)                          # cada línea es JSON válido


def test_guarda_documentos_sin_el_binario(tmp_path):
    doc = types.SimpleNamespace(file_name="cartola.xlsx", mime_type="x/xlsx",
                                file_id="ABC123", file_size=5000)
    guardar_update(_update(caption="la cartola", doc=doc), base=str(tmp_path))
    f = leer_mes("2026-08", base=str(tmp_path))[0]
    assert f["documento"]["file_name"] == "cartola.xlsx"
    assert f["documento"]["file_id"] == "ABC123"
    assert f["caption"] == "la cartola"


def test_guarda_fotos():
    pass  # cubierto por test_foto_guarda_file_id


def test_foto_guarda_file_id(tmp_path):
    foto = [types.SimpleNamespace(file_id="FOTO1")]
    guardar_update(_update(foto=foto), base=str(tmp_path))
    assert leer_mes("2026-08", base=str(tmp_path))[0]["foto"]["file_id"] == "FOTO1"


def test_se_puede_buscar_despues(tmp_path):
    guardar_update(_update(PARTE), base=str(tmp_path))
    guardar_update(_update("otra cosa"), base=str(tmp_path))
    hits = buscar("horometro inicio", base=str(tmp_path))
    assert len(hits) == 1
    assert "3263" in hits[0]["text"]


# ── Nunca puede voltear el bot ─────────────────────────────────────────────

def test_un_update_sin_mensaje_no_rompe(tmp_path):
    upd = types.SimpleNamespace(update_id=1, effective_message=None,
                                effective_user=None, effective_chat=None)
    assert guardar_update(upd, base=str(tmp_path)) is None


def test_si_no_puede_escribir_no_lanza(tmp_path, monkeypatch):
    """Perder el respaldo es malo; perder el mensaje es peor."""
    def explota(*a, **k):
        raise OSError("disco lleno")
    monkeypatch.setattr("builtins.open", explota)
    assert guardar_update(_update("hola"), base=str(tmp_path)) is None


def test_una_linea_corrupta_no_inutiliza_el_resto(tmp_path):
    guardar_update(_update("buena"), base=str(tmp_path))
    with open(ruta_del_mes(base=str(tmp_path)), "a", encoding="utf-8") as fh:
        fh.write("{esto no es json}\n")
    guardar_update(_update("otra buena"), base=str(tmp_path))
    filas = leer_mes("2026-08", base=str(tmp_path))
    assert [f["text"] for f in filas] == ["buena", "otra buena"]


def test_mes_inexistente_devuelve_vacio(tmp_path):
    assert leer_mes("2020-01", base=str(tmp_path)) == []
