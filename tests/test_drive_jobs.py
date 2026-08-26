# -*- coding: utf-8 -*-
"""Avisos de Drive: cuota, y el resumen del comando /drive."""
import pytest

from handlers.drive_jobs import hay_que_avisar_cuota, resumen_cola
from modules.drive.cola import Cola
from tests.drive_falso import DriveFalso

GB = 1024 ** 3


# ── Aviso de cuota: antes de llenarse, no cuando ya no puede subir ──────────

def test_no_avisa_con_espacio_de_sobra():
    d = DriveFalso(cuota_usada=3 * GB, cuota_total=15 * GB)
    assert hay_que_avisar_cuota(d, umbral=0.80) is False


def test_avisa_al_pasar_el_umbral():
    d = DriveFalso(cuota_usada=13 * GB, cuota_total=15 * GB)
    assert hay_que_avisar_cuota(d, umbral=0.80) is True


def test_justo_en_el_umbral_avisa():
    d = DriveFalso(cuota_usada=12 * GB, cuota_total=15 * GB)
    assert hay_que_avisar_cuota(d, umbral=0.80) is True


def test_sin_dato_de_cuota_no_avisa():
    d = DriveFalso(cuota_usada=0, cuota_total=0)
    assert hay_que_avisar_cuota(d, umbral=0.80) is False


def test_si_la_cuota_falla_no_avisa_ni_revienta():
    class DriveRoto(DriveFalso):
        def cuota(self):
            raise ConnectionError("sin internet")
    assert hay_que_avisar_cuota(DriveRoto(), umbral=0.80) is False


# ── Resumen para /drive ────────────────────────────────────────────────────

def test_resumen_de_cola_vacia(tmp_path):
    r = resumen_cola(Cola(str(tmp_path / "c.jsonl")))
    assert r["pendientes"] == 0
    assert r["rendidos"] == 0


def test_resumen_cuenta_pendientes(tmp_path):
    c = Cola(str(tmp_path / "c.jsonl"))
    c.encolar("a.pdf", "F/2026", "a.pdf")
    c.encolar("b.pdf", "F/2026", "b.pdf")
    r = resumen_cola(c)
    assert r["pendientes"] == 2 and r["rendidos"] == 0


def test_resumen_cuenta_rendidos(tmp_path):
    c = Cola(str(tmp_path / "c.jsonl"))
    c.encolar("a.pdf", "F/2026", "a.pdf")
    iid = c.pendientes()[0]["id"]
    for _ in range(5):
        c.marcar_error(iid, "sin internet")
    r = resumen_cola(c)
    assert r["pendientes"] == 0 and r["rendidos"] == 1


def test_el_resumen_trae_el_ultimo_error_para_poder_diagnosticar(tmp_path):
    c = Cola(str(tmp_path / "c.jsonl"))
    c.encolar("a.pdf", "F/2026", "a.pdf")
    iid = c.pendientes()[0]["id"]
    for _ in range(5):
        c.marcar_error(iid, "cuota agotada")
    assert "cuota agotada" in resumen_cola(c)["ultimo_error"]
