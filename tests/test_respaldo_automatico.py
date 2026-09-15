# -*- coding: utf-8 -*-
"""El Master se respalda solo, sin borrar lo que no le toca, y avisa si falla.

🔴 Medido el 14 y 15-sep-2026:
  · La única vía automática, la tarea de Windows DailyBanco-18h, fallaba TODOS
    los días: busca py.exe, que no existe (0x80070002). Nadie se enteraba.
  · _rotate_snapshots borraba por ORDEN ALFABÉTICO. Los 22 'MASTER_pre_*'
    ordenan después de '2026-…' y ocupaban 22 de los 30 cupos para siempre: de
    47 respaldos creados desde el 26-ago, en la carpeta quedaban 8.
  · Nada comprobaba la copia: si el Master se guardaba mientras se copiaba, el
    respaldo podía quedar a medias.
"""
import asyncio
import inspect
import os
import shutil
from datetime import date, timedelta

import pytest

from infrastructure import backups


@pytest.fixture
def entorno(tmp_path):
    master = tmp_path / "master.xlsx"
    master.write_bytes(b"version 1")
    base = tmp_path / "Backups"
    return {"master": str(master), "base": str(base),
            "cola": str(tmp_path / "cola.jsonl"),
            "snaps": base / "Master" / "snapshots"}


def _backup(e):
    return backups.backup_master("test", excel_path=e["master"],
                                 backup_base=e["base"], cola_path=e["cola"])


def _respaldar_si_cambio(e):
    return backups.respaldar_si_cambio("test", excel_path=e["master"],
                                       backup_base=e["base"], cola_path=e["cola"])


# ── Retención ───────────────────────────────────────────────────────────────

def test_la_retencion_nunca_borra_los_respaldos_manuales(entorno):
    """El bug del orden alfabético: los MASTER_pre_* se comían los cupos y lo que
    terminaba borrado eran los respaldos nuevos."""
    snaps = entorno["snaps"]
    snaps.mkdir(parents=True)
    for i in range(25):
        (snaps / f"MASTER_pre_carga{i:02d}_20260609_1900{i:02d}.xlsx").write_bytes(b"x")
    recientes = [(date.today() - timedelta(days=k)).isoformat() + "_10-00.xlsx"
                 for k in range(1, 11)]
    for nombre in recientes:
        (snaps / nombre).write_bytes(b"x")

    _backup(entorno)

    nombres = {p.name for p in snaps.iterdir()}
    assert sum(n.startswith("MASTER_pre_") for n in nombres) == 25
    assert set(recientes) <= nombres


def test_la_retencion_deja_uno_por_mes_de_los_automaticos_viejos(entorno):
    snaps = entorno["snaps"]
    snaps.mkdir(parents=True)
    mes_viejo = date.today().replace(day=1) - timedelta(days=45)
    viejos = [mes_viejo.replace(day=d).isoformat() + "_10-00.xlsx" for d in (1, 10, 20)]
    for nombre in viejos:
        (snaps / nombre).write_bytes(b"x")

    _backup(entorno)

    nombres = {p.name for p in snaps.iterdir()}
    assert nombres & set(viejos) == {viejos[-1]}     # queda el último de ese mes


# ── Respaldar solo si cambió ────────────────────────────────────────────────

def test_la_primera_vez_respalda(entorno):
    snap = _respaldar_si_cambio(entorno)
    assert snap and os.path.exists(snap)


def test_si_el_master_no_cambio_no_repite(entorno):
    _respaldar_si_cambio(entorno)
    assert _respaldar_si_cambio(entorno) is None
    assert len(list(entorno["snaps"].glob("*.xlsx"))) == 1


def test_si_el_master_cambio_respalda_de_nuevo(entorno):
    primero = _respaldar_si_cambio(entorno)
    # Otro nombre, para no depender de que las dos llamadas caigan en el mismo minuto
    os.replace(primero, os.path.join(os.path.dirname(primero), "2020-01-01_10-00.xlsx"))
    with open(entorno["master"], "wb") as f:
        f.write(b"version 2")

    segundo = _respaldar_si_cambio(entorno)

    assert segundo is not None
    with open(segundo, "rb") as f:
        assert f.read() == b"version 2"


# ── La copia tiene que ser buena ────────────────────────────────────────────

def test_si_el_master_cambia_durante_la_copia_no_queda_un_respaldo_a_medias(entorno, monkeypatch):
    copia_real = shutil.copy2

    def copia_mientras_alguien_guarda(origen, destino, *a, **k):
        resultado = copia_real(origen, destino, *a, **k)
        with open(origen, "ab") as f:
            f.write(b"+")
        return resultado

    monkeypatch.setattr(shutil, "copy2", copia_mientras_alguien_guarda)
    monkeypatch.setattr("time.sleep", lambda s: None)

    with pytest.raises(RuntimeError):
        _backup(entorno)
    snaps = entorno["snaps"]
    assert not snaps.exists() or not list(snaps.iterdir())


# ── El job del bot ──────────────────────────────────────────────────────────

class _Bot:
    def __init__(self):
        self.enviados = []

    async def send_message(self, chat_id, text, **kw):
        self.enviados.append((chat_id, text))


class _Ctx:
    def __init__(self):
        self.bot = _Bot()
        self.bot_data = {"owner_chat_id": 8684368429}


def test_si_el_respaldo_falla_le_avisa_al_dueno(monkeypatch):
    """La tarea de Windows falló a diario durante semanas sin que nadie supiera."""
    from handlers import monitoreo

    def falla(*a, **k):
        raise OSError("no encuentro la carpeta de Dropbox")

    monkeypatch.setattr(backups, "respaldar_si_cambio", falla, raising=False)
    ctx = _Ctx()
    asyncio.run(monitoreo.job_respaldo_master(ctx))
    assert len(ctx.bot.enviados) == 1
    destino, texto = ctx.bot.enviados[0]
    assert destino == 8684368429
    assert "respald" in texto.lower()


def test_si_el_respaldo_anda_no_molesta(monkeypatch):
    from handlers import monitoreo
    monkeypatch.setattr(backups, "respaldar_si_cambio",
                        lambda *a, **k: "C:/x/2026-09-15_10-00.xlsx", raising=False)
    ctx = _Ctx()
    asyncio.run(monitoreo.job_respaldo_master(ctx))
    assert ctx.bot.enviados == []


def test_el_bot_programa_el_respaldo():
    import main
    fuente = inspect.getsource(main)
    assert "run_repeating(job_respaldo_master" in fuente
    assert 'name="respaldo_master"' in fuente
