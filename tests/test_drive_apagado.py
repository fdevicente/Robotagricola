# -*- coding: utf-8 -*-
"""Drive se puede apagar, y apagado no molesta.

El 5-oct-2026 la red con Drive empezó a dar "The read operation timed out" en
todo: al revisar _Entrada, al subir y hasta al leer la cuota. Una subida
(CARVALLO_63756.jpg) agotó sus 5 intentos y quedó "rendida"... y a partir de
ahí el job avisó **en cada pasada**: el mismo aviso cada 10 minutos, toda la
noche (`drive_jobs.py:97-104`). El dueño pidió apagar la subida a Drive y
dejar los documentos solo en el PC, como antes.

Apagado quiere decir apagado de verdad: no encolar, no autenticar, no avisar.
Lo local no se toca — el documento ya está guardado en el PC antes de encolarse
y el respaldo del Master sigue yendo a Dropbox.

Se enciende con DRIVE_ACTIVO=1 en el .env. Lo que quedó en la cola sigue ahí:
cuando se encienda, sube.
"""
import asyncio
import types

import pytest

import config


class _Bot:
    def __init__(self):
        self.mensajes = []

    async def send_message(self, chat_id=None, text=None, **kw):
        self.mensajes.append(text)


class _ColaFalsa:
    def __init__(self):
        self.encolados = []

    def encolar(self, ruta, carpeta, nombre):
        self.encolados.append((ruta, carpeta, nombre))
        return "id"


def _contexto():
    return types.SimpleNamespace(bot=_Bot(), bot_data={"owner_chat_id": 42})


def _cola_con_un_rendido(tmp_path):
    """Una cola como la del 5-oct: una subida que agotó sus intentos."""
    from modules.drive.cola import Cola
    ruta = str(tmp_path / "drive_cola.jsonl")
    cola = Cola(ruta, max_intentos=5)
    iid = cola.encolar(str(tmp_path / "factura.jpg"), "Facturas Recibidas/2026",
                       "factura.jpg")
    for _ in range(5):
        cola.marcar_error(iid, "The read operation timed out")
    assert len(cola.rendidos()) == 1
    return ruta


# ── Encolar ───────────────────────────────────────────────────────────────

def test_con_drive_apagado_no_se_encola_el_documento(monkeypatch, tmp_path):
    from handlers import facturas
    monkeypatch.setattr(config, "DRIVE_ACTIVO", False, raising=False)
    cola = _ColaFalsa()

    facturas.encolar_documento(str(tmp_path / "factura.jpg"),
                               fecha_emision="2026-10-05", cola=cola)

    assert cola.encolados == []


def test_con_drive_encendido_se_encola_como_siempre(monkeypatch, tmp_path):
    """Apagarlo no puede significar romperlo: encendido, sigue igual."""
    from handlers import facturas
    monkeypatch.setattr(config, "DRIVE_ACTIVO", True, raising=False)
    cola = _ColaFalsa()

    facturas.encolar_documento(str(tmp_path / "factura.jpg"),
                               fecha_emision="2026-10-05", cola=cola)

    assert len(cola.encolados) == 1
    assert cola.encolados[0][1] == "Facturas Recibidas/2026"


# ── Los jobs ──────────────────────────────────────────────────────────────

def test_apagado_el_job_de_la_cola_no_avisa_ni_se_conecta(monkeypatch, tmp_path):
    """Era el aviso cada 10 minutos: con Drive apagado no hay nada que avisar."""
    from handlers import drive_jobs
    from modules.drive import cliente

    monkeypatch.setattr(config, "DRIVE_ACTIVO", False, raising=False)
    monkeypatch.setattr(drive_jobs, "DRIVE_COLA_PATH", _cola_con_un_rendido(tmp_path))

    def no_deberia(*a, **k):
        raise AssertionError("se conectó a Drive estando apagado")

    monkeypatch.setattr(cliente, "DriveCliente", no_deberia)
    ctx = _contexto()

    asyncio.run(drive_jobs.job_drive_cola(ctx))

    assert ctx.bot.mensajes == []


def test_apagado_el_job_de_entrada_no_mira_drive(monkeypatch):
    from handlers import drive_jobs
    from modules.drive import cliente

    monkeypatch.setattr(config, "DRIVE_ACTIVO", False, raising=False)

    def no_deberia(*a, **k):
        raise AssertionError("se conectó a Drive estando apagado")

    monkeypatch.setattr(cliente, "DriveCliente", no_deberia)

    asyncio.run(drive_jobs.job_drive_entrada(_contexto()))


# ── Lo local sigue igual ──────────────────────────────────────────────────

def test_el_respaldo_del_master_se_hace_aunque_drive_este_apagado(monkeypatch, tmp_path):
    """Lo que importa del respaldo es la copia en Dropbox, no la de Drive."""
    import openpyxl

    from infrastructure.backups import backup_master
    monkeypatch.setattr(config, "DRIVE_ACTIVO", False, raising=False)

    master = tmp_path / "MASTER.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "dato"
    wb.save(master)
    wb.close()
    base = tmp_path / "Backups"
    cola = tmp_path / "drive_cola.jsonl"

    snap = backup_master("prueba", excel_path=str(master),
                         backup_base=str(base), cola_path=str(cola))

    assert snap and len(list((base / "Master" / "snapshots").glob("*.xlsx"))) == 1
    assert not cola.exists(), "encoló para Drive estando apagado"


def test_drive_dice_que_esta_apagado(monkeypatch):
    """Un /drive que muestra 'todo al día' con Drive apagado estaría mintiendo."""
    from handlers import drive_jobs
    monkeypatch.setattr(config, "DRIVE_ACTIVO", False, raising=False)

    respuestas = []

    class _Msg:
        async def reply_text(self, texto, **kw):
            respuestas.append(texto)

    update = types.SimpleNamespace(message=_Msg())

    asyncio.run(drive_jobs.cmd_drive(update, _contexto()))

    assert respuestas and "apagad" in respuestas[0].lower()
