# -*- coding: utf-8 -*-
"""El Master: un solo escritor a la vez, y nunca a medio escribir.

Revisión del 14-sep-2026 (informe "¿Llevar el Robot a una Raspberry Pi?"):

  · `_save_wb` escribía directo sobre el archivo: openpyxl lo vacía y lo vuelve
    a llenar, así que un corte a mitad deja el Master corrupto. El 15-sep a las
    17:29 Windows Update reinició el PC sin preguntarle a nadie.
  · No había lock. 53 funciones cargan el libro entero, lo cambian y lo
    guardan: el bot en varios hilos, sus tareas programadas y el dashboard, que
    es OTRO proceso y escribe en 7 rutas (conciliar, desconciliar, comentarios,
    rechazos, categoría del banco). Si dos se cruzan, el segundo en guardar
    borra lo del primero sin avisar.

El informe dejó escrito cuándo sale bien: "un test con dos escritores a la vez
no pierde cambios, y matar el proceso a mitad de un guardado deja intacto el
Master anterior". Son los dos primeros tests de este archivo.
"""
import ast
import os
import subprocess
import sys
import threading
import time

import openpyxl
import pytest

import config

ROBOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COL_PROVEEDOR, COL_NUMERO, COL_DRIVE = 4, 7, 22


def _master(ruta, facturas):
    """Un Master mínimo: hoja Facturas con (proveedor, número) por fila."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.cell(1, COL_PROVEEDOR, "Proveedor")
    ws.cell(1, COL_NUMERO, "Numero")
    ws.cell(1, COL_DRIVE, "Drive")
    for fila, (proveedor, numero) in enumerate(facturas, start=2):
        ws.cell(fila, COL_PROVEEDOR, proveedor)
        ws.cell(fila, COL_NUMERO, numero)
    wb.save(ruta)
    wb.close()


def _enlaces(ruta):
    wb = openpyxl.load_workbook(ruta)
    ws = wb["Facturas"]
    valores = [ws.cell(f, COL_DRIVE).value for f in range(2, ws.max_row + 1)]
    wb.close()
    return valores


# ── Lo que pide el informe ─────────────────────────────────────────────────

CORTE_A_MEDIAS = r"""
import os, sys, time
sys.path.insert(0, sys.argv[1])
import excel_manager


class LibroQueSeCorta:
    '''Escribe la mitad y se queda colgado: un guardado largo cuando se va la luz.'''
    def save(self, destino):
        with open(destino, "wb") as f:
            f.write(b"PK\x03\x04" + b"\x00" * 200000)
            f.flush()
            os.fsync(f.fileno())
        print("A_MEDIAS", flush=True)
        time.sleep(120)


excel_manager._save_wb(LibroQueSeCorta(), sys.argv[2], intentos=1, espera=0)
"""


def test_matar_el_proceso_a_mitad_del_guardado_deja_intacto_el_master(tmp_path, monkeypatch):
    locks = tmp_path / "locks"
    monkeypatch.setattr(config, "LOCK_DIR", str(locks), raising=False)
    master = tmp_path / "MASTER.xlsx"
    _master(master, [("PROVEEDOR A", 101)])
    antes = master.read_bytes()

    hijo = subprocess.Popen(
        [sys.executable, "-c", CORTE_A_MEDIAS, ROBOT, str(master)],
        stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, cwd=ROBOT,
        env=dict(os.environ, LOCK_DIR=str(locks)))
    reloj = threading.Timer(90, hijo.kill)   # si se cuelga antes, que no cuelgue la suite
    reloj.start()
    try:
        linea = hijo.stdout.readline().strip()
    finally:
        hijo.kill()
        hijo.wait(timeout=30)
        reloj.cancel()
    assert linea == "A_MEDIAS", hijo.stderr.read()

    assert master.read_bytes() == antes, "el Master quedó a medio escribir"


def test_dos_escritores_a_la_vez_no_pierden_cambios(tmp_path, monkeypatch):
    """guardar_enlace dos veces a la vez, para dos facturas distintas.

    Cada uno carga el libro, espera medio segundo (la ventana en que el otro se
    cruza) y guarda. Sin lock, los dos cargan la misma versión y el segundo en
    guardar borra el enlace del primero.
    """
    monkeypatch.setattr(config, "LOCK_DIR", str(tmp_path / "locks"), raising=False)
    master = tmp_path / "MASTER.xlsx"
    _master(master, [("PROVEEDOR A", 101), ("PROVEEDOR B", 202)])

    cargar = openpyxl.load_workbook

    def cargar_y_demorar(*a, **k):
        wb = cargar(*a, **k)
        time.sleep(0.5)
        return wb

    monkeypatch.setattr(openpyxl, "load_workbook", cargar_y_demorar)
    from modules.drive.enlaces import guardar_enlace

    errores = []

    def escribir(numero, file_id, proveedor):
        try:
            guardar_enlace(str(master), numero, file_id, proveedor)
        except Exception as e:  # noqa: BLE001 — un choque también revienta la lectura
            errores.append(repr(e))

    hilos = [threading.Thread(target=escribir, args=("101", "ID-A", "PROVEEDOR A")),
             threading.Thread(target=escribir, args=("202", "ID-B", "PROVEEDOR B"))]
    for h in hilos:
        h.start()
    for h in hilos:
        h.join(timeout=120)

    assert not errores, errores
    enlaces = _enlaces(master)
    assert all(enlaces), f"se perdió un enlace: {enlaces}"


# ── Otro proceso: el dashboard ─────────────────────────────────────────────

TOMA_EL_LOCK = r"""
import os, sys, time
from filelock import FileLock
with FileLock(os.path.join(sys.argv[1], "master.lock")):
    print("TOMADO", flush=True)
    time.sleep(float(sys.argv[2]))
"""


def test_si_otro_proceso_esta_escribiendo_espera_su_turno(tmp_path, monkeypatch):
    """El dashboard es otro proceso: un lock que solo ven los hilos del bot no sirve."""
    locks = tmp_path / "locks"
    locks.mkdir()
    monkeypatch.setattr(config, "LOCK_DIR", str(locks), raising=False)
    master = tmp_path / "MASTER.xlsx"
    _master(master, [("PROVEEDOR A", 101)])
    from modules.drive.enlaces import guardar_enlace

    otro = subprocess.Popen([sys.executable, "-c", TOMA_EL_LOCK, str(locks), "2"],
                            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, text=True)
    try:
        assert otro.stdout.readline().strip() == "TOMADO"
        inicio = time.monotonic()
        guardar_enlace(str(master), "101", "ID-A", "PROVEEDOR A")
        espero = time.monotonic() - inicio
    finally:
        otro.wait(timeout=60)

    assert espero >= 1.0, f"escribió sin esperar al otro proceso ({espero:.2f} s)"
    assert _enlaces(master) == ["https://drive.google.com/file/d/ID-A/view"]


# ── _save_wb, el único punto de escritura ──────────────────────────────────

def _libro(ruta, valor):
    wb = openpyxl.Workbook()
    wb.active["A1"] = valor
    wb.save(ruta)
    wb.close()


def _valor(ruta):
    wb = openpyxl.load_workbook(ruta)
    valor = wb.active["A1"].value
    wb.close()
    return valor


def test_guardar_deja_el_archivo_nuevo_y_ningun_temporal(tmp_path):
    from excel_manager import _save_wb
    master = tmp_path / "MASTER.xlsx"
    _libro(master, "viejo")
    wb = openpyxl.load_workbook(master)
    wb.active["A1"] = "nuevo"
    assert _save_wb(wb, str(master)) is True
    wb.close()
    assert _valor(master) == "nuevo"
    assert [p.name for p in tmp_path.iterdir()] == ["MASTER.xlsx"]


def test_si_el_master_esta_ocupado_reintenta_y_guarda(tmp_path, monkeypatch):
    """Excel abierto, o en Windows un lector a mitad de lectura: os.replace
    falla con PermissionError hasta que lo sueltan."""
    from excel_manager import _save_wb
    master = tmp_path / "MASTER.xlsx"
    _libro(master, "viejo")
    reemplazar, llamadas = os.replace, []

    def ocupado_dos_veces(origen, destino):
        llamadas.append(destino)
        if len(llamadas) <= 2:
            raise PermissionError("ocupado")
        reemplazar(origen, destino)

    monkeypatch.setattr(os, "replace", ocupado_dos_veces)
    wb = openpyxl.load_workbook(master)
    wb.active["A1"] = "nuevo"
    assert _save_wb(wb, str(master), intentos=5, espera=0) is True
    wb.close()
    assert len(llamadas) == 3
    assert _valor(master) == "nuevo"
    assert [p.name for p in tmp_path.iterdir()] == ["MASTER.xlsx"]


def test_si_el_master_sigue_ocupado_falla_sin_tocarlo_ni_dejar_temporales(tmp_path, monkeypatch):
    from excel_manager import _save_wb
    master = tmp_path / "MASTER.xlsx"
    _libro(master, "viejo")
    antes = master.read_bytes()

    def siempre_ocupado(origen, destino):
        raise PermissionError("abierto en Excel")

    monkeypatch.setattr(os, "replace", siempre_ocupado)
    wb = openpyxl.load_workbook(master)
    wb.active["A1"] = "nuevo"
    with pytest.raises(PermissionError):
        _save_wb(wb, str(master), intentos=2, espera=0)
    wb.close()
    assert master.read_bytes() == antes
    assert [p.name for p in tmp_path.iterdir()] == ["MASTER.xlsx"]


# ── Que no aparezca un escritor nuevo sin lock ─────────────────────────────

SALTAR = {"tests", "scripts", "__pycache__", ".git", ".venv"}
IMAGENES = {"img", "pix", "imagen", "image"}
NO_ES_EL_MASTER = {
    ("modules/conciliacion_export.py", "a_excel"): "arma la exportación en memoria",
}


def _es_guardado(call):
    nombre = ast.unparse(call.func)
    if nombre.endswith("_save_wb"):
        return True
    return nombre.endswith(".save") and nombre[:-len(".save")] not in IMAGENES


def _funciones_que_guardan():
    for carpeta, subs, archivos in os.walk(ROBOT):
        subs[:] = [s for s in subs if s not in SALTAR]
        for archivo in archivos:
            if not archivo.endswith(".py"):
                continue
            ruta = os.path.join(carpeta, archivo)
            rel = os.path.relpath(ruta, ROBOT).replace(os.sep, "/")
            with open(ruta, encoding="utf-8") as f:
                arbol = ast.parse(f.read())
            for nodo in ast.walk(arbol):
                if not isinstance(nodo, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                if (rel, nodo.name) in NO_ES_EL_MASTER or nodo.name == "_save_wb":
                    continue
                guardados = [c for c in ast.walk(nodo)
                             if isinstance(c, ast.Call) and _es_guardado(c)]
                if guardados:
                    yield rel, nodo, guardados


def test_toda_funcion_que_guarda_el_master_tiene_el_lock():
    """Un solo escritor sin lock basta para perder cambios."""
    sin_lock = [f"{rel}:{nodo.lineno} {nodo.name}"
                for rel, nodo, _ in _funciones_que_guardan()
                if not any(ast.unparse(d).endswith("escribe_master")
                           for d in nodo.decorator_list)]
    assert not sin_lock, ("Cargan, cambian y guardan el Master sin el lock:\n  "
                          + "\n  ".join(sin_lock))


def test_nadie_guarda_el_master_por_fuera_de_save_wb():
    """wb.save directo se salta el guardado atómico."""
    directos = [f"{rel}:{c.lineno} {nodo.name}"
                for rel, nodo, guardados in _funciones_que_guardan()
                for c in guardados if ast.unparse(c.func).endswith(".save")]
    assert not directos, "Guardan con wb.save directo:\n  " + "\n  ".join(directos)
