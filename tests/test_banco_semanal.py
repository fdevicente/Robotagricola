# -*- coding: utf-8 -*-
"""El sync del banco pasa a ser SEMANAL (viernes am) y siempre pide la cartola.

Decision del dueno 2026-08-24: el scraper es fragil (Akamai lo tumbo en agosto y
volvio solo), asi que en vez de intentarlo 2 veces al dia se intenta UNA vez por
semana, y si falla el bot pide subir el texto de la cartola, que es la via que
siempre funciona.

OJO CON EL DIA: en python-telegram-bot >= 20 el parametro `days` de run_daily va
de 0=domingo a 6=sabado. Antes era 0=lunes. Con el mapeo viejo, days=(4,) caia
en viernes; hoy cae en JUEVES. Por eso hay un test que fija el mapeo.
"""
import pytest

from handlers.finanzas import _aviso_banco_manual


# ── El dia correcto ─────────────────────────────────────────────────────────

def test_en_ptb_moderno_el_viernes_es_el_5():
    """Si una actualizacion de PTB cambia el mapeo, este test avisa."""
    from telegram.ext._jobqueue import JobQueue
    assert JobQueue._CRON_MAPPING[5] == "fri"
    assert JobQueue._CRON_MAPPING.index("fri") == 5


def test_el_job_del_banco_quedo_solo_el_viernes():
    """Lee main.py: un solo run_daily del banco y con days=(5,)."""
    import re
    from pathlib import Path
    src = Path(__file__).resolve().parent.parent / "main.py"
    texto = src.read_text(encoding="utf-8")
    # hasta el name=, porque dentro de la llamada hay parentesis de dtime(...)
    llamadas = re.findall(r"run_daily\(\s*job_sync_banco.*?name=\"[^\"]+\"",
                          texto, re.S)
    assert len(llamadas) == 1, "deberia haber UNA sola programacion del banco"
    assert "days=(5,)" in llamadas[0], llamadas[0]
    assert "hour=18" not in llamadas[0], "ya no corre a las 18:00"
    assert 'name="banco_viernes"' in llamadas[0], llamadas[0]


# ── Todos los jobs semanales corren el dia que dicen ────────────────────────

def _llamada(nombre_job):
    import re
    from pathlib import Path
    src = Path(__file__).resolve().parent.parent / "main.py"
    texto = src.read_text(encoding="utf-8")
    m = re.findall(r"run_daily\(\s*%s.*?name=\"[^\"]+\"" % nombre_job, texto, re.S)
    assert len(m) == 1, "esperaba una sola programacion de %s" % nombre_job
    return m[0]


@pytest.mark.parametrize("job,dia_esperado", [
    ("job_sync_banco",     5),   # viernes
    ("job_resumen_semanal", 1),  # lunes
    ("job_bodega_check",    1),  # lunes
])
def test_los_jobs_semanales_usan_el_indice_correcto(job, dia_esperado):
    """BUG REAL: resumen_semanal y bodega_check decian 'lunes' con days=(0,),
    que en PTB >= 20 es DOMINGO. Llevaban meses corriendo el dia equivocado."""
    from telegram.ext._jobqueue import JobQueue
    llamada = _llamada(job)
    assert "days=(%d,)" % dia_esperado in llamada, llamada
    nombre_dia = {1: "mon", 5: "fri"}[dia_esperado]
    assert JobQueue._CRON_MAPPING[dia_esperado] == nombre_dia


def test_ningun_job_semanal_quedo_en_domingo_por_error():
    """days=(0,) es domingo: si alguien lo escribe pensando en lunes, avisar.

    Mira solo lineas de CODIGO: los comentarios que explican el bug viejo
    mencionan days=(0,) a proposito.
    """
    from pathlib import Path
    src = Path(__file__).resolve().parent.parent / "main.py"
    codigo = [ln for ln in src.read_text(encoding="utf-8").splitlines()
              if not ln.strip().startswith("#")]
    culpables = [ln.strip() for ln in codigo if "days=(0,)" in ln]
    assert not culpables, (
        "days=(0,) en main.py es DOMINGO en PTB >= 20, no lunes: %s" % culpables)


# ── El aviso siempre termina pidiendo la cartola ────────────────────────────

@pytest.mark.parametrize("motivo", ["captcha", "error"])
def test_siempre_pide_la_cartola(motivo):
    txt = _aviso_banco_manual("09:00", motivo=motivo, detalle="lo que sea").lower()
    assert "cartola" in txt
    assert "mánda" in txt or "manda" in txt


def test_con_captcha_explica_que_no_se_va_a_saltar():
    txt = _aviso_banco_manual("09:00", motivo="captcha").lower()
    assert "captcha" in txt or "robot" in txt
    assert "bloque" in txt          # avisa del riesgo de que bloqueen la cuenta


def test_con_error_muestra_el_detalle():
    txt = _aviso_banco_manual("09:00", motivo="error", detalle="Timeout de 30s")
    assert "Timeout de 30s" in txt


def test_el_detalle_se_recorta():
    txt = _aviso_banco_manual("09:00", motivo="error", detalle="x" * 2000)
    assert len(txt) < 1500


@pytest.mark.parametrize("detalle,pista", [
    ("no encontré el campo RUT", "página"),
    ("selector no existe", "página"),
    ("credenciales inválidas", ".env"),
    ("Executable doesn't exist", "playwright"),
])
def test_pistas_segun_el_tipo_de_error(detalle, pista):
    txt = _aviso_banco_manual("09:00", motivo="error", detalle=detalle).lower()
    assert pista.lower() in txt


def test_dice_que_es_semanal_para_no_asustar():
    """Con un solo intento por semana hay que avisar cuando vuelve a probar."""
    txt = _aviso_banco_manual("09:00", motivo="error", detalle="x").lower()
    assert "viernes" in txt or "semana" in txt


def test_incluye_la_hora():
    assert "09:00" in _aviso_banco_manual("09:00", motivo="captcha")
