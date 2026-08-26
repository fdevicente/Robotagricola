# -*- coding: utf-8 -*-
"""El mensaje de "falta autorizar" tiene que ser copiable tal cual.

La primera version escribia la ruta con barras simples. En Python `\b` es el
caracter de retroceso, asi que el mensaje salia como
`%LOCALAPPDATA%\\Pythonin\\python3.11.exe` — una instruccion rota, justo en el
momento en que el robot le pide al dueno que reautorice.
"""
import pytest

from modules.drive.auth import FaltaAutorizacion, cargar_credenciales


def _mensaje(tmp_path):
    with pytest.raises(FaltaAutorizacion) as e:
        cargar_credenciales(token_path=str(tmp_path / "no.json"),
                            client_secret_path=str(tmp_path / "cs.json"))
    return str(e.value)


def test_el_mensaje_no_trae_caracteres_de_control(tmp_path):
    """Un retroceso o tabulacion suelta significa que una barra se leyo mal."""
    msg = _mensaje(tmp_path)
    intrusos = [c for c in msg if ord(c) < 32 and c != "\n"]
    assert not intrusos, "caracteres de control en el mensaje: %r" % intrusos


def test_la_ruta_del_python_queda_completa(tmp_path):
    msg = _mensaje(tmp_path)
    assert r"\Python\bin\python3.11.exe" in msg, msg


def test_menciona_el_script_a_correr(tmp_path):
    assert "scripts/autorizar_drive.py" in _mensaje(tmp_path)
