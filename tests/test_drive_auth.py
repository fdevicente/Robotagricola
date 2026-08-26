# -*- coding: utf-8 -*-
"""La autenticación falla con un mensaje accionable, nunca en silencio.

Con un Gmail común NO sirve una cuenta de servicio: sus archivos no tienen
cuota de Drive y la subida falla. Va OAuth con token de refresco.
"""
import pytest

from modules.drive.auth import FaltaAutorizacion, cargar_credenciales


def test_sin_archivo_de_token_avisa_que_hay_que_autorizar(tmp_path):
    with pytest.raises(FaltaAutorizacion) as e:
        cargar_credenciales(token_path=str(tmp_path / "no-existe.json"),
                            client_secret_path=str(tmp_path / "cs.json"))
    assert "autorizar_drive" in str(e.value)


def test_el_mensaje_dice_como_arreglarlo(tmp_path):
    with pytest.raises(FaltaAutorizacion) as e:
        cargar_credenciales(token_path=str(tmp_path / "no.json"),
                            client_secret_path=str(tmp_path / "cs.json"))
    msg = str(e.value).lower()
    assert "drive" in msg and "python" in msg
