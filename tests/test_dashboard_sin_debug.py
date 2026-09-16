# -*- coding: utf-8 -*-
"""El dashboard no puede levantarse con el depurador abierto a la red.

Revisión del 14-sep-2026 (informe "¿Llevar el Robot a una Raspberry Pi?"):
`dashboard.py` terminaba en

    app.run(host="0.0.0.0", port=port, debug=True)

`debug=True` deja el depurador de Werkzeug en pie: cualquiera que alcance el
puerto puede ver el código cuando algo falla y, con la consola, ejecutar lo que
quiera en el PC. Y `0.0.0.0` lo ofrece a toda la red, no solo a esta máquina.
El bot manda el link como `http://localhost:5000`, así que servir en localhost
no le quita nada a nadie; abrirlo a la red es una decisión aparte, y para eso
hace falta un servidor de verdad (waitress), que hoy ni siquiera está instalado.

Se mira el código y no el servidor levantado: importar el módulo arranca Flask
y toca los archivos de secreto del dashboard.
"""
import ast
import os

RUTA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "src", "dashboard.py")


def _fuente() -> str:
    with open(RUTA, encoding="utf-8") as f:
        return f.read()


def _llamada_app_run():
    """La llamada `app.run(...)` del final del módulo."""
    arbol = ast.parse(_fuente())
    for nodo in ast.walk(arbol):
        if (isinstance(nodo, ast.Call)
                and ast.unparse(nodo.func).endswith("app.run")):
            return {k.arg: ast.unparse(k.value) for k in nodo.keywords}
    return None


def test_el_dashboard_no_arranca_con_el_depurador():
    kwargs = _llamada_app_run()
    assert kwargs is not None, "no encontré app.run en dashboard.py"
    assert kwargs.get("debug", "False") != "True", \
        "el dashboard levanta el depurador de Werkzeug"


def test_por_defecto_escucha_solo_en_esta_maquina():
    """El host puede venir de una variable: lo que importa es su valor por
    defecto, así que se mira el módulo entero y no solo la llamada."""
    kwargs = _llamada_app_run()
    assert "0.0.0.0" not in kwargs.get("host", ""), \
        "app.run se ata a toda la red"
    fuente = _fuente()
    assert "DASHBOARD_HOST" in fuente, \
        "abrir el dashboard a la red tiene que ser una decisión explícita"
    assert "127.0.0.1" in fuente, \
        "el host por defecto tiene que ser esta máquina"
