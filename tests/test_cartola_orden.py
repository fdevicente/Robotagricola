"""El banco entrega la cartola del más NUEVO al más viejo.

Si se importa en ese orden, los varios movimientos de un mismo día quedan
invertidos en el Master y el "último saldo" que se lee es el del primer
movimiento del día, no el de cierre. Eso mostró una vez $94,3M de caja cuando
el portal decía $80,7M.
"""
from datetime import date

from modules.banco_import import _en_orden_cronologico


def _m(d, saldo, desc=""):
    return {"fecha": d, "saldo": saldo, "desc": desc,
            "doc": "", "cargo": 0.0, "abono": 0.0}


def test_invierte_una_cartola_que_viene_del_mas_nuevo_al_mas_viejo():
    """Caso real: 6 cargos el 4-ago, el archivo los lista de saldo mayor a menor."""
    archivo = [
        _m(date(2026, 8, 4), 80_703_080, "Lipigas"),      # cierre del día
        _m(date(2026, 8, 4), 84_579_080, "Comercial alamo"),
        _m(date(2026, 8, 4), 94_322_675, "SmartWays"),    # primero del día
        _m(date(2026, 8, 3), 99_322_675, "Agri For"),
    ]
    ordenado = _en_orden_cronologico(archivo)

    assert [m["saldo"] for m in ordenado] == [
        99_322_675, 94_322_675, 84_579_080, 80_703_080]
    # el último es el saldo de cierre que muestra el portal
    assert ordenado[-1]["saldo"] == 80_703_080


def test_respeta_un_archivo_que_ya_viene_del_mas_viejo_al_mas_nuevo():
    archivo = [
        _m(date(2026, 8, 3), 99_322_675),
        _m(date(2026, 8, 4), 94_322_675),
        _m(date(2026, 8, 4), 80_703_080),
    ]
    assert [m["saldo"] for m in _en_orden_cronologico(archivo)] == [
        99_322_675, 94_322_675, 80_703_080]


def test_no_revienta_con_cero_o_un_movimiento():
    assert _en_orden_cronologico([]) == []
    uno = [_m(date(2026, 8, 4), 1)]
    assert _en_orden_cronologico(uno) == uno


def test_ordena_por_fecha_aunque_el_archivo_venga_mezclado():
    archivo = [
        _m(date(2026, 7, 1), 10),
        _m(date(2026, 8, 4), 30),
        _m(date(2026, 7, 15), 20),
    ]
    fechas = [m["fecha"] for m in _en_orden_cronologico(archivo)]
    assert fechas == sorted(fechas)
