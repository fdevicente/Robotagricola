# -*- coding: utf-8 -*-
"""La columna 'Aplica IVA' de Cosechas solo pesa en las filas 'esperado'.

Contexto real: Valbifrut paga en CLP y agrega 19% (mayo-2026: neto $223.596.523
el dia 4 + IVA $42.483.234 el dia 18). Pacific paga en USD por COMEX, sin IVA.
Antes el proyector calculaba TODA fila esperada sin IVA, subestimando a quien
paga en pesos mientras los F29 si se proyectaban como egreso.

OJO: todos los tests que escriben Excel pasan la ruta EXPLICITA. Un test destruyo
el Master real por confiar en el default de _save_wb.
"""
import openpyxl
import pytest

from modules.cash_flow.projector import load_expected_ingresos

HEADERS = ["Año", "Cultivo", "Kg total", "Exportadora", "Kg asignados",
           "Precio USD/kg", "N° cuotas", "Cuota #", "Fecha estimada",
           "Monto USD estimado", "Tipo cuota", "Estado", "Fecha real recibido",
           "Monto real recibido", "Moneda recibida", "Notas", "Aplica IVA"]


def _fila(**kw):
    base = dict(anio=2027, cultivo="NOGALES", kg=100000, exportadora="X",
                kg_asig=100000, precio=2.1, n_cuotas=1, cuota=1,
                fecha_est="2027-05-15", usd=100000, tipo="adelanto",
                estado="esperado", fecha_real=None, monto_real=None,
                moneda=None, notas="", iva=None)
    base.update(kw)
    return [base[k] for k in ("anio", "cultivo", "kg", "exportadora", "kg_asig",
                              "precio", "n_cuotas", "cuota", "fecha_est", "usd",
                              "tipo", "estado", "fecha_real", "monto_real",
                              "moneda", "notas", "iva")]


@pytest.fixture
def libro(tmp_path, monkeypatch):
    """Crea un Excel de juguete y fija el tipo de cambio en 910."""
    ruta = tmp_path / "cosechas.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cosechas"
    ws.append(HEADERS)
    wb.save(ruta)
    wb.close()

    import config
    monkeypatch.setitem(config.CASH_FLOW_CONFIG, "usd_clp_estimado", 910)
    return ruta


def _agregar(ruta, filas):
    wb = openpyxl.load_workbook(ruta)
    for f in filas:
        wb["Cosechas"].append(f)
    wb.save(ruta)   # ruta EXPLICITA, nunca el default
    wb.close()


def test_esperado_sin_iva_queda_en_el_neto(libro):
    _agregar(libro, [_fila(usd=100000, iva="NO")])
    ing = load_expected_ingresos(str(libro))
    assert len(ing) == 1
    assert ing[0]["monto_clp"] == pytest.approx(100000 * 910)
    assert ing[0]["aplica_iva"] is False


def test_esperado_con_iva_suma_19_por_ciento(libro):
    _agregar(libro, [_fila(usd=100000, iva="SI")])
    ing = load_expected_ingresos(str(libro))
    assert ing[0]["monto_clp"] == pytest.approx(100000 * 910 * 1.19)
    assert ing[0]["aplica_iva"] is True


def test_columna_vacia_se_comporta_como_no(libro):
    """Las filas viejas no tienen la columna: no deben cambiar de valor."""
    _agregar(libro, [_fila(usd=100000, iva=None)])
    ing = load_expected_ingresos(str(libro))
    assert ing[0]["monto_clp"] == pytest.approx(100000 * 910)


@pytest.mark.parametrize("marca", ["si", "Sí", " SI ", "S", "TRUE", "1"])
def test_variantes_de_si(libro, marca):
    _agregar(libro, [_fila(usd=100000, iva=marca)])
    ing = load_expected_ingresos(str(libro))
    assert ing[0]["monto_clp"] == pytest.approx(100000 * 910 * 1.19), marca


def test_recibido_ignora_la_columna(libro):
    """El efectivo real ya trae el IVA adentro: no se le suma otra vez."""
    _agregar(libro, [_fila(estado="recibido", monto_real=266079757,
                           moneda="CLP", fecha_real="2026-05-04", iva="SI")])
    ing = load_expected_ingresos(str(libro))
    assert ing[0]["monto_clp"] == pytest.approx(266079757)


def test_caso_real_valbifrut_vs_pacific(libro):
    """El adelanto de may-2027 repartido como en 2026."""
    _agregar(libro, [
        _fila(exportadora="Valbifrut", usd=210600, iva="SI"),
        _fila(exportadora="Pacific Nuts", usd=149400, iva="NO"),
    ])
    ing = load_expected_ingresos(str(libro))
    porexp = {i["exportadora"]: i["monto_clp"] for i in ing}
    assert porexp["Valbifrut"] == pytest.approx(210600 * 910 * 1.19)
    assert porexp["Pacific Nuts"] == pytest.approx(149400 * 910)
    # el IVA de Valbifrut vale ~36 millones: no es un detalle
    assert porexp["Valbifrut"] - 210600 * 910 > 36_000_000
