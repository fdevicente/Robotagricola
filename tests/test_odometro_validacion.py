"""Un horómetro no retrocede ni salta lo imposible.

Casos reales del 10-ago-2026 que el bot aceptó sin chistar:
  · Juan escribió "Tractor jhon deere 50853200" (modelo y horómetro pegados)
    → guardó 50.853.200 y calculó 50.850.034 "horas del día".
  · "Tractor massey ferguson 4292 3200" cuando venía en 5.222 → retrocedió.
"""
import pytest

from bitacora_manager import validar_odometro


# ── Los dos casos que se colaron ─────────────────────────────────────────

def test_el_modelo_pegado_al_horometro_se_rechaza():
    motivo = validar_odometro("TRACTOR JOHN DEERE 5085", 50_853_200, 3166, dias=60)
    assert motivo and "imposible" in motivo


def test_un_horometro_que_baja_se_rechaza():
    motivo = validar_odometro("TRACTOR MASSEY FERGUSON 4292", 3200, 5222, dias=60)
    assert motivo and "no retrocede" in motivo


# ── Lo que sí debe pasar ─────────────────────────────────────────────────

@pytest.mark.parametrize("maquina, nuevo, previo, dias", [
    ("TRACTOR MASSEY FERGUSON 4275", 3452.5, 3417, 60),   # +35,5 h en 2 meses
    ("TRACTOR JOHN DEERE 5085", 3261, 3166, 60),          # +95 h
    ("TRACTOR MASSEY FERGUSON 6711", 2033.5, 1964, 60),   # +69,5 h
    ("EXCAVADORA", 7240.7, 7214.3, 3),                    # 26,4 h en 3 días
    ("TRACTOR JOHN DEERE 5425", 3200, 3200, 60),          # sin uso
])
def test_las_lecturas_reales_pasan(maquina, nuevo, previo, dias):
    assert validar_odometro(maquina, nuevo, previo, dias) is None


def test_la_primera_lectura_siempre_pasa():
    assert validar_odometro("MÁQUINA NUEVA", 8865.8, None, 0) is None


def test_mismo_dia_con_jornada_larga_pasa():
    """20 h en un día es mucho pero posible (dos turnos)."""
    assert validar_odometro("EXCAVADORA", 7220, 7200, dias=0) is None


# ── Camionetas: se miden en kilómetros ───────────────────────────────────

def test_una_camioneta_puede_hacer_muchos_kilometros():
    """3.000 km en un mes es normal; con el tope de horas se rechazaría."""
    assert validar_odometro("CAMIONETA SSANGYONG GRAN MUSSO (JUAN)",
                            129_000, 126_000, dias=30) is None


def test_pero_no_un_millon_de_kilometros():
    motivo = validar_odometro("CAMIONETA SSANGYONG GRAN MUSSO (JUAN)",
                              1_126_157, 126_157, dias=30)
    assert motivo and "imposible" in motivo


def test_el_tope_crece_con_los_dias():
    """Lo que es imposible en un día es normal en tres meses."""
    assert validar_odometro("TRACTOR X", 3000, 1000, dias=1) is not None
    assert validar_odometro("TRACTOR X", 3000, 1000, dias=90) is None


def test_el_motivo_explica_los_numeros():
    motivo = validar_odometro("TRACTOR X", 3200, 5222, dias=10)
    assert "5,222" in motivo and "3,200" in motivo
