# -*- coding: utf-8 -*-
"""El acumulado de labores sale de la Bitácora del Master.

OJO CON LOS FALSOS HERMANOS: "Poda nogales", "Pintar poda nogales" y "Sacar
restos poda nogales" comparten la palabra pero son TRES trabajos distintos.
Agrupar por "poda" a secas los colapsa y borra justo lo que importa: medido
contra el Master real, sacar los restos costo mas del doble de jornadas que
podar.
"""
from openpyxl import Workbook

from modules.labores import resumen_labores

ENC = ["Fecha", "Hora", "Tipo", "Actividad", "Cultivo", "Sector",
       "Jornadas Hombre", "Trabajadores", "Insumo", "Cantidad", "Unidad",
       "Registro", "Registrado por", "Máquina", "Odómetro", "Horas Día",
       "Superficie ha", "Días Cubiertos"]


def _fila(fecha, actividad, jh, trabajadores, cultivo="NOGALES", sector="",
          tipo="LABOR"):
    r = [None] * len(ENC)
    r[0], r[2], r[3] = fecha, tipo, actividad
    r[4], r[5], r[6], r[7] = cultivo, sector, jh, trabajadores
    return r


def _libro(tmp_path, filas):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(ENC)
    for f in filas:
        ws.append(f)
    ruta = tmp_path / "master.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_la_misma_labor_escrita_distinto_queda_en_un_grupo(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "Felicito Amigo, Ramiro Amigo"),
        _fila("2026-06-11", "poda  nogales", 1, "Felicito Amigo"),
        _fila("2026-06-12", "Poda de nogales", 1, "Felicito Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert len(r) == 1
    assert r[0]["jornadas"] == 4
    assert r[0]["dias"] == 3
    assert set(r[0]["etiquetas"]) == {"Poda nogales", "poda  nogales",
                                      "Poda de nogales"}


def test_poda_pintar_y_sacar_restos_son_tres_labores(tmp_path):
    """Comparten la palabra 'poda' y son tres trabajos distintos."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-06-10", "Pintar poda nogales", 1, "B"),
        _fila("2026-06-10", "Sacar restos poda nogales", 3, "C"),
    ])
    nombres = {x["labor"] for x in resumen_labores(path=ruta)}
    assert len(nombres) == 3


def test_cuenta_personas_distintas_no_apariciones(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Desaguar", 2, "Felicito Amigo, Ramiro Amigo"),
        _fila("2026-06-11", "Desaguar", 1, "Felicito Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert r[0]["personas"] == 2
    assert r[0]["jornadas"] == 3


def test_una_fila_sin_jornadas_suma_dia_pero_no_jornadas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Aseo general", None, "Felicito Amigo"),
        _fila("2026-06-11", "Aseo general", 2, "Felicito Amigo, Ramiro Amigo"),
    ])
    r = resumen_labores(path=ruta)
    assert r[0]["jornadas"] == 2
    assert r[0]["dias"] == 2


def test_la_lectura_de_horometro_no_es_una_labor(tmp_path):
    """La escribe el propio bot al guardar una lectura."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Lectura de horómetro", None, ""),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda"}


def test_sin_hoja_bitacora_devuelve_vacio(tmp_path):
    wb = Workbook()
    ruta = tmp_path / "vacio.xlsx"
    wb.save(ruta)
    assert resumen_labores(path=str(ruta)) == []


def test_un_excel_ilegible_no_revienta(tmp_path):
    assert resumen_labores(path=str(tmp_path / "no_existe.xlsx")) == []


def test_filtra_por_rango_de_fechas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-08-10", "Poda nogales", 3, "A"),
    ])
    r = resumen_labores(desde="2026-08-01", path=ruta)
    assert r[0]["jornadas"] == 3


# ── Lo que NO es trabajo ───────────────────────────────────────────────────
# Medido contra el Master: la hoja mezcla labores con ausencias y con las filas
# que escribe el propio bot. Por Tipo: LABOR 313 jornadas, RIEGO 41,
# APLICACION 24, MAQUINARIA 4 y OTRO 0. Sin filtrar, "Vacaciones" y "Ausente"
# salian en la lista de labores como si fueran trabajo hecho.


def test_las_ausencias_no_son_una_labor(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Vacaciones", None, "Ramiro Amigo", tipo="OTRO"),
        _fila("2026-06-10", "Ausente", None, "Javier Gonzalez", tipo="OTRO"),
        _fila("2026-06-10", "Aucente", None, "Javier Gonzalez"),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda"}


def test_lo_que_escribe_el_bot_por_la_maquina_no_es_una_labor(tmp_path):
    """Carga de combustible y horómetro roto: filas de maquinaria, sin jornadas."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Carga de combustible", None, "", tipo="MAQUINARIA"),
        _fila("2026-06-10", "Horómetro en mal estado", None, "", tipo="MAQUINARIA"),
        _fila("2026-06-10", "Poda nogales", 2, "A"),
    ])
    assert {x["labor"] for x in resumen_labores(path=ruta)} == {"Poda"}


def test_una_fila_de_maquinaria_CON_jornadas_si_es_trabajo(tmp_path):
    """Destroncar nogales entra como MAQUINARIA y es trabajo de verdad."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Destroncar nogales", 2, "A", tipo="MAQUINARIA"),
    ])
    assert resumen_labores(path=ruta)[0]["jornadas"] == 2


def test_una_aplicacion_sin_jornadas_anotadas_igual_cuenta_como_dia(tmp_path):
    """Se hizo el trabajo aunque nadie anotó las jornadas: no se puede borrar."""
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Aplicación fungicida", None, "", tipo="APLICACION"),
    ])
    r = resumen_labores(path=ruta)
    assert len(r) == 1
    assert r[0]["dias"] == 1
    assert r[0]["jornadas"] == 0


# ── Lo que se pagó, cruzado por RUT ────────────────────────────────────────
# Las transferencias a la gente de planta llevan el RUT en la glosa
# ("TEF 9850887-2 FELICITO AMIGO") y esos RUT están en la hoja Personal. Por
# nombre no se puede: la bitácora usa el canónico y Personal el legal.

from modules.labores import pagos_por_mes   # noqa: E402

PERSONAL = ["Nombre", "RUT", "Cargo", "Fecha Ingreso", "Días Pendientes",
            "Días Tomados Total", "Última Vacación", "Notas"]
BANCO = ["Fecha", "Descripcion", "Referencia", "Cargo", "Abono", "Saldo",
         "Tipo", "Categoria", "Cultivo", "Factura_linkeada"]


def _libro_pagos(tmp_path, personal, movs, filas_bit=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bitácora"
    ws.append(ENC)
    for f in filas_bit:
        ws.append(f)
    wp = wb.create_sheet("Personal")
    wp.append(PERSONAL)
    for p in personal:
        wp.append(list(p) + [None] * (len(PERSONAL) - len(p)))
    wbco = wb.create_sheet("Cuenta Banco")
    wbco.append(BANCO)
    for m in movs:
        fila = [None] * len(BANCO)
        fila[0], fila[1], fila[3], fila[7] = m[0], m[1], m[2], m[3]
        wbco.append(fila)
    ruta = tmp_path / "pagos.xlsx"
    wb.save(ruta)
    return str(ruta)


def test_cruza_el_pago_por_RUT_no_por_nombre(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-20", "TEF  9850887-2 FELICITO AMIGO", 885110,
          "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["planta"]["9850887-2"]["2026-06"] == 885110


def test_dos_transferencias_el_mismo_mes_se_suman(tmp_path):
    """Juan cobró dos veces en septiembre: 1.356.322 + 1.080.000."""
    ruta = _libro_pagos(
        tmp_path,
        [("Juan Parada Castillo", "13.373.052-4")],
        [("2026-09-01", "TEF 13373052-4 Juan Parada Cas", 1356322,
          "MANO DE OBRA PLANTA"),
         ("2026-09-01", "TEF 13373052-4 Juan Parada Cas", 1080000,
          "MANO DE OBRA PLANTA")])
    # Pagadas el 1-sep: las dos son el sueldo de AGOSTO.
    assert pagos_por_mes(path=ruta)["planta"]["13373052-4"]["2026-08"] == 2436322


def test_lo_del_dueno_queda_fuera(tmp_path):
    """CRAVE SPA y las remuneraciones de Felix no son costo de labor."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-01", "TEF 77912665-K CRAVE SPA", 2029455,
          "MANO DE OBRA PLANTA"),
         ("2026-06-01", "Remuneracion Mayo Felix De Vicente", 2024358,
          "MANO DE OBRA PLANTA"),
         ("2026-06-01", "TEF  9359341-3 FELIX DE VICENT", 905117,
          "MANO DE OBRA PLANTA")])
    p = pagos_por_mes(path=ruta)
    assert p["planta"] == {}
    assert p["previred"] == {}
    assert p["temporal"] == {}


def test_previred_se_guarda_aparte(tmp_path):
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-06-12", "PAGO COTIZ.PREVIRED", 2042431, "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["previred"]["2026-06"] == 2042431


def test_la_mano_de_obra_temporal_va_a_su_bolsa(tmp_path):
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-08-20", "PAGO CUADRILLA", 1500000, "MANO DE OBRA TEMPORAL")])
    assert pagos_por_mes(path=ruta)["temporal"]["2026-08"] == 1500000


def test_un_abono_no_es_un_pago(tmp_path):
    """Solo los cargos son plata que salió."""
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-08-01", "DEVOLUCION", 0, "MANO DE OBRA TEMPORAL")])
    assert pagos_por_mes(path=ruta)["temporal"] == {}


# ── El mes que se PAGA no es el mes que se TRABAJA ─────────────────────────
# Medido contra el Master: el 1-jul se pagó el sueldo de junio, el 31-jul el de
# julio y el 1-sep el de agosto. Sin corregirlo, agosto quedaba sin pagos y
# julio con el doble, y el costo por jornada salía disparatado en los dos.


def test_el_sueldo_pagado_el_primero_es_del_mes_anterior(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-07-01", "TEF  9850887-2 FELICITO AMIGO", 885110,
          "MANO DE OBRA PLANTA")])
    p = pagos_por_mes(path=ruta)["planta"]["9850887-2"]
    assert p == {"2026-06": 885110}


def test_el_sueldo_pagado_a_fin_de_mes_es_de_ese_mes(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-07-31", "TEF  9850887-2 FELICITO AMIGO", 885110,
          "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["planta"]["9850887-2"] == {"2026-07": 885110}


def test_el_1_de_enero_cae_en_diciembre_del_año_anterior(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2027-01-01", "TEF  9850887-2 FELICITO AMIGO", 885110,
          "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["planta"]["9850887-2"] == {"2026-12": 885110}


def test_previred_tambien_se_corre_al_mes_trabajado(tmp_path):
    ruta = _libro_pagos(
        tmp_path, [],
        [("2026-07-10", "PAGO COTIZ.PREVIRED", 2067109, "MANO DE OBRA PLANTA")])
    assert pagos_por_mes(path=ruta)["previred"] == {"2026-07": 2067109}


# ── Costo por trabajador ───────────────────────────────────────────────────

from modules.labores import costo_por_trabajador   # noqa: E402


def test_el_costo_por_jornada_sale_del_sueldo_dividido_las_jornadas(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-20", "TEF  9850887-2 FELICITO AMIGO", 900000,
          "MANO DE OBRA PLANTA")],
        [_fila("2026-06-%02d" % d, "Poda nogales", 1, "Felicito Amigo")
         for d in range(10, 16)])          # 6 jornadas: por encima del mínimo
    f = {x["persona"]: x for x in costo_por_trabajador(path=ruta)}["Felicito Amigo Soto"]
    assert f["jornadas"] == 6
    assert f["pagado"] == 900000
    assert f["costo_jornada"] == 150000


def test_un_mes_sin_pagos_da_sin_datos_no_cero(tmp_path):
    """Un cero se lee como 'salió gratis' y es mentira."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [],
        [_fila("2026-06-10", "Poda nogales", 1, "Felicito Amigo")])
    f = costo_por_trabajador(path=ruta)[0]
    assert f["costo_jornada"] is None
    assert f["costo"] is None


def test_una_persona_con_sueldo_y_cero_jornadas_no_divide_por_cero(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-20", "TEF  9850887-2 FELICITO AMIGO", 900000,
          "MANO DE OBRA PLANTA")],
        [])
    f = costo_por_trabajador(path=ruta)[0]
    assert f["jornadas"] == 0
    assert f["costo_jornada"] is None


def test_previred_se_reparte_a_prorrata_y_suma_lo_pagado(tmp_path):
    """Lo imputado tiene que ser IGUAL a lo que salió del banco."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2"),
         ("Ramiro Amigo Soto", "11.768.374-5")],
        [("2026-06-20", "TEF  9850887-2 FELICITO AMIGO", 600000,
          "MANO DE OBRA PLANTA"),
         ("2026-06-20", "TEF 11768374-5 RAMIRO AMIGO", 400000,
          "MANO DE OBRA PLANTA"),
         ("2026-06-20", "PAGO COTIZ.PREVIRED", 300000, "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 3, "Felicito Amigo"),
         _fila("2026-06-11", "Poda nogales", 1, "Ramiro Amigo")])
    r = {x["persona"]: x for x in costo_por_trabajador(path=ruta)}
    # 4 jornadas en el mes: Felicito 3, Ramiro 1 → previred 225.000 / 75.000
    assert round(r["Felicito Amigo Soto"]["previred"]) == 225000
    assert round(r["Ramiro Amigo Soto"]["previred"]) == 75000
    assert round(sum(x["previred"] for x in r.values())) == 300000
    assert round(r["Felicito Amigo Soto"]["costo"]) == 825000


def test_el_dueno_no_aparece_entre_los_trabajadores(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-20", "TEF 77912665-K CRAVE SPA", 2029455,
          "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 1, "Felicito Amigo")])
    nombres = {x["persona"] for x in costo_por_trabajador(path=ruta)}
    assert not any("FELIX" in n.upper() or "CRAVE" in n.upper() for n in nombres)


def test_calza_el_nombre_canonico_con_el_legal(tmp_path):
    """La bitácora dice "Felicito Amigo" y Personal "Felicito Amigo Soto"."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-06-20", "TEF  9850887-2 FELICITO AMIGO", 900000,
          "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 2, "Felicito Amigo")])
    assert costo_por_trabajador(path=ruta)[0]["jornadas"] == 2


def test_una_jornada_compartida_se_reparte_entre_los_que_estaban(tmp_path):
    """Dos personas en una fila de 2 jornadas: una cada una, no dos cada una."""
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2"),
         ("Ramiro Amigo Soto", "11.768.374-5")],
        [("2026-06-20", "TEF  9850887-2 FELICITO AMIGO", 100000,
          "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 2, "Felicito Amigo, Ramiro Amigo")])
    r = {x["persona"]: x for x in costo_por_trabajador(path=ruta)}
    assert r["Felicito Amigo Soto"]["jornadas"] == 1
    assert r["Ramiro Amigo Soto"]["jornadas"] == 1


# ── Calzar el nombre de la bitácora con el de Personal ─────────────────────
# ⚠️ Medido contra el Master: calzando por prefijo, "Ramiro Amigo" NO calza con
# "Luis Ramiro Amigo Soto" ni "Patricio Mora" con "Luis Patricio Mora Amigo".
# Tres de los seis quedaban con 0 jornadas y solo se atribuían 113 de 382.
# Probé con "Felicito Amigo", que sí calza por prefijo, y di el filtro por bueno.

from modules.labores import _persona_de   # noqa: E402

_PERSONAL = {
    "felicito amigo soto": ("9850887-2", "Felicito Amigo Soto"),
    "luis ramiro amigo soto": ("11768374-5", "Luis Ramiro Amigo Soto"),
    "luis patricio mora amigo": ("21331792-K", "Luis Patricio Mora Amigo"),
    "agustin segundo mora hernandez": ("12318508-0", "Agustin Segundo Mora Hernandez"),
    "juan parada castillo": ("13373052-4", "Juan Parada Castillo"),
    "javier gonzalez": ("20230894-5", "Javier Gonzalez"),
}


def test_el_nombre_corto_calza_aunque_el_legal_lleve_otro_nombre_delante():
    """'Ramiro Amigo' vive dentro de 'Luis Ramiro Amigo Soto'."""
    assert _persona_de("Ramiro Amigo", _PERSONAL)[0] == "11768374-5"
    assert _persona_de("Patricio Mora", _PERSONAL)[0] == "21331792-K"
    assert _persona_de("Agustin Mora", _PERSONAL)[0] == "12318508-0"


def test_los_que_ya_calzaban_siguen_calzando():
    assert _persona_de("Felicito Amigo", _PERSONAL)[0] == "9850887-2"
    assert _persona_de("Javier Gonzalez", _PERSONAL)[0] == "20230894-5"
    assert _persona_de("Juan Parada", _PERSONAL)[0] == "13373052-4"


def test_un_apellido_solo_no_identifica_a_nadie():
    """'Amigo' está en tres de los seis: no puede elegir uno."""
    assert _persona_de("Amigo", _PERSONAL) == (None, None)
    assert _persona_de("Mora", _PERSONAL) == (None, None)


def test_un_nombre_de_pila_solo_tampoco():
    """Hay CUATRO Juanes en la cuadrilla y ninguno es Juan Parada."""
    assert _persona_de("Juan", _PERSONAL) == (None, None)


def test_alguien_de_la_cuadrilla_no_calza_con_nadie():
    assert _persona_de("Pedro Soto", _PERSONAL) == (None, None)
    assert _persona_de("Josefina Quiroga", _PERSONAL) == (None, None)


def test_dos_apellidos_que_existen_pero_no_juntos_no_calzan():
    """'Mora Amigo' aparece en Patricio, pero 'Amigo Mora' no es nadie nuevo."""
    assert _persona_de("Ramiro Mora", _PERSONAL) == (None, None)


def test_con_muy_pocas_jornadas_no_se_publica_un_costo_por_jornada(tmp_path):
    """Juan es jefe de campo: reporta los partes y casi no se anota a sí mismo.

    Medido contra el Master, tiene 1 jornada y $6.388.889 pagados en tres meses.
    Dividir da $6.412.067 la jornada, que no significa nada. Su COSTO sí se
    muestra —es plata real— pero el costo por jornada queda en "sin datos".
    """
    ruta = _libro_pagos(
        tmp_path,
        [("Juan Parada Castillo", "13.373.052-4")],
        [("2026-06-20", "TEF 13373052-4 Juan Parada Cas", 6000000,
          "MANO DE OBRA PLANTA")],
        [_fila("2026-06-10", "Poda nogales", 1, "Juan Parada")])
    j = costo_por_trabajador(path=ruta)[0]
    assert j["jornadas"] == 1
    assert j["costo"] == 6000000          # su costo es real y se muestra
    assert j["costo_jornada"] is None     # el costo POR JORNADA no significa nada


# ── Costo por labor, por cultivo y por mes ─────────────────────────────────

from modules.labores import evolucion_mensual, por_cultivo_sector   # noqa: E402


def test_el_costo_de_una_labor_suma_planta_y_cuadrilla(tmp_path):
    ruta = _libro_pagos(
        tmp_path,
        [("Felicito Amigo Soto", "9.850.887-2")],
        [("2026-08-20", "TEF  9850887-2 FELICITO AMIGO", 900000,
          "MANO DE OBRA PLANTA"),
         ("2026-08-20", "PAGO CUADRILLA", 600000, "MANO DE OBRA TEMPORAL")],
        [_fila("2026-08-10", "Sacar restos poda nogales", 1, "Felicito Amigo"),
         _fila("2026-08-10", "Sacar restos poda nogales", 3, "Pedro Soto, Ana Ruiz, Luis Paz")])
    r = {x["labor"]: x for x in resumen_labores(path=ruta)}["Sacar restos de poda"]
    assert r["jornadas"] == 4
    # planta 900.000 / 1 jornada · cuadrilla 600.000 / 3 jornadas
    assert round(r["costo"]) == 1500000


def test_sin_pagos_el_costo_de_la_labor_es_sin_datos(tmp_path):
    ruta = _libro(tmp_path, [_fila("2026-06-10", "Poda nogales", 2, "A B")])
    assert resumen_labores(path=ruta)[0]["costo"] is None


def test_por_cultivo_abre_las_jornadas(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A", cultivo="NOGALES", sector="1"),
        _fila("2026-06-10", "Poda avellanos", 3, "B", cultivo="AVELLANOS", sector="2"),
    ])
    r = {(x["cultivo"], x["sector"]): x for x in por_cultivo_sector(path=ruta)}
    assert r[("NOGALES", "1")]["jornadas"] == 2
    assert r[("AVELLANOS", "2")]["jornadas"] == 3


def test_la_evolucion_va_por_mes(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-07-10", "Poda nogales", 5, "A"),
    ])
    meses = {x["mes"]: x for x in evolucion_mensual(path=ruta)}
    assert meses["2026-06"]["jornadas"] == 2
    assert meses["2026-07"]["jornadas"] == 5
    assert [x["mes"] for x in evolucion_mensual(path=ruta)] == ["2026-06", "2026-07"]


def test_la_evolucion_abre_por_labor(tmp_path):
    ruta = _libro(tmp_path, [
        _fila("2026-06-10", "Poda nogales", 2, "A"),
        _fila("2026-06-10", "Aplicación herbicida", 1, "B"),
    ])
    m = evolucion_mensual(path=ruta)[0]
    assert m["por_labor"]["Poda"] == 2
    assert m["por_labor"]["Aplicación herbicida"] == 1
