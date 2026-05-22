"""
dashboard_data.py — Extrae datos del Excel y archivos de exportación para el dashboard.
"""
import os
import logging
from datetime import date, datetime, timedelta
from collections import defaultdict
from openpyxl import load_workbook
from config import EXCEL_PATH

logger = logging.getLogger(__name__)

# Claude/ -> Agricola Santa Elisa/
AGRICOLA_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
EXPORT_2024 = os.path.join(AGRICOLA_ROOT, "Exportacion Espana 2024")
EXPORT_2025 = os.path.join(AGRICOLA_ROOT, "Exportacion Espana 2025")
CAMARICO = os.path.join(AGRICOLA_ROOT, "CAMARICO 2023")


def _open_wb(path=None):
    return load_workbook(path or EXCEL_PATH, read_only=True, data_only=True)


def _parse_date(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.strptime(str(val)[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def _safe_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        s = str(val).replace("$", "").replace(" ", "")
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0


# ─── FACTURAS ───────────────────────────────────────────────

def get_facturas_summary():
    """Resumen general de facturas."""
    wb = _open_wb()
    ws = wb["Facturas"]
    total_facturas = 0
    total_monto = 0
    pagadas = 0
    vencidas = 0
    por_pagar = 0
    hoy = date.today()
    por_mes = defaultdict(float)
    por_proveedor = defaultdict(float)
    por_tipo_doc = defaultdict(int)

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        total_facturas += 1
        monto = _safe_float(row[14])  # col O = Monto/TOTAL
        total_monto += monto
        fecha_emision = _parse_date(row[0])
        fecha_venc = _parse_date(row[1])
        fecha_pago = str(row[2] or "") if row[2] else ""
        proveedor = str(row[3] or "Sin proveedor")
        tipo_doc = str(row[5] or "Factura")

        if fecha_emision:
            key = f"{fecha_emision.year}-{fecha_emision.month:02d}"
            por_mes[key] += monto

        por_proveedor[proveedor] += abs(monto)
        por_tipo_doc[tipo_doc] += 1

        if fecha_pago and fecha_pago.strip():
            pagadas += 1
        elif fecha_venc and fecha_venc < hoy:
            vencidas += 1
        else:
            por_pagar += 1

    wb.close()

    top_proveedores = sorted(por_proveedor.items(), key=lambda x: x[1], reverse=True)[:10]

    return {
        "total_facturas": total_facturas,
        "total_monto": total_monto,
        "pagadas": pagadas,
        "vencidas": vencidas,
        "por_pagar": por_pagar,
        "por_mes": dict(sorted(por_mes.items())),
        "top_proveedores": [{"nombre": n, "monto": m} for n, m in top_proveedores],
        "por_tipo_doc": dict(por_tipo_doc),
    }


def get_facturas_detalle(filtro: str = "todas"):
    """Retorna filas detalladas de facturas filtradas por estado."""
    wb = _open_wb()
    ws = wb["Facturas"]
    hoy = date.today()
    filas = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        fecha_emision = _parse_date(row[0])
        fecha_venc = _parse_date(row[1])
        fecha_pago = str(row[2] or "").strip() if row[2] else ""
        proveedor = str(row[3] or "")
        rut = str(row[4] or "")
        tipo_doc = str(row[5] or "")
        nro = str(row[6] or "")
        glosa = str(row[7] or "")
        monto = _safe_float(row[14])

        # Determinar estado
        if fecha_pago:
            estado = "pagada"
        elif fecha_venc and fecha_venc < hoy:
            estado = "vencida"
        else:
            estado = "por_pagar"

        if filtro != "todas" and estado != filtro:
            continue

        dias_venc = (hoy - fecha_venc).days if fecha_venc else 0

        filas.append({
            "fecha_emision": str(fecha_emision or ""),
            "fecha_venc": str(fecha_venc or ""),
            "fecha_pago": fecha_pago,
            "proveedor": proveedor,
            "rut": rut,
            "documento": tipo_doc,
            "nro": nro,
            "glosa": glosa,
            "monto": monto,
            "estado": estado,
            "dias_vencida": dias_venc if estado == "vencida" else 0,
        })

    wb.close()
    filas.sort(key=lambda x: x.get("dias_vencida", 0), reverse=True)
    return filas


def get_movimientos_banco():
    """Retorna todas las filas de Cuenta Banco."""
    wb = _open_wb()
    if "Cuenta Banco" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["Cuenta Banco"]
    filas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        filas.append({
            "fecha": str(row[0] or ""),
            "descripcion": str(row[1] or ""),
            "referencia": str(row[2] or ""),
            "cargo": _safe_float(row[3]),
            "abono": _safe_float(row[4]),
            "saldo": _safe_float(row[5]),
        })
    wb.close()
    return filas


def get_banco_revisar():
    """Retorna cargos con Categoria=REVISAR para revision manual.

    Columnas Cuenta Banco: A=fecha, B=descripcion, C=referencia, D=cargo,
    E=abono, F=saldo, G=tipo, H=categoria, I=cultivo
    """
    wb = _open_wb()
    if "Cuenta Banco" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["Cuenta Banco"]
    filas = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row[0]:
            continue
        # Categoria col H = index 7
        categoria = row[7] if len(row) > 7 else None
        if categoria != "REVISAR":
            continue
        filas.append({
            "fila": idx,
            "fecha": str(row[0] or "")[:10],
            "descripcion": str(row[1] or ""),
            "referencia": str(row[2] or ""),
            "cargo": _safe_float(row[3]),
            "tipo": str(row[6] or "") if len(row) > 6 else "",
            "categoria": str(categoria or ""),
            "cultivo": str(row[8] or "") if len(row) > 8 else "",
        })
    wb.close()
    filas.sort(key=lambda x: x["cargo"], reverse=True)
    return filas


BANCO_CATEGORIAS_VALIDAS = [
    "Fertilizantes", "Fitosanitarios", "Maquinaria - mantencion",
    "Mano de obra temporal", "Mano de obra planta", "Combustible",
    "Riego", "Inversion / Replante", "Servicios profesionales",
    "Arriendos / Patentes / Seguros", "Caja chica / Imprevistos",
    "SERVICIOS DE EXPORTACION", "INSUMOS AGRICOLAS", "TRANSPORTE",
    "ENERGIA", "COSTO ENERGETICO", "HERRAMIENTAS", "SUMINISTROS",
    "MATERIALES", "SEGURIDAD", "MANTENIMIENTO", "CAMBIO DIVISA",
    "AGUA", "ALIMENTACION", "ARRIENDOS", "SERVICIOS",
]


def update_banco_categoria(fila: int, categoria: str, cultivo: str = "GENERAL"):
    """Actualiza Categoria/Cultivo de una fila de Cuenta Banco."""
    from openpyxl import load_workbook as _lw
    wb = _lw(EXCEL_PATH)
    ws = wb["Cuenta Banco"]
    ws.cell(fila, 8).value = categoria
    ws.cell(fila, 9).value = cultivo
    wb.save(EXCEL_PATH)
    wb.close()
    return {"ok": True, "fila": fila, "categoria": categoria}


# ─── COSTOS POR CULTIVO ────────────────────────────────────

def get_costos_por_cultivo():
    """Analiza costos por cultivo basado en facturas y aplicaciones."""
    wb = _open_wb()
    cultivos = {"Nogales": 0, "Cerezos": 0, "Avellanos": 0, "General": 0}

    # Desde Aplicaciones (uso de insumos)
    if "Aplicaciones" in wb.sheetnames:
        ws = wb["Aplicaciones"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            cultivo = str(row[4] or "General")
            cantidad = _safe_float(row[2])
            # Estimar costo basado en inventario
            if cultivo in cultivos:
                cultivos[cultivo] += cantidad * 100  # estimación base
            else:
                cultivos["General"] += cantidad * 100

    # Desde Facturas - clasificar por glosa
    ws_f = wb["Facturas"]
    keywords_cultivo = {
        "nogal": "Nogales", "nuez": "Nogales", "nueces": "Nogales",
        "cerezo": "Cerezos", "cereza": "Cerezos",
        "avellano": "Avellanos", "avellana": "Avellanos",
    }
    facturas_por_cultivo = {"Nogales": 0, "Cerezos": 0, "Avellanos": 0, "General": 0}

    for row in ws_f.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        glosa = str(row[7] or "").lower() + " " + str(row[8] or "").lower()
        monto = _safe_float(row[14])
        asignado = False
        for kw, cult in keywords_cultivo.items():
            if kw in glosa:
                facturas_por_cultivo[cult] += abs(monto)
                asignado = True
                break
        if not asignado:
            facturas_por_cultivo["General"] += abs(monto)

    wb.close()
    return {
        "aplicaciones": cultivos,
        "facturas": facturas_por_cultivo,
    }


# ─── COMPARACIÓN ANUAL ─────────────────────────────────────

def get_comparacion_anual():
    """Gastos por año y mes para comparación."""
    wb = _open_wb()
    ws = wb["Facturas"]
    por_anio = defaultdict(float)
    por_anio_mes = defaultdict(lambda: defaultdict(float))

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        fecha = _parse_date(row[0])
        monto = _safe_float(row[14])
        if fecha:
            por_anio[fecha.year] += abs(monto)
            por_anio_mes[fecha.year][fecha.month] += abs(monto)

    wb.close()

    # También intentar leer planilla de gastos histórica
    try:
        gastos_path = os.path.join(CAMARICO, "PLANILLA GASTOS CAMARICO 2023-2024.xlsx")
        if os.path.exists(gastos_path):
            wb2 = load_workbook(gastos_path, read_only=True, data_only=True)
            if "DATOS" in wb2.sheetnames:
                ws2 = wb2["DATOS"]
                for row in ws2.iter_rows(min_row=2, values_only=True):
                    if not row:
                        continue
                    # Intentar extraer fecha y monto de la planilla
                    fecha = _parse_date(row[0] if len(row) > 0 else None)
                    monto = _safe_float(row[-1] if row else None)
                    if fecha and monto:
                        if fecha.year not in por_anio or por_anio[fecha.year] == 0:
                            por_anio[fecha.year] += abs(monto)
                            por_anio_mes[fecha.year][fecha.month] += abs(monto)
            wb2.close()
    except Exception as e:
        logger.warning(f"No se pudo leer planilla gastos: {e}")

    result = {}
    for year in sorted(por_anio.keys()):
        meses = [0.0] * 12
        for m, v in por_anio_mes[year].items():
            meses[m - 1] = v
        result[year] = {"total": por_anio[year], "meses": meses}

    return result


# ─── EXPORTACIONES A ESPAÑA ─────────────────────────────────

def get_exportaciones():
    """Extrae datos de exportación desde las carpetas de Exportación España."""
    exports = []

    for year, folder in [(2024, EXPORT_2024), (2025, EXPORT_2025)]:
        if not os.path.isdir(folder):
            continue

        # Proformas
        for fname in os.listdir(folder):
            if "proforma" in fname.lower() and fname.endswith(".xlsx"):
                try:
                    wb = load_workbook(os.path.join(folder, fname), read_only=True, data_only=True)
                    ws = wb[wb.sheetnames[0]]
                    export_data = {"year": year, "tipo": "Proforma", "archivo": fname}
                    found_header = False
                    for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
                        vals = [str(c or "").lower() for c in row]
                        joined = " ".join(vals)
                        # Buscar fila header "Producto, Item, ..., Valor US$"
                        if "valor" in joined and "us" in joined:
                            found_header = True
                            continue
                        # Filas de datos después del header
                        if found_header:
                            # Última columna = Valor US$
                            last_val = _safe_float(row[-1]) if row else 0
                            if last_val > 1000:
                                export_data["valor_usd"] = export_data.get("valor_usd", 0) + last_val
                            # Columna cantidad (index 5 normalmente)
                            for c in row:
                                v = _safe_float(c)
                                if 1000 < v < 500000 and "cantidad_kg" not in export_data:
                                    export_data["cantidad_kg"] = v
                            # "Total" marca fin
                            if "total" in joined:
                                if last_val > 1000:
                                    export_data["valor_usd"] = last_val  # usar Total como definitivo
                                found_header = False
                    wb.close()
                    exports.append(export_data)
                except Exception as e:
                    logger.warning(f"Error leyendo {fname}: {e}")

        # Packing Lists
        for fname in os.listdir(folder):
            if "packing" in fname.lower() and fname.endswith(".xlsx"):
                try:
                    wb = load_workbook(os.path.join(folder, fname), read_only=True, data_only=True)
                    ws = wb[wb.sheetnames[0]]
                    pack_data = {"year": year, "tipo": "Packing", "archivo": fname}
                    for row in ws.iter_rows(min_row=1, max_row=30, values_only=True):
                        vals = [str(c or "") for c in row]
                        for c in row:
                            v = _safe_float(c)
                            if 10000 < v < 200000:
                                if "peso" not in " ".join(vals).lower():
                                    pack_data.setdefault("peso_total_kg", v)
                    wb.close()
                    exports.append(pack_data)
                except Exception as e:
                    logger.warning(f"Error leyendo {fname}: {e}")

        # Tarjas (contenedores)
        tarjas = []
        for fname in os.listdir(folder):
            if "tarja" in fname.lower() and fname.endswith(".xlsx"):
                try:
                    wb = load_workbook(os.path.join(folder, fname), read_only=True, data_only=True)
                    n_contenedores = len(wb.sheetnames)
                    peso_total = 0
                    for sname in wb.sheetnames:
                        ws = wb[sname]
                        for row in ws.iter_rows(min_row=7, max_row=7, values_only=True):
                            for c in row:
                                v = _safe_float(c)
                                if 200 < v < 2000:
                                    peso_total += v
                    wb.close()
                    tarjas.append({
                        "year": year, "tipo": "Tarja", "archivo": fname,
                        "contenedores": n_contenedores, "peso_neto_kg": peso_total
                    })
                except Exception as e:
                    logger.warning(f"Error leyendo tarja {fname}: {e}")
        exports.extend(tarjas)

    # Resumen por año
    resumen = {}
    for year in [2024, 2025]:
        items = [e for e in exports if e.get("year") == year]
        total_kg = sum(e.get("cantidad_kg", 0) or e.get("peso_neto_kg", 0) or e.get("peso_total_kg", 0) for e in items)
        total_usd = sum(e.get("valor_usd", 0) for e in items)
        contenedores = sum(e.get("contenedores", 0) for e in items if e.get("tipo") == "Tarja")
        resumen[year] = {
            "total_kg": total_kg,
            "total_usd": total_usd,
            "contenedores": contenedores,
            "archivos": len(items),
        }

    return {"detalle": exports, "resumen": resumen}


# ─── CAJA CHICA ─────────────────────────────────────────────

def get_caja_chica():
    """Resumen de caja chica."""
    wb = _open_wb()
    if "Caja Chica" not in wb.sheetnames:
        wb.close()
        return {"saldo": 0, "ingresos": 0, "egresos": 0, "movimientos": []}

    ws = wb["Caja Chica"]
    ingresos = 0
    egresos = 0
    movimientos = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        tipo = str(row[1] or "")
        monto = _safe_float(row[2])
        if tipo.lower() in ("ingreso", "depósito", "deposito"):
            ingresos += monto
        else:
            egresos += monto
        movimientos.append({
            "fecha": str(row[0] or ""),
            "tipo": tipo,
            "monto": monto,
            "detalle": str(row[3] or ""),
        })
    wb.close()
    return {
        "saldo": ingresos - egresos,
        "ingresos": ingresos,
        "egresos": egresos,
        "ultimos": movimientos[-10:],
    }


# ─── CUENTA BANCO ───────────────────────────────────────────

def get_cuenta_banco():
    """Resumen de cuenta banco."""
    wb = _open_wb()
    if "Cuenta Banco" not in wb.sheetnames:
        wb.close()
        return {"n_movimientos": 0, "ultimo_saldo": 0, "cargos": 0, "abonos": 0}

    ws = wb["Cuenta Banco"]
    cargos = 0
    abonos = 0
    ultimo_saldo = 0
    n = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        n += 1
        cargos += _safe_float(row[3])
        abonos += _safe_float(row[4])
        s = _safe_float(row[5])
        if s:
            ultimo_saldo = s
    wb.close()
    return {
        "n_movimientos": n,
        "ultimo_saldo": ultimo_saldo,
        "cargos": cargos,
        "abonos": abonos,
    }


# ─── INVENTARIO ─────────────────────────────────────────────

def get_inventario_resumen():
    """Resumen de inventario."""
    wb = _open_wb()
    if "Inventario" not in wb.sheetnames:
        wb.close()
        return {"productos": [], "alertas": 0}

    ws = wb["Inventario"]
    productos = []
    alertas = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        stock = _safe_float(row[3])
        minimo = _safe_float(row[4])
        alerta = stock <= minimo and minimo > 0
        if alerta:
            alertas += 1
        productos.append({
            "producto": str(row[0]),
            "categoria": str(row[1] or ""),
            "unidad": str(row[2] or ""),
            "stock": stock,
            "minimo": minimo,
            "alerta": alerta,
        })
    wb.close()
    return {"productos": productos, "alertas": alertas}


# ─── PERSONAL Y VACACIONES ──────────────────────────────────

def get_personal_resumen():
    """Resumen de personal y vacaciones."""
    wb = _open_wb()
    personal = []
    if "Personal" in wb.sheetnames:
        ws = wb["Personal"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            personal.append({
                "nombre": str(row[0]),
                "cargo": str(row[2] or ""),
                "dias_pendientes": _safe_float(row[4]),
                "dias_tomados": _safe_float(row[5]),
                "ultima_vacacion": str(row[6] or ""),
            })

    vacaciones = []
    if "Vacaciones" in wb.sheetnames:
        ws = wb["Vacaciones"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            vacaciones.append({
                "trabajador": str(row[0]),
                "inicio": str(row[1] or ""),
                "fin": str(row[2] or ""),
                "dias": int(_safe_float(row[3])),
                "estado": str(row[4] or ""),
            })
    wb.close()
    return {"personal": personal, "vacaciones": vacaciones[-10:]}


# ─── TAREAS ─────────────────────────────────────────────────

def get_tareas_resumen():
    """Resumen de tareas."""
    wb = _open_wb()
    if "Tareas" not in wb.sheetnames:
        wb.close()
        return {"pendientes": 0, "en_progreso": 0, "completadas": 0, "tareas": []}

    ws = wb["Tareas"]
    pendientes = 0
    en_progreso = 0
    completadas = 0
    tareas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        estado = str(row[4] or "Pendiente")
        if estado == "Pendiente":
            pendientes += 1
        elif estado == "En Progreso":
            en_progreso += 1
        elif estado == "Hecho":
            completadas += 1
        tareas.append({
            "id": int(_safe_float(row[0])),
            "descripcion": str(row[2] or ""),
            "prioridad": str(row[3] or "Media"),
            "estado": estado,
            "responsable": str(row[5] or ""),
        })
    wb.close()
    return {
        "pendientes": pendientes,
        "en_progreso": en_progreso,
        "completadas": completadas,
        "tareas": tareas,
    }


# ─── DATOS COMPLETOS PARA DASHBOARD ────────────────────────

def get_dashboard_data():
    """Retorna todos los datos necesarios para el dashboard."""
    try:
        facturas = get_facturas_summary()
    except Exception as e:
        logger.error(f"Error facturas: {e}")
        facturas = {}

    try:
        costos = get_costos_por_cultivo()
    except Exception as e:
        logger.error(f"Error costos cultivo: {e}")
        costos = {}

    try:
        anual = get_comparacion_anual()
    except Exception as e:
        logger.error(f"Error comparacion anual: {e}")
        anual = {}

    try:
        exports = get_exportaciones()
    except Exception as e:
        logger.error(f"Error exportaciones: {e}")
        exports = {}

    try:
        caja = get_caja_chica()
    except Exception as e:
        logger.error(f"Error caja chica: {e}")
        caja = {}

    try:
        banco = get_cuenta_banco()
    except Exception as e:
        logger.error(f"Error banco: {e}")
        banco = {}

    try:
        inventario = get_inventario_resumen()
    except Exception as e:
        logger.error(f"Error inventario: {e}")
        inventario = {}

    try:
        personal = get_personal_resumen()
    except Exception as e:
        logger.error(f"Error personal: {e}")
        personal = {}

    try:
        tareas = get_tareas_resumen()
    except Exception as e:
        logger.error(f"Error tareas: {e}")
        tareas = {}

    return {
        "fecha": date.today().strftime("%Y-%m-%d"),
        "facturas": facturas,
        "costos_cultivo": costos,
        "comparacion_anual": anual,
        "exportaciones": exports,
        "caja_chica": caja,
        "cuenta_banco": banco,
        "inventario": inventario,
        "personal": personal,
        "tareas": tareas,
    }
