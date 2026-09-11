"""
inventario_manager.py — Gestión de Inventario y Uso por Cultivo
Hojas: Inventario, Aplicaciones
"""
import logging
import re
import unicodedata
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from config import EXCEL_PATH
from excel_manager import _save_wb  # guardado con reintentos si Excel está abierto

logger = logging.getLogger(__name__)

INVENTARIO_SHEET = "Inventario"
APLICACIONES_SHEET = "Aplicaciones"

INVENTARIO_HEADERS = [
    "Producto", "Categoría", "Unidad", "Stock Actual",
    "Stock Mínimo", "Última Entrada", "Último Uso"
]
APLICACIONES_HEADERS = [
    "Fecha", "Producto", "Cantidad", "Unidad", "Cultivo",
    "Sector", "Responsable", "Observaciones"
]

CATEGORIAS = ["Fertilizante", "Fungicida", "Herbicida", "Insecticida", "Semilla", "Otro"]
CULTIVOS = ["Nogales", "Cerezos", "Avellanos"]


def _open_wb(path=None):
    return load_workbook(path or EXCEL_PATH)


def _create_sheet(wb, name, headers, widths, color):
    if name not in wb.sheetnames:
        ws = wb.create_sheet(name)
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        font = Font(bold=True, color="FFFFFF", size=11)
        thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                      top=Side(style='thin'), bottom=Side(style='thin'))
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.fill = fill
            c.font = font
            c.alignment = Alignment(horizontal="center")
            c.border = thin
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i) if i <= 26 else 'A'].width = w
        logger.info(f"Hoja '{name}' creada")
    return wb[name]


def _ensure_inventario(wb):
    return _create_sheet(wb, INVENTARIO_SHEET, INVENTARIO_HEADERS,
                         [35, 16, 8, 14, 14, 12, 12], "1B5E20")


def _ensure_aplicaciones(wb):
    return _create_sheet(wb, APLICACIONES_SHEET, APLICACIONES_HEADERS,
                         [12, 30, 10, 8, 12, 15, 18, 35], "33691E")


def crear_hojas_inventario():
    """Crea las hojas Inventario y Aplicaciones."""
    wb = _open_wb()
    _ensure_inventario(wb)
    _ensure_aplicaciones(wb)
    _save_wb(wb)


MIN_LETRAS = 3


def _clave_prod(t) -> str:
    """Minusculas, sin tildes y con los espacios colapsados."""
    t = "".join(c for c in unicodedata.normalize("NFD", str(t or "").lower())
                if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", t).split())


def buscar_producto(nombres, consulta):
    """(indice del producto, candidatos si es ambiguo). Puro: recibe la lista.

    🔴 EL 10-SEP-2026 ESTO EXIGIA COINCIDENCIA EXACTA. Juan escribio "Katana",
    el inventario dice "KATANA 1 KG - herbicida", no calzo, y registrar_uso creo
    una fila NUEVA con stock -5 sin avisarle a nadie.

    Ahora el nombre corto calza con el comercial largo. Pero NO SE ADIVINA: hay
    TRES filas de Ripper Full, y elegir una al azar es peor que preguntar. Si
    hay mas de una candidata se devuelven todas y que decida quien pregunte.
    """
    k = _clave_prod(consulta)
    if len(k.replace(" ", "")) < MIN_LETRAS:
        return None, []                      # "k" no identifica nada

    claves = [_clave_prod(n) for n in nombres]
    for i, c in enumerate(claves):
        if c == k:
            return i, []                     # el exacto gana siempre

    cand = [i for i, c in enumerate(claves) if c.startswith(k) or k in c]
    if len(cand) == 1:
        return cand[0], []
    return None, [nombres[i] for i in cand]


def _find_producto(ws, producto: str):
    """Busca un producto en inventario. Retorna row_idx, o None si no hay
    ninguno o si hay VARIOS candidatos --ahi hay que preguntar, no elegir."""
    nombres, filas = [], []
    for row_idx in range(2, ws.max_row + 1):
        val = ws.cell(row=row_idx, column=1).value
        if val:
            nombres.append(str(val))
            filas.append(row_idx)
    idx, _ = buscar_producto(nombres, producto)
    return filas[idx] if idx is not None else None


def candidatos_producto(ws, producto: str) -> list:
    """Los nombres que calzan cuando hay mas de uno. Vacia si no hay duda."""
    nombres = [str(ws.cell(row=r, column=1).value)
               for r in range(2, ws.max_row + 1)
               if ws.cell(row=r, column=1).value]
    return buscar_producto(nombres, producto)[1]


def agregar_stock(producto: str, cantidad: float, categoria: str = "Otro",
                  unidad: str = "L") -> dict:
    """Agrega stock a un producto existente o lo crea."""
    wb = _open_wb()
    ws = _ensure_inventario(wb)
    row_idx = _find_producto(ws, producto)
    hoy = date.today().strftime("%Y-%m-%d")

    if row_idx:
        stock_actual = float(ws.cell(row=row_idx, column=4).value or 0)
        nuevo_stock = stock_actual + cantidad
        ws.cell(row=row_idx, column=4).value = nuevo_stock
        ws.cell(row=row_idx, column=6).value = hoy
        _save_wb(wb)
        return {"producto": producto, "stock_anterior": stock_actual,
                "agregado": cantidad, "stock_nuevo": nuevo_stock, "nuevo": False}
    else:
        ws.append([producto, categoria, unidad, cantidad, 0, hoy, ""])
        _save_wb(wb)
        return {"producto": producto, "stock_anterior": 0,
                "agregado": cantidad, "stock_nuevo": cantidad, "nuevo": True}


def registrar_uso(producto: str, cantidad: float, cultivo: str,
                  sector: str = "", responsable: str = "",
                  observaciones: str = "", fecha: str = "",
                  unidad: str = "", path: str = None) -> dict:
    """Registra uso de un producto y descuenta del inventario.

    `fecha` es la del TRABAJO, no la de hoy: Juan reporta las aplicaciones days
    despues y con `date.today()` quedaban todas con la fecha en que las mando.
    `unidad` la que diga el parte ("5,4 kilos"): la del inventario puede no
    coincidir, y rotular kilos de fungicida como litros no es un detalle.
    Las dos caen a lo de antes si no vienen.
    """
    if not _clave_prod(producto):
        # Asi nacio la fila con el producto en blanco el 10-sep-2026.
        return {"error": "sin_producto", "producto": producto}

    wb = _open_wb(path)
    ws_inv = _ensure_inventario(wb)
    ws_app = _ensure_aplicaciones(wb)

    dudosos = candidatos_producto(ws_inv, producto)
    if dudosos:
        # Hay TRES filas de Ripper Full: elegir una al azar es peor que
        # preguntar. No se escribe nada y que decida quien pregunte.
        wb.close()
        return {"error": "ambiguo", "producto": producto, "candidatos": dudosos}

    row_idx = _find_producto(ws_inv, producto)
    cuando = (fecha or "")[:10] or date.today().strftime("%Y-%m-%d")
    desconocido = row_idx is None

    if row_idx:
        stock_actual = float(ws_inv.cell(row=row_idx, column=4).value or 0)
        unidad = unidad or ws_inv.cell(row=row_idx, column=3).value or "L"
        nuevo_stock = stock_actual - cantidad
        ws_inv.cell(row=row_idx, column=4).value = nuevo_stock
        ws_inv.cell(row=row_idx, column=7).value = cuando
    else:
        # 🔴 ANTES SE CREABA UNA FILA CON STOCK NEGATIVO Y SIN AVISAR: asi
        # nacieron "Katana -5" y "Agrocupper -2,1". El uso SI se anota --el
        # trabajo se hizo-- pero el inventario no se inventa un producto.
        unidad = unidad or "L"
        nuevo_stock = None
        logger.warning("Uso de un producto que no esta en el inventario: %r",
                       producto)

    ws_app.append([
        cuando, producto, cantidad, unidad, cultivo,
        sector, responsable, observaciones
    ])

    _save_wb(wb, path)
    minimo = float(ws_inv.cell(row=row_idx, column=5).value or 0) if row_idx else 0
    return {
        "producto": producto, "cantidad": cantidad, "unidad": unidad,
        "cultivo": cultivo, "sector": sector,
        "stock_restante": nuevo_stock,
        "alerta_bajo": (nuevo_stock is not None and nuevo_stock <= minimo),
        "stock_negativo": (nuevo_stock is not None and nuevo_stock < 0),
        "producto_desconocido": desconocido,
    }


def consultar_inventario() -> list[dict]:
    """Lista todo el inventario actual."""
    wb = _open_wb()
    ws = _ensure_inventario(wb)
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        stock = float(row[3] or 0)
        minimo = float(row[4] or 0)
        items.append({
            "producto": str(row[0]),
            "categoria": str(row[1] or ""),
            "unidad": str(row[2] or "L"),
            "stock": stock,
            "minimo": minimo,
            "alerta": stock <= minimo and minimo > 0,
        })
    wb.close()
    return items


def productos_bajo_stock() -> list[dict]:
    """Retorna productos con stock bajo el mínimo."""
    return [p for p in consultar_inventario() if p["alerta"]]


def consumo_por_cultivo(dias: int = 7) -> dict:
    """Resumen de consumo por cultivo en los últimos N días."""
    wb = _open_wb()
    ws = _ensure_aplicaciones(wb)
    hoy = date.today()
    resumen = {c: [] for c in CULTIVOS}
    resumen["Otro"] = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        try:
            if isinstance(row[0], str):
                fecha = datetime.strptime(row[0][:10], "%Y-%m-%d").date()
            else:
                fecha = row[0] if isinstance(row[0], date) else row[0].date()
            if (hoy - fecha).days <= dias:
                cultivo = str(row[4] or "Otro")
                key = cultivo if cultivo in CULTIVOS else "Otro"
                resumen[key].append({
                    "producto": str(row[1] or ""),
                    "cantidad": float(row[2] or 0),
                    "unidad": str(row[3] or ""),
                    "fecha": str(fecha),
                })
        except Exception:
            continue
    wb.close()
    return resumen


def agregar_stock_desde_factura(items: list):
    """Agrega stock automáticamente desde items de factura procesada.
    Busca productos que parezcan insumos agrícolas."""
    keywords = {
        "fertilizante": "Fertilizante", "urea": "Fertilizante", "npk": "Fertilizante",
        "nitrato": "Fertilizante", "fosfato": "Fertilizante", "potasio": "Fertilizante",
        "abono": "Fertilizante", "compost": "Fertilizante",
        "fungicida": "Fungicida", "captan": "Fungicida", "cobre": "Fungicida",
        "azufre": "Fungicida", "mancozeb": "Fungicida",
        "herbicida": "Herbicida", "glifosato": "Herbicida", "roundup": "Herbicida",
        "insecticida": "Insecticida", "aceite": "Insecticida",
        "intrepid": "Insecticida", "clorpirifos": "Insecticida",
    }
    agregados = []
    for item in items:
        glosa = str(item.get("Detalle / Glosa") or "").lower()
        glosa2 = str(item.get("Glosa II") or "").lower()
        texto = f"{glosa} {glosa2}"

        categoria = None
        for kw, cat in keywords.items():
            if kw in texto:
                categoria = cat
                break

        if categoria:
            nombre = item.get("Detalle / Glosa") or item.get("Glosa II") or "Insumo"
            qty = float(item.get("Cantidad") or 1)
            unidad = "L" if any(u in texto for u in ("litro", "lt", "lts")) else "Kg" if any(
                u in texto for u in ("kilo", "kg")) else "Un"
            result = agregar_stock(nombre, qty, categoria, unidad)
            agregados.append(result)

    return agregados


def agrupar_duplicados(productos) -> list:
    """Grupos de filas que son EL MISMO producto escrito distinto.

    `productos`: dicts con al menos "nombre". Devuelve solo los grupos de 2 o mas.

    ⚠️ LA REGLA ES PREFIJO DE PALABRAS, no primera palabra. Medido contra el
    Master: agrupar por la primera palabra juntaria los CUATRO Defender --Zn,
    Calcio, K (Potasio) y Boro-- que son productos distintos. En cambio
    "Ripper Full" SI es prefijo de "RIPPER FULL SL 20 LT - herbicida", y son el
    mismo.
    """
    def _prefijo(a, b):
        pa, pb = a.split(), b.split()
        return len(pa) <= len(pb) and pb[:len(pa)] == pa

    claves = [_clave_prod(p.get("nombre")) for p in productos]
    grupos, usados = [], set()
    for i, k in enumerate(claves):
        if i in usados or not k:
            continue
        grupo, usados = [productos[i]], usados | {i}
        for j, o in enumerate(claves):
            if j in usados or not o:
                continue
            if _prefijo(k, o) or _prefijo(o, k):
                grupo.append(productos[j])
                usados.add(j)
        if len(grupo) > 1:
            grupos.append(grupo)
    return grupos
