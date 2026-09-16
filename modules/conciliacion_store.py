"""modules/conciliacion_store.py — Registro de conciliaciones banco ↔ documentos.

Hoja `Conciliaciones`: una fila por VÍNCULO (movimiento ↔ documento, con monto
asignado). Esto permite lo que la col J (texto) no podía:
  - conciliación PARCIAL (un pago cubre parte de una factura),
  - N:M (un cargo paga varias facturas; una factura se paga en cuotas).

La col J de Cuenta Banco pasa a ser un resumen legible que este módulo mantiene.
"""
import logging
from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from config import EXCEL_PATH
from excel_manager import _save_wb
from infrastructure.escritura_master import escribe_master

logger = logging.getLogger(__name__)

SHEET = "Conciliaciones"
HEADERS = ["ID", "Fecha Conciliación", "Fila Banco", "Fecha Mov",
           "Descripción Mov", "Monto Mov", "Tipo Doc", "Fila Doc",
           "N° Doc", "Proveedor", "Monto Asignado", "Criterio",
           "Usuario", "Nota"]
_WIDTHS = [6, 15, 10, 12, 40, 13, 14, 9, 13, 30, 14, 16, 14, 30]

# Tipos de documento/destino válidos
TIPOS_DOC = ["FACTURA", "BOLETA", "BALANCE", "TERCEROS", "NO CONCILIABLE"]

COL_J_BANCO = 10   # Factura_linkeada (resumen legible)


def _open():
    return load_workbook(EXCEL_PATH)


@escribe_master
def crear_hoja(wb=None) -> None:
    """Crea la hoja Conciliaciones si no existe. Idempotente."""
    propio = wb is None
    if propio:
        wb = _open()
    if SHEET not in wb.sheetnames:
        ws = wb.create_sheet(SHEET)
        fill = PatternFill("solid", fgColor="1F4E78")
        for i, h in enumerate(HEADERS, 1):
            c = ws.cell(1, i, h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = fill
            c.alignment = Alignment(horizontal="center")
        for i, w in enumerate(_WIDTHS, 1):
            ws.column_dimensions[chr(64 + i)].width = w
        ws.freeze_panes = "A2"
        if propio:
            _save_wb(wb)
            logger.info("Hoja Conciliaciones creada")
    if propio:
        wb.close()


def _next_id(ws) -> int:
    m = 0
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row and isinstance(row[0], (int, float)):
            m = max(m, int(row[0]))
    return m + 1


def _monto_mov(ws_banco, fila_banco: int) -> float:
    try:
        cargo = float(ws_banco.cell(fila_banco, 4).value or 0)
        abono = float(ws_banco.cell(fila_banco, 5).value or 0)
    except (TypeError, ValueError):
        return 0.0
    return cargo if cargo > 0 else abono


def _asignado_por_fila(ws_conc) -> dict:
    """{fila_banco: monto asignado total} desde la hoja Conciliaciones."""
    out = {}
    for row in ws_conc.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            fb = int(row[2])
            m = float(row[10] or 0)
        except (TypeError, ValueError):
            continue
        out[fb] = out.get(fb, 0.0) + m
    return out


def _actualizar_col_j(ws_banco, ws_conc, fila_banco: int) -> None:
    """Reescribe el resumen legible de la col J para un movimiento."""
    vinculos = []
    for row in ws_conc.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            if int(row[2]) != fila_banco:
                continue
        except (TypeError, ValueError):
            continue
        nro = str(row[8] or "").strip()
        prov = str(row[9] or "").strip()
        tipo = str(row[6] or "").strip()
        if tipo == "FACTURA" and nro:
            # Las notas de débito/crédito ya vienen con su propia sigla (ND, NC):
            # anteponerles "F" daba "FND506808".
            pre = "F" if nro[:1].isdigit() else ""
            vinculos.append(f"{pre}{nro} {prov[:24]}")
        elif tipo == "BOLETA" and nro:
            pre = "BH" if nro[:1].isdigit() else ""
            vinculos.append(f"{pre}{nro} {prov[:24]}")
        else:
            vinculos.append(tipo.capitalize() or "Conciliado")
    ws_banco.cell(fila_banco, COL_J_BANCO).value = (
        " + ".join(vinculos) if vinculos else None)


def _fecha_mov(ws_banco, fila_banco: int):
    """La fecha del movimiento del banco."""
    f = ws_banco.cell(fila_banco, 1).value
    return f.date() if isinstance(f, datetime) else f


def _ya_vinculado(ws_conc, fila_banco: int, nro, proveedor) -> bool:
    """Ese mismo movimiento ya paga ese mismo documento.

    Vincularlo de nuevo duplicaría lo asignado y dejaría la factura "pagada"
    con la mitad del dinero.
    """
    objetivo = _clave_doc(nro, proveedor)
    for row in ws_conc.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        try:
            if int(row[2]) != int(fila_banco):
                continue
        except (TypeError, ValueError):
            continue
        if _clave_doc(row[8], row[9]) == objetivo:
            return True
    return False


def _asignado_al_doc(ws_conc, nro, proveedor) -> float:
    """Lo que suman todas las cuotas registradas para ese documento."""
    objetivo = _clave_doc(nro, proveedor)
    total = 0.0
    for row in ws_conc.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None or _clave_doc(row[8], row[9]) != objetivo:
            continue
        try:
            total += float(row[10] or 0)
        except (TypeError, ValueError):
            continue
    return total


def _total_documento(ws_fact, filas_doc) -> float:
    """El total de la factura: manda la col 16 y, si no está, suman los ítems."""
    total_col = suma_items = 0.0
    for r in filas_doc:
        try:
            fila = int(r)
        except (TypeError, ValueError):
            continue
        try:
            total_col = max(total_col, float(ws_fact.cell(fila, 16).value or 0))
        except (TypeError, ValueError):
            pass
        try:
            suma_items += float(ws_fact.cell(fila, 15).value or 0)
        except (TypeError, ValueError):
            pass
    return total_col if total_col > 0 else suma_items


def _cubre(total: float, asignado: float) -> bool:
    """Con $1 de tolerancia, igual que estado_documento: es el redondeo del IVA."""
    return total > 0 and abs(total - asignado) <= 1


@escribe_master
def registrar_vinculos(vinculos: list[dict], usuario: str = "",
                       excel_path: str | None = None) -> dict:
    """Registra una lista de vínculos en UN solo guardado.

    Cada vínculo: {fila_banco, tipo_doc, fila_doc, nro_doc, proveedor,
                   monto_asignado, criterio, nota, fecha_pago(optional),
                   filas_doc(list, para completar Fecha Pago en Facturas)}

    La Fecha Pago se escribe SOLO cuando lo asignado cubre el total del
    documento: una cuota no paga la factura, y con la fecha puesta la factura
    sale de la deuda del dashboard. Un movimiento que ya paga ese documento no
    se vuelve a vincular. Devuelve {registrados, ids, repetidos, fechas}.
    """
    if not vinculos:
        return {"registrados": 0, "ids": [], "repetidos": 0, "fechas": 0}
    ruta = excel_path or EXCEL_PATH
    wb = load_workbook(ruta)
    crear_hoja(wb)
    ws_conc = wb[SHEET]
    ws_banco = wb["Cuenta Banco"]
    ws_fact = wb["Facturas"]

    nid = _next_id(ws_conc)
    ids = []
    hoy = date.today().isoformat()
    filas_tocadas = set()
    anotados = []
    repetidos = 0

    for v in vinculos:
        fb = int(v["fila_banco"])
        if _ya_vinculado(ws_conc, fb, v.get("nro_doc"), v.get("proveedor")):
            repetidos += 1
            logger.warning("El movimiento de la fila %s ya estaba vinculado a "
                           "Nº%s %s: no lo repito.", fb, v.get("nro_doc"),
                           v.get("proveedor"))
            continue
        fecha_mov = _fecha_mov(ws_banco, fb)
        desc_mov = str(ws_banco.cell(fb, 2).value or "")[:60]
        monto_mov = _monto_mov(ws_banco, fb)

        ws_conc.append([
            nid, hoy, fb,
            fecha_mov.isoformat() if isinstance(fecha_mov, date) else str(fecha_mov or ""),
            desc_mov, monto_mov,
            str(v.get("tipo_doc") or "FACTURA").upper(),
            v.get("fila_doc"),
            str(v.get("nro_doc") or ""),
            str(v.get("proveedor") or "")[:40],
            round(float(v.get("monto_asignado") or monto_mov)),
            str(v.get("criterio") or "manual"),
            usuario or str(v.get("usuario") or ""),
            str(v.get("nota") or ""),
        ])
        ids.append(nid)
        nid += 1
        filas_tocadas.add(fb)
        anotados.append(v)

    # La Fecha Pago va SOLO cuando lo asignado cubre el total del documento:
    # una cuota vincula, pero no paga. Se mira DESPUÉS de anotar todos los
    # vínculos, para que dos cuotas de la misma tanda cuenten juntas.
    fechas = 0
    for v in anotados:
        filas_doc = [int(r) for r in (v.get("filas_doc") or [])]
        if not filas_doc:
            continue
        total = _total_documento(ws_fact, filas_doc)
        asignado = _asignado_al_doc(ws_conc, v.get("nro_doc"), v.get("proveedor"))
        if not _cubre(total, asignado):
            logger.info("Nº%s %s: asignado %s de %s, sigue debiendo; no pongo "
                        "fecha de pago.", v.get("nro_doc"), v.get("proveedor"),
                        round(asignado), round(total))
            continue
        fecha_pago = v.get("fecha_pago") or _fecha_mov(ws_banco, int(v["fila_banco"]))
        escritas = 0
        for r in filas_doc:
            cell = ws_fact.cell(r, 3)
            if not (cell.value and str(cell.value).strip()):
                cell.value = fecha_pago
                escritas += 1
        if escritas:
            fechas += 1

    for fb in filas_tocadas:
        _actualizar_col_j(ws_banco, ws_conc, fb)

    _save_wb(wb, ruta)
    wb.close()
    logger.info(f"Conciliaciones registradas: {len(ids)} · "
                f"facturas que quedaron pagadas: {fechas}")
    return {"registrados": len(ids), "ids": ids, "repetidos": repetidos,
            "fechas": fechas}


def _filas_del_documento(ws_fact, nro, proveedor) -> list:
    """Las filas de Facturas de ese documento (proveedor + número)."""
    objetivo = _clave_doc(nro, proveedor)
    return [f for f in range(2, ws_fact.max_row + 1)
            if _clave_doc(ws_fact.cell(f, 7).value,
                          ws_fact.cell(f, 4).value) == objetivo]


def _mismo_dia(uno, otro) -> bool:
    """'2026-04-23', datetime(2026,4,23) y date(2026,4,23) son el mismo día."""
    return bool(uno) and bool(otro) and str(uno).strip()[:10] == str(otro).strip()[:10]


def _soltar_fecha_pago(ws_fact, ws_banco, fila_banco, nro, proveedor,
                       asignado: float) -> int:
    """Suelta la fecha de pago que había puesto la conciliación.

    Solo si el documento dejó de estar cubierto, y solo la fecha del propio
    movimiento: la que el dueño escribió a mano lleva otra fecha y no se toca.
    """
    filas = _filas_del_documento(ws_fact, nro, proveedor)
    if not filas or _cubre(_total_documento(ws_fact, filas), asignado):
        return 0
    fecha_mov = _fecha_mov(ws_banco, fila_banco) if fila_banco else None
    soltadas = 0
    for f in filas:
        if _mismo_dia(ws_fact.cell(f, 3).value, fecha_mov):
            ws_fact.cell(f, 3).value = None
            soltadas += 1
    return soltadas


@escribe_master
def desconciliar(id_vinculo: int, excel_path: str | None = None) -> bool:
    """Elimina un vínculo por ID y actualiza el resumen del movimiento.

    Si con eso el documento deja de estar cubierto, vuelve a deber: se suelta
    la fecha de pago que había puesto la conciliación. Antes quedaba puesta y
    la factura desaparecía de la deuda aunque el vínculo ya no existiera.
    """
    ruta = excel_path or EXCEL_PATH
    wb = load_workbook(ruta)
    if SHEET not in wb.sheetnames:
        wb.close()
        return False
    ws_conc = wb[SHEET]
    ws_banco = wb["Cuenta Banco"]
    ws_fact = wb["Facturas"] if "Facturas" in wb.sheetnames else None
    fila_borrar, fb, nro, proveedor = None, None, None, None
    for r in range(2, ws_conc.max_row + 1):
        if ws_conc.cell(r, 1).value == id_vinculo:
            fila_borrar = r
            try:
                fb = int(ws_conc.cell(r, 3).value)
            except (TypeError, ValueError):
                fb = None
            nro = ws_conc.cell(r, 9).value
            proveedor = ws_conc.cell(r, 10).value
            break
    if fila_borrar is None:
        wb.close()
        return False
    ws_conc.delete_rows(fila_borrar)
    if fb:
        _actualizar_col_j(ws_banco, ws_conc, fb)
    soltadas = 0
    if ws_fact is not None and nro:
        soltadas = _soltar_fecha_pago(ws_fact, ws_banco, fb, nro, proveedor,
                                      _asignado_al_doc(ws_conc, nro, proveedor))
    _save_wb(wb, ruta)
    wb.close()
    logger.info("Conciliación %s eliminada%s", id_vinculo,
                f" · {soltadas} fecha(s) de pago soltada(s)" if soltadas else "")
    return True


def _clave_doc(nro, proveedor) -> tuple:
    """Identifica un documento. Incluye el proveedor porque distintos
    proveedores repiten numeración de facturas."""
    n = str(nro or "").strip().upper()
    if n.endswith(".0"):
        n = n[:-2]
    return (" ".join(str(proveedor or "").upper().split()), n)


def asignado_por_documento(excel_path: str | None = None) -> dict:
    """{(proveedor, n° doc): monto ya asignado} sumando todos los movimientos.

    Es lo que permite pagar una factura en cuotas: cada abono suma contra el
    mismo documento y se sabe cuánto falta.
    """
    wb = load_workbook(excel_path or EXCEL_PATH, read_only=True, data_only=True)
    try:
        if SHEET not in wb.sheetnames:
            return {}
        out = {}
        for row in wb[SHEET].iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            nro = str(row[8] or "").strip()
            if not nro:
                continue
            try:
                m = float(row[10] or 0)
            except (TypeError, ValueError):
                continue
            k = _clave_doc(nro, row[9])
            out[k] = out.get(k, 0.0) + m
        return out
    finally:
        wb.close()


def estado_documento(nro, proveedor, total: float = 0.0,
                     excel_path: str | None = None) -> dict:
    """Cobertura de un documento: cuánto se pagó y cuánto falta.

    `estado`: 'pagado' (saldo ≤ $1) · 'parcial' · 'pendiente'.
    """
    asignado = asignado_por_documento(excel_path).get(
        _clave_doc(nro, proveedor), 0.0)
    total = float(total or 0)
    saldo = round(total - asignado)
    if asignado <= 0:
        estado = "pendiente"
    elif abs(saldo) <= 1:
        estado = "pagado"
    else:
        estado = "parcial"
    return {"total": total, "asignado": round(asignado),
            "saldo": saldo, "estado": estado}


def vinculos_de_documento(nro, proveedor,
                          excel_path: str | None = None) -> list:
    """Movimientos que pagaron ese documento (las cuotas), del más viejo al más nuevo."""
    wb = load_workbook(excel_path or EXCEL_PATH, read_only=True, data_only=True)
    try:
        if SHEET not in wb.sheetnames:
            return []
        objetivo = _clave_doc(nro, proveedor)
        out = []
        for row in wb[SHEET].iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            if _clave_doc(row[8], row[9]) != objetivo:
                continue
            out.append({
                "id": int(row[0]), "fila_banco": row[2], "fecha_mov": row[3],
                "desc_mov": row[4], "monto_mov": row[5],
                "monto_asignado": row[10], "criterio": row[11],
            })
        out.sort(key=lambda x: str(x["fecha_mov"] or ""))
        return out
    finally:
        wb.close()


def saldo_por_asignar(fila_banco: int, excel_path: str | None = None) -> dict:
    """Cuánto queda por asignar de un movimiento del banco."""
    wb = load_workbook(excel_path or EXCEL_PATH, read_only=True, data_only=True)
    try:
        ws_banco = wb["Cuenta Banco"]
        try:
            row = next(ws_banco.iter_rows(min_row=fila_banco, max_row=fila_banco,
                                           values_only=True))
        except StopIteration:
            return {"monto": 0, "asignado": 0, "saldo": 0, "estado": "por conciliar"}
        try:
            cargo, abono = float(row[3] or 0), float(row[4] or 0)
        except (TypeError, ValueError):
            cargo = abono = 0.0
        monto = cargo if cargo > 0 else abono
        asig = 0.0
        if SHEET in wb.sheetnames:
            for r in wb[SHEET].iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None:
                    continue
                try:
                    if int(r[2]) == fila_banco:
                        asig += float(r[10] or 0)
                except (TypeError, ValueError):
                    continue
    finally:
        wb.close()
    saldo = round(monto - asig)
    if asig <= 0:
        estado = "por conciliar"
    elif abs(saldo) <= 1:
        estado = "conciliado"
    else:
        estado = "parcial"
    return {"monto": monto, "asignado": round(asig), "saldo": saldo,
            "estado": estado, "fecha": row[0], "desc": str(row[1] or "")}


def resumen_estados() -> dict:
    """{fila_banco: {monto, asignado, saldo, estado}} para pintar la UI.

    estado: 'conciliado' (saldo ≤ $1), 'parcial' (algo asignado),
    'por conciliar' (nada asignado — las filas sin entrada no aparecen).
    """
    wb = load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        wb.close()
        return {}
    ws_conc = wb[SHEET]
    asignado = _asignado_por_fila(ws_conc)
    ws_banco = wb["Cuenta Banco"]
    out = {}
    for fb, asig in asignado.items():
        try:
            row = next(ws_banco.iter_rows(min_row=fb, max_row=fb, values_only=True))
        except StopIteration:
            continue
        try:
            cargo = float(row[3] or 0)
            abono = float(row[4] or 0)
        except (TypeError, ValueError):
            cargo = abono = 0
        monto = cargo if cargo > 0 else abono
        saldo = round(monto - asig)
        out[fb] = {"monto": monto, "asignado": round(asig), "saldo": saldo,
                   "estado": "conciliado" if abs(saldo) <= 1 else "parcial"}
    wb.close()
    return out
