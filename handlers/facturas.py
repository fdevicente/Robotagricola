"""handlers/facturas.py — Procesamiento de facturas: helpers internos.

Paso 1: solo helpers/constantes. Pasos siguientes agregan comandos y callbacks.
"""
import asyncio
import json
import logging
import os
import re
from datetime import datetime

from config import DOWNLOAD_DIR, BOLETAS_DIR, EXCEL_PATH

logger = logging.getLogger(__name__)


CAMPOS_EDITABLES = {
    "edit_proveedor": ("Nombre Factura / Proveedor",      "🏢 Nombre del proveedor"),
    "edit_rut":       ("Rut",                             "🪪 RUT del proveedor"),
    "edit_fecha":     ("Fecha Emision",                   "📅 Fecha de emisión (YYYY-MM-DD)"),
    "edit_vence":     ("Fecha Vencimiento",               "⏰ Fecha de vencimiento (YYYY-MM-DD)"),
    "edit_nro":       ("Numero Factura / Nro Documento",  "📄 Número de documento"),
    "edit_ref":       ("Referencia Factura",              "🔗 Nº factura referenciada (para NC/ND)"),
    "edit_glosa":     ("Detalle / Glosa",                 "📦 Glosa / descripción corta"),
    "edit_glosa2":    ("Glosa II",                        "📝 Detalle completo"),
    "edit_cantidad":  ("Cantidad",                        "🔢 Cantidad"),
    "edit_unitario":  ("Valor unitario",                  "💲 Valor unitario neto"),
    "edit_total":     ("TOTAL NETO",                      "💰 Total ítem sin IVA"),
}
# "edit_total_factura" se maneja aparte
CAMPOS_COMUNES = {"edit_proveedor", "edit_rut", "edit_fecha", "edit_vence", "edit_nro", "edit_ref"}
CAMPOS_POR_ITEM = {"edit_glosa", "edit_glosa2", "edit_cantidad", "edit_unitario", "edit_total"}


def _save_path(filename):
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    return os.path.join(DOWNLOAD_DIR, filename)


def _save_path_boleta(filename):
    os.makedirs(BOLETAS_DIR, exist_ok=True)
    return os.path.join(BOLETAS_DIR, filename)


def _es_boleta(items):
    """Detecta si los items corresponden a una boleta (caja chica).
    Boletas de Honorarios NO son caja chica — van a Facturas."""
    if not items:
        return False
    doc = str(items[0].get("Documento") or "").lower()
    return "boleta" in doc and "boleta de honorario" not in doc


def _renombrar_archivo(file_path: str, items: list) -> str:
    """Renombra el archivo usando Proveedor + Nº Documento."""
    try:
        item = items[0]
        proveedor = str(item.get("Nombre Factura / Proveedor") or "").strip()
        nro = str(item.get("Numero Factura / Nro Documento") or "").strip()
        if not proveedor and not nro:
            return file_path

        def _limpiar(s):
            s = re.sub(r'[\\/:*?"<>|]', '', s)
            s = re.sub(r'\s+', '_', s.strip())
            return s[:60]

        partes = []
        if proveedor:
            partes.append(_limpiar(proveedor))
        if nro:
            partes.append(nro)

        ext = os.path.splitext(file_path)[1]
        nombre = "_".join(partes) + ext
        dir_path = os.path.dirname(file_path)
        nuevo = os.path.join(dir_path, nombre)

        if os.path.exists(nuevo) and os.path.abspath(nuevo) != os.path.abspath(file_path):
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            nombre = "_".join(partes) + f"_{ts}" + ext
            nuevo = os.path.join(dir_path, nombre)

        os.rename(file_path, nuevo)
        logger.info(f"Archivo renombrado: {os.path.basename(file_path)} → {nombre}")

        # Limpiar archivos derivados (_resized.jpg, _scan.png)
        base_viejo = os.path.splitext(file_path)[0]
        for sufijo in ("_resized.jpg", "_scan.png"):
            derivado = base_viejo + sufijo
            if os.path.exists(derivado):
                try:
                    os.remove(derivado)
                    logger.info(f"Derivado eliminado: {os.path.basename(derivado)}")
                except Exception as e:
                    logger.warning(f"No se pudo eliminar {derivado}: {e}")
        return nuevo
    except Exception as e:
        logger.error(f"No se pudo renombrar archivo {file_path}: {e}")
        return file_path


def _registrar_correccion(item: dict, campo: str, valor_original, valor_nuevo):
    """Guarda la corrección del usuario para aprendizaje futuro."""
    try:
        log_path = os.path.join(DOWNLOAD_DIR, "correcciones_log.json")
        log = []
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                log = json.load(f)
        entrada = {
            "timestamp": datetime.now().isoformat(),
            "rut":        item.get("Rut"),
            "proveedor":  item.get("Nombre Factura / Proveedor"),
            "nro_factura": item.get("Numero Factura / Nro Documento"),
            "campo":      campo,
            "valor_claude": valor_original,
            "valor_usuario": valor_nuevo,
        }
        if campo in ("Monto / TOTAL", "Total Factura", "Valor unitario"):
            try:
                orig = float(valor_original or 0)
                nuevo = float(valor_nuevo or 0)
                if orig > 0 and nuevo > 0:
                    entrada["factor"] = round(nuevo / orig, 6)
            except Exception:
                pass
        log.append(entrada)
        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(log, f, ensure_ascii=False, indent=2)
        logger.info(f"Corrección registrada: {campo} | '{valor_original}' → '{valor_nuevo}' "
                    f"({item.get('Rut')})")
    except Exception as e:
        logger.warning(f"No se pudo registrar corrección: {e}")


def _rut_existe(rut):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        ws = wb["Proveedores"]
        rut_n = rut.replace(".", "").replace("-", "").replace(" ", "").upper()
        for row in ws.iter_rows(min_row=3, values_only=True):
            r = str(row[2]).replace(".", "").replace("-", "").replace(" ", "").upper() if row[2] else ""
            if rut_n and rut_n == r:
                wb.close()
                return True
        wb.close()
    except Exception as e:
        logger.warning(f"Error verificando RUT: {e}")
    return False


def _agregar_proveedor(nombre, rut):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb["Proveedores"]
        ws.append([None, nombre, rut])
        wb.save(EXCEL_PATH)
        logger.info(f"Proveedor agregado: {nombre} — {rut}")
        return True
    except Exception as e:
        logger.error(f"Error agregando proveedor: {e}")
        return False


async def _download_with_retry(f, path, retries=3):
    """Descarga archivo de Telegram con reintentos."""
    for attempt in range(1, retries + 1):
        try:
            await f.download_to_drive(path)
            return True
        except Exception as e:
            logger.warning(f"Descarga intento {attempt}/{retries} falló: {e}")
            if attempt < retries:
                await asyncio.sleep(2 * attempt)
    return False


# ── Preview builder ──────────────────────────────

from utils.formatting import esc as _esc, format_date as _format_date, calc_vencimiento as _calc_vencimiento


def _build_preview(items):
    lines = ["📋 *Datos extraídos — revisa antes de guardar:*\n"]
    first = items[0]
    fecha_venc = first.get("Fecha Vencimiento") or _calc_vencimiento(first.get("Fecha Emision"))

    lines.append(f"🏢 *Proveedor:* {_esc(first.get('Nombre Factura / Proveedor') or '? no detectado')}")
    lines.append(f"🪪 *RUT:* {_esc(first.get('Rut') or '? no detectado')}")
    doc_tipo = _esc(first.get('Documento') or '—')
    doc_nro = _esc(first.get('Numero Factura / Nro Documento') or '—')
    lines.append(f"📄 *Documento:* {doc_tipo}  Nº {doc_nro}")
    ref = first.get("Referencia Factura")
    if ref:
        lines.append(f"🔗 *Ref. Factura:* Nº {_esc(ref)}")
    lines.append(f"📅 *Emisión:* {_format_date(first.get('Fecha Emision'))}   "
                  f"⏰ *Vence:* {_format_date(fecha_venc)}\n")

    doc = str(first.get('Documento') or '').lower()
    es_honorario = "boleta de honorario" in doc
    exenta = any(k in doc for k in ('exenta', 'exento', 'no afecta', 'no afecto')) or es_honorario

    pesos = []
    total_imp_esp = 0.0
    for item in items:
        unitario = float(item.get('Valor unitario') or 0)
        cantidad = float(item.get('Cantidad') or 1)
        pesos.append(unitario * cantidad)
        total_imp_esp += float(item.get('Impuesto Especifico') or 0)
    total_neto_raw = sum(pesos)

    total_factura = round(float(first.get('Total Factura') or 0))
    if total_factura > 0:
        base_iva = total_factura - round(total_imp_esp)
        if exenta:
            iva_total = 0
            neto_anchor = base_iva
        else:
            iva_total = round(base_iva / 1.19 * 0.19)
            neto_anchor = base_iva - iva_total
        total_con_iva = total_factura
    else:
        neto_anchor = round(total_neto_raw)
        iva_total = 0 if exenta else round(total_neto_raw * 0.19)
        total_con_iva = round(neto_anchor * (1.0 if exenta else 1.19) + total_imp_esp)

    n = len(items)
    netos_linea = [0] * n
    if n > 0:
        if total_neto_raw > 0:
            acumulado = 0
            for i, peso in enumerate(pesos):
                if i == n - 1:
                    netos_linea[i] = neto_anchor - acumulado
                else:
                    val = round(neto_anchor * peso / total_neto_raw)
                    netos_linea[i] = val
                    acumulado += val
        else:
            base = neto_anchor // n
            for i in range(n):
                netos_linea[i] = base
            netos_linea[-1] += neto_anchor - base * n

    for i, item in enumerate(items, 1):
        unitario = float(item.get('Valor unitario') or 0)
        cantidad = float(item.get('Cantidad') or 1)
        neto_linea = netos_linea[i - 1]
        if n > 1:
            lines.append(f"*— Ítem {i} —*")
        lines.append(f"📦 *Glosa:* {_esc(item.get('Detalle / Glosa') or '? no detectado')}")
        if item.get('Glosa II'):
            lines.append(f"📝 *Detalle:* {_esc(item.get('Glosa II'))}")
        lines.append(f"🔢 *Cantidad:* {cantidad:g}   💲 *Unit neto:* ${unitario:,.3f}")
        lines.append(f"💵 *Neto línea:* ${neto_linea:,.0f}\n")

    if es_honorario:
        pago_profesional = neto_anchor
        retencion = round(total_imp_esp)
        costo_total = total_con_iva if total_factura > 0 else pago_profesional + retencion
        lines.append(f"💵 *Pago al profesional:* ${pago_profesional:,.0f}")
        if retencion:
            lines.append(f"🏦 *Impto. Retenido (ASE→SII):* ${retencion:,.0f}")
        lines.append(f"💰 *COSTO TOTAL:* ${costo_total:,.0f}\n")
    else:
        lines.append(f"📊 *NETO:* ${neto_anchor:,.0f}")
        if not exenta:
            lines.append(f"📊 *IVA 19%:* ${iva_total:,.0f}")
        if total_imp_esp:
            lines.append(f"⛽ *Imp. Específico:* ${total_imp_esp:,.0f}")
        lines.append(f"💰 *TOTAL:* ${total_con_iva:,.0f}\n")
    lines.append("¿Qué deseas hacer?")
    return "\n".join(lines)


async def _show_preview(query, context):
    from utils.keyboards import main_keyboard
    items = context.user_data.get("pending_items", [])
    await query.edit_message_text(_build_preview(items), parse_mode="Markdown",
                                    reply_markup=main_keyboard())
