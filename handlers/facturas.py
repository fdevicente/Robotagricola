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
