"""
processors/extractor.py
Motor: Escaneo tipo CamScanner → OCR → IA (Anthropic Claude o Ollama)
"""
import os, re, json, logging, base64, tempfile
from datetime import datetime
import requests

logger = logging.getLogger(__name__)

# Ruta explícita a Tesseract (necesario en Windows cuando no está en PATH del subprocess)
try:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
except ImportError:
    pass


# ─────────────────────────────────────────────
# PROVEEDORES DESDE EXCEL
# ─────────────────────────────────────────────
def _load_proveedores_rows():
    try:
        from config import EXCEL_PATH
        import openpyxl
        wb   = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        ws   = wb["Proveedores"]
        rows = [(str(r[1]).strip(), str(r[2]).strip())
                for r in ws.iter_rows(min_row=3, values_only=True) if r[1] and r[2]]
        wb.close()
        return rows
    except Exception as e:
        logger.warning(f"No se pudo cargar proveedores: {e}")
        return []


def _match_proveedor(rut, nombre):
    rows    = _load_proveedores_rows()
    rut_n   = rut.replace(".", "").replace(" ", "").upper() if rut else ""
    nomb_n  = nombre.upper().strip() if nombre else ""
    for n_ex, r_ex in rows:
        if rut_n and rut_n == r_ex.replace(".", "").replace(" ", "").upper():
            logger.info(f"Match RUT {rut} → {n_ex}")
            return n_ex, r_ex
    for n_ex, r_ex in rows:
        if nomb_n and (nomb_n in n_ex.upper() or n_ex.upper() in nomb_n):
            logger.info(f"Match nombre '{nombre}' → {n_ex}")
            return n_ex, r_ex
    return None, None


# ─────────────────────────────────────────────
# ESCANEO TIPO CAMSCANNER
# ─────────────────────────────────────────────
def _order_points(pts):
    import numpy as np
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect


def _four_point_transform(image, pts):
    import numpy as np, cv2
    rect = _order_points(pts)
    tl, tr, br, bl = rect
    maxW = max(int(np.linalg.norm(br-bl)), int(np.linalg.norm(tr-tl)))
    maxH = max(int(np.linalg.norm(tr-br)), int(np.linalg.norm(tl-bl)))
    dst  = np.array([[0,0],[maxW-1,0],[maxW-1,maxH-1],[0,maxH-1]], dtype="float32")
    M    = cv2.getPerspectiveTransform(rect, dst)
    return cv2.warpPerspective(image, M, (maxW, maxH))


def _scan_document(image_path: str) -> str:
    """Escaneo tipo CamScanner: detecta esquinas, corrige perspectiva, umbral adaptativo."""
    try:
        import cv2, numpy as np
        from PIL import Image

        img = cv2.imread(image_path)
        if img is None: raise ValueError("No se pudo leer imagen")
        orig   = img.copy()
        h, w   = img.shape[:2]
        ratio  = 800.0 / max(h, w)
        small  = cv2.resize(img, (int(w*ratio), int(h*ratio)))

        gray    = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray, (5,5), 0)
        edges   = cv2.Canny(blurred, 30, 120)
        kernel  = np.ones((5,5), np.uint8)
        dilated = cv2.dilate(edges, kernel, iterations=2)

        contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        contours    = sorted(contours, key=cv2.contourArea, reverse=True)[:5]

        doc_pts = None
        for cnt in contours:
            peri   = cv2.arcLength(cnt, True)
            approx = cv2.approxPolyDP(cnt, 0.02*peri, True)
            if len(approx) == 4:
                doc_pts = approx.reshape(4,2) / ratio
                logger.info("Documento detectado con 4 esquinas — corrigiendo perspectiva")
                break

        if doc_pts is not None:
            candidate = _four_point_transform(orig, doc_pts)
            # Validar que el resultado sea razonable (mínimo 200x200 px)
            if candidate.shape[0] >= 200 and candidate.shape[1] >= 200:
                warped = candidate
                logger.info(f"Perspectiva corregida: {candidate.shape[1]}x{candidate.shape[0]}")
            else:
                logger.warning(f"Warp inválido ({candidate.shape[1]}x{candidate.shape[0]}), usando imagen completa")
                warped = orig
                doc_pts = None  # forzar fallback
        if doc_pts is None:
            # Recorte del área más grande
            best_rect, best_area = None, 0
            for cnt in contours:
                area = cv2.contourArea(cnt)
                if area < h*w*ratio**2*0.1: continue
                peri   = cv2.arcLength(cnt, True)
                approx = cv2.approxPolyDP(cnt, 0.02*peri, True)
                if len(approx) >= 4 and area > best_area:
                    best_area = area
                    x,y,cw,ch = cv2.boundingRect(approx)
                    best_rect  = (int(x/ratio), int(y/ratio), int(cw/ratio), int(ch/ratio))
            if best_rect:
                x,y,cw,ch = best_rect
                m = 20
                x=max(0,x-m); y=max(0,y-m)
                cw=min(w-x,cw+m*2); ch=min(h-y,ch+m*2)
                warped = orig[y:y+ch, x:x+cw]
                logger.info(f"Recorte simple: {cw}x{ch}")
            else:
                warped = orig
                logger.info("Usando imagen completa")

        # Umbral adaptativo (efecto escaneo nítido)
        gray_w   = cv2.cvtColor(warped, cv2.COLOR_BGR2GRAY)
        scanned  = cv2.adaptiveThreshold(gray_w, 255,
                      cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 21, 10)

        scan_path = os.path.splitext(image_path)[0] + "_scan.png"
        cv2.imwrite(scan_path, scanned)
        logger.info(f"Scan guardado: {scan_path} ({scanned.shape[1]}x{scanned.shape[0]})")
        return scan_path

    except ImportError:
        logger.warning("OpenCV no disponible")
        return _resize_image(image_path)
    except Exception as e:
        logger.warning(f"Escaneo falló: {e}")
        return _resize_image(image_path)


def _resize_image(image_path, max_size=1600):
    try:
        from PIL import Image
        # Usar context manager para que Windows libere el handle
        # antes de que main.py intente renombrar el archivo original.
        with Image.open(image_path) as img:
            img.thumbnail((max_size, max_size), Image.LANCZOS)
            out = os.path.splitext(image_path)[0] + "_resized.jpg"
            img.save(out, quality=85)
        return out
    except Exception as e:
        logger.warning(f"Resize falló: {e}")
        return image_path


# ─────────────────────────────────────────────
# OCR
# ─────────────────────────────────────────────
def _ocr_text(image_path):
    try:
        import pytesseract
        from PIL import Image
        # Context manager para liberar el handle (Windows lo necesita
        # para que después se pueda borrar/renombrar la imagen).
        with Image.open(image_path) as img:
            return pytesseract.image_to_string(img, lang="eng").strip()
    except Exception as e:
        logger.warning(f"Tesseract falló: {e}")
        return ""


def _combinar_ocr(tess: str, surya: str) -> str:
    """Combina textos de dos motores OCR. Si ambos tienen contenido, los etiqueta."""
    t = (tess or "").strip()
    s = (surya or "").strip()
    if t and s:
        return f"=== OCR Tesseract ===\n{t}\n\n=== OCR Surya ===\n{s}"
    return t or s


# ─────────────────────────────────────────────
# EXTRACCIÓN TOTAL DESDE OCR (chequeo cruzado)
# ─────────────────────────────────────────────
def _extraer_total_ocr(ocr_text: str) -> float:
    """Busca el TOTAL de la factura en el texto OCR como chequeo cruzado.
    Ignora SUBTOTAL y TOTAL NETO. Retorna el mayor valor encontrado en líneas con 'TOTAL'.
    Los puntos en números se interpretan siempre como separadores de miles (los totales son enteros).
    """
    if not ocr_text:
        return 0.0
    candidatos = []
    excluir = ("subtotal", "total neto", "total exento", "total exenta",
               "totalizado", "total descuento", "total afecto")
    for line in ocr_text.split("\n"):
        lu = line.lower().strip()
        if "total" not in lu:
            continue
        if any(x in lu for x in excluir):
            continue
        # Extraer todos los números de la línea: puntos siempre = miles en totales
        for m in re.finditer(r"\d[\d.]*\d|\d", line):
            raw = m.group().replace(".", "")
            try:
                val = float(raw)
                if val >= 1000:
                    candidatos.append(val)
            except ValueError:
                pass
    return max(candidatos) if candidatos else 0.0


# ─────────────────────────────────────────────
# PDF → IMÁGENES
# ─────────────────────────────────────────────
def _pdf_to_images(pdf_path):
    try:
        import fitz
        doc     = fitz.open(pdf_path)
        tmp_dir = tempfile.mkdtemp()
        paths   = []
        for i, page in enumerate(doc):
            pix  = page.get_pixmap(dpi=200)
            path = os.path.join(tmp_dir, f"page_{i+1}.png")
            pix.save(path)
            paths.append(path)
        doc.close()
        return paths
    except Exception as e:
        logger.error(f"Error convirtiendo PDF: {e}")
        return []


# ─────────────────────────────────────────────
# CONSTANTES COMBUSTIBLE
# ─────────────────────────────────────────────
COMBUSTIBLE_KEYWORDS = (
    "petroleo", "petróleo", "diesel", "gasolina", "bencina",
    "combustible", "kerosene", "glp", "gas licuado", "lubricante",
)

# Patrones OCR para impuestos específicos a combustibles (IEF, IEV, FEPP)
IMP_PATTERNS = [
    r"ief[\s\-:$]*([\d.,]+)",
    r"iev[\s\-/:a-záéíóú]*[\s:$]*([\d.,]+)",
    r"fepp[\s\-:$]*([\d.,]+)",
    r"imp[.\s]*esp[a-záéíóú]*[.\s:$]*([\d.,]+)",
    r"impuesto\s+espec[íi]fico[.\s:$]*([\d.,]+)",
    r"i\.?e\.?c\.?[\s:$]*([\d.,]+)",
    r"imp\.?\s*comb[a-záéíóú]*[\s:$]*([\d.,]+)",
    r"impuesto\s+al\s+combustible[\s:$]*([\d.,]+)",
]


# ─────────────────────────────────────────────
# PROMPT
# ─────────────────────────────────────────────
PROMPT = """Eres un lector experto de facturas chilenas. Analiza PRIMERO la imagen y usa el OCR como apoyo.
Responde SOLO con JSON válido, sin texto adicional antes ni después.

TEXTO OCR (puede tener errores, prioriza lo que VES en la imagen):
{ocr_text}

JSON requerido (exactamente este formato):
{{"items":[{{"Fecha Emision":"YYYY-MM-DD","Fecha Vencimiento":"YYYY-MM-DD","Nombre Factura / Proveedor":"texto","Rut":"XX.XXX.XXX-X","Documento":"tipo","Numero Factura / Nro Documento":"numero","Referencia Factura":null,"Detalle / Glosa":"texto corto","Glosa II":"texto completo","Valor unitario":0,"Cantidad":1,"Monto / TOTAL":0,"Impuesto Especifico":null,"Total Factura":0}}]}}

REGLAS IMPORTANTES:

0. TOTAL FACTURA — LO PRIMERO QUE DEBES LEER:
   Antes de extraer cualquier otro dato, localiza el TOTAL que aparece al pie del documento
   (puede llamarse "TOTAL", "TOTAL A PAGAR", "TOTAL FACTURA", "TOTAL DOCUMENTO", etc.).
   - Escribe ese valor en el campo "Total Factura" de TODOS los ítems (mismo valor en cada uno).
   - NUNCA lo dejes en 0 o null si el número está visible en el documento.
   - Este valor es el anchor: el sistema lo usará para calcular NETO e IVA hacia atrás.
     Factura CON IVA:  NETO = Total Factura ÷ 1.19
     Factura EXENTA:   NETO = Total Factura (sin IVA)
   - Para facturas multi-ítem: "Monto / TOTAL" de cada línea = total de ESA línea con IVA.
     La suma de todos los "Monto / TOTAL" debe ser igual a "Total Factura".
   - REGLA CRÍTICA DE FORMATO: Los totales en pesos chilenos son SIEMPRE enteros (sin centavos).
     El punto en el TOTAL es SIEMPRE separador de miles. Devuelve el número SIN puntos.
     Ejemplos: "771.784" → 771784 | "1.234.567" → 1234567 | "648.553" → 648553
     Verifica con el texto "SON: … pesos" si hay dudas sobre el valor real.

1. EMISOR (proveedor) = quien cobra = razón social y RUT en el encabezado SUPERIOR de la factura.
   Receptor = bajo SEÑOR(ES)/CLIENTE — NO usarlo como proveedor.

   CASO ESPECIAL — ESTACIONES DE SERVICIO (Copec, Shell, ENEX, PETROBRAS, etc.):
   Estas facturas tienen DOS razones sociales:
   - La razón social del encabezado (ej: "COMERCIAL DOMINGA LIMITADA") con su RUT → ese ES el proveedor.
   - "POR CUENTA DE: COPEC S.A." → es solo el operador de la red, NO el proveedor.
   Usa SIEMPRE la razón social del encabezado y SU RUT como proveedor.
   Ejemplo: R.Social = "COMERCIAL DOMINGA LIMITADA", RUT = 76.379.987-5 → proveedor correcto.

2. TIPO DE DOCUMENTO y NUMERO:
   - "Documento" = tipo exacto del recuadro superior derecho. Valores posibles:
     "Factura Electronica", "Factura Exenta Electronica", "Factura No Afecta",
     "Nota de Credito", "Nota de Debito", "Boleta Electronica", "Boleta Exenta".
     IMPORTANTE: si dice "EXENTA", "NO AFECTA", o "NO AFECTA O EXENTA", inclúyelo EXACTO en el tipo.
     Ejemplo: "FACTURA NO AFECTA O EXENTA ELECTRONICA" → Documento = "Factura No Afecta o Exenta Electronica".
     Las BOLETAS son documentos más simples (tickets de compra, supermercados, ferreterías,
     bencineras, etc.) sin RUT del comprador. Si es boleta, el RUT y datos del receptor no importan.
   - "Numero Factura / Nro Documento" = el número del documento (N° o Nº). Solo dígitos.
   - "Referencia Factura" = si es Nota de Crédito o Nota de Débito, busca el número de la
     FACTURA ORIGINAL a la que hace referencia (suele decir "Ref.", "Documento referencia",
     "Factura afectada", etc.). Solo dígitos. Si es una Factura normal, usa null.

3. GLOSA: Dos campos distintos:
   - "Detalle / Glosa" = descripción CORTA y útil del producto/servicio. Si el nombre es técnico
     (ej: INTREPID, CORAGEN, ABAMECTIN), agrega qué tipo de producto es
     (ej: "INTREPID - insecticida", "ABAMECTIN 18 EC - acaricida/insecticida", "Diesel combustible").
     Usa tu conocimiento de productos agrícolas para describir.
   - Para CUENTAS DE SERVICIOS (electricidad CGE, agua, gas, etc.): busca el PERIODO DE CONSUMO
     (fechas de lectura de medidor) y pon en Detalle/Glosa:
     "Consumo electrico periodo: DD MMM a DD MMM" (ej: "Consumo electrico periodo: 27 NOV a 29 DIC").
   - "Glosa II" = descripción completa tal como aparece en la factura.
   NO uses el giro/actividad económica del emisor ni del receptor como glosa.

4. MONTOS Y DECIMALES — REGLA CRÍTICA:
   Usa PUNTO como separador decimal en el JSON (nunca coma).
   Valor unitario = precio neto SIN IVA por unidad.

   CÓMO DISTINGUIR punto-decimal vs punto-miles:
   a) Si el número tiene MÁS DE UN punto → es separador de miles. Elimínalos.
      Ejemplo: "1.689.786" → 1689786
   b) Si el número tiene UN SOLO punto:
      - Multiplica unitario × cantidad.
      - Si el resultado SE ACERCA al neto de la línea → el punto ES decimal. CONSÉRVALO.
        Ejemplo: 105.612 × 16000 = 1.689.792 ≈ neto línea → unitario = 105.612
      - Si el resultado NO cuadra → el punto es miles. Elimínalo.
        Ejemplo: 105.612 como miles = 105612 → 105612 × 1 = 105612 ≈ total → unitario = 105612
   c) Precios por litro de combustible/lubricante SIEMPRE tienen decimales (2 o 3 cifras tras el punto).
   d) Proveedores como Martínez y Valdivieso (agroquímicos/fertilizantes) usan precios con
      hasta 3 decimales por unidad (kg, litro, unidad). Si ves "2.345,678" → 2345.678.
      Conserva SIEMPRE los 3 decimales en el JSON para estos casos.
   e) NUNCA redondees el valor unitario a entero si tiene decimales reales.
   f) En facturas de ferretería/materiales (Sodimac, Easy, Construmart, etc.) los precios
      suelen tener el sufijo "C/U" (costo por unidad) junto al valor: "35.285,71C/U".
      Ese sufijo no es parte del número — ignóralo y extrae solo el valor numérico: 35285.71.
      La columna "VALOR" en estas facturas muestra el NETO de la línea (sin IVA).
      El IVA aparece solo al final del documento como un total.

4b. IMPUESTO ESPECÍFICO (combustibles y lubricantes):
   APLICA A CUALQUIER PROVEEDOR que venda petróleo, diesel, gasolina, bencina o lubricantes
   (bencineras, COPEVAL, distribuidoras agrícolas, etc.). Estructura especial:
     SUBTOTAL NETO  → es el neto real afecto a IVA
     IVA (19%)      → 19% del neto
     IEF            → Impuesto Específico a los Combustibles
     IEV/FEPP       → otro impuesto al combustible
     TOTAL          → suma de todo lo anterior

   REGLAS:
   a) "Impuesto Especifico" = IEF + IEV/FEPP sumados. Ambos son impuestos específicos.
      Ejemplo: IEF=$7.270 + IEV/FEPP=$390 → Impuesto Especifico = 7660
   b) "Valor unitario" = SUBTOTAL NETO ÷ Cantidad (precio neto real por litro, SIN impuestos).
      NO uses el $/Unit del cuadro de productos (ese es precio all-in con todos los impuestos).
      Ejemplo: SUBTOTAL NETO=$10.370, Cantidad=17.406 litros → Valor unitario = 10370/17.406 = 595.8
   c) "Cantidad" = litros despachados.
      CASO GENERAL: el punto en la cantidad ES decimal. "17.406" → 17.406 litros.
      EXCEPCIÓN COPEVAL: en facturas de COPEVAL, la columna "Cant" usa punto como miles.
        "2.000" → 2,000 litros. "5.000" → 5,000 litros.
        Reconócelo porque la parte decimal son todos ceros ("000").
   d) "Monto / TOTAL" = el TOTAL de la factura (suma de todo).
      COPEVAL COMBUSTIBLE — formato especial:
      Precio Unitario usa COMA como decimal: "746,8905" = 746.8905 $/litro (4 decimales).
      Cantidad usa PUNTO como miles cuando termina en ceros: "2.000" = 2,000 litros.
      Verificación: Valor Neto línea = Cantidad × Precio = 2000 × 746.8905 = 1.493.781 ✓
   e) "Detalle / Glosa" = tipo de combustible: "Gasolina 93 - combustible", "Diesel - combustible", etc.
   f) Otros impuestos específicos no combustible: busca líneas "Imp. Específico", "Imp. Esp.", "I.E.C."
      y súmalos todos en el campo Impuesto Especifico.

5. BOLETA DE HONORARIOS ELECTRONICA:
   Documento emitido por profesionales independientes. NO tiene IVA.
   Estructura:
     Total Honorarios  → monto bruto del servicio (lo que ASE debe pagar en total)
     17.5% Impto. Retenido → impuesto que ASE retiene y paga al SII
     Total             → lo que recibe el profesional (Total Honorarios − Retenido)

   REGLAS:
   a) "Documento" = "Boleta de Honorarios Electronica"
   b) "Monto / TOTAL" = campo "Total" de la boleta = lo que recibe el profesional
      (= Total Honorarios − Impto. Retenido). Ej: 1.169.591 − 204.678 = 964.913
   c) "Total Factura"  = Total Honorarios (el bruto, costo total para ASE). Ej: 1.169.591
   d) "Impuesto Especifico" = monto de Impto. Retenido (lo que ASE retiene y paga al SII). Ej: 204.678
   e) "Valor unitario" = mismo valor que "Monto / TOTAL" (pago neto al profesional), "Cantidad" = 1
   f) NO hay IVA; NO apliques 1.19 a ningún valor.
   g) Emisor = el profesional (RUT y nombre en el encabezado de la boleta).

6. MULTI-ITEM: Si la factura tiene varios productos en la tabla, crea UN objeto por cada ítem.
   Cada ítem lleva los mismos datos de cabecera (fecha, proveedor, RUT, N° factura)
   pero su propia descripción, cantidad, valor unitario y su propio Monto/TOTAL.
   IMPORTANTE: Monto/TOTAL de cada ítem = total NETO de esa línea SIN IVA (unitario × cantidad).
   NO multipliques por 1.19 — el sistema calcula el IVA automáticamente.
   NO pongas el total general de la factura en cada ítem.
   Si la factura tiene un solo ítem, Monto/TOTAL = unitario × cantidad (NETO, sin IVA).

   ⚠️ VERIFICACIÓN OBLIGATORIA ANTES DE RESPONDER (multi-ítem):
   - Suma todos los (Valor unitario × Cantidad) → debe ser igual al NETO impreso al pie
     del documento (campo "NETO" o "Total Neto"). NO al TOTAL, al NETO.
   - Si la suma NO cuadra con el NETO impreso (±1% tolerancia), NO devuelvas aún.
     Vuelve a leer la imagen columna por columna (Cantidad | P.Unitario | Valor Total)
     y corrige los valores hasta que la suma = NETO.
   - Los sellos ("PAGADO", "DESPACHADO") y manos/dedos que cubren números son frecuentes.
     Deduce el valor oculto usando la restricción: Valor Total de la línea = Cantidad × P.Unitario,
     y la suma de todos los Valor Total = NETO impreso.
   - NUNCA inventes valores plausibles si no estás seguro. Si una línea es totalmente ilegible
     incluso después de releer, pon Valor unitario=0 y deja claro en Detalle/Glosa "ilegible".
   - Confirma también que la Cantidad leída es razonable (entero típico: 3, 15, 30, 32, 39, 50)
     y el P.Unitario tiene el orden de magnitud correcto (tubería 90mm cuesta más que 75mm,
     reductores cuestan menos que tubos, etc.). La lógica de negocio es una pista útil.

7. Si un campo no existe o no es legible, usa null.

SOLO JSON, sin explicaciones."""


# ─────────────────────────────────────────────
# PARSER JSON ROBUSTO
# ─────────────────────────────────────────────
def _parse_json(raw: str) -> list:
    """Extrae lista de items de la respuesta, tolerando formatos imperfectos."""
    # Quitar markdown
    raw = re.sub(r"```(?:json)?", "", raw).strip()

    # Buscar el primer { o el patrón "items":
    brace = raw.find("{")
    items_match = re.search(r'"items"\s*:\s*\[', raw)

    if brace == -1 and items_match:
        # Respuesta es directamente "items": [...]
        raw = "{" + raw[items_match.start():]
    elif brace > 0:
        raw = raw[brace:]

    # Reparar truncado
    if not raw.rstrip().endswith("}"):
        raw = re.sub(r",\s*$", "", raw.rstrip())
        raw += "}" * max(0, raw.count("{") - raw.count("}"))
        raw += "]" * max(0, raw.count("[") - raw.count("]"))
        raw += "}" * max(0, raw.count("{") - raw.count("}"))
        logger.warning("JSON reparado (truncado)")

    try:
        return json.loads(raw).get("items", [])
    except json.JSONDecodeError:
        # Último intento: buscar cualquier {...}
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if m:
            try: return json.loads(m.group(0)).get("items", [])
            except: pass
        logger.error(f"JSON no parseable:\n{raw[:300]}")
        return []


# ─────────────────────────────────────────────
# MOTOR ANTHROPIC CLAUDE
# ─────────────────────────────────────────────
def _call_claude(image_path: str, ocr_text: str) -> list:
    from config import ANTHROPIC_API_KEY
    if not ANTHROPIC_API_KEY:
        logger.warning("ANTHROPIC_API_KEY no configurada")
        return []

    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    ext = os.path.splitext(image_path)[1].lower()
    media_type = "image/png" if ext == ".png" else "image/jpeg"
    prompt = PROMPT.format(ocr_text=ocr_text or "(sin OCR)")

    content = [
        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
        {"type": "text", "text": prompt}
    ]
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }

    import time

    # Sonnet primero, Haiku como fallback
    models = ["claude-sonnet-4-6", "claude-haiku-4-5-20251001"]

    for attempt, model in enumerate(models):
        try:
            payload = {
                "model": model,
                "max_tokens": 8192,   # subido de 4096 — evita truncado en facturas largas
                "messages": [{"role": "user", "content": content}]
            }
            logger.info(f"Llamando a {model} ({len(img_b64)//1024} KB) — intento {attempt+1}/3...")
            t0 = time.time()
            resp = requests.post("https://api.anthropic.com/v1/messages",
                                 json=payload, headers=headers, timeout=150)
            resp.raise_for_status()
            resp_data = resp.json()
            raw = resp_data["content"][0]["text"].strip()
            stop = resp_data.get("stop_reason", "unknown")
            usage = resp_data.get("usage", {})
            elapsed = time.time() - t0
            logger.info(f"Claude {model} OK en {elapsed:.1f}s — stop={stop}, tokens_out={usage.get('output_tokens')}")
            logger.info(f"Claude respondió: {raw[:500]}")
            if stop == "max_tokens":
                logger.warning("Respuesta truncada por max_tokens — intentando modelo siguiente")
                continue   # respuesta cortada → probar siguiente modelo
            items = _parse_json(raw)
            if items:
                return items
            # JSON vacío o inválido → probar siguiente modelo
            logger.warning(f"JSON vacío o inválido en intento {attempt+1} — reintentando...")
            continue
        except requests.exceptions.Timeout:
            logger.warning(f"Timeout en {model} (intento {attempt+1}/3)")
            continue
        except requests.exceptions.HTTPError as e:
            status = e.response.status_code if e.response is not None else 0
            if status == 429:
                wait = min(10 * (2 ** attempt), 60)
                logger.warning(f"Rate limit (429) — esperando {wait}s antes de reintentar...")
                time.sleep(wait)
                continue
            logger.error(f"HTTP error {status} en {model}: {e}")
            if attempt < len(models) - 1:
                continue
            return []
        except requests.exceptions.RequestException as e:
            logger.error(f"Error {model}: {e}")
            if attempt < len(models) - 1:
                continue
            return []

    logger.error("Todos los intentos de Claude API fallaron")
    return []


# ─────────────────────────────────────────────
# MOTOR OLLAMA
# ─────────────────────────────────────────────
def _call_ollama(image_path: str, ocr_text: str) -> list:
    from config import OLLAMA_HOST, OLLAMA_MODEL, OLLAMA_TIMEOUT
    try:
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")
        # Prompt con ejemplos reales de Claude (few-shot learning)
        prompt = _prompt_ollama_con_ejemplos(ocr_text)
        payload = {
            "model": OLLAMA_MODEL,
            "messages": [{"role": "user", "content": prompt, "images": [img_b64]}],
            "stream": False,
            "options": {"temperature": 0.1, "num_ctx": 4096}
        }
        logger.info(f"Llamando a Ollama/{OLLAMA_MODEL} ({len(img_b64)//1024} KB)...")
        resp = requests.post(f"{OLLAMA_HOST}/api/chat", json=payload,
                             timeout=OLLAMA_TIMEOUT)
        resp.raise_for_status()
        raw = resp.json()["message"]["content"].strip()
        logger.info(f"Ollama respondió: {raw[:300]}")
        return _parse_json(raw)
    except Exception as e:
        logger.warning(f"Ollama no disponible o tardó demasiado: {e}")
        return []


def _call_claude_simple(image_path: str, ocr_text: str) -> list:
    """Prompt mínimo de emergencia — más tolerante, menos reglas."""
    import time
    from config import ANTHROPIC_API_KEY
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")
    ext = os.path.splitext(image_path)[1].lower()
    media_type = "image/png" if ext == ".png" else "image/jpeg"
    prompt = PROMPT_SIMPLE.format(ocr_text=ocr_text or "(sin OCR)")
    content = [
        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": img_b64}},
        {"type": "text", "text": prompt}
    ]
    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
    }
    for model in ["claude-sonnet-4-6", "claude-haiku-4-5-20251001"]:
        try:
            payload = {"model": model, "max_tokens": 4096,
                       "messages": [{"role": "user", "content": content}]}
            resp = requests.post("https://api.anthropic.com/v1/messages",
                                 json=payload, headers=headers, timeout=120)
            resp.raise_for_status()
            raw = resp.json()["content"][0]["text"].strip()
            logger.info(f"Prompt simple {model}: {raw[:300]}")
            items = _parse_json(raw)
            if items:
                logger.info(f"Prompt simple exitoso con {model}")
                return items
        except Exception as e:
            logger.warning(f"Prompt simple {model} falló: {e}")
    return []


PROMPT_SIMPLE = """Extrae los datos de este documento y responde SOLO con JSON válido.
Texto OCR de apoyo: {ocr_text}

{{"items":[{{"Fecha Emision":"YYYY-MM-DD","Fecha Vencimiento":"YYYY-MM-DD",
"Nombre Factura / Proveedor":"texto","Rut":"XX.XXX.XXX-X","Documento":"tipo",
"Numero Factura / Nro Documento":"numero","Referencia Factura":null,
"Detalle / Glosa":"descripcion corta","Glosa II":"descripcion completa",
"Valor unitario":0,"Cantidad":1,"Monto / TOTAL":0,"Impuesto Especifico":null,
"Total Factura":0}}]}}

Reglas mínimas:
- Proveedor = quien emite (encabezado superior), NO el receptor.
- Monto/TOTAL y Total Factura = TOTAL del documento (con IVA si aplica).
- Si hay varios ítems, crea un objeto por ítem con los mismos datos de cabecera.
- NUNCA uses separadores de miles en los números. Esto es JSON, no un recibo:
  el punto de "$116.463" es separador de miles y en JSON se lee como decimal,
  así que $116.463 se convertiría en ciento dieciséis coma cuatro.
    correcto:   "Total Factura":116463
    incorrecto: "Total Factura":116.463
  Escribe los montos como enteros pelados, sin puntos, sin comas y sin "$".
- Solo JSON, sin texto adicional."""


def _ollama_disponible() -> bool:
    """Verifica si Ollama está corriendo y tiene el modelo de visión instalado."""
    from config import OLLAMA_HOST, OLLAMA_MODEL, OLLAMA_TIMEOUT
    try:
        resp = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=3)
        modelos = [m["name"] for m in resp.json().get("models", [])]
        return any(OLLAMA_MODEL.split(":")[0] in m for m in modelos)
    except Exception:
        return False


# ─────────────────────────────────────────────
# APRENDIZAJE: Claude enseña a Ollama
# ─────────────────────────────────────────────
def _ruta_ejemplos() -> str:
    from config import DOWNLOAD_DIR
    return os.path.join(DOWNLOAD_DIR, "ollama_ejemplos.json")


def _cargar_ejemplos() -> list:
    path = _ruta_ejemplos()
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _guardar_ejemplo(items: list):
    """Guarda el resultado correcto de Claude como ejemplo para Ollama."""
    if not items:
        return
    try:
        ejemplos = _cargar_ejemplos()
        primer = items[0]
        doc_tipo = str(primer.get("Documento") or "").lower()
        # Guardar solo los campos clave para el ejemplo (sin imagen)
        ejemplo = {
            "tipo": doc_tipo,
            "proveedor": primer.get("Nombre Factura / Proveedor"),
            "json": json.dumps({"items": items}, ensure_ascii=False)
        }
        # Mantener máximo 20 ejemplos, 1 por tipo de documento para diversidad
        tipos_existentes = {e["tipo"] for e in ejemplos}
        if doc_tipo in tipos_existentes:
            # Reemplazar el ejemplo existente de ese tipo
            ejemplos = [e for e in ejemplos if e["tipo"] != doc_tipo]
        ejemplos.append(ejemplo)
        ejemplos = ejemplos[-20:]  # máximo 20
        with open(_ruta_ejemplos(), "w", encoding="utf-8") as f:
            json.dump(ejemplos, f, ensure_ascii=False, indent=2)
        logger.info(f"Ejemplo guardado para Ollama (tipo: {doc_tipo})")
    except Exception as e:
        logger.warning(f"No se pudo guardar ejemplo: {e}")


def _prompt_ollama_con_ejemplos(ocr_text: str) -> str:
    """Genera el prompt simple con ejemplos reales de Claude para few-shot learning."""
    ejemplos = _cargar_ejemplos()
    base = PROMPT_SIMPLE.format(ocr_text=ocr_text or "(sin OCR)")
    if not ejemplos:
        return base
    # Agregar hasta 2 ejemplos relevantes
    shots = "\n\nEJEMPLOS DE RESPUESTAS CORRECTAS (aprende de estos):\n"
    for ej in ejemplos[-2:]:
        shots += f"\nTipo: {ej['tipo']}\nJSON correcto:\n{ej['json']}\n"
    return base + shots


def _comparar_resultados(claude: list, ollama: list) -> dict:
    """Compara campos clave entre Claude y Ollama para medir precisión."""
    if not claude or not ollama:
        return {}
    c, o = claude[0], ollama[0]
    campos = ["Nombre Factura / Proveedor", "Total Factura", "Documento",
              "Numero Factura / Nro Documento"]
    aciertos = sum(1 for k in campos if str(c.get(k) or "").strip().lower()
                   == str(o.get(k) or "").strip().lower())
    return {"aciertos": aciertos, "total": len(campos),
            "porcentaje": round(aciertos / len(campos) * 100)}


def _call_ia(image_path: str, ocr_text: str) -> list:
    """Estrategia paralela:
    - Claude y Ollama corren simultáneamente
    - Se usa siempre el resultado de Claude (preciso)
    - El resultado de Claude se guarda para que Ollama aprenda (few-shot)
    - Si Claude falla, Ollama actúa como fallback
    """
    import concurrent.futures
    from config import ANTHROPIC_API_KEY

    ollama_ok = _ollama_disponible()
    futures = {}

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
        if ANTHROPIC_API_KEY:
            futures["claude"] = ex.submit(_call_claude, image_path, ocr_text)
        if ollama_ok:
            futures["ollama"] = ex.submit(_call_ollama, image_path, ocr_text)

        claude_items, ollama_items = [], []

        if "claude" in futures:
            try:
                # 330s = 2 intentos × 150s + 30s buffer — evita cortar Claude a la mitad
                claude_items = futures["claude"].result(timeout=330) or []
            except Exception as e:
                logger.warning(f"Claude falló en paralelo: {e}")

        if "ollama" in futures:
            try:
                ollama_items = futures["ollama"].result(timeout=90) or []
            except Exception as e:
                logger.warning(f"Ollama falló en paralelo: {e}")

    # ── Claude tiene prioridad absoluta ───────────────────────────────────────
    if claude_items:
        if ollama_items:
            stats = _comparar_resultados(claude_items, ollama_items)
            logger.info(f"Ollama vs Claude: {stats.get('aciertos')}/{stats.get('total')} "
                        f"campos correctos ({stats.get('porcentaje')}%)")
        # Guardar resultado correcto de Claude para que Ollama aprenda
        _guardar_ejemplo(claude_items)
        return claude_items

    # ── Claude falló — usar Ollama como fallback ──────────────────────────────
    if ollama_items:
        logger.warning("Claude falló — usando resultado de Ollama como fallback")
        return ollama_items

    # ── Todo falló — prompt simplificado de emergencia ────────────────────────
    if ANTHROPIC_API_KEY:
        logger.warning("Todo falló — reintentando con prompt simplificado...")
        return _call_claude_simple(image_path, ocr_text)

    logger.error("Sin resultados de ningún motor")
    return []


# ─────────────────────────────────────────────
# REGISTRO DE FACTURAS PROCESADAS
# ─────────────────────────────────────────────
def _registrar_factura(items: list, file_path: str):
    """Guarda un registro de facturas procesadas en facturas_log.json."""
    try:
        from config import DOWNLOAD_DIR
        log_path = os.path.join(DOWNLOAD_DIR, "facturas_log.json")
        if os.path.exists(log_path):
            with open(log_path, "r", encoding="utf-8") as f:
                log = json.load(f)
        else:
            log = []

        for item in items:
            log.append({
                "timestamp":   datetime.now().isoformat(),
                "archivo":     os.path.basename(file_path),
                "proveedor":   item.get("Nombre Factura / Proveedor"),
                "rut":         item.get("Rut"),
                "nro_factura": item.get("Numero Factura / Nro Documento"),
                "fecha":       item.get("Fecha Emision"),
                "total":       item.get("Monto / TOTAL"),
            })

        with open(log_path, "w", encoding="utf-8") as f:
            json.dump(log, f, ensure_ascii=False, indent=2)
        logger.info(f"Factura registrada en log: {log_path}")
    except Exception as e:
        logger.warning(f"No se pudo registrar en log: {e}")


def _factura_duplicada(items: list) -> bool:
    """Verifica si la factura ya fue procesada (mismo proveedor + número)."""
    try:
        from config import DOWNLOAD_DIR
        log_path = os.path.join(DOWNLOAD_DIR, "facturas_log.json")
        if not os.path.exists(log_path):
            return False
        with open(log_path, "r", encoding="utf-8") as f:
            log = json.load(f)
        for item in items:
            nro = str(item.get("Numero Factura / Nro Documento") or "")
            rut = str(item.get("Rut") or "")
            for entry in log:
                if (str(entry.get("nro_factura") or "") == nro and
                    str(entry.get("rut") or "") == rut and nro and rut):
                    logger.warning(f"Factura duplicada detectada: {rut} Nº{nro}")
                    return True
    except Exception as e:
        logger.warning(f"Error verificando duplicado: {e}")
    return False


# ─────────────────────────────────────────────
# APRENDIZAJE DESDE CORRECCIONES DEL USUARIO
# ─────────────────────────────────────────────
def _aplicar_correcciones_aprendidas(items: list) -> list:
    """Lee correcciones_log.json y aplica patrones aprendidos de correcciones previas.
    - Nombres de proveedor: si el usuario corrigió el nombre para este RUT, lo aplica.
    - Totales: si hay un factor de corrección consistente (>=2 veces) para este RUT, lo aplica.
    """
    try:
        from config import DOWNLOAD_DIR
        log_path = os.path.join(DOWNLOAD_DIR, "correcciones_log.json")
        if not os.path.exists(log_path):
            return items
        with open(log_path, "r", encoding="utf-8") as f:
            log = json.load(f)
        if not log:
            return items

        from collections import defaultdict
        factores_total   = defaultdict(list)   # rut → [factores de corrección de totales]
        nombres_corr     = {}                  # rut → último nombre correcto
        glosas_corr      = defaultdict(list)   # rut → [(glosa_claude, glosa_usuario)]

        for entrada in log:
            rut   = entrada.get("rut") or ""
            campo = entrada.get("campo") or ""
            if not rut:
                continue
            if campo == "Nombre Factura / Proveedor":
                nombres_corr[rut] = entrada.get("valor_usuario")
            elif campo in ("Total Factura", "Monto / TOTAL"):
                factor = entrada.get("factor")
                if factor and abs(factor - 1.0) > 0.02:   # ignorar correcciones menores al 2%
                    factores_total[rut].append(factor)
            elif campo == "Detalle / Glosa":
                factores_total  # no usamos glosa aún

        for item in items:
            rut = item.get("Rut") or ""

            # Aplicar nombre corregido
            if rut in nombres_corr:
                nombre_aprendido = nombres_corr[rut]
                if nombre_aprendido and nombre_aprendido != item.get("Nombre Factura / Proveedor"):
                    logger.info(f"Aprendizaje: proveedor {item.get('Nombre Factura / Proveedor')!r} → {nombre_aprendido!r}")
                    item["Nombre Factura / Proveedor"] = nombre_aprendido

            # Aplicar factor de corrección de total si es consistente
            if rut in factores_total:
                factores = factores_total[rut]
                if len(factores) >= 2:
                    factor_med = sum(factores) / len(factores)
                    varianza   = max(abs(f - factor_med) / factor_med for f in factores)
                    if varianza < 0.05:   # factor consistente (< 5% de dispersión)
                        total_actual = float(item.get("Total Factura") or 0)
                        if total_actual > 0:
                            total_corr = round(total_actual * factor_med)
                            logger.info(
                                f"Aprendizaje: factor {factor_med:.3f} aplicado al total de {rut}: "
                                f"${total_actual:,.0f} → ${total_corr:,.0f}"
                            )
                            item["Total Factura"] = total_corr

    except Exception as e:
        logger.warning(f"Error aplicando correcciones aprendidas: {e}")
    return items


# ─────────────────────────────────────────────
# FUNCIÓN PRINCIPAL
# ─────────────────────────────────────────────
ASE_RUTS    = {"798577107", "79857710K", "798577100"}
ASE_NOMBRES = {"AGRICOLA SANTA ELISA", "AGRICOLA STA ELISA", "AGRICOLA STA. ELISA"}


def _limpiar_items(items: list) -> list:
    """Limpia ASE, hace match de proveedores y normaliza numéricos."""
    for item in items:
        rut    = item.get("Rut") or ""
        nombre = item.get("Nombre Factura / Proveedor") or ""

        # Limpiar si RUT es ASE
        rut_n = rut.replace(".", "").replace("-", "").replace(" ", "").upper()
        if rut_n in ASE_RUTS:
            logger.warning(f"RUT {rut} es ASE — limpiando")
            item["Rut"] = None; rut = ""

        # Limpiar si nombre es ASE
        if any(n in nombre.upper() for n in ASE_NOMBRES):
            logger.warning(f"Nombre '{nombre}' es ASE — limpiando")
            item["Nombre Factura / Proveedor"] = None
            item["Rut"] = None; rut = ""; nombre = ""

        # Match contra lista proveedores
        n_ex, r_ex = _match_proveedor(rut, nombre)
        if n_ex:
            item["Nombre Factura / Proveedor"] = n_ex
            item["Rut"] = r_ex

        # Limpiar número factura
        nro = item.get("Numero Factura / Nro Documento")
        if nro:
            nro_l = re.sub(r"[^\d]", "", str(nro))
            item["Numero Factura / Nro Documento"] = nro_l or None

        # Detectar M&V para tratamiento especial de totales
        nombre_u = str(item.get("Nombre Factura / Proveedor") or "").upper()
        es_mv = "MARTINEZ" in nombre_u or "VALDIVIESO" in nombre_u

        # Campos donde el valor NUNCA tiene decimales reales (siempre enteros)
        CAMPOS_TOTAL = {"Monto / TOTAL", "Total Factura", "Impuesto Especifico"}

        # Normalizar numéricos
        for field in ("Valor unitario", "Cantidad", "Monto / TOTAL", "Impuesto Especifico", "Total Factura"):
            val = item.get(field)
            if val is not None:
                try:
                    # Si ya es número del JSON (int/float)
                    if isinstance(val, (int, float)):
                        # Totales siempre enteros (pesos chilenos sin centavos)
                        if field in CAMPOS_TOTAL:
                            fval = float(val)
                            # CORRECCIÓN: float con exactamente 3 decimales → punto es miles chileno
                            # Ejemplo: 771.784 → 771784, 42.792 → 42792, 123.225 → 123225
                            # (En CLP los totales son siempre enteros, nunca tienen centavos)
                            if not isinstance(val, int):
                                s = repr(val)
                                if '.' in s:
                                    dec = s.split('.')[1].rstrip('0')
                                    if len(dec) == 3:
                                        fval = fval * 1000
                                        logger.info(f"{field}: float 3-dec → ×1000 ({val} → {round(fval)})")
                            item[field] = round(fval)
                        else:
                            item[field] = float(val)
                        continue
                    # Si es string, limpiar formato
                    v = str(val).replace("$","").replace(" ","").strip()
                    # Eliminar sufijos no numéricos: "C/U", "kg", "*", etc.
                    v = re.sub(r"[^\d.,\-]", "", v)
                    # Coma como decimal chileno: "105.612,50" → "105612.50"
                    if "," in v:
                        v = v.replace(".", "").replace(",", ".")
                    else:
                        if v.count(".") > 1:
                            # Múltiples puntos = miles siempre: "1.689.786" → 1689786
                            v = v.replace(".", "")
                        elif v.count(".") == 1:
                            partes = v.split(".")
                            decimal_part = partes[1]
                            todos_ceros = all(c == "0" for c in decimal_part)
                            es_copeval = "COPEVAL" in nombre_u
                            # SOLO COPEVAL: Cantidad "2.000" → 2000 litros (punto = miles)
                            # Para otros proveedores el punto en Cantidad puede ser decimal real
                            if field == "Cantidad" and todos_ceros and len(decimal_part) == 3 and es_copeval:
                                v = v.replace(".", "")
                                logger.info(f"COPEVAL Cantidad: punto tratado como miles → {v}")
                            # Campos de total: punto = miles si 3+ decimales o es M&V
                            elif field in CAMPOS_TOTAL and (es_mv or len(decimal_part) >= 3):
                                v = v.replace(".", "")
                                if es_mv:
                                    logger.info(f"M&V total: punto tratado como miles en {field} → {v}")
                            # Valor unitario con decimales reales: conservar
                    item[field] = float(v) if v else None
                except: item[field] = None

    # Corregir totales por línea
    doc0 = str(items[0].get("Documento") or "").lower()
    es_exenta = any(k in doc0 for k in ("exenta", "exento", "no afecta", "no afecto", "boleta de honorario"))

    # CORRECCIÓN: cross-check unitario × cantidad vs Monto/TOTAL
    # Si Claude escribió el precio en formato miles chileno (ej: 1.783 → float 1.783),
    # el cálculo quedó ~1000 veces menor que el Monto/TOTAL registrado.
    # Detectar ratio ≈ 1000 y corregir multiplicando el unitario × 1000.
    for item in items:
        unitario = item.get("Valor unitario")
        cantidad = item.get("Cantidad")
        monto    = item.get("Monto / TOTAL")
        if unitario and cantidad and monto and unitario > 0 and cantidad > 0 and monto > 0:
            calc = unitario * cantidad
            if calc > 0:
                ratio = monto / calc
                if 950 <= ratio <= 1050:
                    nuevo = round(unitario * 1000)
                    logger.info(f"Auto-corrección unitario: {unitario} → {nuevo} "
                                f"(ratio={ratio:.1f}, punto era separador miles)")
                    item["Valor unitario"] = nuevo

    # CORRECCIÓN: si Claude todavía dividió unitarios por 1.19 (hábito viejo),
    # sum(unit × cant) ≈ TotalFactura/1.19 en vez de ≈ NETO.
    # Detectar ratio ~0.84 y corregir unitarios = Monto/TOTAL / Cantidad.
    if not es_exenta and items:
        total_factura = float(items[0].get("Total Factura") or 0)
        imp_esp_total = sum(float(it.get("Impuesto Especifico") or 0) for it in items)
        base_iva = total_factura - imp_esp_total
        neto_esperado = base_iva / (1 + 0.19) if base_iva > 0 else 0
        suma_netos = sum(
            float(it.get("Monto / TOTAL") or 0) for it in items
        )
        if neto_esperado > 0 and suma_netos > 0:
            ratio_global = suma_netos / neto_esperado
            # ratio ≈ 0.84 → Claude dividió por 1.19; los Monto/TOTAL son neto/1.19
            if 0.80 <= ratio_global <= 0.88:
                logger.warning(
                    f"Auto-corrección global: Monto/TOTAL divididos por 1.19 "
                    f"(ratio={ratio_global:.3f}, suma={suma_netos:.0f}, "
                    f"neto_esperado={neto_esperado:.0f}). Corrigiendo ×1.19."
                )
                for it in items:
                    mt = float(it.get("Monto / TOTAL") or 0)
                    if mt > 0:
                        mt_corr = round(mt * 1.19)
                        it["Monto / TOTAL"] = mt_corr
                        cant = float(it.get("Cantidad") or 1)
                        if cant > 0:
                            u_new = mt_corr / cant
                            if abs(u_new - round(u_new)) < 0.01:
                                it["Valor unitario"] = int(round(u_new))
                            else:
                                it["Valor unitario"] = round(u_new, 6)
            # ratio ≈ 1.0 → Monto/TOTAL ya son neto correctos, todo OK
            elif 0.97 <= ratio_global <= 1.03:
                logger.info(f"Monto/TOTAL OK: suma netos={suma_netos:.0f} ≈ neto esperado={neto_esperado:.0f}")

    es_honorario = "boleta de honorario" in doc0

    if len(items) == 1:
        # Ítem único: si tenemos Total Factura verificado, usarlo como Monto/TOTAL
        # (Monto/TOTAL para ítem único = Total Factura, _compute_totals calcula hacia atrás)
        # EXCEPTO en honorarios: allí Monto/TOTAL = pago al profesional (< Total Factura)
        tf = float(items[0].get("Total Factura") or 0)
        if tf > 0 and not es_honorario:
            items[0]["Monto / TOTAL"] = round(tf)
            logger.info(f"Ítem único: Monto/TOTAL fijado desde Total Factura → {round(tf)}")
        elif float(items[0].get("Monto / TOTAL") or 0) == 0:
            unitario = float(items[0].get("Valor unitario") or 0)
            cantidad = float(items[0].get("Cantidad") or 1)
            if unitario > 0:
                items[0]["Monto / TOTAL"] = round(unitario * cantidad)

    if len(items) > 1:
        totals = [float(i.get("Monto / TOTAL") or 0) for i in items]
        # Si todos los ítems tienen el mismo total, es el total general → recalcular por línea
        if len(set(totals)) == 1 and totals[0] > 0:
            logger.info("Todos los ítems tienen el mismo total — recalculando por línea")
            for item in items:
                unitario = float(item.get("Valor unitario") or 0)
                cantidad = float(item.get("Cantidad") or 1)
                if unitario > 0:
                    total_linea = round(unitario * cantidad)
                    item["Monto / TOTAL"] = total_linea
                    logger.info(f"  {item.get('Detalle / Glosa')}: {unitario} × {cantidad} = {total_linea}")
        else:
            # Redondear los totales por línea a entero (pesos chilenos son enteros)
            for item in items:
                mt = item.get("Monto / TOTAL")
                if mt is not None:
                    item["Monto / TOTAL"] = round(float(mt))

    # Nota de Crédito → montos negativos (descuenta de la factura original)
    for item in items:
        doc = str(item.get("Documento") or "").lower()
        if "nota de credito" in doc or "nota de crédito" in doc or "nota credito" in doc:
            for field in ("Valor unitario", "Monto / TOTAL"):
                val = item.get(field)
                if val is not None and float(val) > 0:
                    item[field] = -abs(float(val))
            logger.info(f"Nota de Crédito: montos convertidos a negativos")

    return items


def process_file(file_path: str) -> dict:
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        image_paths = _pdf_to_images(file_path)
        if not image_paths:
            return {"status": "error", "message": "No se pudo convertir PDF a imágenes."}
    elif ext in (".jpg", ".jpeg", ".png", ".webp", ".bmp"):
        image_paths = [file_path]
    else:
        return {"status": "error", "message": f"Formato no soportado: {ext}"}

    all_items = []

    for img_path in image_paths:
        if ext != ".pdf":
            # Escaneo tipo CamScanner (para OCR)
            scan_path = _scan_document(img_path)
            # Imagen original redimensionada para Claude (sin umbral B/N)
            claude_img = _resize_image(img_path, max_size=1600)
        else:
            scan_path = img_path
            claude_img = img_path

        # OCR sobre imagen escaneada (B/N nítido ayuda a Tesseract)
        ocr_text = _ocr_text(scan_path)
        logger.info(f"=== OCR ===\n{ocr_text}\n=== FIN OCR ===")

        # Llamar a IA con imagen ORIGINAL (Claude lee mejor el color/gris)
        items = _call_ia(claude_img, ocr_text)

        # Validar Total Factura contra OCR (chequeo cruzado)
        total_ocr = _extraer_total_ocr(ocr_text)
        if total_ocr > 0 and items:
            for item in items:
                total_claude = float(item.get("Total Factura") or 0)
                if total_claude == 0:
                    item["Total Factura"] = total_ocr
                    logger.info(f"Total Factura fijado desde OCR: ${total_ocr:,.0f}")
                else:
                    diferencia = abs(total_claude - total_ocr) / max(total_ocr, total_claude)
                    if diferencia > 0.50:
                        # Diferencia > 50%: posible confusión con otro documento en la foto
                        # (ej: boleta de pago encima de la factura). Confiar en Claude.
                        logger.warning(
                            f"Total OCR (${total_ocr:,.0f}) difiere >50% del de Claude "
                            f"(${total_claude:,.0f}) — posible documento mezclado, se mantiene Claude"
                        )
                    elif diferencia > 0.02:
                        # Diferencia pequeña-moderada: OCR corrige a Claude
                        logger.warning(
                            f"Total Factura Claude (${total_claude:,.0f}) difiere del OCR "
                            f"(${total_ocr:,.0f}) — usando OCR como anchor"
                        )
                        item["Total Factura"] = total_ocr
                    else:
                        logger.info(f"Total Factura Claude OK: ${total_claude:,.0f} ≈ OCR ${total_ocr:,.0f}")

        # Respaldo OCR: si el texto dice "exenta"/"no afecta" pero Claude no lo detectó
        ocr_lower = ocr_text.lower()
        if any(k in ocr_lower for k in ("exenta", "no afecta")):
            for item in items:
                doc = str(item.get("Documento") or "").lower()
                if "exenta" not in doc and "no afecta" not in doc:
                    # Determinar tipo correcto desde el OCR
                    if "no afecta" in ocr_lower and "exenta" in ocr_lower:
                        item["Documento"] = "Factura No Afecta o Exenta Electronica"
                    elif "no afecta" in ocr_lower:
                        item["Documento"] = "Factura No Afecta Electronica"
                    else:
                        item["Documento"] = "Factura Exenta Electronica"
                    logger.info(f"OCR detectó documento exento — corregido a: {item['Documento']}")

        # Respaldo OCR: detectar Impuesto Específico (IEF + IEV + FEPP → un solo valor)
        # Se fuerza la búsqueda si la glosa menciona combustible, aunque Claude ya puso un valor
        for item in items:
            glosa = (str(item.get("Detalle / Glosa") or "") + " " +
                     str(item.get("Glosa II") or "")).lower()
            es_combustible = any(k in glosa for k in COMBUSTIBLE_KEYWORDS)
            imp_actual = float(item.get("Impuesto Especifico") or 0)

            # Buscar si no hay valor O si es combustible (puede que Claude lo omitió)
            if imp_actual == 0 or (es_combustible and imp_actual == 0):
                total_imp = 0.0
                for pat in IMP_PATTERNS:
                    for m in re.finditer(pat, ocr_lower):
                        val_str = m.group(1).replace(".", "").replace(",", ".")
                        try:
                            v = float(val_str)
                            if v > 0:
                                total_imp += v
                        except ValueError:
                            pass
                if total_imp > 0:
                    item["Impuesto Especifico"] = round(total_imp)
                    tag = " (combustible)" if es_combustible else ""
                    logger.info(f"OCR detectó Impuesto Específico{tag}: ${total_imp:,.0f}")

        # Va DESPUÉS del barrido de OCR de arriba, que es justo quien inventa
        # los impuestos que no existen: si la factura ya cuadra sin él, se saca.
        sanear_impuesto_especifico(items)

        all_items.extend(items)

    # Limpiar temporales PDF
    if ext == ".pdf":
        for p in image_paths:
            try: os.remove(p)
            except: pass

    if not all_items:
        return {"status": "error", "message": "La IA no pudo extraer datos de la factura."}

    all_items = _aplicar_correcciones_aprendidas(all_items)
    all_items = _limpiar_items(all_items)

    # Verificar duplicado
    duplicado = _factura_duplicada(all_items)

    # Registrar en log
    _registrar_factura(all_items, file_path)

    return {"status": "ok", "items": all_items, "duplicado": duplicado}


# Un peso arriba o abajo es redondeo del emisor, no un impuesto.
TOLERANCIA_CUADRE = 3


def sanear_impuesto_especifico(items: list) -> None:
    """Borra el impuesto específico cuando la factura cuadra sin él.

    El impuesto específico es un cargo REAL de combustible (IEF, IEV, FEPP),
    no el lugar donde se esconde lo que no cuadra. Pero `IMP_PATTERNS` barre el
    OCR con patrones laxos y suma TODO lo que pesca, así que se colaban valores
    inventados: en el Master quedó una factura con impuesto de $1 (ruido puro) y
    otra donde `28.084 × 1,19 = 33.420` daba el total EXACTO y aun así tenía
    $11.619 de impuesto encima.

    La prueba es aritmética: si el neto por el IVA ya llega al total, no hay
    impuesto que agregar. Si en cambio el impuesto es justo lo que falta para
    llegar, es de verdad y se respeta.

    Sin `Total Factura` no hay contra qué comparar y no se toca nada: mejor
    dejar un impuesto dudoso que borrar uno bueno adivinando.
    """
    if not items:
        return
    try:
        total = float(items[0].get("Total Factura") or 0)
    except (TypeError, ValueError):
        return
    if total <= 0:
        return

    doc = str(items[0].get("Documento") or "").lower()
    sin_iva = any(k in doc for k in ("exenta", "exento", "no afecta",
                                     "no afecto", "boleta de honorario"))
    iva_factor = 1.0 if sin_iva else 1.19

    neto = 0.0
    imp = 0.0
    for it in items:
        try:
            neto += float(it.get("Valor unitario") or 0) * float(it.get("Cantidad") or 1)
            imp += float(it.get("Impuesto Especifico") or 0)
        except (TypeError, ValueError):
            return
    if imp == 0:
        return

    sin_impuesto = abs(neto * iva_factor - total)
    con_impuesto = abs(neto * iva_factor + imp - total)
    if sin_impuesto <= TOLERANCIA_CUADRE and sin_impuesto < con_impuesto:
        logger.warning(
            "Impuesto Específico descartado: la factura cuadra sin él "
            "(neto %.0f x %.2f = %.0f = total %.0f). Venía $%.0f.",
            neto, iva_factor, neto * iva_factor, total, imp)
        for it in items:
            it["Impuesto Especifico"] = 0
