import os
import logging
from .ollama_engine import extract_invoice_data
from .traditional_engine import extract_text_from_pdf, extract_text_from_image, parse_raw_text
from pdf2image import convert_from_path
import tempfile

logger = logging.getLogger(__name__)

def process_file(file_path: str) -> dict:
    """
    Main entry point for processing an invoice file.
    Handles both PDFs and Images.
    Runs Ollama extraction (converting PDF to image if needed).
    Runs traditional extraction as fallback/cross-reference.
    """
    is_pdf = file_path.lower().endswith('.pdf')
    
    ollama_results = []
    traditional_text_blocks = []

    if is_pdf:
        # 1. Traditional text extraction from PDF
        logger.info(f"Extracting raw text from PDF: {file_path}")
        raw_text = extract_text_from_pdf(file_path)
        traditional_text_blocks.append(raw_text)

        # 2. Convert PDF pages to images for Ollama Vision
        logger.info(f"Converting PDF {file_path} to images for Vision...")
        try:
            # Requires poppler installed on Windows and added to PATH, or passed explicitly
            # If poppler is not installed, this will throw an error
            images = convert_from_path(file_path, dpi=200)
            
            for i, img in enumerate(images):
                with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                    img.save(tmp.name, "JPEG")
                    logger.info(f"Processing PDF page {i+1} with Ollama...")
                    
                    page_data = extract_invoice_data(tmp.name, ocr_text=raw_text)
                    
                    if isinstance(page_data, list):
                        ollama_results.extend(page_data)
                    elif isinstance(page_data, dict) and "error" not in page_data:
                        ollama_results.append(page_data)
                    else:
                         logger.warning(f"Error extracting PDF page {i+1}: {page_data}")

                os.remove(tmp.name) # Clean up temp image
                
        except Exception as e:
            logger.error(f"Error converting PDF to images: {e}. Ensure poppler is installed.")
            return {"status": "error", "message": f"PDF to Image conversion failed: {e}", "raw_pdf_text": raw_text}

    else:
        # It's an image
        logger.info(f"Processing image {file_path}...")
        
        # 1. Traditional OCR
        raw_text = extract_text_from_image(file_path)
        traditional_text_blocks.append(raw_text)
        
        # 2. Ollama Vision (With OCR Text Hint)
        data = extract_invoice_data(file_path, ocr_text=raw_text)
        if isinstance(data, list):
            ollama_results.extend(data)
        elif isinstance(data, dict) and "error" not in data:
            ollama_results.append(data)
        else:
             logger.warning(f"Error extracting image: {data}")
    
    # Cross-reference logic (Basic approach for now: check if Ollama got results)
    # A more advanced logic would verify if 'Total' in Ollama JSON exists in traditional OCR raw text
    
    if not ollama_results:
         return {"status": "failed", "message": "Ollama could not extract structured data.", "raw_text": traditional_text_blocks}

    # Ensure each dictionary has EXACTLY the 15 keys expected, filling missing ones with None
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    
    expected_keys = [
        "Fecha Emision / Fecha", "Fecha Vencimiento", "Fecha Pago", 
        "Nombre Factura / Proveedor", "Rut", "Documento", 
        "Numero Factura / Nro Documento", "Detalle / Glosa (solamente el nombre del producto) / NOTA", 
        "Glosa II ( Detalle completo del producto)", "Valor unitario producto sin iva  /  $ UNIT", 
        "Cantidad", "TOTAL NETO", "IVA", "ESPECIFICO ( Impuesto especifico) ", "Monto / TOTAL"
    ]
    
    cleaned_items = []
    total_calculated = 0
    
    for item in ollama_results:
        clean_item = {}
        for key in expected_keys:
            clean_item[key] = item.get(key, None)
            
        # 1. Date Logic
        fecha_emision = clean_item.get("Fecha Emision / Fecha")
        fecha_venc = clean_item.get("Fecha Vencimiento")
        
        if fecha_emision and (not fecha_venc or str(fecha_venc).strip() == "" or str(fecha_venc).lower() == "null"):
            try:
                # Assuming YYYY-MM-DD or DD-MM-YYYY
                formats = ["%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"]
                parsed_date = None
                for fmt in formats:
                    try:
                        parsed_date = datetime.strptime(str(fecha_emision).strip(), fmt)
                        break
                    except ValueError:
                        continue
                if parsed_date:
                    next_month = parsed_date + relativedelta(months=1)
                    clean_item["Fecha Vencimiento"] = next_month.strftime("%Y-%m-%d")
            except Exception as e:
                logger.warning(f"Could not calculate due date from {fecha_emision}: {e}")
                
        # 2. Total validation & Math calculation
        # Since LLMs are bad at math, we recalculate Neto, IVA and Total explicitly in Python
        try:
            # Clean and get base values
            def clean_number(val):
                if not val: return 0.0
                return float(str(val).replace(",","").replace("$","").replace(".","").strip())

            valor_unitario = clean_number(clean_item.get("Valor unitario producto sin iva  /  $ UNIT"))
            cantidad = clean_number(clean_item.get("Cantidad"))
            if cantidad == 0: cantidad = 1.0 # Default fallback
            
            # Recompute
            neto_calculado = valor_unitario * cantidad
            iva_calculado = round(neto_calculado * 0.19)
            
            # Use original total if explicit, else calculate (assuming 19% IVA for standard "Factura Electronica")
            is_exempt = "exenta" in str(clean_item.get("Documento", "")).lower()
            if is_exempt:
                iva_calculado = 0
                
            especifico = clean_number(clean_item.get("ESPECIFICO ( Impuesto especifico) "))
            monto_total_item = neto_calculado + iva_calculado + especifico
            
            # Overwrite the json fields with the correct Python math
            clean_item["Valor unitario producto sin iva  /  $ UNIT"] = valor_unitario
            clean_item["Cantidad"] = cantidad
            clean_item["TOTAL NETO"] = neto_calculado
            clean_item["IVA"] = iva_calculado
            clean_item["ESPECIFICO ( Impuesto especifico) "] = especifico
            clean_item["Monto / TOTAL"] = monto_total_item
            
            total_calculated += monto_total_item
            
        except Exception as e:
            logger.warning(f"Error calculating math: {e}")

        cleaned_items.append(clean_item)

    # Note: We can add logic to compare total_calculated with a global extracted total here if we want to reject bad parses.

    return {
        "status": "success",
        "items": cleaned_items,
        "raw_text_cross_reference": "\n---\n".join(traditional_text_blocks)
    }
