import json
import base64
import ollama
import logging
from config import OLLAMA_MODEL, OLLAMA_HOST

logger = logging.getLogger(__name__)

# Configure the ollama client to use the defined host
client = ollama.Client(host=OLLAMA_HOST)

def extract_invoice_data(image_path: str, ocr_text: str = "") -> dict:
    """
    Takes an image path and optional OCR text, sends it to local Ollama (Llama 3.2 Vision), 
    and returns parsed JSON containing the invoice data.
    """
    logger.info(f"Extracting data from {image_path} using Ollama ({OLLAMA_MODEL})...")
    
    prompt = """
Eres un experto auditor de Facturas Electrónicas Chilenas (SII).
Tu objetivo es leer ATENTAMENTE esta imagen/PDF y convertir su contenido en N objetos JSON, uno por cada línea de detalle (producto/servicio) facturado.

REGLAS ESTRICTAS QUE DEBES CUMPLIR O FALLARÁS:
1. El "Numero Factura / Nro Documento" suele estar en la esquina superior derecha, en rojo, debajo del RUT (Folio). NUNCA inventes "001" a menos que ese sea realmente el folio oficial.
2. Identifica TODAS las filas en la tabla de detalles/productos. Si ves "total kilometraje por 42.000" y "desmontar y montar... por 100.000", son DOS filas.
3. Para cada fila de producto, devuelve un objeto JSON independiente. En cada uno repite la info del proveedor, rut, folio, y las fechas.
4. "Valor unitario producto sin iva  /  $ UNIT" debe ser el valor neto individual sin puntos ni signos peso.
5. "TOTAL NETO" debe ser la suma de los valores netos (generalmente al fondo de la factura). 
6. "Monto / TOTAL" a nivel de ítem: Toma el valor unitario de la línea y súmale el 19% de IVA (Valor Unitario * 1.19), a menos de que la factura indique que son valores exentos.
8. IMPORTANTE: EL SIGUIENTE ES SOLO UN EJEMPLO DE FORMATO. ¡NO COPIES LA INFORMACIÓN DE ESTE EJEMPLO! DEBES LEER LA IMAGEN REAL Y EXTRAER SUS DATOS.

EJEMPLO DE FORMATO ESPERADO (CON DATOS FALSOS INVENTADOS):
[
  {
    "Fecha Emision / Fecha": "2050-01-01",
    "Fecha Vencimiento": "2050-02-01",
    "Fecha Pago": null,
    "Nombre Factura / Proveedor": "EMPRESA FALSA DE EJEMPLO SPA",
    "Rut": "99.999.999-9",
    "Documento": "Factura Electronica",
    "Numero Factura / Nro Documento": "123456789",
    "Detalle / Glosa (solamente el nombre del producto) / NOTA": "Producto Falso 1",
    "Glosa II ( Detalle completo del producto)": "Descripción inventada 1",
    "Valor unitario producto sin iva  /  $ UNIT": 1000,
    "Cantidad": 2,
    "TOTAL NETO": 3000,
    "IVA": 570,
    "ESPECIFICO ( Impuesto especifico) ": 0,
    "Monto / TOTAL": 1190
  },
  {
    "Fecha Emision / Fecha": "2050-01-01",
    "Nombre Factura / Proveedor": "EMPRESA FALSA DE EJEMPLO SPA",
    "Numero Factura / Nro Documento": "123456789",
    "Detalle / Glosa (solamente el nombre del producto) / NOTA": "Producto Falso 2",
    "Valor unitario producto sin iva  /  $ UNIT": 1000,
    "TOTAL NETO": 3000,
    "Monto / TOTAL": 1190
  }
]
"""
    if ocr_text and len(ocr_text.strip()) > 20:
        prompt += f"\n\n--- IMPORTANTE ---\nAdicionalmente, he escaneado la imagen con un software OCR tradicional. Aquí tienes los textos crudos que encontró. ÚSALOS ESTRICTAMENTE como ayuda o 'pista' para encontrar los números exactos, nombres de ítems y valores reales de la factura, pero guíate siempre por tu visión para entender qué número corresponde a qué campo. TEXTO OCR:\n\n{ocr_text}"

    try:
        with open(image_path, "rb") as image_file:
            image_bytes = image_file.read()
            # Ojo: la API de Ollama espera la imagen en bytes directamente
            
        response = client.chat(
            model=OLLAMA_MODEL,
            messages=[{
                'role': 'user',
                'content': prompt,
                'images': [image_bytes]
            }],
            options={
                "temperature": 0.0 # Lowest temperature for consistent JSON extraction
            }
        )
        
        raw_content = response['message']['content']
        
        # Simple cleanup in case the LLM wrapped it in markdown code blocks
        if "```json" in raw_content:
            raw_content = raw_content.split("```json")[1].split("```")[0].strip()
        elif "```" in raw_content:
             raw_content = raw_content.split("```")[1].split("```")[0].strip()
             
        parsed_json = json.loads(raw_content)
        return parsed_json

    except Exception as e:
        logger.error(f"Error extracting data with Ollama: {e}")
        return {"error": str(e), "raw": raw_content if 'raw_content' in locals() else None}
