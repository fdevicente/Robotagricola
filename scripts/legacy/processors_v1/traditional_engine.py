import logging
import pytesseract
from PIL import Image
import pdf2image
import os
import fitz # PyMuPDF
import json
import re

# Set the Tesseract data prefix explicitly so it finds the spa.traineddata
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TESSDATA_DIR = os.path.join(BASE_DIR, "tessdata")
os.environ["TESSDATA_PREFIX"] = TESSDATA_DIR

logger = logging.getLogger(__name__)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract embedded text from a PDF."""
    text = ""
    try:
        with fitz.open(pdf_path) as doc:
            for page in doc:
                text += page.get_text()
    except Exception as e:
        logger.error(f"Error reading PDF {pdf_path}: {e}")
    return text

def extract_text_from_image(image_path: str) -> str:
    """Extract text from an image using Tesseract OCR."""
    try:
        image = Image.open(image_path)
        
        # Point to the local tessdata folder we just created
        custom_config = f'--tessdata-dir "{TESSDATA_DIR}"'
        
        text = pytesseract.image_to_string(image, lang='spa', config=custom_config) 
        return text
    except Exception as e:
        logger.error(f"Error reading image {image_path}: {e}")
        return ""

def parse_raw_text(text: str) -> dict:
    """
    Takes raw OCR/PDF text and structures it into the target format.
    Usually we'd use a lightweight LLM call here, but for fallback/cross-check
    we can use regex or basic parsing if it's highly standardized, 
    or a simple LLM prompt. To keep it fully local and simple, 
    we will rely heavily on Ollama for primary.
    """
    # Placeholder for simple regex extraction or calling Ollama with just text.
    # For a robust approach, we can ask Ollama to structure the RAW text as well.
    # We will return the raw text for now so `extractor.py` can cross-reference it.
    
    return {"raw_text": text}
