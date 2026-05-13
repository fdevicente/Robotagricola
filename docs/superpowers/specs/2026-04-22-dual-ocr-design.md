# Diseño: OCR Dual Paralelo (Tesseract CPU + Surya GPU)

**Fecha:** 2026-04-22  
**Archivo objetivo:** `Robot/processors/extractor.py`

---

## Contexto

El bot actualmente usa solo Tesseract (CPU) como OCR. El texto extraído cumple tres roles:
1. Contexto para Claude en el prompt
2. Chequeo cruzado del total (`_extraer_total_ocr`)
3. Detección de keywords (exenta, impuesto específico)

Se agrega Surya OCR (GPU/CUDA) corriendo en paralelo para enriquecer el contexto y mejorar la precisión, especialmente en facturas complejas (M&V, combustibles).

---

## Arquitectura

```
imagen escaneada (scan_path)
  ├── Thread 1 (CPU) → pytesseract  → texto_tess
  └── Thread 2 (GPU) → surya-ocr   → texto_surya
              ↓ (concurrent.futures, max_workers=2)
        _combinar_ocr(texto_tess, texto_surya)
              ↓
        "=== OCR Tesseract ===\n...\n=== OCR Surya ===\n..."
              ↓
        Claude (prompt con ocr_text enriquecido)
        _extraer_total_ocr (opera sobre texto combinado, sin cambios)
```

---

## Cambios en extractor.py

### Funciones nuevas

**`_ocr_surya(image_path: str) -> str`**
- Importa `surya.ocr` dentro de la función (import lazy, no falla si no está instalado)
- Carga modelo una sola vez usando variable de módulo `_SURYA_MODEL = None`
- Retorna texto concatenado de todas las líneas detectadas
- En caso de error: warning en log, retorna `""`

**`_ocr_dual(image_path: str) -> tuple[str, str]`**
- Lanza `_ocr_text()` y `_ocr_surya()` en paralelo con `ThreadPoolExecutor(max_workers=2)`
- Retorna `(texto_tess, texto_surya)`
- Timeout 30s por motor; si uno falla retorna `""` para ese motor

**`_combinar_ocr(tess: str, surya: str) -> str`**
- Si ambos tienen texto: retorna texto etiquetado con separadores
- Si solo uno tiene texto: retorna ese texto sin etiqueta (compatibilidad con lógica actual)
- Si ambos vacíos: retorna `""`

### Funciones modificadas

**`process_file()`**
- Reemplaza `ocr_text = _ocr_text(scan_path)` por:
  ```python
  tess_text, surya_text = _ocr_dual(scan_path)
  ocr_text = _combinar_ocr(tess_text, surya_text)
  ```
- `_extraer_total_ocr` recibe el texto combinado (sin cambio de firma)
- El texto combinado contiene ambas lecturas; `_extraer_total_ocr` opera igual sobre él (regex sobre todas las líneas)

**`PROMPT`**
- Cambia el label de `TEXTO OCR` a `TEXTO OCR (dos motores: Tesseract CPU + Surya GPU)`
- Sin cambio en el placeholder `{ocr_text}`

---

## Manejo de errores

| Escenario | Comportamiento |
|---|---|
| `surya-ocr` no instalado | ImportError silencioso, solo Tesseract |
| Surya falla en runtime | Warning log, solo Tesseract |
| Tesseract falla | Warning log, solo Surya |
| Ambos fallan | `""` — mismo comportamiento actual |
| Solo un motor produce texto | Se usa sin etiqueta (sin ruido extra) |

---

## Carga del modelo Surya

Surya carga modelos pesados (~500 MB) en la primera llamada. Para evitar recargar en cada factura, se cachea en una variable de módulo:

```python
_SURYA_OCR_PREDICTOR = None

def _get_surya_predictor():
    global _SURYA_OCR_PREDICTOR
    if _SURYA_OCR_PREDICTOR is None:
        from surya.recognition import RecognitionPredictor
        _SURYA_OCR_PREDICTOR = RecognitionPredictor()
    return _SURYA_OCR_PREDICTOR
```

CUDA se detecta automáticamente por Surya/PyTorch.

---

## Lo que NO cambia

- `_ocr_text()` — sin modificaciones
- `_extraer_total_ocr()` — misma lógica, recibe texto combinado
- `_call_ia()`, `_call_claude()`, `_limpiar_items()` — sin cambios
- Toda la lógica de M&V, COPEVAL, honorarios — sin cambios
- El prompt de Claude — solo cambia el label del bloque OCR

---

## Instalación

```bash
py -3.11 -m pip install surya-ocr
```

Surya requiere PyTorch con CUDA. Si PyTorch ya está instalado sin CUDA, Surya corre en CPU (más lento pero funcional).
