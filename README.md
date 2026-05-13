# Bitácora de Desarrollo - Bot Agrícola Santa Elisa

## Objetivo General
Automatizar la recepción y extracción de datos de facturas (imágenes y PDFs) enviadas por Telegram, utilizando IA para leer los datos y consolidarlos automáticamente en un archivo Excel.

---

## Estado actual del proyecto (24 Marzo 2026)

### Estructura de archivos
```
bot_agricola/
├── main.py                  ← Bot Telegram con handlers completos
├── config.py                ← Variables de entorno
├── excel_manager.py         ← Escribe en Facturas recibidas.xlsx
├── .env                     ← Credenciales y rutas
├── requirements.txt         ← Dependencias
├── iniciar_bot.bat          ← Doble clic para arrancar en Windows
└── processors/
    ├── __init__.py
    └── extractor.py         ← Motor de extracción (CamScanner + IA)
```

### Excel: Facturas recibidas.xlsx
- **Hoja:** `Facturas`
- **Columnas (15):** Fecha Emision, Fecha Vencimiento, Fecha Pago, Nombre Factura/Proveedor, Rut, Documento, Numero Factura, Detalle/Glosa, Glosa II, Valor unitario, Cantidad, TOTAL NETO, IVA, Impuesto Especifico, Monto/TOTAL
- **Hoja:** `Proveedores` — columnas: [vacío], Proveedor, RUT (121 proveedores)

---

## Flujo completo del bot

```
Foto/PDF recibida por Telegram
    ↓
Escaneo tipo CamScanner (OpenCV):
  - Detecta 4 esquinas del documento
  - Corrige perspectiva (warp)
  - Umbral adaptativo blanco/negro
    ↓
OCR con Tesseract (texto crudo)
    ↓
IA (Claude API si hay key, Ollama si no):
  - Extrae JSON estructurado
  - Parser robusto que repara JSON truncado/incompleto
    ↓
Post-procesamiento Python:
  - Bloquea RUT/nombre de Agrícola Santa Elisa (receptor, no emisor)
  - Match proveedor por RUT exacto o nombre similar contra hoja Proveedores
  - Limpia número de factura (quita "N°", letras)
  - Normaliza montos (quita puntos de miles, signos $)
    ↓
Detección de duplicados (facturas_log.json)
    ↓
Preview en Telegram con todos los campos + fecha vencimiento calculada
  [✅ Guardar] [✏️ Editar campo] [❌ Cancelar]
    ↓
Si proveedor nuevo → ofrece agregar a hoja Proveedores
    ↓
Guarda en Excel + registra en facturas_log.json
```

---

## Motor de IA

### Claude API (activo si ANTHROPIC_API_KEY está en .env)
- Modelo: `claude-haiku-4-5-20251001`
- Velocidad: ~2-5 segundos por factura
- Costo: ~$0.003 USD por factura
- Precisión: muy alta

### Ollama (fallback si no hay API key)
- Modelo: `llama3.2-vision-fast` (versión con num_ctx=2048 creada manualmente)
- Velocidad: ~40 segundos por factura
- Requiere: `ollama serve` corriendo en segundo plano
- GPU: NVIDIA GeForce (8GB VRAM) — el modelo de 11GB no cabe completo

---

## Reglas de negocio implementadas

1. **Agrícola Santa Elisa** (RUTs: 79857710-7 y variantes) → siempre es el RECEPTOR, nunca el emisor. Si aparece como proveedor, se limpia el RUT y se busca por nombre.
2. **Fecha Vencimiento** → si no viene en la factura, se calcula como Fecha Emisión + 1 mes exacto (tanto en preview como en Excel).
3. **Cálculo matemático** → Python calcula NETO = cantidad × unitario, IVA = neto × 19%, TOTAL = neto + IVA. No depende de la IA.
4. **Multi-ítem** → si la factura tiene múltiples productos, se crea una fila por ítem en el Excel.
5. **Duplicados** → se detectan por RUT + número de factura contra `facturas_log.json`.
6. **Proveedor nuevo** → si el RUT no está en la hoja Proveedores, el bot pregunta si agregarlo antes de guardar.

---

## Comandos Telegram

| Comando | Función |
|---------|---------|
| `/start` | Bienvenida |
| `/ayuda` | Ayuda |
| `/deshacer` | Elimina última factura del Excel y borra el archivo |

### Botones inline
- **✅ Guardar en Excel** → guarda y verifica proveedor nuevo
- **✏️ Editar campo** → menú con 10 campos editables desde Telegram
- **❌ Cancelar** → descarta sin guardar

---

## Variables .env necesarias

```
TELEGRAM_TOKEN=...
OLLAMA_MODEL=llama3.2-vision-fast
OLLAMA_HOST=http://localhost:11434
EXCEL_PATH=C:\ruta\completa\Facturas recibidas.xlsx
DOWNLOAD_DIR=C:\ruta\completa\Facturas Recibidas por Telegram
ANTHROPIC_API_KEY=sk-ant-...  (opcional pero recomendado)
```

---

## Dependencias (requirements.txt)

```
python-telegram-bot==21.3
openpyxl>=3.1
python-dateutil>=2.9
requests>=2.31
python-dotenv>=1.0
Pillow>=10.0
pytesseract>=0.3.10
PyMuPDF>=1.24
opencv-python>=4.8
```

---

## Pendientes / próximos pasos

- [ ] Probar con más facturas variadas para afinar el prompt
- [ ] Considerar agregar comando `/proveedores` para listar/buscar proveedores desde Telegram
- [ ] Agregar comando `/resumen` para ver totales del mes desde Telegram
- [ ] Evaluar si el escaneo CamScanner detecta bien documentos con fondo oscuro o muy iluminados
- [ ] El archivo `facturas_log.json` podría crecer — considerar limpieza periódica o migrar a SQLite