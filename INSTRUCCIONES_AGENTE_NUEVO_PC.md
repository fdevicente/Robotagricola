# INSTRUCCIONES PARA AGENTE — BOT ERP AGRÍCOLA SANTA ELISA

> **AGENTE:** Lee este archivo completo antes de responder. Al terminar di:
> *"Hola! Leí INSTRUCCIONES_AGENTE_NUEVO_PC.md y estoy contextualizado. ¿En qué continuamos?"*

---

## 1. Contexto del Negocio
- **Agrícola Santa Elisa:** 70 hectáreas en Chile (52 Nogales, 4 Cerezas, 11 Avellanos)
- **Misión:** Bot de Telegram con IA que automatiza ingreso de facturas, control de pagos, reportes y gestión del campo
- **Usuario:** Administrador del campo, opera desde el celular vía Telegram

---

## 2. Estado Actual del Sistema — LEER CON ATENCIÓN

### ✅ YA ESTÁ FUNCIONANDO
- Bot corriendo con `py -3.11 main.py` desde la carpeta `Robot/`
- OCR de facturas con Claude API (Anthropic) — usa visión multimodal
- Guardado automático en `MASTER Agricola Santa Elisa.xlsx`
- Alertas de vencimiento a las 08:00 y sincronización banco a las 18:00
- Edición de campos antes de guardar (proveedor, monto, glosa, etc.)
- Detección de facturas duplicadas
- Módulos: tareas, inventario bodega, vacaciones personal, dashboard web, caja chica

### 🖥️ CONFIGURACIÓN DEL PC (Windows 10, usuario: Windows)
- **Python:** 3.11.9 — ejecutar SIEMPRE con `py -3.11` (NO `python`)
- **Ruta del proyecto:** `C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\Robot\`
- **Excel principal:** `C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx`
- **Credenciales:** guardadas en Windows Credential Manager via `credential_manager.py`
  - Si el bot dice `PROTECTED_BY_CREDENTIAL_MANAGER`, ejecutar: `py -3.11 credential_manager.py`

### 📁 Estructura de la carpeta Robot/
```
Robot/
├── main.py                  ← Bot principal (arranque)
├── config.py                ← Lee .env y Credential Manager
├── credential_manager.py    ← Seguridad con Windows Keyring
├── excel_manager.py         ← CRUD Excel con reintentos si está abierto
├── scotiabank_scraper.py    ← Scraping banco (parcial)
├── chat_inteligente.py      ← Chat IA con Claude
├── inventario_manager.py    ← Bodega / insumos
├── tareas_manager.py        ← Tareas y bitácora
├── vacaciones_manager.py    ← RRHH vacaciones
├── processors/
│   └── extractor.py         ← OCR + Claude API (motor principal)
├── src/
│   └── dashboard.py         ← Dashboard web Flask
├── iniciar_bot.bat          ← Doble clic para arrancar
├── .env                     ← Rutas y config (NO credenciales)
└── requirements.txt
```

---

## 3. Cómo Arrancar el Bot
```cmd
cd "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\Robot"
py -3.11 main.py
```
O doble clic en `iniciar_bot.bat`

---

## 4. Fixes Recientes Aplicados (no revertir)

### excel_manager.py
- `_save_wb()`: reintentos automáticos (5 intentos, 4s espera) si Excel está abierto

### main.py — _build_preview()
- Muestra `Impuesto Específico` (⛽) sumado al total
- Formato unitario con 3 decimales (`:.3f`)
- Facturas exentas no muestran IVA
- Fix edición: usa `reply_text` nuevo en vez de `msg.edit_text` (que fallaba silenciosamente)

### processors/extractor.py — PROMPT de Claude
- **Estaciones de servicio (Copec/Shell/ENEX):** usar razón social del encabezado como proveedor, NO "POR CUENTA DE"
- **IEF + IEV/FEPP:** ambos son impuesto específico, sumarlos en campo `Impuesto Especifico`
- **Cantidad litros:** 17.406 = diecisiete PUNTO cuatro litros (decimal, nunca miles)
- **Valor unitario bencineras:** = SUBTOTAL NETO ÷ Cantidad (no el $/Unit all-in)
- **Martínez y Valdivieso:** precios con 3 decimales reales, conservarlos
- **Respaldo OCR:** detecta IEF, IEV, FEPP por regex si Claude los omite

---

## 5. Problemas Conocidos / Pendientes

| Tema | Estado |
|------|--------|
| Scotiabank scraper | Parcialmente implementado, no testeado |
| Facturas combustible 3 decimales | Fix aplicado, pendiente prueba usuario |
| Botón "Editar Total" no refrescaba | ✅ Corregido |
| Excel abierto bloqueaba guardado | ✅ Corregido con reintentos |

---

## 6. Dependencias Instaladas (py -3.11)
```
python-telegram-bot==21.3  openpyxl  python-dateutil  requests
python-dotenv  Pillow  pytesseract  PyMuPDF  opencv-python
numpy  flask  anthropic  APScheduler  pytz  tzdata  keyring
```
Instalar todo con:
```cmd
py -3.11 -m pip install python-telegram-bot[job-queue]==21.3 openpyxl python-dateutil requests python-dotenv Pillow pytesseract PyMuPDF opencv-python numpy flask anthropic keyring tzdata
```

**Tesseract OCR** (instalador externo):
→ https://github.com/UB-Mannheim/tesseract/wiki — instalar con idioma Spanish

---

## 7. Archivos Excel del Negocio (NO mover)
```
Agricola Santa Elisa/
├── MASTER Agricola Santa Elisa.xlsx  ← Excel principal del bot (Facturas, Boletas, Banco)
├── FXP.xlsx                          ← Control de flujo y pagos
├── Flujo Campo.xlsx
└── CAMARICO 2023/
    ├── PLANILLA GASTOS CAMARICO 2023-2024.xlsx
    └── FUERA/BODEGA ENTRADAS-SALIDAS.xlsx
```

---

## 8. Próxima Acción
Al iniciar nueva sesión, el agente debe:
1. Leer este archivo
2. Preguntar al usuario en qué quiere trabajar
3. Leer los archivos relevantes ANTES de modificar cualquier cosa

---

## 9. Instalación en PC Nuevo — Guía Completa

> **AGENTE:** Si el usuario menciona que está en un PC nuevo o distinto, indica proactivamente esta sección completa y guíalo paso a paso.

### 9.1 Programas a Instalar (en este orden)

| # | Programa | Versión | Dónde descargar |
|---|----------|---------|-----------------|
| 1 | **Python 3.11** | 3.11.x (NO 3.12+) | https://www.python.org/downloads/release/python-3119/ |
| 2 | **Tesseract OCR** | 5.x | https://github.com/UB-Mannheim/tesseract/wiki |
| 3 | **Microsoft Excel** | Cualquier versión moderna | Licencia Office del usuario |
| 4 | **Git** *(opcional)* | Última | https://git-scm.com/download/win |

#### Notas de instalación:

**Python 3.11:**
- Marcar ✅ "Add Python to PATH" durante la instalación
- Verificar: `py -3.11 --version` → debe mostrar `3.11.x`
- Usar SIEMPRE `py -3.11` (nunca `python` a secas)

**Tesseract OCR:**
- Instalar en la ruta por defecto: `C:\Program Files\Tesseract-OCR\`
- Durante la instalación, en "Additional language data" seleccionar **Spanish** y **English**
- Si no se instaló español, descargar `spa.traineddata` desde:
  https://github.com/tesseract-ocr/tessdata/raw/main/spa.traineddata
  y copiarlo a: `C:\Program Files\Tesseract-OCR\tessdata\`
- ⚠️ El bot ya tiene la ruta fija en `processors/extractor.py` — no mover el ejecutable

---

### 9.2 APIs y Credenciales Necesarias

| Credencial | Para qué sirve | Dónde obtenerla | Obligatoria |
|------------|---------------|-----------------|-------------|
| **TELEGRAM_TOKEN** | El bot de Telegram | Hablar con @BotFather en Telegram → `/newbot` | ✅ Sí |
| **ANTHROPIC_API_KEY** | OCR de facturas con Claude (motor principal) | https://console.anthropic.com | ✅ Sí |
| **BANCO_RUT_EMPRESA** | Scraper Scotiabank | Credenciales del banco del cliente | ⬜ Opcional |
| **BANCO_RUT_USUARIO** | Scraper Scotiabank | Credenciales del banco del cliente | ⬜ Opcional |
| **BANCO_CLAVE** | Scraper Scotiabank | Credenciales del banco del cliente | ⬜ Opcional |

---

### 9.3 Configuración del archivo `.env`

Crear el archivo `.env` en la carpeta `Robot/` con este contenido:

```env
TELEGRAM_TOKEN=token_del_bot_aqui
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxx
EXCEL_PATH=C:\Users\TU_USUARIO\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx
DOWNLOAD_DIR=Facturas Recibidas por Telegram
BOLETAS_DIR=Boletas Recibidas por Telegram
TELEGRAM_CHAT_ID=

# Banco (opcional)
BANCO_RUT_EMPRESA=
BANCO_RUT_USUARIO=
BANCO_CLAVE=

# Ollama (alternativa gratuita a Claude, opcional)
OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=llama3.2-vision-fast
```

> ⚠️ Cambiar `TU_USUARIO` por el nombre de usuario real de Windows.
> Luego ejecutar `py -3.11 credential_manager.py` para mover las credenciales al Credential Manager de Windows y dejar el `.env` limpio.

---

### 9.4 Instalar Dependencias Python

```cmd
py -3.11 -m pip install python-telegram-bot[job-queue]==21.3 openpyxl python-dateutil requests python-dotenv Pillow pytesseract PyMuPDF opencv-python numpy flask anthropic keyring tzdata APScheduler pytz
```

Versiones confirmadas funcionando en PC original:
- `python-telegram-bot` 21.3
- `anthropic` 0.96.0
- `Flask` 3.1.3
- `APScheduler` 3.10.4
- `keyring` 25.7.0

---

### 9.5 Verificación Final (antes de arrancar el bot)

Ejecutar estos comandos uno a uno para confirmar que todo está OK:

```cmd
py -3.11 --version
```
→ Debe mostrar `Python 3.11.x`

```cmd
py -3.11 -c "import pytesseract; pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'; print(pytesseract.get_tesseract_version())"
```
→ Debe mostrar `5.x.x`

```cmd
py -3.11 -c "import anthropic; print('anthropic OK')"
```
→ Debe mostrar `anthropic OK`

```cmd
py -3.11 -c "import openpyxl, flask, telegram; print('deps OK')"
```
→ Debe mostrar `deps OK`

Si todo pasa, arrancar con:
```cmd
cd "C:\Users\TU_USUARIO\Desktop\Workflow\Agricola Santa Elisa\Robot"
py -3.11 main.py
```
