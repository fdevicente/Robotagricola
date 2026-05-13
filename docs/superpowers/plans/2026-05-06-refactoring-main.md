# Refactoring main.py — Modularización del Bot Agrícola

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reorganizar main.py (1680 líneas) en módulos por dominio, extraer utilidades duplicadas, y eliminar prompts duplicados en extractor.py — sin cambiar funcionalidad.

**Architecture:** Se crean dos paquetes nuevos: `utils/` (funciones de parseo y formato compartidas) y `handlers/` (handlers de Telegram separados por dominio). main.py queda como punto de entrada (~80 líneas). extractor.py elimina los prompts SYSTEM_PROMPT duplicados y usa solo PROMPT + PROMPT_SIMPLE.

**Tech Stack:** Python 3.14, python-telegram-bot, openpyxl, anthropic API (requests directo)

---

## File Structure

### Nuevos archivos:
- `utils/__init__.py` — vacío
- `utils/parsing.py` — `parsear_fecha()`, `parsear_monto()`
- `utils/formatting.py` — `esc()`, `format_date()`, `calc_vencimiento()`
- `utils/keyboards.py` — todos los InlineKeyboardMarkup builders
- `handlers/__init__.py` — `register_all(app)` que registra todos los handlers
- `handlers/facturas.py` — handle_photo, handle_document, preview, edición, guardar, deshacer
- `handlers/finanzas.py` — pagado, deposito, saldo, banco, reporte, jobs banco
- `handlers/tareas.py` — tarea, tareas, hecho, bitacora
- `handlers/inventario_h.py` — inventario, uso (sufijo _h para no colisionar con inventario_manager)
- `handlers/personal.py` — personal, vacaciones, agregar_trabajador
- `handlers/chat.py` — chat inteligente fallback

### Archivos modificados:
- `main.py` — reducido a arranque + registro de handlers
- `processors/extractor.py` — eliminar SYSTEM_PROMPT y SYSTEM_PROMPT_SIMPLE duplicados

### Archivos sin cambios:
- `config.py`, `excel_manager.py`, `tareas_manager.py`, `inventario_manager.py`, `vacaciones_manager.py`, `chat_inteligente.py`, `credential_manager.py`, `scotiabank_scraper.py`

---

## Task 1: Crear utils/parsing.py — Funciones de parseo compartidas

**Files:**
- Create: `utils/__init__.py`
- Create: `utils/parsing.py`

- [ ] **Step 1: Crear utils/__init__.py vacío**

```python
# utils/__init__.py
```

- [ ] **Step 2: Crear utils/parsing.py con parsear_fecha y parsear_monto**

```python
"""
utils/parsing.py — Funciones de parseo reutilizadas en múltiples handlers.
- parsear_fecha: texto libre (DD/MM/YYYY, YYYY-MM-DD, "hoy") → "YYYY-MM-DD"
- parsear_monto: texto con formato chileno ($, puntos miles, coma decimal) → float
"""
from datetime import datetime


def parsear_fecha(texto: str) -> str | None:
    """Convierte texto a fecha YYYY-MM-DD. Retorna None si no es válido.

    Formatos aceptados:
      - "hoy"           → fecha actual
      - "15/03/2026"    → DD/MM/YYYY
      - "2026-03-15"    → ISO
      - Cualquier otro  → intenta dateutil.parser
    """
    texto = texto.strip()
    if texto.lower() == "hoy":
        return datetime.now().strftime("%Y-%m-%d")
    try:
        if "/" in texto:
            parts = texto.split("/")
            if len(parts) == 3 and len(parts[0]) <= 2:
                return datetime.strptime(texto, "%d/%m/%Y").strftime("%Y-%m-%d")
        from dateutil import parser as date_parser
        return date_parser.parse(texto).strftime("%Y-%m-%d")
    except Exception:
        return None


def parsear_monto(texto: str) -> float | None:
    """Convierte texto con formato chileno a float. Retorna None si no es válido.

    Formatos aceptados:
      - "$1.234.567"  → 1234567.0  (puntos = miles)
      - "7.164,32"    → 7164.32    (coma = decimal)
      - "500000"      → 500000.0
    """
    try:
        n = texto.replace("$", "").replace(" ", "").strip()
        if "," in n:
            # Formato chileno: puntos = miles, coma = decimal
            n = n.replace(".", "").replace(",", ".")
        elif n.count(".") > 1:
            # Múltiples puntos = separadores de miles
            n = n.replace(".", "")
        return float(n)
    except (ValueError, AttributeError):
        return None
```

- [ ] **Step 3: Verificar que importa correctamente**

Run: `python -c "from utils.parsing import parsear_fecha, parsear_monto; print(parsear_fecha('15/03/2026'), parsear_monto('1.234.567'))"`
Expected: `2026-03-15 1234567.0`

- [ ] **Step 4: Commit**

```
git add utils/
git commit -m "refactor: crear utils/parsing.py con parsear_fecha y parsear_monto"
```

---

## Task 2: Crear utils/formatting.py — Funciones de formato Telegram

**Files:**
- Create: `utils/formatting.py`

- [ ] **Step 1: Crear utils/formatting.py**

Extraer de main.py las funciones `_esc` (línea 130), `_format_date` (línea 137), `_calc_vencimiento` (línea 144):

```python
"""
utils/formatting.py — Helpers de formato para mensajes de Telegram.
- esc: escapa caracteres Markdown
- format_date: fecha legible
- calc_vencimiento: emisión + 30 días
"""
from datetime import datetime


def esc(text) -> str:
    """Escapa caracteres especiales de Telegram Markdown."""
    if not text:
        return text
    for ch in ('*', '_', '`', '[', ']'):
        text = str(text).replace(ch, '')
    return text


def format_date(val) -> str:
    """Formatea fecha para mostrar en Telegram."""
    if val is None:
        return "—"
    try:
        if isinstance(val, str):
            return val[:10]
        return val.strftime("%d/%m/%Y")
    except Exception:
        return str(val)


def calc_vencimiento(fecha_emision) -> str | None:
    """Calcula fecha de vencimiento = emisión + 1 mes."""
    try:
        from dateutil.relativedelta import relativedelta
        if not fecha_emision:
            return None
        dt = datetime.strptime(str(fecha_emision)[:10], "%Y-%m-%d")
        return (dt + relativedelta(months=1)).strftime("%Y-%m-%d")
    except Exception:
        return None
```

- [ ] **Step 2: Verificar imports**

Run: `python -c "from utils.formatting import esc, format_date, calc_vencimiento; print(esc('*hola*'), format_date('2026-01-15'))"`
Expected: `hola 2026-01-15`

- [ ] **Step 3: Commit**

```
git add utils/formatting.py
git commit -m "refactor: crear utils/formatting.py con esc, format_date, calc_vencimiento"
```

---

## Task 3: Crear utils/keyboards.py — Teclados de Telegram

**Files:**
- Create: `utils/keyboards.py`

- [ ] **Step 1: Crear utils/keyboards.py**

Extraer de main.py: `_main_keyboard` (línea 294), `_edit_keyboard` (línea 304), `_proveedor_nuevo_keyboard` (línea 329).

```python
"""
utils/keyboards.py — Constructores de InlineKeyboardMarkup para Telegram.
"""
from telegram import InlineKeyboardButton, InlineKeyboardMarkup


def main_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Guardar en Excel", callback_data="confirm_save"),
         InlineKeyboardButton("✏️ Editar campo",    callback_data="edit_menu")],
        [InlineKeyboardButton("❌ Cancelar",         callback_data="cancel_save")],
    ])


def edit_keyboard(items=None):
    n_items = len(items) if items else 1
    kb = [
        [InlineKeyboardButton("🏢 Proveedor",  callback_data="edit_proveedor"),
         InlineKeyboardButton("🪪 RUT",         callback_data="edit_rut")],
        [InlineKeyboardButton("📅 Fecha",       callback_data="edit_fecha"),
         InlineKeyboardButton("⏰ Vencimiento", callback_data="edit_vence")],
        [InlineKeyboardButton("📄 Nº Documento", callback_data="edit_nro"),
         InlineKeyboardButton("🔗 Ref. Factura", callback_data="edit_ref")],
        [InlineKeyboardButton("📦 Glosa",       callback_data="edit_glosa")],
        [InlineKeyboardButton("📝 Detalle",     callback_data="edit_glosa2"),
         InlineKeyboardButton("🔢 Cantidad",    callback_data="edit_cantidad")],
        [InlineKeyboardButton("💲 Unit neto",   callback_data="edit_unitario"),
         InlineKeyboardButton("💰 Total ítem",  callback_data="edit_total")],
        [InlineKeyboardButton("💰 TOTAL FACTURA", callback_data="edit_total_factura")],
    ]
    item_row = [InlineKeyboardButton("➕ Agregar ítem", callback_data="add_item")]
    if n_items > 1:
        item_row.append(InlineKeyboardButton("🗑️ Eliminar ítem", callback_data="del_item"))
    kb.append(item_row)
    kb.append([InlineKeyboardButton("« Volver", callback_data="back_preview")])
    return InlineKeyboardMarkup(kb)


def proveedor_nuevo_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Sí, agregar a la lista", callback_data="add_proveedor_yes"),
         InlineKeyboardButton("❌ No",                    callback_data="add_proveedor_no")],
    ])


def back_button(callback_data="edit_menu"):
    """Botón solitario '« Volver'."""
    return InlineKeyboardMarkup([[InlineKeyboardButton("« Volver", callback_data=callback_data)]])
```

- [ ] **Step 2: Verificar imports**

Run: `python -c "from utils.keyboards import main_keyboard, edit_keyboard; print(type(main_keyboard()))"`
Expected: `<class 'telegram.InlineKeyboardMarkup'>`

- [ ] **Step 3: Commit**

```
git add utils/keyboards.py
git commit -m "refactor: crear utils/keyboards.py con todos los teclados inline"
```

---

## Task 4: Crear handlers/__init__.py y handlers/facturas.py

**Files:**
- Create: `handlers/__init__.py`
- Create: `handlers/facturas.py`

- [ ] **Step 1: Crear handlers/__init__.py con register_all**

```python
"""
handlers/__init__.py — Registra todos los handlers del bot en la app de Telegram.
"""
from telegram.ext import CommandHandler, MessageHandler, CallbackQueryHandler, filters


def register_all(app):
    """Registra comandos, mensajes y callbacks de TODOS los módulos."""
    from handlers import facturas, finanzas, tareas, inventario_h, personal, chat

    # ── Comandos ──
    app.add_handler(CommandHandler("start",    facturas.cmd_start))
    app.add_handler(CommandHandler("ayuda",    facturas.cmd_start))
    app.add_handler(CommandHandler("deshacer", facturas.cmd_deshacer))
    app.add_handler(CommandHandler("cancelar", facturas.cmd_cancelar))

    app.add_handler(CommandHandler("pagado",   finanzas.cmd_pagado))
    app.add_handler(CommandHandler("deposito", finanzas.cmd_deposito))
    app.add_handler(CommandHandler("saldo",    finanzas.cmd_saldo))
    app.add_handler(CommandHandler("reporte",  finanzas.cmd_reporte))
    app.add_handler(CommandHandler("dashboard", finanzas.cmd_dashboard))
    app.add_handler(CommandHandler("banco",    finanzas.cmd_banco))

    app.add_handler(CommandHandler("tarea",    tareas.cmd_tarea))
    app.add_handler(CommandHandler("tareas",   tareas.cmd_tareas))
    app.add_handler(CommandHandler("hecho",    tareas.cmd_hecho))
    app.add_handler(CommandHandler("bitacora", tareas.cmd_bitacora))

    app.add_handler(CommandHandler("inventario", inventario_h.cmd_inventario))
    app.add_handler(CommandHandler("uso",        inventario_h.cmd_uso))

    app.add_handler(CommandHandler("personal",           personal.cmd_personal))
    app.add_handler(CommandHandler("vacaciones",         personal.cmd_vacaciones))
    app.add_handler(CommandHandler("agregar_trabajador", personal.cmd_agregar_trabajador))

    # ── Mensajes ──
    app.add_handler(MessageHandler(filters.Document.ALL,            facturas.handle_document))
    app.add_handler(MessageHandler(filters.PHOTO,                   facturas.handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, chat.handle_text))

    # ── Callbacks ──
    # Facturas
    app.add_handler(CallbackQueryHandler(facturas.cb_confirm_save,      pattern="^confirm_save$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_cancel_save,       pattern="^cancel_save$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_edit_menu,         pattern="^edit_menu$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_back_preview,      pattern="^back_preview$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_add_proveedor_yes, pattern="^add_proveedor_yes$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_add_proveedor_no,  pattern="^add_proveedor_no$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_add_item,          pattern="^add_item$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_del_item_menu,     pattern="^del_item$"))
    app.add_handler(CallbackQueryHandler(facturas.cb_del_item_confirm,  pattern="^delitem_"))
    app.add_handler(CallbackQueryHandler(facturas.cb_edit_total_factura, pattern="^edit_total_factura$"))
    # Finanzas
    app.add_handler(CallbackQueryHandler(finanzas.cb_medio_pago,        pattern="^pago_"))
    app.add_handler(CallbackQueryHandler(finanzas.cb_calce_verificacion, pattern="^calce_"))
    app.add_handler(CallbackQueryHandler(finanzas.cb_reporte,           pattern="^rep_"))
    # Tareas
    app.add_handler(CallbackQueryHandler(tareas.cb_tarea_prioridad,     pattern="^prior_"))
    # Inventario
    app.add_handler(CallbackQueryHandler(inventario_h.cb_cultivo,       pattern="^cult_"))
    # Edición (debe ir AL FINAL porque "^edit_" matchea muchos patrones)
    app.add_handler(CallbackQueryHandler(facturas.cb_select_item,       pattern="^selitem_"))
    app.add_handler(CallbackQueryHandler(facturas.cb_edit_field,        pattern="^edit_"))
```

- [ ] **Step 2: Crear handlers/facturas.py**

Mover desde main.py: constantes CAMPOS_EDITABLES/CAMPOS_COMUNES/CAMPOS_POR_ITEM, funciones _save_path, _save_path_boleta, _renombrar_archivo, _es_boleta, _registrar_correccion, _build_preview, _rut_existe, _agregar_proveedor, cmd_start, cmd_deshacer, cmd_cancelar, handle_document, handle_photo, _process_and_reply, _show_preview, _download_with_retry, y todos los callbacks de factura (cb_confirm_save, cb_cancel_save, cb_edit_menu, cb_edit_field, cb_select_item, cb_edit_total_factura, cb_back_preview, cb_add_item, cb_del_item_menu, cb_del_item_confirm, cb_add_proveedor_yes, cb_add_proveedor_no, _guardar_excel, handle_text_edit_factura).

El archivo completo se escribe en el Step de implementación (Task 8). Aquí se crea el esqueleto vacío para validar imports.

```python
"""
handlers/facturas.py — Procesamiento de facturas, preview, edición y guardado en Excel.
"""
# El contenido completo se escribe en Task 8 (migración desde main.py)
```

- [ ] **Step 3: Commit**

```
git add handlers/
git commit -m "refactor: crear handlers/__init__.py con register_all y esqueleto facturas.py"
```

---

## Task 5: Crear handlers/finanzas.py

**Files:**
- Create: `handlers/finanzas.py`

- [ ] **Step 1: Crear handlers/finanzas.py**

Mover desde main.py: cmd_pagado, cmd_deposito, cmd_saldo, cmd_dashboard, cmd_reporte, cmd_banco, cb_reporte, cb_medio_pago, cb_calce_verificacion, _sync_banco_core, job_sync_banco, y los flujos de texto de pagado/deposito (se extraen como `handle_text_pagado` y `handle_text_deposito`).

```python
"""
handlers/finanzas.py — Comandos financieros: pagado, depósito, saldo, banco, reportes.
"""
# El contenido completo se escribe en Task 8 (migración desde main.py)
```

- [ ] **Step 2: Commit**

```
git add handlers/finanzas.py
git commit -m "refactor: esqueleto handlers/finanzas.py"
```

---

## Task 6: Crear handlers restantes (tareas, inventario_h, personal, chat)

**Files:**
- Create: `handlers/tareas.py`
- Create: `handlers/inventario_h.py`
- Create: `handlers/personal.py`
- Create: `handlers/chat.py`

- [ ] **Step 1: Crear handlers/tareas.py**

```python
"""
handlers/tareas.py — Comandos de tareas y bitácora.
"""
# El contenido completo se escribe en Task 8
```

- [ ] **Step 2: Crear handlers/inventario_h.py**

```python
"""
handlers/inventario_h.py — Comandos de inventario y registro de uso.
(Sufijo _h para no colisionar con inventario_manager.py)
"""
# El contenido completo se escribe en Task 8
```

- [ ] **Step 3: Crear handlers/personal.py**

```python
"""
handlers/personal.py — Comandos de personal y vacaciones.
"""
# El contenido completo se escribe en Task 8
```

- [ ] **Step 4: Crear handlers/chat.py**

```python
"""
handlers/chat.py — Dispatcher de texto: detecta flujo activo o delega a chat inteligente.
Este es el handler central de texto que revisa si hay un flujo en curso
(edición factura, pagado, depósito, tarea, etc.) y delega al handler correcto.
"""
# El contenido completo se escribe en Task 8
```

- [ ] **Step 5: Commit**

```
git add handlers/
git commit -m "refactor: esqueletos handlers tareas, inventario, personal, chat"
```

---

## Task 7: Limpiar prompts duplicados en extractor.py

**Files:**
- Modify: `processors/extractor.py` — eliminar SYSTEM_PROMPT (líneas 636-781) y SYSTEM_PROMPT_SIMPLE (líneas 783-796)

- [ ] **Step 1: Eliminar SYSTEM_PROMPT y SYSTEM_PROMPT_SIMPLE**

Estas dos constantes son copias casi idénticas de PROMPT y PROMPT_SIMPLE respectivamente. No se usan en ningún lugar del código (grep confirma que solo están definidas, nunca importadas ni referenciadas fuera del archivo).

Eliminar las líneas 636-796 de extractor.py (desde `SYSTEM_PROMPT = """Eres un lector...` hasta el cierre de `SYSTEM_PROMPT_SIMPLE`).

- [ ] **Step 2: Verificar que extractor sigue funcionando**

Run: `python -c "from processors.extractor import process_file; print('OK')"`
Expected: `OK`

- [ ] **Step 3: Commit**

```
git add processors/extractor.py
git commit -m "refactor: eliminar SYSTEM_PROMPT duplicados en extractor.py"
```

---

## Task 8: Migración completa — Llenar handlers y reescribir main.py

**Esta es la tarea principal.** Se hace en un solo paso atómico para que el bot nunca quede en estado inconsistente.

**Files:**
- Modify: `handlers/facturas.py` — contenido completo
- Modify: `handlers/finanzas.py` — contenido completo
- Modify: `handlers/tareas.py` — contenido completo
- Modify: `handlers/inventario_h.py` — contenido completo
- Modify: `handlers/personal.py` — contenido completo
- Modify: `handlers/chat.py` — contenido completo
- Modify: `main.py` — reducir a ~80 líneas

- [ ] **Step 1: Escribir handlers/facturas.py completo**

Contiene:
- Constantes: CAMPOS_EDITABLES, CAMPOS_COMUNES, CAMPOS_POR_ITEM
- Funciones privadas: _save_path, _save_path_boleta, _renombrar_archivo, _es_boleta, _registrar_correccion, _build_preview, _rut_existe, _agregar_proveedor, _download_with_retry, _process_and_reply, _show_preview, _guardar_excel
- Comandos: cmd_start, cmd_deshacer, cmd_cancelar
- Callbacks: cb_confirm_save, cb_cancel_save, cb_edit_menu, cb_edit_field, cb_select_item, cb_edit_total_factura, cb_back_preview, cb_add_item, cb_del_item_menu, cb_del_item_confirm, cb_add_proveedor_yes, cb_add_proveedor_no
- Flujo texto: handle_text_edit_factura(update, context) — solo la parte de edición de campos pendientes

Imports necesarios:
```python
import os, re, logging, asyncio
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from config import DOWNLOAD_DIR, BOLETAS_DIR, EXCEL_PATH, TELEGRAM_CHAT_ID
from processors.extractor import process_file
from excel_manager import (append_to_excel, delete_last_rows, append_boleta,
                          delete_last_boletas, consultar_saldo_caja)
from inventario_manager import agregar_stock_desde_factura
from utils.formatting import esc, format_date, calc_vencimiento
from utils.parsing import parsear_monto
from utils.keyboards import main_keyboard, edit_keyboard, proveedor_nuevo_keyboard, back_button
```

IMPORTANTE: Todas las funciones mantienen EXACTAMENTE la misma lógica. Solo cambian:
- `_esc(x)` → `esc(x)` (importado de utils)
- `_format_date(x)` → `format_date(x)` (importado de utils)
- `_calc_vencimiento(x)` → `calc_vencimiento(x)` (importado de utils)
- `_main_keyboard()` → `main_keyboard()` (importado de utils)
- `_edit_keyboard(x)` → `edit_keyboard(x)` (importado de utils)
- `_proveedor_nuevo_keyboard()` → `proveedor_nuevo_keyboard()` (importado de utils)
- Parseo de montos inline → `parsear_monto(texto)` (importado de utils)

- [ ] **Step 2: Escribir handlers/finanzas.py completo**

Contiene:
- Comandos: cmd_pagado, cmd_deposito, cmd_saldo, cmd_dashboard, cmd_reporte, cmd_banco
- Callbacks: cb_reporte, cb_medio_pago, cb_calce_verificacion
- Jobs: job_sync_banco, _sync_banco_core
- Flujos texto: handle_text_pagado(update, context), handle_text_deposito(update, context)

Imports necesarios:
```python
import os, sys, logging, asyncio
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from config import TELEGRAM_CHAT_ID
from excel_manager import (buscar_factura, registrar_pago, registrar_deposito_caja,
                          consultar_saldo_caja, reporte_diario, reporte_semanal,
                          reporte_mensual, guardar_movimientos_banco, obtener_resumen_banco)
from utils.formatting import esc
from utils.parsing import parsear_fecha, parsear_monto
```

- [ ] **Step 3: Escribir handlers/tareas.py completo**

Contiene:
- Comandos: cmd_tarea, cmd_tareas, cmd_hecho, cmd_bitacora
- Callbacks: cb_tarea_prioridad
- Flujos texto: handle_text_tarea(update, context), handle_text_bitacora(update, context)

Imports:
```python
import logging, asyncio
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from tareas_manager import (crear_tarea, listar_tareas, actualizar_tarea,
                            registrar_bitacora, resumen_tareas_semana)
from utils.formatting import esc
```

- [ ] **Step 4: Escribir handlers/inventario_h.py completo**

Contiene:
- Comandos: cmd_inventario, cmd_uso
- Callbacks: cb_cultivo
- Flujos texto: handle_text_uso(update, context)

Imports:
```python
import logging, asyncio
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
from inventario_manager import (consultar_inventario, registrar_uso, productos_bajo_stock)
from utils.formatting import esc
```

- [ ] **Step 5: Escribir handlers/personal.py completo**

Contiene:
- Comandos: cmd_personal, cmd_vacaciones, cmd_agregar_trabajador
- Flujos texto: handle_text_vacacion(update, context), handle_text_trabajador(update, context)

Imports:
```python
import logging, asyncio
from telegram import Update
from telegram.ext import ContextTypes
from vacaciones_manager import (listar_personal, registrar_vacacion, agregar_trabajador)
from utils.formatting import esc
from utils.parsing import parsear_fecha
```

- [ ] **Step 6: Escribir handlers/chat.py completo**

Este es el dispatcher central de texto. La función `handle_text` revisa flujos activos y delega:

```python
"""
handlers/chat.py — Dispatcher central de mensajes de texto.

Revisa si hay un flujo conversacional activo (edición factura, pagado, etc.)
y delega al handler correcto. Si no hay flujo, usa el chat inteligente.
"""
import logging, asyncio
from telegram import Update
from telegram.ext import ContextTypes
from chat_inteligente import responder_chat

logger = logging.getLogger(__name__)


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Punto de entrada para TODOS los mensajes de texto (no comandos)."""
    # Importar handlers de flujo
    from handlers import facturas, finanzas, tareas, inventario_h, personal

    # ── Flujos activos (orden de prioridad) ──
    if context.user_data.get("deposito_state"):
        return await finanzas.handle_text_deposito(update, context)

    if context.user_data.get("pagado_state"):
        return await finanzas.handle_text_pagado(update, context)

    if context.user_data.get("tarea_state"):
        return await tareas.handle_text_tarea(update, context)

    if context.user_data.get("bitacora_state"):
        return await tareas.handle_text_bitacora(update, context)

    if context.user_data.get("uso_state"):
        return await inventario_h.handle_text_uso(update, context)

    if context.user_data.get("vacacion_state"):
        return await personal.handle_text_vacacion(update, context)

    if context.user_data.get("trabajador_state"):
        return await personal.handle_text_trabajador(update, context)

    if context.user_data.get("editing_field"):
        return await facturas.handle_text_edit_factura(update, context)

    # ── Sin flujo activo → Chat inteligente ──
    texto = update.message.text.strip()
    if len(texto) > 2:
        msg = await update.message.reply_text("💬 Pensando...")
        respuesta = await responder_chat(texto)
        try:
            await msg.edit_text(respuesta, parse_mode="Markdown")
        except Exception:
            await msg.edit_text(respuesta)
```

- [ ] **Step 7: Reescribir main.py reducido**

```python
"""
main.py — Bot Agrícola Santa Elisa
Punto de entrada: configura logging, registra handlers, arranca polling.
"""
import os
import logging
from datetime import time as dtime

from telegram.ext import ApplicationBuilder

from config import TELEGRAM_TOKEN, TELEGRAM_CHAT_ID
from handlers import register_all
from excel_manager import crear_hoja_banco
from tareas_manager import crear_hojas_tareas
from inventario_manager import crear_hojas_inventario
from vacaciones_manager import crear_hojas_vacaciones
from handlers.finanzas import job_sync_banco, job_vacaciones_mensuales

_LOG_FMT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
logging.basicConfig(format=_LOG_FMT, level=logging.INFO)
logger = logging.getLogger(__name__)

# FileHandler para persistir logs
try:
    _LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot.log")
    _fh = logging.FileHandler(_LOG_FILE, encoding="utf-8")
    _fh.setLevel(logging.INFO)
    _fh.setFormatter(logging.Formatter(_LOG_FMT))
    logging.getLogger().addHandler(_fh)
except Exception as _e:
    logger.warning(f"No se pudo crear FileHandler de log: {_e}")


def main():
    if not TELEGRAM_TOKEN:
        logger.error("TELEGRAM_TOKEN no configurado")
        return

    app = (ApplicationBuilder()
           .token(TELEGRAM_TOKEN)
           .read_timeout(120).write_timeout(120)
           .connect_timeout(120).pool_timeout(120)
           .build())

    # Registrar todos los handlers
    register_all(app)

    # Crear hojas Excel si no existen
    for init_fn in (crear_hoja_banco, crear_hojas_tareas, crear_hojas_inventario, crear_hojas_vacaciones):
        try:
            init_fn()
        except Exception:
            pass

    # Jobs programados
    try:
        from zoneinfo import ZoneInfo
    except ImportError:
        from backports.zoneinfo import ZoneInfo

    tz_chile = ZoneInfo("America/Santiago")
    app.job_queue.run_daily(job_sync_banco, time=dtime(hour=8,  minute=0, tzinfo=tz_chile), name="banco_8am")
    app.job_queue.run_daily(job_sync_banco, time=dtime(hour=18, minute=0, tzinfo=tz_chile), name="banco_6pm")
    app.job_queue.run_monthly(job_vacaciones_mensuales,
                              when=dtime(hour=7, minute=0, tzinfo=tz_chile),
                              day=1, name="vacaciones_mensuales")
    logger.info("⏰ Jobs programados: banco 08:00/18:00, vacaciones día 1")

    logger.info("✅ Bot iniciado. Esperando mensajes...")
    app.run_polling()


if __name__ == "__main__":
    main()
```

- [ ] **Step 8: Verificar que el bot arranca sin errores**

Run: `python -c "from handlers import register_all; from processors.extractor import process_file; print('Imports OK')"`
Expected: `Imports OK`

Run: `python main.py` (verificar que arranca y responde a /start)

- [ ] **Step 9: Commit**

```
git add handlers/ utils/ main.py
git commit -m "refactor: modularizar main.py en handlers/ y utils/"
```

---

## Task 9: Mover job_vacaciones_mensuales a handlers/personal.py

**Files:**
- Modify: `handlers/finanzas.py` — eliminar job_vacaciones_mensuales
- Modify: `handlers/personal.py` — agregar job_vacaciones_mensuales
- Modify: `main.py` — cambiar import

- [ ] **Step 1: Mover la función**

`job_vacaciones_mensuales` pertenece lógicamente a personal, no a finanzas. Moverla a `handlers/personal.py` y actualizar el import en `main.py`:

```python
# main.py — cambiar:
from handlers.finanzas import job_sync_banco
from handlers.personal import job_vacaciones_mensuales
```

- [ ] **Step 2: Verificar**

Run: `python -c "from handlers.finanzas import job_sync_banco; from handlers.personal import job_vacaciones_mensuales; print('OK')"`

- [ ] **Step 3: Commit**

```
git add handlers/finanzas.py handlers/personal.py main.py
git commit -m "refactor: mover job_vacaciones_mensuales a handlers/personal.py"
```

---

## Task 10: Limpieza final y verificación

- [ ] **Step 1: Eliminar archivos residuales si existen**

Archivos `py` y `cd` sueltos en la raíz que parecen basura:
- `Robot/py`
- `Robot/cd`

- [ ] **Step 2: Verificar que no quedan imports rotos**

Run: `python -c "import main; print('main OK')"`
Run: `python -c "from handlers import facturas, finanzas, tareas, inventario_h, personal, chat; print('All handlers OK')"`
Run: `python -c "from utils import parsing, formatting, keyboards; print('All utils OK')"`
Run: `python -c "from processors.extractor import process_file; print('Extractor OK')"`

- [ ] **Step 3: Verificar conteo de líneas**

Run: `wc -l main.py handlers/*.py utils/*.py`
Expected: main.py ~80 líneas, cada handler 100-400 líneas, total similar al original.

- [ ] **Step 4: Commit final**

```
git add -A
git commit -m "refactor: limpieza final — bot modularizado completamente"
```
