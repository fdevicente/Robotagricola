> # 🔒 DOCUMENTO HISTÓRICO — CONGELADO EL 2026-08-17. NO ACTUALIZAR.
>
> Esto es el registro de **una sesión puntual del 11-junio-2026**, no el estado del
> proyecto. Se conserva como memoria de por qué el código quedó así. **No lo edites,
> no lo uses para decidir nada y no corras los scripts que menciona.**
>
> **El estado real vive en la memoria:**
> `~/.claude/projects/C--Users-Windows-Desktop-Workflow/memory/project_pendientes_roadmap.md`
> (lista de arranque — leer primero) y el resto de `memory/`.
>
> **Advertencias concretas sobre lo que hay aquí abajo:**
> - ⛔ **NO correr `fix_unitario.py`.** Dice "ejecutar una sola vez" y ya se ejecutó
>   en junio-2026. Volver a correrlo recalcularía el valor unitario de TODAS las
>   filas sobre datos que después se corrigieron a mano (66 montos alineados a FXP
>   el 28-jul). Además el comando que muestra (`python fix_unitario.py`) no
>   funciona: **no hay `python` en el PATH**, es
>   `%LOCALAPPDATA%\Python\bin\python3.11.exe`.
> - ⚠️ **De aquí nace la fórmula `=J*K` de la columna L (TOTAL NETO)**, y hoy esas
>   **1.475 fórmulas no tienen valor en caché**: cualquier lectura con
>   `data_only=True` devuelve `None`. openpyxl las vacía al guardar y el bot guarda
>   seguido. El proyector no se afecta porque lee la **columna O (Total por Item)**.
>   Si algo llega a necesitar TOTAL NETO, hay que calcularlo o recalcular el libro.
> - Los fragmentos de código son de junio-2026 y `main.py` fue refactorizado
>   después: no asumir que las líneas siguen donde dice.

---

# Cambios Robot Agrícola Santa Elisa — Sesión 2026-06-11

## 1. Precio unitario como fórmula en Excel (`excel_manager.py`)

**Problema:** El precio unitario que extraía la IA desde las facturas siempre salía incorrecto (más bajo de lo real). La cantidad y el total neto de línea eran correctos.

**Solución:**
- **Col J (Valor unitario):** se escribe como valor hardcodeado (lo que la IA extrajo, o lo que el usuario corrigió en Telegram).
- **Col L (TOTAL NETO):** se escribe como fórmula dinámica `=J{fila}*K{fila}` (precio × cantidad, sin IVA). Se recalcula automáticamente si se corrige J o K directamente en Excel.

```python
# excel_manager.py — dentro de append_to_excel()
row[9]  = item.get("Valor unitario")   # Col J: valor
row[11] = None                          # Col L: se llena como fórmula abajo

ws.append(row)
cur_row = ws.max_row

# TOTAL NETO = Valor unitario × Cantidad (sin IVA)
cell_neto = ws.cell(row=cur_row, column=12)
cell_neto.value = f"=J{cur_row}*K{cur_row}"
cell_neto.number_format = '#,##0'
```

---

## 2. Script de corrección de filas existentes (`fix_unitario.py`)

Archivo creado en la carpeta del Robot. Ejecutar **una sola vez** para corregir todas las filas ya guardadas en el Excel.

**Qué hace:**
1. Lee el valor actual de TOTAL NETO (correcto) y Cantidad de cada fila.
2. Calcula Valor unitario = TOTAL NETO / Cantidad y lo escribe como número en col J.
3. Reemplaza col L con la fórmula `=J{fila}*K{fila}`.

**Cómo ejecutar:**
```bash
cd "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\Robot"
python fix_unitario.py
```

---

## 3. Botón "Total ítem" en Telegram (`main.py`)

**Problema:** El botón "💰 Total ítem" pedía el valor **con IVA**, lo que era confuso porque internamente el sistema trabaja con netos.

**Cambio:** El botón ahora pide el **neto de línea** (precio × cantidad, sin IVA).

### Cambios en `CAMPOS_EDITABLES`:
```python
# Antes:
"edit_total": ("Monto / TOTAL", "💰 Total ítem con IVA")

# Después:
"edit_total": ("TOTAL NETO",   "💰 Total ítem sin IVA")
```

### Lógica al guardar el valor editado:
Cuando el usuario escribe el neto, el bot deriva:
- `Valor unitario = neto / Cantidad`
- `Monto / TOTAL = neto × 1.19 + Impuesto Específico` (o sin IVA si es exenta)
- `Total Factura` se actualiza sumando los `Monto / TOTAL` de todos los ítems.

```python
elif campo == "TOTAL NETO":
    qty      = float(item.get("Cantidad") or 1)
    imp_esp  = float(item.get("Impuesto Especifico") or 0)
    sin_iva  = any(k in doc for k in ("exenta", "exento", "no afecta", ...))
    iva_factor = 1.0 if sin_iva else 1.19
    item["Valor unitario"] = val / qty if qty else 0
    item["Monto / TOTAL"]  = round(val * iva_factor + imp_esp)
```

### Set `NUMERICOS` actualizado:
```python
NUMERICOS = {"Valor unitario", "Cantidad", "Monto / TOTAL", "TOTAL NETO"}
```

### Propagación de cambios:
```python
# No recalcular Monto/TOTAL si ya fue calculado en el bloque TOTAL NETO
if campo not in ("Monto / TOTAL", "TOTAL NETO"):
    ...
```

---

## 4. Renombrado automático de documentos (`main.py`)

**Problema:** Los archivos se guardaban con el nombre genérico que venía de Telegram (ej: `document.pdf`, `file_123.jpg`), sin referencia al proveedor ni al número de factura.

**Solución:** Se agrega la función `_renombrar_archivo()` que renombra el archivo en disco usando el nombre del proveedor y el número de documento, inmediatamente después de que la IA extrae los datos.

### Formato del nombre:
```
NombreProveedor_NroFactura.ext
```
Ejemplo: `COPEVAL_123456.pdf`, `Martinez_y_Valdivieso_98765.jpg`

### Función agregada:
```python
def _renombrar_archivo(file_path: str, items: list) -> str:
    """Renombra el archivo usando Proveedor + Nº Documento.
    Devuelve la nueva ruta (o la original si algo falla)."""
    ...
```

### Comportamiento:
- Caracteres inválidos en Windows (`\ / : * ? " < > |`) se eliminan del nombre.
- Espacios se reemplazan por guiones bajos.
- Si ya existe un archivo con ese nombre, agrega timestamp para no sobreescribir.
- Si la IA no detectó ni empresa ni número, el archivo conserva su nombre original.

### Llamada en `_process_and_reply`:
```python
items = result.get("items", [])
# Renombrar el archivo con Proveedor + Nº Documento
file_path = _renombrar_archivo(file_path, items)
context.user_data["pending_file_path"] = file_path
```

---

## Resumen de archivos modificados

| Archivo | Cambio |
|---|---|
| `excel_manager.py` | Col L (TOTAL NETO) escrita como fórmula `=J*K` |
| `main.py` | Botón "Total ítem" → sin IVA + lógica derivación unitario |
| `main.py` | Función `_renombrar_archivo()` + llamada en `_process_and_reply` |
| `fix_unitario.py` | Nuevo script para corregir filas existentes (ejecutar 1 vez) |
