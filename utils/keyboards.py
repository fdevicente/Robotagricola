"""
utils/keyboards.py - Constructores de InlineKeyboardMarkup para Telegram.
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
        [InlineKeyboardButton("\U0001f3e2 Proveedor",  callback_data="edit_proveedor"),
         InlineKeyboardButton("\U0001faaa RUT",         callback_data="edit_rut")],
        [InlineKeyboardButton("\U0001f4c5 Fecha",       callback_data="edit_fecha"),
         InlineKeyboardButton("⏰ Vencimiento", callback_data="edit_vence")],
        [InlineKeyboardButton("\U0001f4c4 Nº Documento", callback_data="edit_nro"),
         InlineKeyboardButton("\U0001f517 Ref. Factura", callback_data="edit_ref")],
        [InlineKeyboardButton("\U0001f4e6 Glosa",       callback_data="edit_glosa")],
        [InlineKeyboardButton("\U0001f4dd Detalle",     callback_data="edit_glosa2"),
         InlineKeyboardButton("\U0001f522 Cantidad",    callback_data="edit_cantidad")],
        [InlineKeyboardButton("\U0001f4b2 Unit neto",   callback_data="edit_unitario"),
         InlineKeyboardButton("\U0001f4b0 Total ítem",  callback_data="edit_total")],
        [InlineKeyboardButton("\U0001f4b0 TOTAL FACTURA", callback_data="edit_total_factura")],
    ]
    item_row = [InlineKeyboardButton("➕ Agregar ítem", callback_data="add_item")]
    if n_items > 1:
        item_row.append(InlineKeyboardButton("\U0001f5d1️ Eliminar ítem", callback_data="del_item"))
    kb.append(item_row)
    kb.append([InlineKeyboardButton("« Volver", callback_data="back_preview")])
    return InlineKeyboardMarkup(kb)


def proveedor_nuevo_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Sí, agregar a la lista", callback_data="add_proveedor_yes"),
         InlineKeyboardButton("❌ No",                    callback_data="add_proveedor_no")],
    ])


def back_button(callback_data="edit_menu"):
    """Boton solitario Volver."""
    return InlineKeyboardMarkup([[InlineKeyboardButton("« Volver", callback_data=callback_data)]])
