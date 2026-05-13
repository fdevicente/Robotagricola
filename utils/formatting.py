"""
utils/formatting.py - Helpers de formato para mensajes de Telegram.
- esc: escapa caracteres Markdown
- format_date: fecha legible
- calc_vencimiento: emision + 1 mes
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
    """Calcula fecha de vencimiento = emision + 1 mes."""
    try:
        from dateutil.relativedelta import relativedelta
        if not fecha_emision:
            return None
        dt = datetime.strptime(str(fecha_emision)[:10], "%Y-%m-%d")
        return (dt + relativedelta(months=1)).strftime("%Y-%m-%d")
    except Exception:
        return None
