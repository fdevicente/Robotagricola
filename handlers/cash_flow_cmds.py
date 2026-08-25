"""Comandos Telegram para cash flow."""
import logging
from modules.cash_flow.projector import get_cash_flow

logger = logging.getLogger(__name__)

_MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
          "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _label(y, m):
    return f"{_MESES[m]}-{str(y)[-2:]}"


def _fmt_money(v):
    return f"${v:,.0f}"


def format_proyeccion(cf: dict) -> str:
    """Formato texto para Telegram."""
    lines = ["📊 *Proyeccion flujo de caja*", ""]
    lines.append(f"`{'Mes':<8} {'Ingr':>12} {'Egr':>12} {'Saldo':>14}`")
    for ym in cf["months"]:
        s = cf["saldo"][ym]
        emoji = "🔴" if s["saldo_cierre"] < 0 else ""
        lines.append(
            f"`{_label(*ym):<8} {_fmt_money(s['ingresos']):>12} "
            f"{_fmt_money(s['egresos']):>12} {_fmt_money(s['saldo_cierre']):>14}` {emoji}"
        )
    return "\n".join(lines)


async def cmd_proyeccion(update, context):
    """`/proyeccion [meses]` (default 6)."""
    from modules.cuentas import caja_total

    args = context.args or []
    n_meses = int(args[0]) if args and args[0].isdigit() else 6
    caja = caja_total()                 # cuenta corriente + cuenta dólar
    saldo_actual = caja["total"]

    from datetime import date
    today = date.today()
    sy, sm = today.year, today.month
    ey, em = sy, sm
    for _ in range(n_meses - 1):
        em += 1
        if em > 12:
            em = 1
            ey += 1

    cf = get_cash_flow(start=(sy, sm), end=(ey, em),
                       saldo_inicial=saldo_actual)
    from modules.cuentas import formato as _fmt_caja
    text = _fmt_caja(caja) + "\n\n" + format_proyeccion(cf)
    await update.message.reply_text(text, parse_mode="Markdown")


def format_categoria(cat_name: str, egresos: dict, months: list) -> str:
    """Texto Telegram con detalle de una categoria."""
    lines = [f"📋 *Categoria: {cat_name}*", ""]
    total = 0
    for ym in months:
        m_total = sum(v for (y, mo, c, _cul), v in egresos.items()
                       if y == ym[0] and mo == ym[1] and c == cat_name)
        total += m_total
        lines.append(f"`{_label(*ym):<8} {_fmt_money(m_total):>14}`")
    lines.append("")
    lines.append(f"`Total      {_fmt_money(total):>14}`")
    return "\n".join(lines)


async def cmd_categoria(update, context):
    """`/categoria <nombre>` muestra gasto mensual."""
    args = context.args or []
    if not args:
        await update.message.reply_text(
            "Uso: /categoria <nombre>. Ej: /categoria Fertilizantes")
        return
    cat = " ".join(args)
    from datetime import date
    today = date.today()
    months = [(today.year, m) for m in range(1, today.month + 1)]
    cf = get_cash_flow(start=months[0], end=months[-1],
                       saldo_inicial=0)
    text = format_categoria(cat, cf["egresos"], months)
    await update.message.reply_text(text, parse_mode="Markdown")


from telegram.ext import (
    ConversationHandler, CommandHandler, MessageHandler, filters,
)
from modules.cash_flow.cosecha_wizard import CosechaWizard, save_to_cosechas

WIZARD_STATE = 1


async def cmd_cosecha_start(update, context):
    args = context.args or []
    if not args:
        await update.message.reply_text("Uso: /cosecha <NOGALES|CEREZOS|AVELLANOS>")
        return ConversationHandler.END
    cultivo = args[0].upper()
    if cultivo not in ("NOGALES", "CEREZOS", "AVELLANOS"):
        await update.message.reply_text("Cultivo invalido")
        return ConversationHandler.END
    w = CosechaWizard(cultivo=cultivo)
    context.user_data["cosecha_wizard"] = w
    await update.message.reply_text(w.prompt)
    return WIZARD_STATE


async def cb_cosecha_resp(update, context):
    w = context.user_data.get("cosecha_wizard")
    if not w:
        return ConversationHandler.END
    w.responder(update.message.text)
    if w.estado == "resumen":
        from datetime import date
        year = date.today().year
        added = save_to_cosechas(w.data, year=year)
        await update.message.reply_text(
            f"Listo. {added} filas guardadas en Cosechas.")
        context.user_data.pop("cosecha_wizard", None)
        return ConversationHandler.END
    await update.message.reply_text(w.prompt)
    return WIZARD_STATE


async def cmd_cosecha_cancel(update, context):
    context.user_data.pop("cosecha_wizard", None)
    await update.message.reply_text("Wizard cancelado.")
    return ConversationHandler.END


cosecha_conv = ConversationHandler(
    entry_points=[CommandHandler("cosecha", cmd_cosecha_start)],
    states={WIZARD_STATE: [MessageHandler(filters.TEXT & ~filters.COMMAND,
                                            cb_cosecha_resp)]},
    fallbacks=[CommandHandler("cancelar", cmd_cosecha_cancel)],
)
