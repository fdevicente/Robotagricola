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
    args = context.args or []
    n_meses = int(args[0]) if args and args[0].isdigit() else 6
    saldo_actual = 130_600_000

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
    text = format_proyeccion(cf)
    await update.message.reply_text(text, parse_mode="Markdown")
