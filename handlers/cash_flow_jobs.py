"""Jobs programados de cash flow (resumen semanal, etc.)."""
import logging
from datetime import date

logger = logging.getLogger(__name__)

_MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
          "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _label(y, m):
    return f"{_MESES[m]}-{str(y)[-2:]}"


def format_resumen_semanal(cf: dict, alertas: list) -> str:
    lines = ["📅 *Resumen semanal*", ""]
    ym = cf["months"][0] if cf["months"] else None
    if ym:
        s = cf["saldo"][ym]
        lines.append(f"Mes {_label(*ym)}: saldo cierre ${s['saldo_cierre']:,.0f}")
        lines.append(f"  ingresos ${s['ingresos']:,.0f}, "
                      f"egresos ${s['egresos']:,.0f}")
    lines.append("")
    if alertas:
        lines.append("*Alertas:*")
        for a in alertas:
            lines.append(f"  {a['mensaje']}")
    else:
        lines.append("✅ Sin alertas activas")
    return "\n".join(lines)


async def job_resumen_semanal(context):
    """Envia resumen al chat configurado. Trigger: cron Lunes 8am."""
    from config import TELEGRAM_CHAT_ID
    from modules.cash_flow.projector import get_cash_flow
    from modules.cash_flow.alerts import detect_factura_por_vencer
    from excel_manager import read_facturas_pendientes

    if not TELEGRAM_CHAT_ID:
        logger.warning("TELEGRAM_CHAT_ID no configurado, skip resumen semanal")
        return

    today = date.today()
    cf = get_cash_flow(start=(today.year, today.month),
                       end=(today.year, today.month),
                       saldo_inicial=130_600_000)
    facturas = read_facturas_pendientes()
    alertas = detect_factura_por_vencer(facturas, hoy=today, dias=7)
    text = format_resumen_semanal(cf, alertas)
    await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID,
                                    text=text, parse_mode="Markdown")
