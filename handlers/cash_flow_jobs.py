"""Jobs programados de cash flow (resumen semanal, etc.)."""
import logging
from datetime import date

logger = logging.getLogger(__name__)

_MESES = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
          "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _label(y, m):
    return f"{_MESES[m]}-{str(y)[-2:]}"


def format_resumen_semanal(cf: dict, alertas: list, caja: dict | None = None) -> str:
    lines = ["📅 *Resumen semanal*", ""]
    if caja:
        # Las dos cuentas por separado: la caja no vive solo en pesos
        lines.append(f"💰 Caja total ${caja['total']:,.0f}")
        lines.append(f"    corriente ${caja['clp']:,.0f} · "
                      f"dólar US${caja['usd']:,.0f}")
        lines.append("")
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


async def job_reporte_mensual(context):
    """Genera el reporte mensual del mes anterior en PDF y lo envía al grupo.

    Trigger: día 1 de cada mes, 08:00 hora Chile.
    Envía al chat donde se suben las facturas (banco_chat_id) o al
    TELEGRAM_CHAT_ID configurado como fallback.
    """
    from config import TELEGRAM_CHAT_ID
    from modules.cash_flow.reporte_pdf import generar_reporte_pdf

    today = date.today()
    # Mes anterior
    if today.month == 1:
        y, m = today.year - 1, 12
    else:
        y, m = today.year, today.month - 1

    # Destino: dueño (/soydueno) primero; fallback banco_chat_id / config
    chat_id = (context.bot_data.get("owner_chat_id")
               or context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID)
    if not chat_id:
        logger.warning("Sin chat configurado, skip reporte mensual")
        return

    try:
        logger.info(f"Generando reporte mensual {y}-{m:02d}...")
        pdf_path = generar_reporte_pdf(y, m)
        mes_nombre = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                       "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"][m]
        with open(pdf_path, "rb") as f:
            await context.bot.send_document(
                chat_id=chat_id,
                document=f,
                filename=f"Reporte_{mes_nombre}_{y}.pdf",
                caption=f"📊 *Reporte Mensual {mes_nombre} {y}*\n"
                        f"Agrícola Santa Elisa",
                parse_mode="Markdown",
            )
        logger.info(f"Reporte mensual {y}-{m:02d} enviado a {chat_id}")
    except Exception as e:
        logger.error(f"Error generando/enviando reporte mensual: {e}")
        try:
            await context.bot.send_message(
                chat_id=chat_id,
                text=f"⚠️ No pude generar el reporte mensual de {y}-{m:02d}: {e}")
        except Exception:
            pass


async def job_resumen_semanal(context):
    """Envia resumen al chat configurado. Trigger: cron Lunes 8am."""
    from config import TELEGRAM_CHAT_ID
    from modules.cash_flow.projector import get_cash_flow
    from modules.cash_flow.alerts import detect_factura_por_vencer
    from excel_manager import read_facturas_pendientes

    chat_id = (context.bot_data.get("owner_chat_id")
               or context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID)
    if not chat_id:
        logger.warning("Sin chat destino (correr /soydueno), skip resumen semanal")
        return

    from modules.cuentas import caja_total

    today = date.today()
    caja = caja_total()          # cuenta corriente + cuenta dólar
    cf = get_cash_flow(start=(today.year, today.month),
                       end=(today.year, today.month),
                       saldo_inicial=caja["total"])
    facturas = read_facturas_pendientes()
    alertas = detect_factura_por_vencer(facturas, hoy=today, dias=7)
    text = format_resumen_semanal(cf, alertas, caja)
    try:
        await context.bot.send_message(chat_id=chat_id,
                                        text=text, parse_mode="Markdown")
    except Exception:
        await context.bot.send_message(chat_id=chat_id, text=text)
