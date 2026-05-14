"""Recalcula la hoja Flujo Caja con la proyeccion actual.

Uso: py -3.11 recalc_flujo_caja.py
     py -3.11 recalc_flujo_caja.py --saldo 130600000
"""
import argparse
import logging

from infrastructure.backups import backup_master
from modules.cash_flow.projector import get_cash_flow, write_flujo_caja

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logger = logging.getLogger(__name__)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--saldo", type=float, default=130_600_000,
                         help="Saldo banco actual CLP (default 130.6M)")
    parser.add_argument("--start", default="2026-05")
    parser.add_argument("--end", default="2027-04")
    parser.add_argument("--base-year", type=int, default=2025)
    args = parser.parse_args()

    sy, sm = map(int, args.start.split("-"))
    ey, em = map(int, args.end.split("-"))

    logger.info("Backup pre-recalc...")
    backup_master(reason="pre-recalc-flujo")

    logger.info(f"Calculando proyeccion {args.start} -> {args.end}...")
    result = get_cash_flow(
        start=(sy, sm), end=(ey, em),
        saldo_inicial=args.saldo, base_year=args.base_year,
    )

    logger.info(f"Escribiendo {len(result['months'])} meses en Flujo Caja...")
    write_flujo_caja(
        saldo_data=result["saldo"], egresos=result["egresos"],
        ingresos=result["ingresos"], months=result["months"],
    )

    logger.info("Backup post-recalc...")
    backup_master(reason="post-recalc-flujo")

    total_ing = sum(s["ingresos"] for s in result["saldo"].values())
    total_eg = sum(s["egresos"] for s in result["saldo"].values())
    saldo_final = list(result["saldo"].values())[-1]["saldo_cierre"]
    logger.info("=== Proyeccion ===")
    logger.info(f"Ingresos totales: ${total_ing:>15,.0f}")
    logger.info(f"Egresos totales:  ${total_eg:>15,.0f}")
    logger.info(f"Saldo final:      ${saldo_final:>15,.0f}")


if __name__ == "__main__":
    main()
