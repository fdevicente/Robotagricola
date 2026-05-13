"""
Script de onboarding historico - Fase 2 Cash Flow.
Ejecutar UNA VEZ despues de setup_cash_flow.py.

Acciones:
1. Backup preventivo
2. Detectar patrones de ingreso en Cuenta Banco
3. Batch categorize Facturas con Claude (~USD $5 para 1300 facturas)
4. Backup post-import
5. Reportar # de filas a revisar (confianza < 0.85)

Uso:
  py -3.11 onboarding_cash_flow.py
  py -3.11 onboarding_cash_flow.py --limit 50
  py -3.11 onboarding_cash_flow.py --skip-claude
"""
import argparse
import logging
import sys

from infrastructure.backups import backup_master
from modules.cash_flow.historical_importer import detect_income_patterns
from modules.cash_flow.categorizer import batch_categorize_history
from config import EXCEL_PATH

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logger = logging.getLogger(__name__)


def _progress(done: int, total: int):
    if done % 25 == 0 or done == total:
        pct = 100 * done / total if total else 100
        logger.info(f"   ... {done}/{total} ({pct:.0f}%)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=None,
                         help="Maximo de facturas a categorizar (None=todas)")
    parser.add_argument("--skip-claude", action="store_true",
                         help="Solo correr patrones de banco, no llamar Claude")
    args = parser.parse_args()

    logger.info("=== Onboarding Cash Flow - Fase 2 ===")
    logger.info(f"Master: {EXCEL_PATH}")

    logger.info("1/4 Backup preventivo...")
    backup_master(reason="pre-fase-2-onboarding")

    logger.info("2/4 Detectando patrones de ingreso en banco...")
    bank_report = detect_income_patterns()
    logger.info(f"   Banco: {bank_report}")

    if args.skip_claude:
        logger.info("3/4 SKIP - flag --skip-claude")
        cat_report = {"processed": 0, "low_confidence": 0}
    else:
        logger.info(f"3/4 Categorizando facturas con Claude (limit={args.limit})...")
        logger.info("   Costo estimado: ~USD $5 para 1300 facturas")
        cat_report = batch_categorize_history(
            limit=args.limit, progress_cb=_progress,
        )
        logger.info(f"   Categorizacion: {cat_report}")

    logger.info("4/4 Backup post-onboarding...")
    backup_master(reason="post-fase-2-onboarding")

    logger.info("=== Onboarding completo ===")
    logger.info(f"Facturas categorizadas: {cat_report.get('processed', 0)}")
    logger.info(f"REVISAR (baja confianza): {cat_report.get('low_confidence', 0)}")
    logger.info(f"Banco - venta_dolares: {bank_report.get('venta_dolares', 0)}")
    logger.info(f"Banco - ingreso_clp:   {bank_report.get('ingreso_clp', 0)}")
    logger.info(f"Banco - sueldo:        {bank_report.get('sueldo', 0)}")
    logger.info(f"Banco - sin clasificar: {bank_report.get('no_match', 0)}")
    logger.info("")
    logger.info("Siguiente: revisar filas Categoria=REVISAR en Master.Facturas")
    logger.info("Despues: Plan 3 (banco + matching)")


if __name__ == "__main__":
    sys.exit(main())
