"""
Script de setup inicial para Fase 1 - Cash Flow.
Ejecutar UNA VEZ antes de usar el modulo cash_flow.

Uso: py -3.11 setup_cash_flow.py
"""
import logging
from infrastructure.backups import backup_master
from excel_manager import ensure_cash_flow_sheets, ensure_new_columns
from config import EXCEL_PATH

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
logger = logging.getLogger(__name__)


def main():
    logger.info("=== Setup Cash Flow - Fase 1 ===")

    logger.info("1/3 Backup preventivo...")
    backup_master(reason="pre-fase-1-setup")

    logger.info("2/3 Creando hojas nuevas en Master...")
    ensure_cash_flow_sheets()
    logger.info("   Hojas creadas: Cosechas, Guias Despacho, Flujo Caja, "
                "Ajustes Manuales, Config, Hectareas")

    logger.info("3/3 Agregando columnas nuevas...")
    ensure_new_columns()
    logger.info("   Facturas: +Categoria, +Cultivo, +Confianza, +Categorizado_por")
    logger.info("   Cuenta Banco: +Tipo, +Categoria, +Cultivo, +Factura_linkeada")

    backup_master(reason="post-fase-1-setup")

    logger.info("=== Setup completo ===")
    logger.info(f"Master: {EXCEL_PATH}")
    logger.info("Siguiente paso: ejecutar Plan 2 (categorizacion del historico)")


if __name__ == "__main__":
    main()
