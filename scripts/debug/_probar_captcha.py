"""Abre el login del banco y reporta si detecta el verificador anti-robot.

No envía credenciales: la detección corre antes de rellenar el formulario.
"""
import logging
import sys

sys.stdout.reconfigure(encoding="utf-8")
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

from scotiabank_scraper import CaptchaRequerido, sync_scotiabank_movements

try:
    movs = sync_scotiabank_movements()
    print(f"\n✅ Entró sin CAPTCHA — {len(movs)} movimientos extraídos.")
except CaptchaRequerido as e:
    print("\n🔐 CAPTCHA DETECTADO — el scraper cortó como corresponde:\n")
    print(e)
except Exception as e:
    print(f"\n❌ Falló por otra razón ({type(e).__name__}):\n{e}")
