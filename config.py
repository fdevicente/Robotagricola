import os
from dotenv import load_dotenv

load_dotenv()

# Intentar usar Credential Manager para claves sensibles
try:
    from credential_manager import get_secret
    _USE_CM = True
except ImportError:
    _USE_CM = False
    def get_secret(key, fallback=""):
        return os.getenv(key, fallback)

TELEGRAM_TOKEN    = get_secret("TELEGRAM_TOKEN")
OLLAMA_MODEL      = os.getenv("OLLAMA_MODEL", "llama3.2-vision-fast")
OLLAMA_HOST       = os.getenv("OLLAMA_HOST",  "http://localhost:11434")
EXCEL_PATH        = os.getenv("EXCEL_PATH",   os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "MASTER Agricola Santa Elisa.xlsx"))
DOWNLOAD_DIR      = os.getenv("DOWNLOAD_DIR", "Facturas Recibidas por Telegram")
BOLETAS_DIR       = os.getenv("BOLETAS_DIR",  os.path.join(os.path.dirname(DOWNLOAD_DIR), "Boletas Recibidas por Telegram"))
ANTHROPIC_API_KEY = get_secret("ANTHROPIC_API_KEY")

# Scotiabank scraper credentials
BANCO_RUT_EMPRESA = get_secret("BANCO_RUT_EMPRESA")
BANCO_RUT_USUARIO = get_secret("BANCO_RUT_USUARIO")
BANCO_CLAVE       = get_secret("BANCO_CLAVE")

# Chat ID para notificaciones automáticas (se obtiene con /start y se guarda)
TELEGRAM_CHAT_ID  = os.getenv("TELEGRAM_CHAT_ID", "")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(BOLETAS_DIR, exist_ok=True)

# --- Cash Flow (Fase 1) ---
DROPBOX_BASE = os.getenv("DROPBOX_BASE",
    r"C:\Users\Windows\Dropbox\Agricola Santa Elisa")
DROPBOX_BACKUP_PATH = os.path.join(DROPBOX_BASE, "Backups")
FXP_PATH = os.path.join(DROPBOX_BASE, "FXP.xlsx")
GUIAS_DIR = os.getenv("GUIAS_DIR",
    os.path.join(os.path.dirname(EXCEL_PATH), "Guias Recibidas por Telegram"))
DOCUMENTOS_DIR = os.path.join(DROPBOX_BASE, "Documentos")
REPORTES_DIR = os.path.join(os.path.dirname(EXCEL_PATH), "Reportes")

for _d in [GUIAS_DIR, DOCUMENTOS_DIR, REPORTES_DIR,
           os.path.join(DOCUMENTOS_DIR, "Guias Despacho"),
           os.path.join(DROPBOX_BACKUP_PATH, "Master", "snapshots"),
           os.path.join(DROPBOX_BACKUP_PATH, "Robot")]:
    os.makedirs(_d, exist_ok=True)

CASH_FLOW_CONFIG = {
    'saldo_minimo_pct': 0.10,
    'umbral_alerta_cat_pct': 0.90,
    'umbral_confianza': 0.85,
    'ventana_match_dias': 15,
    'fecha_limite_cerezas': '12-15',
    'fecha_limite_nueces': '05-30',
    'dias_sin_guia_cierre': 7,
    'usd_clp_estimado': 1000,
}
