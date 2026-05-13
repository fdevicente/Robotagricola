"""
scotiabank_scraper.py
Automates login to Scotiabank Chile Portal Empresas and extracts last movements.
Credentials loaded from .env (never hardcoded).
"""

import logging
import os
import re
import time
from datetime import datetime

from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

logger = logging.getLogger(__name__)

# === Credentials: Credential Manager > .env ===
try:
    from credential_manager import get_secret
    RUT_EMPRESA = get_secret("BANCO_RUT_EMPRESA")
    RUT_USUARIO = get_secret("BANCO_RUT_USUARIO")
    CLAVE       = get_secret("BANCO_CLAVE")
except ImportError:
    RUT_EMPRESA = os.getenv("BANCO_RUT_EMPRESA", "")
    RUT_USUARIO = os.getenv("BANCO_RUT_USUARIO", "")
    CLAVE       = os.getenv("BANCO_CLAVE", "")

# === Selectors (verified) ===
LOGIN_URL = "https://appservtrx.scotiabank.cl/portalempresas/login"
SEL_RUT_EMPRESA = "#INP_COMMON_RUT_CLIENTRUT"
SEL_RUT_USUARIO = "#INP_COMMON_RUT_USERRUT"
SEL_CLAVE = "#INP_COMMON_PASSWORD_PASS"
SEL_BTN_LOGIN = "#BTN_COMMON_LOGIN"


def _parse_amount(text: str) -> float | None:
    """Parse Chilean amount strings like '$1.234.567' or '-$ 500.000' to float."""
    text = text.strip().replace("$", "").replace(" ", "")
    text = re.sub(r'\.(?=\d{3})', '', text)
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _close_modals(page):
    """Cierra modales de T&C, avisos y overlays via JavaScript."""
    for _ in range(5):
        closed = page.evaluate("""() => {
            const btns = document.querySelectorAll(
                'button.close, .modal .close, [class*="close"], button[aria-label="Close"]'
            );
            let closed = 0;
            btns.forEach(b => {
                if (b.offsetParent !== null) { b.click(); closed++; }
            });
            document.querySelectorAll('button, a').forEach(el => {
                const txt = el.textContent.trim().toLowerCase();
                if ((txt === 'aceptar' || txt === 'acepto' || txt === 'cerrar' ||
                     txt.includes('acepto los') || txt.includes('términos'))
                    && el.offsetParent !== null) {
                    el.click(); closed++;
                }
            });
            document.querySelectorAll('.modal-backdrop, .overlay, [class*="modal"]').forEach(m => {
                if (m.classList.contains('show') || m.style.display !== 'none') {
                    m.remove(); closed++;
                }
            });
            return closed;
        }""")
        if closed > 0:
            logger.info(f"Modales cerrados: {closed}")
            time.sleep(1)
        else:
            break


def _navigate_to_cuentas(page):
    """Navega al menú Cuentas de la SPA via click por coordenadas."""
    menu_items = page.evaluate("""() => {
        const items = [];
        document.querySelectorAll('a, span, div, li').forEach(el => {
            const txt = el.textContent.trim();
            const rect = el.getBoundingClientRect();
            if (txt === 'Cuentas' && rect.top < 100 && rect.width > 0) {
                items.push({x: rect.x + rect.width/2, y: rect.y + rect.height/2});
            }
        });
        return items;
    }""")
    if menu_items:
        x, y = menu_items[0]['x'], menu_items[0]['y']
        page.mouse.click(x, y)
        time.sleep(3)
        logger.info("Navegado a Cuentas")
        return True
    return False


def sync_scotiabank_movements() -> list[dict]:
    """
    Opens Scotiabank Portal Empresa, logs in and extracts movements.
    Returns a list of dicts: [{fecha, descripcion, referencia, cargo, abono, saldo}]
    """
    if not all([RUT_EMPRESA, RUT_USUARIO, CLAVE]):
        raise RuntimeError(
            "Credenciales bancarias no configuradas. "
            "Agrega BANCO_RUT_EMPRESA, BANCO_RUT_USUARIO y BANCO_CLAVE al .env"
        )

    from playwright.sync_api import sync_playwright

    movimientos = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/119.0.0.0 Safari/537.36"
            )
        )
        page = context.new_page()

        try:
            logger.info(f"Navegando a {LOGIN_URL} ...")
            page.goto(LOGIN_URL, timeout=60_000)
            page.wait_for_selector(SEL_RUT_EMPRESA, timeout=15_000)

            logger.info("Ingresando credenciales...")
            page.fill(SEL_RUT_EMPRESA, RUT_EMPRESA)
            page.fill(SEL_RUT_USUARIO, RUT_USUARIO)
            page.fill(SEL_CLAVE, CLAVE)

            page.click(SEL_BTN_LOGIN)
            page.wait_for_load_state("domcontentloaded", timeout=60_000)
            time.sleep(5)

            if "error" in page.title().lower() or page.locator("text=Clave incorrecta").count() > 0:
                raise RuntimeError("Login fallido: credenciales incorrectas o error en el banco.")

            logger.info("Login exitoso.")
            _close_modals(page)
            time.sleep(2)

            # Navegar a Cuentas (SPA Vue.js — usar click por coordenadas)
            _navigate_to_cuentas(page)

            page.wait_for_load_state("domcontentloaded", timeout=30_000)
            time.sleep(4)

            # Extraer tabla de movimientos
            # Formato: Fecha | Descripción | Nº OPERACIÓN | MONTO | SALDO
            logger.info("Extrayendo tabla de movimientos...")
            rows = page.locator("table tbody tr").all()
            if not rows:
                rows = page.locator("table tr").all()
                if rows and len(rows) > 1:
                    rows = rows[1:]

            if not rows:
                logger.warning("No se encontraron filas en la tabla de movimientos.")

            for row in rows:
                cells = row.locator("td").all_inner_texts()
                if len(cells) < 3:
                    continue

                fecha_raw = cells[0].strip()
                desc = cells[1].strip()

                # Limpiar fecha: quitar "Cargo", "Abono", "Created with Sketch."
                fecha_str = re.sub(r'(Cargo|Abono|Created with Sketch\.?)', '', fecha_raw).strip()

                # Limpiar descripción
                desc = desc.replace('\n', ' ').strip()

                if len(cells) >= 5:
                    referencia = cells[2].strip()
                    monto = _parse_amount(cells[3]) or 0
                    saldo = _parse_amount(cells[4]) or 0
                elif len(cells) >= 4:
                    referencia = ""
                    monto = _parse_amount(cells[2]) or 0
                    saldo = _parse_amount(cells[3]) or 0
                else:
                    referencia = ""
                    monto = _parse_amount(cells[2]) or 0
                    saldo = 0

                # Separar cargo y abono por signo del monto
                cargo = abs(monto) if monto < 0 else 0
                abono = monto if monto > 0 else 0

                movimientos.append({
                    "fecha": fecha_str,
                    "descripcion": desc,
                    "referencia": referencia,
                    "cargo": cargo,
                    "abono": abono,
                    "saldo": saldo,
                })

            logger.info(f"Se extrajeron {len(movimientos)} movimientos del banco.")

        except Exception as e:
            logger.error(f"Error en scraping Scotiabank: {e}")
            try:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                screenshot_dir = os.path.join(os.path.dirname(__file__), "Claude")
                page.screenshot(path=os.path.join(screenshot_dir, f"bank_error_{ts}.png"))
            except Exception:
                pass
            raise

        finally:
            browser.close()

    return movimientos


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    movs = sync_scotiabank_movements()
    print(f"\n=== {len(movs)} Movimientos extraídos ===")
    for m in movs:
        cargo_s = f"-${m['cargo']:,.0f}" if m['cargo'] else ""
        abono_s = f"+${m['abono']:,.0f}" if m['abono'] else ""
        print(f"  {m['fecha']} | {m['descripcion'][:40]:40s} | {cargo_s:>12s} {abono_s:>12s} | Saldo: ${m['saldo']:,.0f}")
