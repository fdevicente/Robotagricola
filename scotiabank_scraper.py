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

# === Selectores ===
# El banco migró el login en ago-2026 a un portal nuevo (banco.scotiabank.cl/
# mfe-login) con IDs distintos. Se listan varios candidatos por campo: se usa
# el primero que exista, así sigue funcionando si vuelven a cambiarlo.
LOGIN_URL = "https://appservtrx.scotiabank.cl/portalempresas/login"

SEL_RUT_EMPRESA = [
    "input[name='bussinessId']",                                    # portal nuevo
    "#login-business-content-card-form-input-dni-business-input",
    "#INP_COMMON_RUT_CLIENTRUT",                                    # portal viejo
]
SEL_RUT_USUARIO = [
    "input[name='userId']",
    "#login-business-content-card-form-input-dni-input",
    "#INP_COMMON_RUT_USERRUT",
]
SEL_CLAVE = [
    "input[name='pass']",
    "#login-business-content-card-form-input-password-input",
    "#INP_COMMON_PASSWORD_PASS",
]
SEL_BTN_LOGIN = [
    "button[type=submit]:has-text('Ingresar')",
    "#BTN_COMMON_LOGIN",
    "button:has-text('Ingresar')",
]


class CaptchaRequerido(RuntimeError):
    """El banco pide una verificación anti-robot que debe resolver una persona.

    El CAPTCHA existe precisamente para impedir el acceso automatizado: el bot
    NO lo resuelve ni lo elude. Cuando aparece, la vía correcta es que el dueño
    descargue la cartola del portal y se la mande al bot por Telegram
    (`modules/banco_import.py` la importa con preview y deduplicación).
    """


# Marcadores de los verificadores más usados. Se busca por proveedor y por
# texto visible, porque el banco puede cambiar de proveedor sin avisar.
SEL_CAPTCHA = [
    "iframe[src*='recaptcha']",
    "iframe[src*='hcaptcha']",
    "iframe[src*='challenges.cloudflare.com']",
    "div.g-recaptcha",
    "div.h-captcha",
    "div.cf-turnstile",
    "#captcha",
    "[id*='captcha' i]",
    "[class*='captcha' i]",
]
TEXTO_CAPTCHA = [
    "no soy un robot",
    "not a robot",
    "verificación de seguridad",
    "verifica que eres humano",
    "completa la verificación",
]

# Cookies que delatan un sistema de detección de bots. Scotiabank montó
# Akamai Bot Manager en ago-2026 (_abck / bm_sz / bm_sv / bmuid) junto con
# huella digital de dispositivo (iframes cd__fontDetectionFrame): el botón
# "Ingresar" queda deshabilitado hasta que el verificador aprueba la sesión.
COOKIES_ANTIBOT = {"_abck", "bm_sz", "bm_sv", "bmuid",       # Akamai
                   "datadome",                                # DataDome
                   "cf_clearance", "__cf_bm",                 # Cloudflare
                   "_px", "_pxhd"}                            # PerimeterX


def _antibot_presente(page) -> bool:
    """True si la sesión está bajo un sistema de detección de bots."""
    try:
        nombres = {c.get("name", "") for c in page.context.cookies()}
    except Exception:
        return False
    encontradas = nombres & COOKIES_ANTIBOT
    if encontradas:
        logger.warning(f"Detección de bots activa (cookies: {sorted(encontradas)})")
        return True
    return False


def _hay_captcha(page) -> bool:
    """True si la página muestra un verificador anti-robot."""
    for sel in SEL_CAPTCHA:
        try:
            if page.locator(sel).first.count() > 0:
                logger.warning(f"CAPTCHA detectado por selector: {sel}")
                return True
        except Exception:
            continue
    try:
        cuerpo = (page.inner_text("body") or "").lower()
    except Exception:
        return False
    for frase in TEXTO_CAPTCHA:
        if frase in cuerpo:
            logger.warning(f"CAPTCHA detectado por texto: {frase!r}")
            return True
    return False


def _abortar_por_captcha(motivo=""):
    raise CaptchaRequerido(
        "El banco agregó una verificación anti-robot en el login"
        + (f" ({motivo})" if motivo else "") + ". "
        "El bot no la resuelve —está puesta justamente para impedir el acceso "
        "automático y forzarla arriesga que bloqueen la cuenta.\n\n"
        "Para actualizar el banco: entra tú al portal, descarga la cartola y "
        "mándamela por Telegram; la importo con preview y sin duplicar.")


def _esperar_boton_habilitado(page, selector, timeout=15_000) -> bool:
    """El botón 'Ingresar' se habilita solo cuando el verificador aprueba.

    Si sigue deshabilitado con el formulario completo, es la señal de que el
    sistema anti-bot marcó la sesión.
    """
    fin = time.time() + timeout / 1000
    while time.time() < fin:
        try:
            if page.locator(selector).first.is_enabled():
                return True
        except Exception:
            pass
        page.wait_for_timeout(500)
    return False


def _primer_selector(page, candidatos, timeout=45_000, descripcion=""):
    """Espera a que aparezca alguno de los selectores y devuelve el que sirvió."""
    if isinstance(candidatos, str):
        candidatos = [candidatos]
    fin = time.time() + timeout / 1000
    while time.time() < fin:
        for sel in candidatos:
            try:
                loc = page.locator(sel).first
                if loc.count() > 0 and loc.is_visible():
                    return sel
            except Exception:
                continue
        page.wait_for_timeout(500)
    raise RuntimeError(
        f"No encontré el campo {descripcion or candidatos[0]} en el portal del banco. "
        f"Puede que hayan vuelto a cambiar la página de login.")


_MESES = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
          "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12}


def _parse_fecha(texto: str):
    """'24 Jul, 2026' → date(2026, 7, 24). Devuelve None si no se entiende.

    El portal muestra la fecha con el mes abreviado en español; hay que
    normalizarla a date para que calce con el resto del Master (si se guarda
    como texto, la deduplicación falla y se reinsertan los movimientos).
    """
    from datetime import date as _date
    t = (texto or "").strip()
    m = re.search(r"(\d{1,2})\s*[-/ ]\s*([A-Za-zÁÉÍÓÚáéíóú]+)[,\s]+(\d{4})", t)
    if m:
        dia, mes_txt, anio = m.group(1), m.group(2)[:4].lower(), m.group(3)
        mes = _MESES.get(mes_txt) or _MESES.get(mes_txt[:3])
        if mes:
            try:
                return _date(int(anio), mes, int(dia))
            except ValueError:
                return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(t[:10], fmt).date()
        except ValueError:
            continue
    return None


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


_LAUNCH_ARGS = ["--no-sandbox", "--disable-blink-features=AutomationControlled"]


def _lanzar_chromium(p):
    """Lanza Chromium tolerando instalaciones incompletas de Playwright.

    Playwright ≥1.49 usa por defecto el binario 'chrome-headless-shell'. Si ese
    binario falta (típico tras actualizar playwright sin re-instalar navegadores),
    se reintenta con el Chromium completo y, como último recurso, con el Chrome
    del sistema. Así /banco no se cae por un navegador faltante.
    """
    intentos = [
        ("headless shell (default)", dict(headless=True, args=_LAUNCH_ARGS)),
        ("chromium completo",        dict(headless=True, args=_LAUNCH_ARGS,
                                          channel="chromium")),
        ("chrome del sistema",       dict(headless=True, args=_LAUNCH_ARGS,
                                          channel="chrome")),
    ]
    errores = []
    for nombre, kwargs in intentos:
        try:
            browser = p.chromium.launch(**kwargs)
            if errores:  # solo avisar si hubo que recurrir a un fallback
                logger.warning(f"Chromium lanzado con fallback: {nombre}")
            return browser
        except Exception as e:
            errores.append(f"{nombre}: {str(e)[:120]}")
            logger.warning(f"Launch falló con {nombre}: {str(e)[:150]}")
    raise RuntimeError(
        "No pude abrir ningún navegador para conectarme al banco.\n"
        "Solución: ejecuta en una terminal:\n"
        "  python -m playwright install chromium\n\n"
        "Detalle: " + " | ".join(errores)
    )


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
        browser = _lanzar_chromium(p)
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
            page.goto(LOGIN_URL, timeout=60_000, wait_until="domcontentloaded")
            time.sleep(3)   # el verificador se monta después que el HTML

            # Si el banco pide verificación anti-robot no hay nada que intentar:
            # se corta acá con un mensaje claro en vez de un timeout confuso.
            if _hay_captcha(page):
                _abortar_por_captcha()

            # El portal nuevo redirige y monta el formulario por JS: hay que
            # esperar a que el campo exista, no basta con que cargue la página.
            try:
                sel_emp = _primer_selector(page, SEL_RUT_EMPRESA, 45_000, "RUT empresa")
                sel_usr = _primer_selector(page, SEL_RUT_USUARIO, 15_000, "RUT usuario")
                sel_cla = _primer_selector(page, SEL_CLAVE, 15_000, "clave")
            except RuntimeError:
                # El CAPTCHA puede aparecer recién ahora y tapar el formulario
                if _hay_captcha(page):
                    _abortar_por_captcha()
                raise
            logger.info(f"Formulario detectado ({sel_emp})")

            logger.info("Ingresando credenciales...")
            page.fill(sel_emp, RUT_EMPRESA)
            page.fill(sel_usr, RUT_USUARIO)
            page.fill(sel_cla, CLAVE)

            sel_btn = _primer_selector(page, SEL_BTN_LOGIN, 10_000, "botón Ingresar")
            # Con el formulario lleno el botón debería habilitarse solo. Si no
            # lo hace, es el verificador anti-robot reteniendo la sesión: se
            # corta acá en vez de reintentar el click durante 30 s.
            if not _esperar_boton_habilitado(page, sel_btn):
                if _antibot_presente(page):
                    _abortar_por_captcha("el botón Ingresar quedó bloqueado "
                                          "por el sistema anti-bot del portal")
                raise RuntimeError(
                    "El botón 'Ingresar' nunca se habilitó con el formulario "
                    "completo. El banco pudo cambiar la validación del login.")
            page.click(sel_btn)
            page.wait_for_load_state("domcontentloaded", timeout=60_000)
            time.sleep(5)

            # Algunos bancos muestran el desafío recién al enviar el formulario
            if _hay_captcha(page):
                _abortar_por_captcha()

            if "error" in page.title().lower() or page.locator("text=Clave incorrecta").count() > 0:
                raise RuntimeError("Login fallido: credenciales incorrectas o error en el banco.")

            logger.info("Login exitoso.")
            _close_modals(page)
            time.sleep(2)

            # Navegar a Cuentas (SPA Vue.js — usar click por coordenadas)
            if not _navigate_to_cuentas(page):
                raise RuntimeError(
                    "No encontré el menú 'Cuentas' en el portal. "
                    "Puede que el banco haya cambiado la navegación.")

            page.wait_for_load_state("domcontentloaded", timeout=30_000)

            # La tabla la renderiza la SPA DESPUÉS de cargar la página: hay que
            # esperarla explícitamente. Con una espera fija se leían 0 filas y
            # el bot informaba "sin movimientos nuevos" sin ningún error.
            logger.info("Esperando que cargue la tabla de movimientos...")
            try:
                page.wait_for_selector("table tbody tr", timeout=45_000)
            except Exception:
                raise RuntimeError(
                    "El portal cargó pero la tabla de movimientos nunca apareció "
                    "(45 s). Revisa si el banco pide un paso extra.")
            time.sleep(2)   # dejar que termine de pintar todas las filas

            # Extraer tabla de movimientos
            # Formato: Fecha | Descripción | Nº OPERACIÓN | MONTO | SALDO
            logger.info("Extrayendo tabla de movimientos...")
            rows = page.locator("table tbody tr").all()
            if not rows:
                rows = page.locator("table tr").all()
                if rows and len(rows) > 1:
                    rows = rows[1:]

            if not rows:
                raise RuntimeError("La tabla de movimientos quedó vacía.")

            for row in rows:
                cells = row.locator("td").all_inner_texts()
                if len(cells) < 3:
                    continue

                fecha_raw = cells[0].strip()
                desc = cells[1].strip()

                # Limpiar fecha: quitar "Cargo", "Abono", "Created with Sketch."
                fecha_str = re.sub(r'(Cargo|Abono|Created with Sketch\.?)', '', fecha_raw).strip()
                # Normalizar a date real (si no, la deduplicación falla)
                fecha_val = _parse_fecha(fecha_str) or fecha_str

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
                    "fecha": fecha_val,
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
