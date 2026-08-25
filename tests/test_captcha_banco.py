"""El scraper debe DETECTAR el CAPTCHA y cortar con un mensaje útil.

El bot no resuelve ni elude la verificación anti-robot: está puesta justamente
para impedir el acceso automático. Lo correcto es avisar y usar la carga manual
de cartola, que ya existe (`modules/banco_import.py`).
"""
import pytest

from scotiabank_scraper import (COOKIES_ANTIBOT, SEL_CAPTCHA, TEXTO_CAPTCHA,
                                 CaptchaRequerido, _abortar_por_captcha,
                                 _antibot_presente, _esperar_boton_habilitado,
                                 _hay_captcha)


class _Loc:
    def __init__(self, n, habilitado=True):
        self._n = n
        self._hab = habilitado
    @property
    def first(self): return self
    def count(self): return self._n
    def is_enabled(self): return self._hab


class _Ctx:
    def __init__(self, cookies): self._c = cookies
    def cookies(self): return [{"name": n} for n in self._c]


class _Page:
    """Página falsa: `presentes` son los selectores que existen."""
    def __init__(self, presentes=(), texto="", cookies=(), boton_habilitado=True):
        self.presentes = set(presentes)
        self.texto = texto
        self.context = _Ctx(cookies)
        self._boton_hab = boton_habilitado

    def locator(self, sel):
        return _Loc(1 if sel in self.presentes else 0, self._boton_hab)

    def inner_text(self, _sel):
        return self.texto

    def wait_for_timeout(self, _ms):
        pass


def test_detecta_recaptcha_por_iframe():
    assert _hay_captcha(_Page(presentes=["iframe[src*='recaptcha']"]))


def test_detecta_turnstile_de_cloudflare():
    assert _hay_captcha(_Page(presentes=["div.cf-turnstile"]))


@pytest.mark.parametrize("frase", TEXTO_CAPTCHA)
def test_detecta_por_texto_visible(frase):
    assert _hay_captcha(_Page(texto=f"Ingresa tus datos. {frase.upper()}"))


def test_login_normal_no_se_confunde_con_captcha():
    """Un formulario de login corriente NO debe dar falso positivo."""
    pagina = _Page(texto="Ingresa tu RUT de empresa, RUT de usuario y clave")
    assert not _hay_captcha(pagina)


def test_no_revienta_si_no_puede_leer_el_body():
    class Rota(_Page):
        def inner_text(self, _sel): raise RuntimeError("detached")
    assert not _hay_captcha(Rota())


def test_abortar_explica_la_via_manual():
    with pytest.raises(CaptchaRequerido) as exc:
        _abortar_por_captcha()
    mensaje = str(exc.value).lower()
    assert "cartola" in mensaje          # dice qué hacer
    assert "no la resuelve" in mensaje   # y deja claro que no la elude


def test_captcha_es_runtime_error():
    """Los llamadores viejos que capturan RuntimeError siguen funcionando."""
    assert issubclass(CaptchaRequerido, RuntimeError)


def test_hay_selectores_de_los_proveedores_principales():
    juntos = " ".join(SEL_CAPTCHA)
    for proveedor in ("recaptcha", "hcaptcha", "turnstile"):
        assert proveedor in juntos


# ── Detección de bots sin CAPTCHA visible ───────────────────────────────
# Scotiabank montó Akamai Bot Manager en ago-2026: no hay widget en pantalla,
# el síntoma es que el botón "Ingresar" queda deshabilitado para siempre.

def test_detecta_akamai_por_cookies():
    """Lo que trae el portal real: _abck / bm_sz / bmuid."""
    assert _antibot_presente(_Page(cookies=["_abck", "bm_sz", "bmuid", "_ga"]))


@pytest.mark.parametrize("cookie", sorted(COOKIES_ANTIBOT))
def test_detecta_cada_proveedor_antibot(cookie):
    assert _antibot_presente(_Page(cookies=[cookie]))


def test_cookies_normales_no_son_antibot():
    assert not _antibot_presente(_Page(cookies=["_ga", "_fbp", "UUID", "brand"]))


def test_boton_habilitado_devuelve_true_rapido():
    assert _esperar_boton_habilitado(_Page(boton_habilitado=True), "button")


def test_boton_deshabilitado_agota_el_plazo():
    """Sin esto el scraper reintentaba el click 30 s y moría con un timeout
    ilegible, en vez de decir que hay un verificador."""
    assert not _esperar_boton_habilitado(
        _Page(boton_habilitado=False), "button", timeout=10)


def test_sin_contexto_de_cookies_no_revienta():
    class Rota(_Page):
        @property
        def context(self): raise RuntimeError("closed")
        @context.setter
        def context(self, _v): pass
    assert not _antibot_presente(Rota())
