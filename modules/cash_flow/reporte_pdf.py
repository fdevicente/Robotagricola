"""Genera el reporte mensual en PDF usando playwright (Chromium headless).

Levanta el dashboard Flask temporalmente si no está corriendo, navega a la
ruta /reporte/<year>/<month> y exporta a PDF.
"""
import os
import time
import socket
import subprocess
from datetime import date


def _puerto_abierto(host="127.0.0.1", port=5000):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        return s.connect_ex((host, port)) == 0


def generar_reporte_pdf(year: int, month: int, output_path: str | None = None,
                         base_url: str = "http://127.0.0.1:5000") -> str:
    """Genera el PDF del reporte mensual. Devuelve la ruta del archivo."""
    from playwright.sync_api import sync_playwright

    robot_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if output_path is None:
        out_dir = os.path.join(robot_dir, "files", "reportes")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f"reporte_{year}_{month:02d}.pdf")

    # Iniciar dashboard si no está corriendo
    dashboard_proc = None
    if not _puerto_abierto():
        import sys
        py = sys.executable
        dashboard_proc = subprocess.Popen(
            [py, os.path.join(robot_dir, "src", "dashboard.py")],
            cwd=robot_dir,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            env={**os.environ, "PYTHONPATH": robot_dir},
        )
        # Esperar a que levante
        for _ in range(30):
            if _puerto_abierto():
                break
            time.sleep(0.5)

    try:
        # El dashboard exige login; el generador entra con el token interno.
        token = os.getenv("DASHBOARD_TOKEN", "")
        if not token:
            tf = os.path.join(robot_dir, ".dashboard_token")
            if os.path.exists(tf):
                with open(tf) as f:
                    token = f.read().strip()
        url = f"{base_url}/reporte/{year}/{month}"
        if token:
            url += f"?token={token}"
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(url, wait_until="networkidle", timeout=30000)
            # Esperar a que el chart se renderice
            page.wait_for_timeout(1500)
            page.pdf(
                path=output_path,
                format="A4",
                print_background=True,
                margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"},
            )
            browser.close()
    finally:
        if dashboard_proc:
            dashboard_proc.terminate()

    return output_path


if __name__ == "__main__":
    import sys
    y = int(sys.argv[1]) if len(sys.argv) > 1 else date.today().year
    m = int(sys.argv[2]) if len(sys.argv) > 2 else date.today().month
    path = generar_reporte_pdf(y, m)
    print(f"PDF generado: {path}")
