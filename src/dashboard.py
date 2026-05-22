"""
dashboard.py — Servidor web Flask para el dashboard de Agrícola Santa Elisa.
Ejecutar: python src/dashboard.py
"""
import os
import sys
import json
import logging
from datetime import date

# Agregar directorio padre al path para importar config
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, render_template, jsonify, request
from dashboard_data import get_dashboard_data, get_facturas_summary, get_comparacion_anual, \
    get_costos_por_cultivo, get_exportaciones, get_caja_chica, get_cuenta_banco, \
    get_inventario_resumen, get_personal_resumen, get_tareas_resumen, \
    get_facturas_detalle, get_movimientos_banco, \
    get_banco_revisar, update_banco_categoria, BANCO_CATEGORIAS_VALIDAS
from modules.cash_flow.projector import get_cash_flow
from modules.cash_flow.replante import afford_check

_MESES_LBL = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _month_label(y, m):
    return f"{_MESES_LBL[m]}-{str(y)[-2:]}"

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__,
            template_folder=os.path.join(os.path.dirname(__file__), "templates"),
            static_folder=os.path.join(os.path.dirname(__file__), "static"))


class DateEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, date):
            return obj.isoformat()
        return super().default(obj)


app.json_encoder = DateEncoder


@app.route("/")
def index():
    data = get_dashboard_data()
    return render_template("dashboard.html", data=data)


@app.route("/api/dashboard")
def api_dashboard():
    return jsonify(get_dashboard_data())


@app.route("/api/facturas")
def api_facturas():
    return jsonify(get_facturas_summary())


@app.route("/api/costos-cultivo")
def api_costos():
    return jsonify(get_costos_por_cultivo())


@app.route("/api/comparacion-anual")
def api_anual():
    return jsonify(get_comparacion_anual())


@app.route("/api/exportaciones")
def api_exports():
    return jsonify(get_exportaciones())


@app.route("/api/caja-chica")
def api_caja():
    return jsonify(get_caja_chica())


@app.route("/api/banco")
def api_banco():
    return jsonify(get_cuenta_banco())


@app.route("/api/inventario")
def api_inventario():
    return jsonify(get_inventario_resumen())


@app.route("/api/personal")
def api_personal():
    return jsonify(get_personal_resumen())


@app.route("/api/tareas")
def api_tareas():
    return jsonify(get_tareas_resumen())


@app.route("/api/facturas/detalle")
def api_facturas_detalle():
    filtro = request.args.get("filtro", "todas")
    return jsonify(get_facturas_detalle(filtro))


@app.route("/api/banco/movimientos")
def api_banco_movimientos():
    return jsonify(get_movimientos_banco())


@app.route("/cash-flow")
def cash_flow_page():
    return render_template("cash_flow.html")


@app.route("/api/cash-flow")
def api_cash_flow():
    saldo = float(request.args.get("saldo", 130_600_000))
    meses = int(request.args.get("meses", 12))
    from datetime import date as _date
    today = _date.today()
    sy, sm = today.year, today.month
    ey, em = sy, sm
    for _ in range(meses - 1):
        em += 1
        if em > 12:
            em = 1
            ey += 1
    cf = get_cash_flow(start=(sy, sm), end=(ey, em),
                       saldo_inicial=saldo)
    labels = [_month_label(y, m) for y, m in cf["months"]]
    return jsonify({
        "labels": labels,
        "saldo_inicio": [cf["saldo"][ym]["saldo_inicio"] for ym in cf["months"]],
        "ingresos": [cf["saldo"][ym]["ingresos"] for ym in cf["months"]],
        "egresos": [cf["saldo"][ym]["egresos"] for ym in cf["months"]],
        "saldo_cierre": [cf["saldo"][ym]["saldo_cierre"] for ym in cf["months"]],
        "ingresos_detail": cf["ingresos"],
    })


@app.route("/api/cash-flow/replante", methods=["POST"])
def api_replante():
    body = request.get_json() or {}
    result = afford_check(
        cultivo=body.get("cultivo", "AVELLANOS"),
        hc=float(body.get("hc", 0)),
        saldo_proyectado=float(body.get("saldo_proyectado", 0)),
        saldo_minimo=float(body.get("saldo_minimo", 0)),
        costo_por_hc=float(body.get("costo_por_hc", 5_000_000)),
    )
    return jsonify(result)


@app.route("/banco/revisar")
def banco_revisar_page():
    return render_template("banco_revisar.html",
                            categorias=BANCO_CATEGORIAS_VALIDAS)


@app.route("/api/banco/revisar")
def api_banco_revisar():
    return jsonify(get_banco_revisar())


@app.route("/api/banco/revisar/<int:fila>", methods=["POST"])
def api_banco_revisar_update(fila):
    body = request.get_json() or {}
    categoria = body.get("categoria", "")
    cultivo = body.get("cultivo", "GENERAL")
    if not categoria:
        return jsonify({"error": "categoria requerida"}), 400
    try:
        result = update_banco_categoria(fila, categoria, cultivo)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("DASHBOARD_PORT", 5000))
    print(f"\n{'='*50}")
    print(f"  Dashboard Agrícola Santa Elisa")
    print(f"  http://localhost:{port}")
    print(f"{'='*50}\n")
    app.run(host="0.0.0.0", port=port, debug=True)
