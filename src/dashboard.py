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
    get_facturas_detalle, get_movimientos_banco

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


if __name__ == "__main__":
    port = int(os.environ.get("DASHBOARD_PORT", 5000))
    print(f"\n{'='*50}")
    print(f"  Dashboard Agrícola Santa Elisa")
    print(f"  http://localhost:{port}")
    print(f"{'='*50}\n")
    app.run(host="0.0.0.0", port=port, debug=True)
