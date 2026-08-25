#!/usr/bin/env python3
"""Desarma el flujo proyectado para TEMP 26/27 (jun-2026 → may-2027):
- Muestra para cada categoría y mes el monto histórico de base
- Factor de escalamiento aplicado
- Ajustes manuales
- Total proyectado final
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
from datetime import date
from collections import defaultdict
from openpyxl import load_workbook
from config import EXCEL_PATH
from modules.cash_flow.projector import (
    load_historical_egresos, load_ajustes_manuales, load_hectareas,
    load_expected_ingresos, compute_factor_hc, EXCLUIR_PROYECCION,
)

# Temporada 26/27 (corte junio) = jun-2026 → may-2027
MESES = [
    (2026, 6), (2026, 7), (2026, 8), (2026, 9), (2026, 10), (2026, 11),
    (2026, 12), (2027, 1), (2027, 2), (2027, 3), (2027, 4), (2027, 5),
]
LABELS = [f"{['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][m]}-{str(y)[-2:]}" for y, m in MESES]

print("=" * 100)
print("DESARMADO DEL FLUJO PROYECTADO TEMP 26/27 (jun-2026 → may-2027)")
print("=" * 100)
print()
print("MÉTODO DE PROYECCIÓN:")
print(" - Toma egresos REALES de TEMP 25/26 (jun-2025 a may-2026) como base.")
print(" - Para cada mes m del histórico, proyecta al mismo mes del año siguiente.")
print(" - Aplica factor de escalamiento por HECTAREAS (cultivo × año).")
print(" - Suma ajustes manuales (helicoptero anulado, replante avellanos, etc).")
print(" - Excluye categoría MANTENIMIENTO HELICOPTERO (operación descontinuada).")
print()

# Cargar datos
historicos = load_historical_egresos()
ajustes = load_ajustes_manuales()
hc = load_hectareas()

print("HECTAREAS (factor de escalamiento):")
for cultivo, datos in hc.items():
    print(f"  {cultivo}: {datos}")
print()

# Mostrar factores
print("FACTORES DE ESCALAMIENTO (base 2025 → target 2026):")
for cultivo in ["NOGALES", "CEREZOS", "AVELLANOS", "GENERAL"]:
    f = compute_factor_hc(hc, cultivo, 2025, 2026)
    print(f"  {cultivo}: ×{f}")
print()

# Calcular proyectado: cat → (y,m) → monto
proj = defaultdict(lambda: defaultdict(float))
hist_used = defaultdict(lambda: defaultdict(float))  # antes de factor

base_year = 2025

for (y_h, m_h, cat, cul), monto in historicos.items():
    if y_h != base_year: continue
    if (cat or "").upper() in EXCLUIR_PROYECCION: continue
    factor = compute_factor_hc(hc, cul, base_year, base_year + 1)
    # Proyecta al mismo mes del año siguiente
    target_y = base_year + 1
    target_m = m_h
    if (target_y, target_m) in MESES:
        proj[cat][(target_y, target_m)] += monto * factor
        hist_used[cat][(target_y, target_m)] += monto

# Aplicar ajustes manuales
ajustes_aplicados = defaultdict(lambda: defaultdict(float))
for a in ajustes:
    ym = a["mes_proyectado"]
    if ym in MESES:
        if (a["categoria"] or "").upper() in EXCLUIR_PROYECCION: continue
        proj[a["categoria"]][ym] += a["monto"]
        ajustes_aplicados[a["categoria"]][ym] += a["monto"]

# Ingresos proyectados (de Cosechas)
ingresos = defaultdict(float)
for i in load_expected_ingresos():
    ym = (i["year"], i["month"])
    if ym in MESES:
        ingresos[ym] += i["monto_clp"]

# ─── Imprimir resumen por categoría ────────────────────
print()
print("=" * 100)
print(f"{'CATEGORIA':40} {'TOTAL':>15} | distribución por mes")
print("=" * 100)

# Header de meses
print(f"{'':40} {'':>15} | " + " ".join(f"{l[:6]:>8}" for l in LABELS))

cat_totals = [(c, sum(montos.values())) for c, montos in proj.items()]
cat_totals.sort(key=lambda x: -x[1])

total_general = 0
for cat, total in cat_totals:
    if total == 0: continue
    total_general += total
    print(f"\n{cat:40} ${total:>13,.0f} | " + " ".join(f"{proj[cat].get(ym, 0)/1e6:>7.2f}M" for ym in MESES))
    # Sub-detalle
    hist_total = sum(hist_used[cat].values())
    ajuste_total = sum(ajustes_aplicados[cat].values())
    if hist_total > 0:
        print(f"  ├─ Histórico TEMP 25/26 (mismo mes año pasado): ${hist_total:,.0f}")
    if ajuste_total != 0:
        print(f"  └─ Ajuste manual: ${ajuste_total:+,.0f}")

print()
print("=" * 100)
print(f"{'TOTAL EGRESOS':40} ${total_general:>13,.0f}")
print()

total_ing = sum(ingresos.values())
print(f"{'INGRESOS PROYECTADOS (Cosechas)':40} ${total_ing:>13,.0f}")
print(f"{'':40} {'':>15} | " + " ".join(f"{ingresos.get(ym,0)/1e6:>7.2f}M" for ym in MESES))
print()
print(f"{'SALDO PROYECTADO':40} ${total_ing - total_general:>+13,.0f}")
print()
print("=" * 100)
print("DETALLE DE AJUSTES MANUALES APLICADOS:")
print("=" * 100)
for a in ajustes:
    ym = a["mes_proyectado"]
    if ym in MESES:
        print(f"  {ym[0]}-{ym[1]:02d} | {a['categoria']:35} {a['cultivo']:10} ${a['monto']:>+15,.0f} | {a['razon']}")
