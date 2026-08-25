"""Colchón desde HOY hasta el cierre de temporada, con la caja de las DOS cuentas."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, "src")
from datetime import date
from modules.cash_flow.projector import get_cash_flow
from modules.cuentas import caja_total, formato

FIN = (2027, 5)
MES = ["", "ene", "feb", "mar", "abr", "may", "jun",
       "jul", "ago", "sep", "oct", "nov", "dic"]

HOY = date.today()
INI = (HOY.year, HOY.month)
caja = caja_total()
print(formato(caja).replace("*", ""), "\n")

cf = get_cash_flow(start=INI, end=FIN, saldo_inicial=caja["total"])

print(f"{'mes':9}{'ingresos':>15}{'egresos':>15}{'neto':>15}{'saldo cierre':>17}")
print("-" * 71)
peor, tot_in, tot_eg = None, 0, 0
for ym in cf["months"]:
    s = cf["saldo"][ym]
    neto = s["ingresos"] - s["egresos"]
    tot_in += s["ingresos"]; tot_eg += s["egresos"]
    alerta = " ⚠️" if s["saldo_cierre"] < 0 else ""
    print(f"{MES[ym[1]]}-{str(ym[0])[-2:]:4}{s['ingresos']:>15,.0f}"
          f"{s['egresos']:>15,.0f}{neto:>15,.0f}{s['saldo_cierre']:>17,.0f}{alerta}")
    if peor is None or s["saldo_cierre"] < peor[1]:
        peor = (ym, s["saldo_cierre"])
print("-" * 71)
print(f"{'TOTAL':9}{tot_in:>15,.0f}{tot_eg:>15,.0f}{tot_in - tot_eg:>15,.0f}")

final = cf["saldo"][cf["months"][-1]]["saldo_cierre"]
sin_caja = next((f"{MES[ym[1]]}-{str(ym[0])[-2:]}" for ym in cf["months"]
                 if cf["saldo"][ym]["saldo_cierre"] < 0), None)
print(f"\n  COLCHÓN al cierre : ${final:>15,.0f}")
print(f"  Sin caja en       : {sin_caja or 'nunca ✅'}")
