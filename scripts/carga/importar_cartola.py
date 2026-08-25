"""Importa una cartola del banco (cuenta corriente o cuenta dólar).

Detecta sola de qué cuenta es: si los saldos traen decimales y son chicos, es
la cuenta en dólares.

Uso:
    python scripts/carga/importar_cartola.py "ruta\\cartola.txt" [...]  # preview
    python scripts/carga/importar_cartola.py "ruta\\cartola.txt" --aplicar
"""
import shutil
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding="utf-8")
from config import EXCEL_PATH
from modules.banco_import import analizar_cartola, importar_cartola, parsear_cartola
from modules.cuentas import DOLAR_SHEET, caja_total, formato

APLICAR = "--aplicar" in sys.argv
RUTAS = [a for a in sys.argv[1:] if not a.startswith("--")]

if not RUTAS:
    print(__doc__)
    sys.exit(1)


def es_cuenta_dolar(ruta: str) -> bool:
    """La cuenta dólar maneja centavos y montos chicos; la CLP, enteros grandes."""
    movs = parsear_cartola(ruta)
    if not movs:
        return False
    saldos = [m["saldo"] for m in movs if m.get("saldo") is not None]
    if not saldos:
        return False
    con_decimales = sum(1 for s in saldos if abs(s - round(s)) > 0.001)
    return con_decimales > len(saldos) * 0.3 or max(abs(s) for s in saldos) < 1_000_000


if APLICAR:
    resp = shutil.copy2(EXCEL_PATH,
                        EXCEL_PATH.replace(".xlsx", f"_bak_{datetime.now():%Y%m%d_%H%M%S}.xlsx"))
    print(f"Respaldo: {resp}\n")

for ruta in RUTAS:
    dolar = es_cuenta_dolar(ruta)
    hoja = DOLAR_SHEET if dolar else None
    dec = 2 if dolar else 0
    simbolo = "US$" if dolar else "$"
    etiqueta = "Cuenta dólar (USD)" if dolar else "Cuenta corriente (CLP)"

    res = (importar_cartola(ruta, hoja, dec) if APLICAR
           else analizar_cartola(ruta, hoja, dec))

    print("=" * 64)
    print(f"{etiqueta}   ·   {ruta.split(chr(92))[-1]}")
    print("=" * 64)
    print(f"  En el archivo   : {res['total_archivo']}")
    print(f"  Ya estaban      : {res['duplicados']}")
    print(f"  NUEVOS          : {len(res['nuevos'])}")
    if res.get("ultima_fecha_master"):
        print(f"  Master llegaba a: {res['ultima_fecha_master']}")
    if res.get("saldo_archivo") is not None:
        print(f"  Saldo cartola   : {simbolo}{res['saldo_archivo']:,.2f}")
    for m in res["nuevos"]:
        signo = -m["cargo"] if m["cargo"] else m["abono"]
        print(f"     {m['fecha']}  {signo:>15,.2f}  {m['desc'][:44]}")
    if APLICAR:
        print(f"\n  ✅ {res['agregados']} agregados.\n")
    else:
        print()

if APLICAR:
    print(formato(caja_total()).replace("*", ""))
else:
    print("(simulación — nada se escribió; agrega --aplicar para importar)")
