"""Lecturas de horómetro que mandó Juan por Telegram el 10-ago-2026.

Fecha tomada del display del remecedor Moresil, que la muestra: 10/08/2026 17:11.

Dos máquinas nuevas que no estaban en el sistema: el remecedor Moresil y la
grúa horquilla Toyota. Y las dos camionetas SsangYong Gran Musso se separan
por dueño, porque "SSANGYONG 1 / 2" no decía cuál era cuál.

Uso:  python scripts/carga/lecturas_maquinaria_10ago.py [--aplicar]
"""
import shutil
import sys
from datetime import date, datetime

sys.stdout.reconfigure(encoding="utf-8")

from config import EXCEL_PATH
from modules.maquinaria import guardar_ficha, maquinas_conocidas, unidad_de

APLICAR = "--aplicar" in sys.argv
FECHA = date(2026, 8, 10)

# (máquina, lectura, nota)
LECTURAS = [
    ("TRACTOR MASSEY FERGUSON 4275", 3452.5, ""),
    ("TRACTOR MASSEY FERGUSON 4292", 5231.0, ""),
    ("TRACTOR MASSEY FERGUSON 6711", 2033.5, ""),
    ("TRACTOR JOHN DEERE 5085", 3261.0, ""),
    ("TRACTOR JOHN DEERE 5425", 3200.0, "⚠️ horómetro en mal estado — valor estimado"),
    ("CAMIONETA SSANGYONG GRAN MUSSO (JUAN)", 126157, ""),
    ("CAMIONETA SSANGYONG GRAN MUSSO (FELIX)", 106200, ""),
    ("REMECEDOR MORESIL", 483.2, "Display marca 00483:12 · motor 1600 RPM"),
    ("GRUA HORQUILLA TOYOTA", 8865.8, ""),
]

# Fichas de lo que sabemos por las fotos
FICHAS = [
    {"maquina": "REMECEDOR MORESIL", "tipo": "OTRO", "marca": "Moresil",
     "notas": "Remecedor. Display digital con fecha y RPM de motor."},
    {"maquina": "GRUA HORQUILLA TOYOTA", "tipo": "OTRO", "marca": "Toyota",
     "notas": "Grúa horquilla."},
    {"maquina": "CAMIONETA SSANGYONG GRAN MUSSO (JUAN)", "tipo": "CAMIONETA",
     "marca": "SsangYong", "modelo": "Gran Musso",
     "notas": "La que usa Juan Parada."},
    {"maquina": "CAMIONETA SSANGYONG GRAN MUSSO (FELIX)", "tipo": "CAMIONETA",
     "marca": "SsangYong", "modelo": "Gran Musso",
     "notas": "La del dueño."},
    {"maquina": "TRACTOR JOHN DEERE 5425", "tipo": "TRACTOR", "marca": "John Deere",
     "modelo": "5425", "estado": "Horómetro en mal estado — hay que repararlo"},
]

previas = {m["maquina"]: m for m in maquinas_conocidas()}

print("=" * 78)
print(f"LECTURAS DEL {FECHA:%d-%m-%Y}")
print("=" * 78)
print(f"{'máquina':40}{'lectura':>12}{'anterior':>11}{'diferencia':>12}")
print("-" * 78)
for nombre, valor, nota in LECTURAS:
    u = unidad_de(nombre)
    p = previas.get(nombre)
    ant = p["ultimo_odometro"] if p and p["ultimo_odometro"] is not None else None
    if ant is None:
        dif = "🆕 primera"
    else:
        d = valor - float(ant)
        dif = f"{d:+,.1f} {u}" + ("  ⚠️ BAJÓ" if d < 0 else "")
    print(f"{nombre[:40]:40}{valor:>10,.1f}{u:>2}"
          f"{(f'{float(ant):,.1f}' if ant is not None else '—'):>11}{dif:>12}")
    if nota:
        print(f"{'':40}{nota}")

if not APLICAR:
    print("\n(simulación — nada se escribió; agrega --aplicar)")
    sys.exit(0)

resp = shutil.copy2(EXCEL_PATH,
                    EXCEL_PATH.replace(".xlsx", f"_bak_{datetime.now():%Y%m%d_%H%M%S}.xlsx"))
print(f"\nRespaldo: {resp}\n")

from bitacora_manager import registrar_bitacora_estructurada

for nombre, valor, nota in LECTURAS:
    campos = {
        "fecha": FECHA.strftime("%Y-%m-%d"), "tipo": "MAQUINARIA",
        "actividad": "Lectura de horómetro", "cultivo": "GENERAL", "sector": "",
        "jornadas_hombre": None, "trabajadores": [], "insumo": "",
        "cantidad": None, "unidad": "", "maquina": nombre, "odometro": valor,
        "superficie_ha": None,
        "texto_original": f"Lectura enviada por Juan el 10-ago-2026. {nota}".strip(),
    }
    res = registrar_bitacora_estructurada(campos, "Juan Parada")
    extra = ""
    if isinstance(res, dict):
        if res.get("es_baseline"):
            extra = " (primera lectura)"
        elif res.get("horas_dia") is not None:
            extra = f" → {res['horas_dia']:g} {unidad_de(nombre)} desde la anterior"
    print(f"  ✅ {nombre[:44]:44} {valor:>10,.1f}{extra}")

print()
for f in FICHAS:
    guardar_ficha(f)
    print(f"  🪪 ficha: {f['maquina']}")

print("\n✅ Listo.")
