"""Prueba el extractor de maquinaria con mensajes como los escribe Juan."""
import json
import sys
from datetime import date

sys.stdout.reconfigure(encoding="utf-8")
from modules.maquinaria import detectar_maquina, extraer_odometro, maquinas_conocidas
from modules.maquinaria_extractor import extraer

HOY = date.today().strftime("%Y-%m-%d")
conocidas = [m["maquina"] for m in maquinas_conocidas()]
print("Máquinas conocidas:", ", ".join(conocidas), "\n")

LECTURAS = [
    "MF 6711 horómetro 1980",
    "el 5085 va en 3.240 horas",
    "excavadora 7250,5",
    "camioneta ssangyong 2 kilometraje 145000",
]
print("=" * 68)
print("LECTURAS SIMPLES (sin IA, todo local)")
print("=" * 68)
for t in LECTURAS:
    m = detectar_maquina(t, [{"maquina": c} for c in conocidas])
    o = extraer_odometro(t)
    print(f"  «{t}»\n      → {m or '❌ no reconocida'} = {o}")

MENSAJES = [
    "Al John Deere 5085 le cambiaron aceite y filtros el 20 de julio "
    "a las 3100 horas, lo hizo Comercial Álamos y salió 180 mil",
    "El 5085 es John Deere del 2018, patente ABCD12, es propio. "
    "El 6711 es Massey Ferguson 2015 arrendado.",
    "A la excavadora le hicieron engrase la semana pasada y el 4275 "
    "necesita neumáticos nuevos adelante",
]
print("\n" + "=" * 68)
print("MANTENCIONES Y FICHAS (con IA)")
print("=" * 68)
for t in MENSAJES:
    print(f"\n«{t[:72]}…»")
    try:
        r = extraer(t, conocidas, HOY)
    except Exception as e:
        print(f"   ❌ {e}")
        continue
    for m in r["mantenciones"]:
        est = str(m.get("estado") or "HECHA").upper()
        print(f"   {'⏳' if est == 'PENDIENTE' else '🔧'} [{est}] {m.get('maquina')} · "
              f"{m.get('tipo')} · {m.get('descripcion')}")
        print(f"      fecha={m.get('fecha')} odo={m.get('odometro')} "
              f"prov={m.get('proveedor')} costo={m.get('costo')}")
    for f in r["fichas"]:
        campos = {k: v for k, v in f.items() if k != "maquina" and v not in (None, "")}
        print(f"   🪪 {f.get('maquina')} · {json.dumps(campos, ensure_ascii=False)}")
    if not r["mantenciones"] and not r["fichas"]:
        print("   (nada extraído)")
