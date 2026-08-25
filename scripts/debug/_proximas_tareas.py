"""¿Qué tarea corre y cuándo, tras el reinicio? (¿revive alguna atrasada?)"""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from datetime import datetime
import pytz
from apscheduler.triggers.cron import CronTrigger

tz = pytz.timezone("America/Santiago")
ahora = datetime.now(tz)

TAREAS = [
    ("reporte_mensual  (mes anterior)", CronTrigger(day=1, hour=8, minute=0, timezone=tz)),
    ("vacaciones_mensuales",            CronTrigger(day=1, hour=7, minute=0, timezone=tz)),
    ("resumen_semanal  (lunes)",        CronTrigger(day_of_week=0, hour=8, minute=0, timezone=tz)),
    ("heartbeat diario",                CronTrigger(hour=20, minute=0, timezone=tz)),
]

print(f"Ahora: {ahora:%Y-%m-%d %H:%M} (Chile)\n")
print(f"{'Tarea':36} {'Próxima ejecución':22} Cubre")
print("-" * 78)
for nombre, trig in TAREAS:
    prox = trig.get_next_fire_time(None, ahora)
    cubre = ""
    if "reporte_mensual" in nombre:
        m = prox.month - 1 or 12
        y = prox.year if prox.month > 1 else prox.year - 1
        cubre = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                 "agosto", "septiembre", "octubre", "noviembre", "diciembre"][m] + f" {y}"
    print(f"{nombre:36} {prox:%Y-%m-%d %H:%M} ({prox:%a})   {cubre}")

print("\nJobQueue de python-telegram-bot usa MemoryJobStore (sin persistencia):")
print("  → al reiniciar, cada tarea se reprograma desde AHORA hacia adelante.")
print("  → las ejecuciones perdidas con el PC apagado NO se recuperan solas.")
