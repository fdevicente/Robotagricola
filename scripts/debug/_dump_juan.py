import sys, json, re
sys.stdout.reconfigure(encoding="utf-8")
from datetime import datetime, timedelta, timezone

with open(r"C:\Users\Windows\AppData\Local\Temp\claude\juan_msgs.json", encoding="utf-8") as f:
    msgs = json.load(f)

CHILE = timezone(timedelta(hours=-4))


def parse_fecha(s):
    # "09.06.2026 22:54:39 UTC+08:00"
    m = re.match(r"(\d{2})\.(\d{2})\.(\d{4}) (\d{2}):(\d{2}):(\d{2}) UTC([+-]\d{2}):(\d{2})", s or "")
    if not m:
        return None
    d, mo, y, hh, mm, ss, oh, om = m.groups()
    off = timezone(timedelta(hours=int(oh), minutes=int(om) * (1 if int(oh) >= 0 else -1)))
    dt = datetime(int(y), int(mo), int(d), int(hh), int(mm), int(ss), tzinfo=off)
    return dt.astimezone(CHILE)


juan = []
for m in msgs:
    if m["from"] != "Juan Parada":
        continue
    dt = parse_fecha(m["date"])
    juan.append({"dt": dt, "text": m["text"], "media": m["media"]})

juan.sort(key=lambda x: x["dt"] or datetime(1900, 1, 1, tzinfo=CHILE))
print(f"Mensajes de Juan: {len(juan)}")
if juan:
    print(f"Rango (hora Chile): {juan[0]['dt']:%Y-%m-%d %H:%M} → {juan[-1]['dt']:%Y-%m-%d %H:%M}\n")

solo_media = 0
for m in juan:
    f = m["dt"].strftime("%Y-%m-%d %a %H:%M") if m["dt"] else "?"
    t = (m["text"] or "").replace("\n", " | ")
    if not t and m["media"]:
        solo_media += 1
        continue
    print(f"[{f}] {t}")

print(f"\n(mensajes solo con foto/archivo, sin texto: {solo_media})")
