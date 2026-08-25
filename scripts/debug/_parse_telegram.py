"""Extrae mensajes de un export HTML de Telegram Desktop."""
import sys, re, html, json
sys.stdout.reconfigure(encoding="utf-8")
from html.parser import HTMLParser

PATH = r"C:\Users\Windows\Downloads\messages.html"


class TgParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.msgs = []
        self.cur = None
        self.capture = None
        self.buf = []
        self.depth = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        cls = a.get("class", "")
        if tag == "div" and "message" in cls and "default" in cls:
            self.cur = {"from": "", "date": "", "text": "", "media": ""}
        if self.cur is not None and tag == "div":
            if "from_name" in cls:
                self.capture, self.buf, self.depth = "from", [], 1
            elif "date" in cls and "details" in cls:
                self.cur["date"] = a.get("title", "")
            elif cls.strip() == "text":
                self.capture, self.buf, self.depth = "text", [], 1
            elif "media_wrap" in cls or "title bold" in cls:
                self.capture, self.buf, self.depth = "media", [], 1
            elif self.capture:
                self.depth += 1
        elif self.capture and tag in ("span", "a", "strong", "em"):
            self.depth += 1
        if tag == "br" and self.capture:
            self.buf.append("\n")

    def handle_endtag(self, tag):
        if self.capture and tag in ("div", "span", "a", "strong", "em"):
            self.depth -= 1
            if self.depth <= 0:
                val = "".join(self.buf).strip()
                if self.cur is not None and val:
                    if self.capture == "media":
                        self.cur["media"] = (self.cur["media"] + " " + val).strip()
                    else:
                        self.cur[self.capture] = val
                self.capture, self.buf = None, []

    def handle_data(self, data):
        if self.capture:
            self.buf.append(data)

    def close_msg(self):
        if self.cur and (self.cur["text"] or self.cur["media"]):
            self.msgs.append(self.cur)
        self.cur = None


with open(PATH, encoding="utf-8") as f:
    contenido = f.read()

# Separar por bloques de mensaje para no perder ninguno
bloques = re.split(r'(?=<div class="message default)', contenido)
todos = []
for b in bloques[1:]:
    p = TgParser()
    p.feed(b)
    p.close_msg()
    todos.extend(p.msgs)

# Heredar remitente en mensajes "joined" (consecutivos del mismo autor)
ultimo = ""
for m in todos:
    if m["from"]:
        ultimo = m["from"]
    else:
        m["from"] = ultimo

print(f"TOTAL MENSAJES: {len(todos)}\n")
remitentes = {}
for m in todos:
    remitentes[m["from"]] = remitentes.get(m["from"], 0) + 1
print("Remitentes:", remitentes)
print(f"Rango: {todos[0]['date'] if todos else '?'} → {todos[-1]['date'] if todos else '?'}\n")

with open(r"C:\Users\Windows\AppData\Local\Temp\claude\juan_msgs.json", "w", encoding="utf-8") as f:
    json.dump(todos, f, ensure_ascii=False, indent=1)
print("Guardado en juan_msgs.json")
