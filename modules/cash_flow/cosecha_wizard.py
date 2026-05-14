"""FSM para wizard de cosecha (cierre por cultivo)."""
import re


PROMPTS = {
    "esperando_kg_totales": "Kg totales cosechados? (numero)",
    "esperando_exportadoras": ("Exportadoras y kg. Formato: "
                                "'Nombre1 kg1, Nombre2 kg2'"),
    "esperando_precio": "Precio USD/kg de {exp}?",
    "esperando_cuotas": "Cuantas cuotas? (1-5)",
    "esperando_cuota_data": "Cuota {n}: fecha YYYY-MM-DD y monto USD",
    "esperando_liquidacion": "Liquidacion final? (si/no)",
    "esperando_liquidacion_data": "Fecha y monto USD estimado",
    "resumen": "Resumen listo. /guardar para confirmar, /cancelar para descartar.",
}


class CosechaWizard:
    def __init__(self, cultivo: str):
        self.cultivo = cultivo
        self.data = {"cultivo": cultivo, "exportadoras": []}
        self.estado = "esperando_kg_totales"
        self._exp_idx = 0
        self._cuota_idx = 0
        self.prompt = PROMPTS["esperando_kg_totales"]
        self._error = ""

    def _set_error(self, msg):
        self._error = msg
        self.prompt = f"Error ({msg}). " + PROMPTS[self.estado]

    def _advance(self, nuevo_estado, **fmt):
        self.estado = nuevo_estado
        self._error = ""
        self.prompt = PROMPTS[nuevo_estado].format(**fmt)

    def responder(self, texto: str):
        texto = (texto or "").strip()

        if self.estado == "esperando_kg_totales":
            try:
                self.data["kg_total"] = int(float(texto))
                self._advance("esperando_exportadoras")
            except ValueError:
                self._set_error("numero invalido")

        elif self.estado == "esperando_exportadoras":
            exps = []
            for chunk in texto.split(","):
                m = re.match(r"(.+?)\s+(\d+)\s*$", chunk.strip())
                if m:
                    exps.append({"nombre": m.group(1).strip(),
                                 "kg": int(m.group(2))})
            if not exps:
                self._set_error("formato invalido")
                return
            self.data["exportadoras"] = exps
            self._exp_idx = 0
            self._advance("esperando_precio",
                          exp=exps[0]["nombre"])

        elif self.estado == "esperando_precio":
            try:
                price = float(texto)
            except ValueError:
                self._set_error("precio invalido")
                return
            self.data["exportadoras"][self._exp_idx]["precio_usd_kg"] = price
            self.data["exportadoras"][self._exp_idx]["cuotas"] = []
            self._advance("esperando_cuotas")

        elif self.estado == "esperando_cuotas":
            try:
                n = int(texto)
                if n < 1 or n > 5:
                    raise ValueError
            except ValueError:
                self._set_error("entre 1 y 5")
                return
            self.data["exportadoras"][self._exp_idx]["n_cuotas"] = n
            self._cuota_idx = 0
            self._advance("esperando_cuota_data", n=1)

        elif self.estado == "esperando_cuota_data":
            m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d+(?:\.\d+)?)\s*$", texto)
            if not m:
                self._set_error("formato fecha + monto")
                return
            self.data["exportadoras"][self._exp_idx]["cuotas"].append({
                "fecha": m.group(1), "usd": float(m.group(2)),
            })
            self._cuota_idx += 1
            total = self.data["exportadoras"][self._exp_idx]["n_cuotas"]
            if self._cuota_idx < total:
                self._advance("esperando_cuota_data", n=self._cuota_idx + 1)
            else:
                self._exp_idx += 1
                if self._exp_idx < len(self.data["exportadoras"]):
                    self._cuota_idx = 0
                    self._advance("esperando_precio",
                                  exp=self.data["exportadoras"][self._exp_idx]["nombre"])
                else:
                    self._advance("esperando_liquidacion")

        elif self.estado == "esperando_liquidacion":
            if texto.lower() in ("si", "sí", "s", "y"):
                self._advance("esperando_liquidacion_data")
            else:
                self.data["liquidacion"] = None
                self._advance("resumen")

        elif self.estado == "esperando_liquidacion_data":
            m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d+(?:\.\d+)?)\s*$", texto)
            if not m:
                self._set_error("formato fecha + monto USD")
                return
            self.data["liquidacion"] = {"fecha": m.group(1),
                                         "usd": float(m.group(2))}
            self._advance("resumen")


def save_to_cosechas(data: dict, year: int,
                       excel_path: str | None = None) -> int:
    """Escribe filas de wizard en Master.Cosechas. Devuelve # filas agregadas."""
    from openpyxl import load_workbook
    from config import EXCEL_PATH
    from excel_manager import COSECHAS_SHEET, _save_wb
    excel_path = excel_path or EXCEL_PATH

    wb = load_workbook(excel_path)
    ws = wb[COSECHAS_SHEET]
    added = 0
    has_liq = bool(data.get("liquidacion"))
    for exp in data["exportadoras"]:
        n_cuotas_total = exp.get("n_cuotas", len(exp.get("cuotas", [])))
        if has_liq:
            n_cuotas_total += 1
        for i, cuota in enumerate(exp.get("cuotas", []), start=1):
            ws.append([
                year, data["cultivo"], data["kg_total"], exp["nombre"],
                exp["kg"], exp["precio_usd_kg"],
                n_cuotas_total, i, cuota["fecha"], cuota["usd"],
                "adelanto", "esperado",
                None, None, "", "",
            ])
            added += 1

    if has_liq:
        liq = data["liquidacion"]
        exp = data["exportadoras"][0]
        n_total = exp.get("n_cuotas", 0) + 1
        ws.append([
            year, data["cultivo"], data["kg_total"], exp["nombre"],
            exp["kg"], exp["precio_usd_kg"], n_total, n_total,
            liq["fecha"], liq["usd"], "liquidacion final", "esperado",
            None, None, "", "",
        ])
        added += 1

    _save_wb(wb, excel_path)
    wb.close()
    return added
