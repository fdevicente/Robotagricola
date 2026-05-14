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
