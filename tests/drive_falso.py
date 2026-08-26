# -*- coding: utf-8 -*-
"""Drive falso para las pruebas. Nunca toca la red."""


class DriveFalso:
    def __init__(self, cuota_usada=0, cuota_total=15 * 1024 ** 3):
        self.archivos = {}      # file_id -> {nombre, carpeta_id, bytes}
        self.carpetas = {"raiz": None}
        self._n = 0
        self.fallar_con = None  # poner una excepción para simular caídas
        self._cuota = (cuota_usada, cuota_total)

    def _id(self, pre):
        self._n += 1
        return "%s-%d" % (pre, self._n)

    def subir(self, ruta_local, carpeta_id, nombre):
        if self.fallar_con:
            raise self.fallar_con
        fid = self._id("file")
        self.archivos[fid] = {"nombre": nombre, "carpeta_id": carpeta_id,
                              "ruta_origen": ruta_local}
        return fid

    def crear_carpeta(self, nombre, padre_id):
        cid = self._id("dir")
        self.carpetas[cid] = {"nombre": nombre, "padre": padre_id}
        return cid

    def buscar_carpeta(self, nombre, padre_id):
        for cid, c in self.carpetas.items():
            if c and c.get("nombre") == nombre and c.get("padre") == padre_id:
                return cid
        return None

    def buscar_archivo(self, nombre, carpeta_id):
        for fid, a in self.archivos.items():
            if a["nombre"] == nombre and a["carpeta_id"] == carpeta_id:
                return fid
        return None

    def listar(self, carpeta_id):
        return [{"id": fid, "nombre": a["nombre"]}
                for fid, a in self.archivos.items()
                if a["carpeta_id"] == carpeta_id]

    def mover(self, file_id, carpeta_destino_id):
        self.archivos[file_id]["carpeta_id"] = carpeta_destino_id

    def cuota(self):
        usado, total = self._cuota
        return {"usado": usado, "total": total}

    def descargar(self, file_id, ruta_local):
        with open(ruta_local, "w", encoding="utf-8") as fh:
            fh.write("contenido falso")
        return ruta_local
