# Documentos en Google Drive — Plan de implementación

> **Para quien ejecute esto:** SUB-SKILL REQUERIDA: usa `superpowers:subagent-driven-development` (recomendada) o `superpowers:executing-plans` para implementar tarea por tarea. Los pasos usan casillas (`- [ ]`) para ir marcando.

**Objetivo:** Archivar en Google Drive los documentos que llegan por Telegram, los respaldos del Master y la base, y una carpeta de entrada donde el dueño suelte archivos — sin que una caída de internet pueda costar un documento.

**Arquitectura:** El archivo se guarda **primero en disco local** y recién después se encola para subir a Drive. La cola se persiste en disco porque el watchdog reinicia el bot seguido. La base guarda el enlace de Drive, nunca el archivo. Todo el acceso a Drive pasa por una clase con una interfaz chica, para poder sustituirla por una falsa en las pruebas y no tocar la red nunca en los tests.

**Stack:** Python 3.11 · `google-api-python-client` + `google-auth-oauthlib` · openpyxl · pytest · python-telegram-bot 21.3

**Diseño de referencia:** `docs/superpowers/specs/2026-08-25-drive-documentos-design.md`

---

## Estructura de archivos

| Archivo | Responsabilidad |
|---|---|
| `modules/drive/auth.py` | Obtener credenciales OAuth y refrescarlas. Nada más. |
| `modules/drive/cliente.py` | Envoltura fina de la API de Drive: subir, mover, listar, buscar, crear carpeta, cuota. |
| `modules/drive/cola.py` | Cola de subidas persistida en disco, con reintentos. |
| `modules/drive/carpetas.py` | Resolver y crear la estructura de carpetas; cachear los IDs. |
| `handlers/drive_entrada.py` | Job que revisa `_Entrada/` y procesa lo que aparezca. |
| `scripts/carga/migrar_documentos_a_drive.py` | Migración única de los ~200 MB. Se corre a mano. |
| `infrastructure/backups.py` *(modificar)* | Subir respaldos a Drive + retención. |
| `handlers/facturas.py` *(modificar)* | Encolar la subida después de guardar en disco. |
| `main.py` *(modificar)* | Registrar los jobs de cola y de entrada. |
| `config.py` *(modificar)* | Rutas y parámetros de Drive. |

---

### Task 1: Dependencias y configuración

**Files:**
- Modify: `requirements.txt`
- Modify: `config.py`
- Modify: `.gitignore`

- [ ] **Step 1: Agregar las librerías**

En `requirements.txt`, al final:

```
google-api-python-client>=2.140
google-auth-oauthlib>=1.2
```

- [ ] **Step 2: Instalarlas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pip install "google-api-python-client>=2.140" "google-auth-oauthlib>=1.2"
```

Esperado: `Successfully installed ...`

- [ ] **Step 3: Agregar la configuración**

En `config.py`, después del bloque de `DROPBOX_BACKUP_PATH`:

```python
# ── Google Drive ────────────────────────────────────────────────────────────
# Credenciales OAuth de la cuenta dedicada del robot. NUNCA al repositorio.
_ROBOT_DIR = os.path.dirname(os.path.abspath(__file__))
DRIVE_CLIENT_SECRET = os.getenv("DRIVE_CLIENT_SECRET",
                                 os.path.join(_ROBOT_DIR, ".drive_client_secret.json"))
DRIVE_TOKEN_PATH = os.getenv("DRIVE_TOKEN_PATH",
                              os.path.join(_ROBOT_DIR, ".drive_token.json"))
# Carpeta raíz en Drive. Se resuelve por nombre la primera vez y se cachea.
DRIVE_RAIZ = os.getenv("DRIVE_RAIZ", "Agrícola Santa Elisa")
# Cola de subidas pendientes (se persiste: el watchdog reinicia el bot seguido)
DRIVE_COLA_PATH = os.getenv("DRIVE_COLA_PATH",
                             os.path.join(_ROBOT_DIR, "files", "drive_cola.jsonl"))
DRIVE_MAX_INTENTOS = int(os.getenv("DRIVE_MAX_INTENTOS", "5"))
# Avisar cuando el Drive pase de este porcentaje de uso
DRIVE_UMBRAL_AVISO = float(os.getenv("DRIVE_UMBRAL_AVISO", "0.80"))
```

- [ ] **Step 4: Excluir los secretos del repositorio**

En `.gitignore`, al final:

```
# Credenciales de Google Drive
.drive_client_secret.json
.drive_token.json
```

- [ ] **Step 5: Verificar que quedaron ignorados**

```bash
cd "C:/Users/Windows/Desktop/Workflow/Agricola Santa Elisa/Robot"
touch .drive_token.json && git check-ignore -v .drive_token.json && rm .drive_token.json
```

Esperado: `.gitignore:NN:.drive_token.json	.drive_token.json`

- [ ] **Step 6: Commit**

```bash
git add requirements.txt config.py .gitignore
git commit -m "Configuración de Google Drive: dependencias, rutas y secretos ignorados"
```

---

### Task 2: Cola de subidas persistida

Esta es la pieza que evita perder documentos. Va primero porque todo lo demás la usa.

**Files:**
- Create: `modules/drive/__init__.py`
- Create: `modules/drive/cola.py`
- Test: `tests/test_drive_cola.py`

- [ ] **Step 1: Escribir las pruebas que fallan**

Crear `tests/test_drive_cola.py`:

```python
# -*- coding: utf-8 -*-
"""La cola de subidas a Drive sobrevive reinicios.

El bot corre bajo watchdog y se reinicia seguido. Una cola en memoria perdería
las subidas pendientes en cada reinicio, y con ellas el vínculo entre el archivo
que ya está en disco y su lugar en Drive.
"""
import pytest

from modules.drive.cola import Cola


@pytest.fixture
def cola(tmp_path):
    return Cola(str(tmp_path / "cola.jsonl"))


def test_encolar_y_leer_pendientes(cola):
    cola.encolar("C:/docs/factura.pdf", "Facturas Recibidas/2026", "COPEVAL_123.pdf")
    p = cola.pendientes()
    assert len(p) == 1
    assert p[0]["ruta_local"] == "C:/docs/factura.pdf"
    assert p[0]["carpeta"] == "Facturas Recibidas/2026"
    assert p[0]["nombre"] == "COPEVAL_123.pdf"
    assert p[0]["intentos"] == 0


def test_la_cola_sobrevive_un_reinicio(tmp_path):
    ruta = str(tmp_path / "cola.jsonl")
    Cola(ruta).encolar("a.pdf", "Facturas Recibidas/2026", "a.pdf")
    # otro proceso, misma ruta
    assert len(Cola(ruta).pendientes()) == 1


def test_marcar_ok_saca_el_item(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    item = cola.pendientes()[0]
    cola.marcar_ok(item["id"], "drive-file-id-123")
    assert cola.pendientes() == []


def test_marcar_error_incrementa_intentos_y_lo_deja_pendiente(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    item = cola.pendientes()[0]
    cola.marcar_error(item["id"], "sin internet")
    p = cola.pendientes()
    assert len(p) == 1
    assert p[0]["intentos"] == 1
    assert p[0]["ultimo_error"] == "sin internet"


def test_tras_demasiados_intentos_deja_de_reintentar(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    item_id = cola.pendientes()[0]["id"]
    for _ in range(5):
        cola.marcar_error(item_id, "falla")
    assert cola.pendientes() == []
    rendidos = cola.rendidos()
    assert len(rendidos) == 1
    assert rendidos[0]["intentos"] == 5


def test_no_encola_dos_veces_el_mismo_archivo(cola):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    assert len(cola.pendientes()) == 1


def test_una_linea_corrupta_no_inutiliza_la_cola(cola, tmp_path):
    cola.encolar("a.pdf", "F/2026", "a.pdf")
    with open(cola.ruta, "a", encoding="utf-8") as fh:
        fh.write("{no es json}\n")
    cola.encolar("b.pdf", "F/2026", "b.pdf")
    assert len(cola.pendientes()) == 2


def test_cola_vacia_no_falla(tmp_path):
    assert Cola(str(tmp_path / "no-existe.jsonl")).pendientes() == []
```

- [ ] **Step 2: Correr las pruebas para verlas fallar**

```bash
cd "C:/Users/Windows/Desktop/Workflow/Agricola Santa Elisa/Robot"
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_cola.py -q
```

Esperado: `ModuleNotFoundError: No module named 'modules.drive'`

- [ ] **Step 3: Crear el paquete**

Crear `modules/drive/__init__.py` vacío (un archivo de cero bytes).

- [ ] **Step 4: Implementar la cola**

Crear `modules/drive/cola.py`:

```python
# -*- coding: utf-8 -*-
"""Cola de subidas a Drive, persistida en disco.

Se persiste a propósito: el bot corre bajo watchdog y se reinicia seguido, así
que una cola en memoria perdería lo pendiente en cada reinicio. El formato es
append-only (una línea JSON por evento) para que una escritura interrumpida no
corrompa lo anterior.
"""
import json
import logging
import os
import uuid
from datetime import datetime, timezone

logger = logging.getLogger(__name__)


class Cola:
    def __init__(self, ruta: str, max_intentos: int = 5):
        self.ruta = ruta
        self.max_intentos = max_intentos

    # ── lectura ────────────────────────────────────────────────────────────
    def _eventos(self) -> list[dict]:
        if not os.path.exists(self.ruta):
            return []
        out = []
        with open(self.ruta, encoding="utf-8") as fh:
            for ln in fh:
                ln = ln.strip()
                if not ln:
                    continue
                try:
                    out.append(json.loads(ln))
                except json.JSONDecodeError:
                    # una línea corrupta no puede inutilizar el resto
                    logger.warning("Línea ilegible en la cola de Drive")
        return out

    def _estado(self) -> dict:
        """Reconstruye el estado actual aplicando los eventos en orden."""
        items: dict = {}
        for e in self._eventos():
            tipo, iid = e.get("evento"), e.get("id")
            if not iid:
                continue
            if tipo == "encolado":
                items.setdefault(iid, {
                    "id": iid, "ruta_local": e["ruta_local"],
                    "carpeta": e["carpeta"], "nombre": e["nombre"],
                    "intentos": 0, "ultimo_error": "", "file_id": None,
                    "listo": False})
            elif iid in items:
                if tipo == "ok":
                    items[iid]["listo"] = True
                    items[iid]["file_id"] = e.get("file_id")
                elif tipo == "error":
                    items[iid]["intentos"] += 1
                    items[iid]["ultimo_error"] = e.get("motivo", "")
        return items

    def pendientes(self) -> list[dict]:
        return [i for i in self._estado().values()
                if not i["listo"] and i["intentos"] < self.max_intentos]

    def rendidos(self) -> list[dict]:
        """Los que agotaron los reintentos. El archivo sigue en disco."""
        return [i for i in self._estado().values()
                if not i["listo"] and i["intentos"] >= self.max_intentos]

    # ── escritura ──────────────────────────────────────────────────────────
    def _append(self, fila: dict) -> None:
        os.makedirs(os.path.dirname(self.ruta) or ".", exist_ok=True)
        fila["cuando"] = datetime.now(timezone.utc).isoformat()
        with open(self.ruta, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(fila, ensure_ascii=False) + "\n")

    def encolar(self, ruta_local: str, carpeta: str, nombre: str) -> str:
        """Encola una subida. Si ese archivo ya está pendiente, no duplica."""
        for i in self.pendientes():
            if i["ruta_local"] == ruta_local and i["carpeta"] == carpeta:
                return i["id"]
        iid = uuid.uuid4().hex
        self._append({"evento": "encolado", "id": iid, "ruta_local": ruta_local,
                      "carpeta": carpeta, "nombre": nombre})
        return iid

    def marcar_ok(self, iid: str, file_id: str) -> None:
        self._append({"evento": "ok", "id": iid, "file_id": file_id})

    def marcar_error(self, iid: str, motivo: str) -> None:
        self._append({"evento": "error", "id": iid, "motivo": str(motivo)[:200]})
```

- [ ] **Step 5: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_cola.py -q
```

Esperado: `8 passed`

- [ ] **Step 6: Commit**

```bash
git add modules/drive/__init__.py modules/drive/cola.py tests/test_drive_cola.py
git commit -m "Cola de subidas a Drive persistida en disco, sobrevive reinicios"
```

---

### Task 3: Cliente de Drive y su doble para pruebas

**Files:**
- Create: `modules/drive/cliente.py`
- Create: `tests/drive_falso.py`
- Test: `tests/test_drive_cliente.py`

- [ ] **Step 1: Escribir el doble de pruebas**

Crear `tests/drive_falso.py`:

```python
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
```

- [ ] **Step 2: Escribir la prueba del contrato**

Crear `tests/test_drive_cliente.py`:

```python
# -*- coding: utf-8 -*-
"""El cliente real y el falso exponen la MISMA interfaz.

Si divergen, las pruebas pasan contra el falso y la producción se rompe.
"""
import inspect

from modules.drive.cliente import DriveCliente
from tests.drive_falso import DriveFalso

METODOS = ["subir", "crear_carpeta", "buscar_carpeta", "buscar_archivo",
           "listar", "mover", "cuota"]


def test_el_falso_implementa_todos_los_metodos_del_real():
    for m in METODOS:
        assert hasattr(DriveCliente, m), "falta %s en el real" % m
        assert hasattr(DriveFalso, m), "falta %s en el falso" % m


def test_las_firmas_coinciden():
    for m in METODOS:
        real = inspect.signature(getattr(DriveCliente, m))
        falso = inspect.signature(getattr(DriveFalso, m))
        assert list(real.parameters) == list(falso.parameters), \
            "%s: %s vs %s" % (m, list(real.parameters), list(falso.parameters))


def test_el_falso_sube_y_encuentra():
    d = DriveFalso()
    cid = d.crear_carpeta("Facturas Recibidas", "raiz")
    fid = d.subir("C:/x/COPEVAL_1.pdf", cid, "COPEVAL_1.pdf")
    assert d.buscar_archivo("COPEVAL_1.pdf", cid) == fid
    assert d.listar(cid) == [{"id": fid, "nombre": "COPEVAL_1.pdf"}]


def test_el_falso_puede_simular_una_caida():
    d = DriveFalso()
    d.fallar_con = ConnectionError("sin internet")
    try:
        d.subir("a.pdf", "dir-1", "a.pdf")
        assert False, "debió lanzar"
    except ConnectionError:
        pass
```

- [ ] **Step 3: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_cliente.py -q
```

Esperado: `ModuleNotFoundError: No module named 'modules.drive.cliente'`

- [ ] **Step 4: Implementar el cliente real**

Crear `modules/drive/cliente.py`:

```python
# -*- coding: utf-8 -*-
"""Envoltura fina de la API de Google Drive.

Se mantiene chica a propósito: todo lo que el robot necesita son siete
operaciones. Con una interfaz así de acotada, `tests/drive_falso.py` puede
sustituirla y las pruebas nunca tocan la red.
"""
import logging
import os

logger = logging.getLogger(__name__)

CARPETA_MIME = "application/vnd.google-apps.folder"


class DriveCliente:
    def __init__(self, servicio=None):
        """`servicio` se inyecta en pruebas; en producción lo construye auth."""
        if servicio is None:
            from modules.drive.auth import construir_servicio
            servicio = construir_servicio()
        self._s = servicio

    def subir(self, ruta_local, carpeta_id, nombre):
        from googleapiclient.http import MediaFileUpload
        media = MediaFileUpload(ruta_local, resumable=True)
        meta = {"name": nombre, "parents": [carpeta_id]}
        res = self._s.files().create(body=meta, media_body=media,
                                     fields="id").execute()
        return res["id"]

    def crear_carpeta(self, nombre, padre_id):
        meta = {"name": nombre, "mimeType": CARPETA_MIME, "parents": [padre_id]}
        return self._s.files().create(body=meta, fields="id").execute()["id"]

    def buscar_carpeta(self, nombre, padre_id):
        q = ("name = '%s' and mimeType = '%s' and '%s' in parents "
             "and trashed = false" % (nombre.replace("'", "\\'"),
                                       CARPETA_MIME, padre_id))
        r = self._s.files().list(q=q, fields="files(id)", pageSize=1).execute()
        f = r.get("files") or []
        return f[0]["id"] if f else None

    def buscar_archivo(self, nombre, carpeta_id):
        q = ("name = '%s' and '%s' in parents and trashed = false"
             % (nombre.replace("'", "\\'"), carpeta_id))
        r = self._s.files().list(q=q, fields="files(id)", pageSize=1).execute()
        f = r.get("files") or []
        return f[0]["id"] if f else None

    def listar(self, carpeta_id):
        q = "'%s' in parents and trashed = false" % carpeta_id
        r = self._s.files().list(q=q, fields="files(id,name)",
                                 pageSize=1000).execute()
        return [{"id": a["id"], "nombre": a["name"]} for a in r.get("files") or []]

    def mover(self, file_id, carpeta_destino_id):
        actual = self._s.files().get(fileId=file_id,
                                     fields="parents").execute()
        previos = ",".join(actual.get("parents") or [])
        self._s.files().update(fileId=file_id, addParents=carpeta_destino_id,
                               removeParents=previos, fields="id").execute()

    def cuota(self):
        r = self._s.about().get(fields="storageQuota").execute()
        q = r["storageQuota"]
        return {"usado": int(q.get("usage", 0)),
                "total": int(q.get("limit", 0)) or None}
```

- [ ] **Step 5: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_cliente.py -q
```

Esperado: `4 passed`

- [ ] **Step 6: Commit**

```bash
git add modules/drive/cliente.py tests/drive_falso.py tests/test_drive_cliente.py
git commit -m "Cliente de Drive con interfaz acotada y doble de pruebas sin red"
```

---

### Task 4: Autenticación OAuth

**Files:**
- Create: `modules/drive/auth.py`
- Create: `scripts/autorizar_drive.py`
- Test: `tests/test_drive_auth.py`

- [ ] **Step 1: Escribir las pruebas**

Crear `tests/test_drive_auth.py`:

```python
# -*- coding: utf-8 -*-
"""La autenticación falla con un mensaje accionable, nunca en silencio.

Con un Gmail común NO sirve una cuenta de servicio: sus archivos no tienen
cuota de Drive y la subida falla. Va OAuth con token de refresco.
"""
import pytest

from modules.drive.auth import FaltaAutorizacion, cargar_credenciales


def test_sin_archivo_de_token_avisa_que_hay_que_autorizar(tmp_path):
    with pytest.raises(FaltaAutorizacion) as e:
        cargar_credenciales(token_path=str(tmp_path / "no-existe.json"),
                            client_secret_path=str(tmp_path / "cs.json"))
    assert "autorizar_drive" in str(e.value)


def test_el_mensaje_dice_como_arreglarlo(tmp_path):
    with pytest.raises(FaltaAutorizacion) as e:
        cargar_credenciales(token_path=str(tmp_path / "no.json"),
                            client_secret_path=str(tmp_path / "cs.json"))
    msg = str(e.value).lower()
    assert "drive" in msg and "python" in msg
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_auth.py -q
```

Esperado: `ModuleNotFoundError: No module named 'modules.drive.auth'`

- [ ] **Step 3: Implementar**

Crear `modules/drive/auth.py`:

```python
# -*- coding: utf-8 -*-
"""Credenciales OAuth para la cuenta de Drive del robot.

POR QUÉ OAUTH Y NO UNA CUENTA DE SERVICIO
Con un Gmail común, los archivos que sube una cuenta de servicio quedan a
nombre de ella, y las cuentas de servicio no tienen cuota de Drive en cuentas
de consumidor: la subida falla. Ese camino solo sirve con Google Workspace y
unidades compartidas.

El permiso es `drive` completo (no `drive.file`) porque la carpeta de entrada
necesita leer archivos que el robot NO creó. Eso hace que Google muestre una
advertencia de "app no verificada" en la primera autorización: se acepta una
vez, es una herramienta interna de un solo usuario.
"""
import logging
import os

logger = logging.getLogger(__name__)

ALCANCES = ["https://www.googleapis.com/auth/drive"]


class FaltaAutorizacion(RuntimeError):
    """No hay token utilizable. Trae el paso a paso para arreglarlo."""


def cargar_credenciales(token_path: str = None, client_secret_path: str = None):
    """Devuelve credenciales válidas, refrescándolas si hace falta."""
    from config import DRIVE_TOKEN_PATH, DRIVE_CLIENT_SECRET
    token_path = token_path or DRIVE_TOKEN_PATH
    client_secret_path = client_secret_path or DRIVE_CLIENT_SECRET

    if not os.path.exists(token_path):
        raise FaltaAutorizacion(
            "Falta autorizar el Google Drive del robot.\n"
            "Corre una vez, desde la carpeta Robot:\n"
            "  %LOCALAPPDATA%\\Python\\bin\\python3.11.exe "
            "scripts/autorizar_drive.py\n"
            "Se abre el navegador, aceptas la advertencia de app no verificada "
            "y queda listo.")

    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    creds = Credentials.from_authorized_user_file(token_path, ALCANCES)
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open(token_path, "w", encoding="utf-8") as fh:
                fh.write(creds.to_json())
        except Exception as e:
            raise FaltaAutorizacion(
                "El permiso de Drive dejó de servir (%s).\n"
                "Vuelve a autorizar: python scripts/autorizar_drive.py" % e)
    if not creds or not creds.valid:
        raise FaltaAutorizacion(
            "El token de Drive no sirve. Corre: "
            "python scripts/autorizar_drive.py")
    return creds


def construir_servicio():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=cargar_credenciales(),
                 cache_discovery=False)
```

Crear `scripts/autorizar_drive.py`:

```python
# -*- coding: utf-8 -*-
"""Autorización única del Drive del robot. Se corre a mano, abre el navegador."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google_auth_oauthlib.flow import InstalledAppFlow

from config import DRIVE_CLIENT_SECRET, DRIVE_TOKEN_PATH
from modules.drive.auth import ALCANCES

if not os.path.exists(DRIVE_CLIENT_SECRET):
    print("Falta el archivo de credenciales:", DRIVE_CLIENT_SECRET)
    print("Bájalo de Google Cloud Console > Credenciales > ID de cliente OAuth")
    print("(tipo: Aplicación de escritorio) y guárdalo con ese nombre.")
    raise SystemExit(1)

flow = InstalledAppFlow.from_client_secrets_file(DRIVE_CLIENT_SECRET, ALCANCES)
creds = flow.run_local_server(port=0)
with open(DRIVE_TOKEN_PATH, "w", encoding="utf-8") as fh:
    fh.write(creds.to_json())
print("Listo. Token guardado en", DRIVE_TOKEN_PATH)
```

- [ ] **Step 4: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_auth.py -q
```

Esperado: `2 passed`

- [ ] **Step 5: Commit**

```bash
git add modules/drive/auth.py scripts/autorizar_drive.py tests/test_drive_auth.py
git commit -m "Autorización OAuth de Drive con mensaje accionable si falta el token"
```

---

### Task 5: Estructura de carpetas

**Files:**
- Create: `modules/drive/carpetas.py`
- Test: `tests/test_drive_carpetas.py`

- [ ] **Step 1: Escribir las pruebas**

Crear `tests/test_drive_carpetas.py`:

```python
# -*- coding: utf-8 -*-
"""Resolver rutas tipo 'Facturas Recibidas/2026' a IDs de Drive, creando lo que falte."""
import pytest

from modules.drive.carpetas import Carpetas
from tests.drive_falso import DriveFalso


@pytest.fixture
def carpetas():
    return Carpetas(DriveFalso(), raiz_id="raiz")


def test_crea_la_ruta_completa(carpetas):
    cid = carpetas.id_de("Facturas Recibidas/2026")
    assert cid is not None
    # la intermedia también existe
    assert carpetas.drive.buscar_carpeta("Facturas Recibidas", "raiz") is not None


def test_no_crea_dos_veces_la_misma_carpeta(carpetas):
    a = carpetas.id_de("Facturas Recibidas/2026")
    b = carpetas.id_de("Facturas Recibidas/2026")
    assert a == b
    assert len([c for c in carpetas.drive.carpetas.values()
                if c and c["nombre"] == "2026"]) == 1


def test_reutiliza_la_carpeta_padre_entre_anios(carpetas):
    carpetas.id_de("Facturas Recibidas/2025")
    carpetas.id_de("Facturas Recibidas/2026")
    padres = [c for c in carpetas.drive.carpetas.values()
              if c and c["nombre"] == "Facturas Recibidas"]
    assert len(padres) == 1


def test_una_sola_carpeta_sin_barras(carpetas):
    assert carpetas.id_de("Respaldos") is not None


def test_ruta_vacia_devuelve_la_raiz(carpetas):
    assert carpetas.id_de("") == "raiz"
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_carpetas.py -q
```

Esperado: `ModuleNotFoundError: No module named 'modules.drive.carpetas'`

- [ ] **Step 3: Implementar**

Crear `modules/drive/carpetas.py`:

```python
# -*- coding: utf-8 -*-
"""Traduce rutas legibles ('Facturas Recibidas/2026') a IDs de carpeta de Drive.

Cachea los IDs en memoria: resolver la misma ruta en cada factura serían dos
llamadas a la API por documento.
"""
import logging

logger = logging.getLogger(__name__)


class Carpetas:
    def __init__(self, drive, raiz_id: str):
        self.drive = drive
        self.raiz_id = raiz_id
        self._cache = {"": raiz_id}

    def id_de(self, ruta: str) -> str:
        """ID de la carpeta, creando los tramos que falten."""
        ruta = (ruta or "").strip("/")
        if ruta in self._cache:
            return self._cache[ruta]
        padre = self.raiz_id
        acumulada = []
        for tramo in ruta.split("/"):
            if not tramo:
                continue
            acumulada.append(tramo)
            clave = "/".join(acumulada)
            if clave in self._cache:
                padre = self._cache[clave]
                continue
            cid = self.drive.buscar_carpeta(tramo, padre)
            if cid is None:
                cid = self.drive.crear_carpeta(tramo, padre)
                logger.info("Drive: carpeta creada %s", clave)
            self._cache[clave] = cid
            padre = cid
        return padre
```

- [ ] **Step 4: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_carpetas.py -q
```

Esperado: `5 passed`

- [ ] **Step 5: Commit**

```bash
git add modules/drive/carpetas.py tests/test_drive_carpetas.py
git commit -m "Resolución de carpetas de Drive con caché de IDs"
```

---

### Task 6: El trabajador que vacía la cola

**Files:**
- Create: `modules/drive/subidor.py`
- Test: `tests/test_drive_subidor.py`

- [ ] **Step 1: Escribir las pruebas**

Crear `tests/test_drive_subidor.py`:

```python
# -*- coding: utf-8 -*-
"""El subidor vacía la cola sin perder nunca el archivo local."""
import pytest

from modules.drive.carpetas import Carpetas
from modules.drive.cola import Cola
from modules.drive.subidor import procesar_cola
from tests.drive_falso import DriveFalso


@pytest.fixture
def entorno(tmp_path):
    doc = tmp_path / "COPEVAL_1.pdf"
    doc.write_text("contenido", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    drive = DriveFalso()
    return cola, drive, Carpetas(drive, "raiz"), str(doc)


def test_sube_lo_pendiente_y_vacia_la_cola(entorno):
    cola, drive, carpetas, doc = entorno
    cola.encolar(doc, "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    res = procesar_cola(cola, drive, carpetas)
    assert res["subidos"] == 1
    assert cola.pendientes() == []
    cid = carpetas.id_de("Facturas Recibidas/2026")
    assert drive.buscar_archivo("COPEVAL_1.pdf", cid) is not None


def test_si_drive_falla_el_archivo_queda_encolado_y_en_disco(entorno):
    cola, drive, carpetas, doc = entorno
    drive.fallar_con = ConnectionError("sin internet")
    cola.encolar(doc, "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    res = procesar_cola(cola, drive, carpetas)
    assert res["subidos"] == 0
    assert res["fallidos"] == 1
    assert len(cola.pendientes()) == 1          # sigue pendiente
    import os
    assert os.path.exists(doc)                   # y el archivo NO se borró


def test_no_sube_dos_veces_el_mismo_documento(entorno):
    cola, drive, carpetas, doc = entorno
    cid = carpetas.id_de("Facturas Recibidas/2026")
    drive.subir(doc, cid, "COPEVAL_1.pdf")       # ya estaba en Drive
    cola.encolar(doc, "Facturas Recibidas/2026", "COPEVAL_1.pdf")
    procesar_cola(cola, drive, carpetas)
    assert len(drive.listar(cid)) == 1


def test_un_archivo_que_ya_no_existe_no_bloquea_la_cola(entorno, tmp_path):
    cola, drive, carpetas, doc = entorno
    cola.encolar(str(tmp_path / "borrado.pdf"), "F/2026", "borrado.pdf")
    res = procesar_cola(cola, drive, carpetas)
    assert res["fallidos"] == 1
    assert cola.pendientes()[0]["intentos"] == 1


def test_cola_vacia_no_hace_nada(entorno):
    cola, drive, carpetas, _ = entorno
    assert procesar_cola(cola, drive, carpetas) == {"subidos": 0, "fallidos": 0}
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_subidor.py -q
```

Esperado: `ModuleNotFoundError: No module named 'modules.drive.subidor'`

- [ ] **Step 3: Implementar**

Crear `modules/drive/subidor.py`:

```python
# -*- coding: utf-8 -*-
"""Vacía la cola de subidas. El archivo local NUNCA se borra acá.

Que el archivo siga en disco después de subir es a propósito: Drive no puede
ser el único lugar donde vivió un documento.
"""
import logging
import os

logger = logging.getLogger(__name__)


def procesar_cola(cola, drive, carpetas) -> dict:
    subidos = fallidos = 0
    for item in cola.pendientes():
        ruta = item["ruta_local"]
        try:
            if not os.path.exists(ruta):
                raise FileNotFoundError("ya no está en disco: %s" % ruta)
            cid = carpetas.id_de(item["carpeta"])
            existente = drive.buscar_archivo(item["nombre"], cid)
            if existente:
                # Juan reenvía cosas; no duplicar
                cola.marcar_ok(item["id"], existente)
                subidos += 1
                continue
            file_id = drive.subir(ruta, cid, item["nombre"])
            cola.marcar_ok(item["id"], file_id)
            subidos += 1
            logger.info("Drive: subido %s -> %s", item["nombre"], item["carpeta"])
        except Exception as e:
            cola.marcar_error(item["id"], str(e))
            fallidos += 1
            logger.warning("Drive: falló %s (%s)", item["nombre"], e)
    return {"subidos": subidos, "fallidos": fallidos}
```

- [ ] **Step 4: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_subidor.py -q
```

Esperado: `5 passed`

- [ ] **Step 5: Commit**

```bash
git add modules/drive/subidor.py tests/test_drive_subidor.py
git commit -m "Subidor: vacía la cola sin borrar nunca el archivo local"
```

---

### Task 7: Enganchar las facturas que llegan por Telegram

**Files:**
- Modify: `handlers/facturas.py` (después de `_renombrar_archivo`, línea ~520)
- Modify: `excel_manager.py` (encabezados de `Facturas`)
- Test: `tests/test_facturas_drive.py`

- [ ] **Step 1: Escribir la prueba**

Crear `tests/test_facturas_drive.py`:

```python
# -*- coding: utf-8 -*-
"""Al guardar una factura se encola su subida, sin bloquear la respuesta."""
import pytest

from modules.drive.cola import Cola
from handlers.facturas import encolar_documento


def test_encola_en_la_carpeta_del_anio(tmp_path):
    doc = tmp_path / "COPEVAL_123.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision="2026-08-25", cola=cola)
    p = cola.pendientes()
    assert len(p) == 1
    assert p[0]["carpeta"] == "Facturas Recibidas/2026"
    assert p[0]["nombre"] == "COPEVAL_123.pdf"


def test_sin_fecha_usa_el_anio_actual(tmp_path):
    from datetime import date
    doc = tmp_path / "X_1.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision=None, cola=cola)
    assert cola.pendientes()[0]["carpeta"].endswith(str(date.today().year))


def test_una_boleta_de_honorarios_va_a_su_carpeta(tmp_path):
    doc = tmp_path / "DONOSO_9.pdf"
    doc.write_text("x", encoding="utf-8")
    cola = Cola(str(tmp_path / "cola.jsonl"))
    encolar_documento(str(doc), fecha_emision="2026-08-25", cola=cola,
                      tipo="boleta")
    assert cola.pendientes()[0]["carpeta"] == "Boletas Honorarios"


def test_si_falla_al_encolar_no_revienta_el_flujo(tmp_path):
    """Perder el enlace es malo; perder la factura es peor."""
    cola = Cola("Z:/ruta/que/no/existe/cola.jsonl")
    # no debe lanzar
    encolar_documento(str(tmp_path / "no-existe.pdf"),
                      fecha_emision="2026-08-25", cola=cola)
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_facturas_drive.py -q
```

Esperado: `ImportError: cannot import name 'encolar_documento'`

- [ ] **Step 3: Implementar en `handlers/facturas.py`**

Agregar cerca del inicio del archivo, después de los imports existentes:

```python
def encolar_documento(file_path: str, fecha_emision=None, cola=None,
                       tipo: str = "factura") -> None:
    """Encola la subida del documento a Drive. Nunca lanza.

    El archivo ya está guardado en disco cuando esto corre: si encolar falla,
    se pierde el enlace, no el documento.
    """
    import os
    from datetime import date
    try:
        if cola is None:
            from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS
            from modules.drive.cola import Cola
            cola = Cola(DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS)
        if tipo == "boleta":
            carpeta = "Boletas Honorarios"
        else:
            anio = str(fecha_emision or "")[:4]
            if not anio.isdigit():
                anio = str(date.today().year)
            carpeta = "Facturas Recibidas/%s" % anio
        cola.encolar(file_path, carpeta, os.path.basename(file_path))
    except Exception as e:
        logger.warning("No pude encolar %s para Drive: %s", file_path, e)
```

Y en `_process_and_reply`, justo después de la línea `file_path = _renombrar_archivo(file_path, items)`:

```python
    encolar_documento(file_path,
                      fecha_emision=(items[0].get("Fecha Emision")
                                      if items else None))
```

- [ ] **Step 4: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_facturas_drive.py -q
```

Esperado: `4 passed`

- [ ] **Step 5: Agregar la columna del enlace**

En `excel_manager.py`, agregar `"Archivo Drive"` al final de la lista de encabezados de la hoja `Facturas` (queda como columna 22).

- [ ] **Step 6: Correr la suite completa**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/ -q
```

Esperado: todo pasando (420 + los nuevos)

- [ ] **Step 7: Commit**

```bash
git add handlers/facturas.py excel_manager.py tests/test_facturas_drive.py
git commit -m "Las facturas que llegan por Telegram se encolan para Drive"
```

---

### Task 8: Respaldos a Drive con retención

**Files:**
- Modify: `infrastructure/backups.py`
- Test: `tests/test_backups_retencion.py`

- [ ] **Step 1: Escribir las pruebas de retención**

Crear `tests/test_backups_retencion.py`:

```python
# -*- coding: utf-8 -*-
"""Retención de respaldos: diarios 30 días, mensuales del año, anuales siempre.

Sin esto, un Master de ~530 KB diario suma ~190 MB al año de copias casi
idénticas.
"""
from datetime import date

from infrastructure.backups import cuales_borrar


def _snaps(fechas):
    return [{"nombre": "master_%s.xlsx" % f, "fecha": date.fromisoformat(f)}
            for f in fechas]


def test_conserva_todos_los_de_los_ultimos_30_dias():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2026-08-25", "2026-08-20", "2026-08-01"])
    assert cuales_borrar(snaps, hoy=hoy) == []


def test_de_un_mes_viejo_conserva_solo_uno():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2026-03-01", "2026-03-15", "2026-03-28"])
    borrar = cuales_borrar(snaps, hoy=hoy)
    assert len(borrar) == 2
    # se conserva el más reciente del mes
    assert "2026-03-28" not in " ".join(b["nombre"] for b in borrar)


def test_de_un_anio_viejo_conserva_solo_uno():
    hoy = date(2026, 8, 25)
    snaps = _snaps(["2024-01-10", "2024-06-10", "2024-12-10"])
    borrar = cuales_borrar(snaps, hoy=hoy)
    assert len(borrar) == 2
    assert "2024-12-10" not in " ".join(b["nombre"] for b in borrar)


def test_lista_vacia_no_falla():
    assert cuales_borrar([], hoy=date(2026, 8, 25)) == []
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_backups_retencion.py -q
```

Esperado: `ImportError: cannot import name 'cuales_borrar'`

- [ ] **Step 3: Implementar la retención**

Agregar en `infrastructure/backups.py`:

```python
def cuales_borrar(snapshots: list[dict], hoy=None) -> list[dict]:
    """Cuáles respaldos sobran.

    Regla: todos los de los últimos 30 días · uno por mes del año en curso ·
    uno por año hacia atrás.
    """
    from datetime import date, timedelta
    hoy = hoy or date.today()
    limite_diario = hoy - timedelta(days=30)

    recientes, por_mes, por_anio = [], {}, {}
    for s in snapshots:
        f = s["fecha"]
        if f >= limite_diario:
            recientes.append(s)
        elif f.year == hoy.year:
            por_mes.setdefault((f.year, f.month), []).append(s)
        else:
            por_anio.setdefault(f.year, []).append(s)

    borrar = []
    for grupo in list(por_mes.values()) + list(por_anio.values()):
        grupo.sort(key=lambda s: s["fecha"])
        borrar.extend(grupo[:-1])          # se conserva el más reciente
    return borrar
```

- [ ] **Step 4: Encolar el respaldo hacia Drive**

En `backup_master`, después de `shutil.copy2(excel_path, ...)` del snapshot, agregar:

```python
    try:
        from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS
        from modules.drive.cola import Cola
        Cola(DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS).encolar(
            snap_path, "Respaldos/Master", os.path.basename(snap_path))
    except Exception as e:
        logger.warning("No pude encolar el respaldo para Drive: %s", e)
```

- [ ] **Step 5: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_backups_retencion.py -q
```

Esperado: `4 passed`

- [ ] **Step 6: Commit**

```bash
git add infrastructure/backups.py tests/test_backups_retencion.py
git commit -m "Respaldos a Drive con retención: 30 diarios, mensuales del año, anuales"
```

---

### Task 9: Carpeta de entrada

**Files:**
- Create: `handlers/drive_entrada.py`
- Test: `tests/test_drive_entrada.py`

- [ ] **Step 1: Escribir las pruebas**

Crear `tests/test_drive_entrada.py`:

```python
# -*- coding: utf-8 -*-
"""La carpeta _Entrada: procesar y mover. El movimiento ES la marca de procesado."""
import pytest

from handlers.drive_entrada import revisar_entrada
from modules.drive.carpetas import Carpetas
from tests.drive_falso import DriveFalso


@pytest.fixture
def entorno():
    drive = DriveFalso()
    carpetas = Carpetas(drive, "raiz")
    entrada = carpetas.id_de("_Entrada")
    return drive, carpetas, entrada


def test_mueve_a_su_carpeta_lo_que_pudo_procesar(entorno):
    drive, carpetas, entrada = entorno
    fid = drive.subir("x", entrada, "COPEVAL_55.pdf")
    res = revisar_entrada(drive, carpetas, procesar=lambda n: "Facturas Recibidas/2026")
    assert res["procesados"] == 1
    assert drive.archivos[fid]["carpeta_id"] == carpetas.id_de("Facturas Recibidas/2026")


def test_lo_que_no_pudo_leer_va_a_sin_procesar(entorno):
    drive, carpetas, entrada = entorno
    fid = drive.subir("x", entrada, "foto_borrosa.jpg")
    def falla(nombre):
        raise ValueError("no se pudo leer")
    res = revisar_entrada(drive, carpetas, procesar=falla)
    assert res["sin_procesar"] == 1
    assert drive.archivos[fid]["carpeta_id"] == carpetas.id_de("_Entrada/Sin procesar")


def test_no_vuelve_a_tomar_lo_que_ya_movio(entorno):
    drive, carpetas, entrada = entorno
    drive.subir("x", entrada, "A.pdf")
    revisar_entrada(drive, carpetas, procesar=lambda n: "Facturas Recibidas/2026")
    res = revisar_entrada(drive, carpetas, procesar=lambda n: "Facturas Recibidas/2026")
    assert res["procesados"] == 0


def test_ignora_la_subcarpeta_sin_procesar(entorno):
    drive, carpetas, entrada = entorno
    sp = carpetas.id_de("_Entrada/Sin procesar")
    drive.subir("x", sp, "viejo.pdf")
    res = revisar_entrada(drive, carpetas, procesar=lambda n: "F/2026")
    assert res["procesados"] == 0


def test_entrada_vacia_no_hace_nada(entorno):
    drive, carpetas, _ = entorno
    assert revisar_entrada(drive, carpetas,
                            procesar=lambda n: "F/2026") == {
        "procesados": 0, "sin_procesar": 0}
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_entrada.py -q
```

Esperado: `ModuleNotFoundError: No module named 'handlers.drive_entrada'`

- [ ] **Step 3: Implementar**

Crear `handlers/drive_entrada.py`:

```python
# -*- coding: utf-8 -*-
"""Revisa `_Entrada/` en Drive y procesa lo que aparezca.

Mover el archivo a su carpeta definitiva ES la marca de procesado: no hace
falta llevar una lista aparte, y un reinicio a mitad de camino no duplica nada.
"""
import logging

logger = logging.getLogger(__name__)

ENTRADA = "_Entrada"
SIN_PROCESAR = "_Entrada/Sin procesar"


def revisar_entrada(drive, carpetas, procesar) -> dict:
    """`procesar(nombre)` devuelve la carpeta destino, o lanza si no pudo."""
    entrada_id = carpetas.id_de(ENTRADA)
    sin_procesar_id = carpetas.id_de(SIN_PROCESAR)
    procesados = fallidos = 0
    for archivo in drive.listar(entrada_id):
        try:
            destino = procesar(archivo["nombre"])
            drive.mover(archivo["id"], carpetas.id_de(destino))
            procesados += 1
            logger.info("Drive entrada: %s -> %s", archivo["nombre"], destino)
        except Exception as e:
            drive.mover(archivo["id"], sin_procesar_id)
            fallidos += 1
            logger.warning("Drive entrada: no pude con %s (%s)",
                            archivo["nombre"], e)
    return {"procesados": procesados, "sin_procesar": fallidos}
```

- [ ] **Step 4: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_entrada.py -q
```

Esperado: `5 passed`

- [ ] **Step 5: Commit**

```bash
git add handlers/drive_entrada.py tests/test_drive_entrada.py
git commit -m "Carpeta de entrada de Drive: procesar y mover"
```

---

### Task 10: Registrar los jobs y el aviso de cuota

**Files:**
- Modify: `main.py`
- Create: `handlers/drive_jobs.py`
- Test: `tests/test_drive_cuota.py`

- [ ] **Step 1: Escribir la prueba del aviso de cuota**

Crear `tests/test_drive_cuota.py`:

```python
# -*- coding: utf-8 -*-
"""Avisar ANTES de que el Drive se llene, no cuando ya no puede subir."""
from handlers.drive_jobs import hay_que_avisar_cuota
from tests.drive_falso import DriveFalso

GB = 1024 ** 3


def test_no_avisa_con_espacio_de_sobra():
    d = DriveFalso(cuota_usada=3 * GB, cuota_total=15 * GB)
    assert hay_que_avisar_cuota(d, umbral=0.80) is False


def test_avisa_al_pasar_el_umbral():
    d = DriveFalso(cuota_usada=13 * GB, cuota_total=15 * GB)
    assert hay_que_avisar_cuota(d, umbral=0.80) is True


def test_justo_en_el_umbral_avisa():
    d = DriveFalso(cuota_usada=12 * GB, cuota_total=15 * GB)
    assert hay_que_avisar_cuota(d, umbral=0.80) is True


def test_sin_dato_de_cuota_no_avisa():
    d = DriveFalso(cuota_usada=0, cuota_total=0)
    assert hay_que_avisar_cuota(d, umbral=0.80) is False
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_cuota.py -q
```

Esperado: `ModuleNotFoundError: No module named 'handlers.drive_jobs'`

- [ ] **Step 3: Implementar**

Crear `handlers/drive_jobs.py`:

```python
# -*- coding: utf-8 -*-
"""Jobs de Drive: vaciar la cola y vigilar la cuota."""
import logging

from config import TELEGRAM_CHAT_ID

logger = logging.getLogger(__name__)


def hay_que_avisar_cuota(drive, umbral: float = 0.80) -> bool:
    try:
        q = drive.cuota()
        total = q.get("total") or 0
        if total <= 0:
            return False
        return (q.get("usado", 0) / total) >= umbral
    except Exception as e:
        logger.warning("No pude leer la cuota de Drive: %s", e)
        return False


async def job_drive_cola(context):
    """Vacía la cola de subidas. Avisa si la autorización se cayó."""
    import asyncio

    from config import (DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS, DRIVE_RAIZ,
                         DRIVE_UMBRAL_AVISO)
    from modules.drive.auth import FaltaAutorizacion
    from modules.drive.carpetas import Carpetas
    from modules.drive.cliente import DriveCliente
    from modules.drive.cola import Cola
    from modules.drive.subidor import procesar_cola

    cola = Cola(DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS)
    if not cola.pendientes() and not cola.rendidos():
        return

    chat_id = (context.bot_data.get("owner_chat_id")
               or context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID)
    try:
        drive = await asyncio.to_thread(DriveCliente)
    except FaltaAutorizacion as e:
        if chat_id and context.bot_data.get("drive_aviso_auth") != "si":
            context.bot_data["drive_aviso_auth"] = "si"
            await context.bot.send_message(chat_id=int(chat_id),
                                            text="🔑 %s" % e)
        return
    context.bot_data["drive_aviso_auth"] = "no"

    raiz = await asyncio.to_thread(_raiz_id, drive, DRIVE_RAIZ)
    carpetas = Carpetas(drive, raiz)
    res = await asyncio.to_thread(procesar_cola, cola, drive, carpetas)
    if res["subidos"]:
        logger.info("Drive: %d subidos, %d fallidos",
                     res["subidos"], res["fallidos"])

    rendidos = cola.rendidos()
    if rendidos and chat_id:
        await context.bot.send_message(
            chat_id=int(chat_id),
            text=("⚠️ %d documento(s) no pudieron subir a Drive tras varios "
                  "intentos. Siguen guardados en el PC." % len(rendidos)))

    if hay_que_avisar_cuota(drive, DRIVE_UMBRAL_AVISO) and chat_id:
        if context.bot_data.get("drive_aviso_cuota") != "si":
            context.bot_data["drive_aviso_cuota"] = "si"
            await context.bot.send_message(
                chat_id=int(chat_id),
                text="📦 El Drive del robot pasó el 80% de su espacio.")


def _raiz_id(drive, nombre_raiz: str) -> str:
    """ID de la carpeta raíz, creándola la primera vez."""
    cid = drive.buscar_carpeta(nombre_raiz, "root")
    return cid or drive.crear_carpeta(nombre_raiz, "root")
```

- [ ] **Step 4: Registrar el job en `main.py`**

Junto a los demás `run_repeating`, después del job `latido`:

```python
    # Vaciar la cola de subidas a Drive cada 10 min. Es barato: si la cola
    # está vacía, ni siquiera se autentica.
    from handlers.drive_jobs import job_drive_cola
    app.job_queue.run_repeating(job_drive_cola, interval=600, first=60,
                                name="drive_cola")
```

- [ ] **Step 5: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_cuota.py -q
```

Esperado: `4 passed`

- [ ] **Step 6: Correr la suite completa y verificar que `main.py` compila**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -c "import ast; ast.parse(open('main.py',encoding='utf-8').read()); print('main.py OK')"
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/ -q
```

Esperado: `main.py OK` y toda la suite pasando

- [ ] **Step 7: Commit**

```bash
git add main.py handlers/drive_jobs.py tests/test_drive_cuota.py
git commit -m "Job de Drive cada 10 min: vacía la cola, avisa por token y por cuota"
```

---

### Task 11: Escribir el enlace de Drive en la fila de la factura

Sin esto el archivo llega a Drive pero es inencontrable desde el Master.

**Files:**
- Create: `modules/drive/enlaces.py`
- Modify: `modules/drive/subidor.py`
- Test: `tests/test_drive_enlaces.py`

- [ ] **Step 1: Escribir las pruebas**

Crear `tests/test_drive_enlaces.py`:

```python
# -*- coding: utf-8 -*-
"""El enlace de Drive vuelve a la fila de la factura.

OJO: ruta EXPLÍCITA al Excel en todos los tests. Un test destruyó el Master
real por confiar en el default de _save_wb.
"""
import openpyxl
import pytest

from modules.drive.enlaces import guardar_enlace

COL_NUM, COL_DRIVE = 7, 22


@pytest.fixture
def libro(tmp_path):
    ruta = tmp_path / "master.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"
    ws.append([None] * 21 + ["Archivo Drive"])
    fila = [None] * 22
    fila[COL_NUM - 1] = 2777
    ws.append(fila)
    wb.save(ruta)
    wb.close()
    return str(ruta)


def test_guarda_el_enlace_en_la_fila_del_numero(libro):
    guardar_enlace(libro, "2777", "abc123")
    wb = openpyxl.load_workbook(libro)
    assert "abc123" in str(wb["Facturas"].cell(2, COL_DRIVE).value)
    wb.close()


def test_el_enlace_es_una_url_abrible(libro):
    guardar_enlace(libro, "2777", "abc123")
    wb = openpyxl.load_workbook(libro)
    assert str(wb["Facturas"].cell(2, COL_DRIVE).value).startswith("https://")
    wb.close()


def test_un_numero_que_no_existe_no_rompe(libro):
    guardar_enlace(libro, "9999", "abc123")
    wb = openpyxl.load_workbook(libro)
    assert wb["Facturas"].cell(2, COL_DRIVE).value is None
    wb.close()


def test_no_pisa_un_enlace_ya_puesto(libro):
    guardar_enlace(libro, "2777", "primero")
    guardar_enlace(libro, "2777", "segundo")
    wb = openpyxl.load_workbook(libro)
    assert "primero" in str(wb["Facturas"].cell(2, COL_DRIVE).value)
    wb.close()
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_enlaces.py -q
```

Esperado: `ModuleNotFoundError: No module named 'modules.drive.enlaces'`

- [ ] **Step 3: Implementar**

Crear `modules/drive/enlaces.py`:

```python
# -*- coding: utf-8 -*-
"""Escribe en la hoja Facturas el enlace del documento en Drive."""
import logging

logger = logging.getLogger(__name__)

COL_NUMERO, COL_DRIVE = 7, 22
URL = "https://drive.google.com/file/d/%s/view"


def guardar_enlace(excel_path: str, numero_factura: str, file_id: str) -> bool:
    """Pone el enlace en TODAS las líneas de esa factura. No pisa lo existente."""
    from openpyxl import load_workbook
    wb = load_workbook(excel_path)
    try:
        ws = wb["Facturas"]
        objetivo = str(numero_factura).strip().rstrip(".0")
        tocadas = 0
        for f in range(2, ws.max_row + 1):
            actual = str(ws.cell(f, COL_NUMERO).value or "").strip().rstrip(".0")
            if actual != objetivo:
                continue
            if ws.cell(f, COL_DRIVE).value:
                continue
            ws.cell(f, COL_DRIVE).value = URL % file_id
            tocadas += 1
        if tocadas:
            wb.save(excel_path)          # ruta EXPLÍCITA siempre
        return tocadas > 0
    finally:
        wb.close()
```

- [ ] **Step 4: Enganchar en el subidor**

En `modules/drive/subidor.py`, después de cada `cola.marcar_ok(item["id"], file_id)` agregar `_enlazar(item, file_id)`, y al final del archivo:

```python
def _enlazar(item: dict, file_id: str) -> None:
    """Deja el enlace en la fila de la factura. Nunca lanza."""
    import re
    try:
        m = re.search(r"_(\d+)\.[A-Za-z0-9]+$", item["nombre"])
        if not m:
            return
        from config import EXCEL_PATH
        from modules.drive.enlaces import guardar_enlace
        guardar_enlace(EXCEL_PATH, m.group(1), file_id)
    except Exception as e:
        logger.warning("No pude enlazar %s: %s", item["nombre"], e)
```

- [ ] **Step 5: Correr las pruebas**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_enlaces.py tests/test_drive_subidor.py -q
```

Esperado: `9 passed`

- [ ] **Step 6: Commit**

```bash
git add modules/drive/enlaces.py modules/drive/subidor.py tests/test_drive_enlaces.py
git commit -m "El enlace de Drive vuelve a la fila de la factura en el Master"
```

---

### Task 12: Conectar la carpeta de entrada al lector de documentos

La Task 9 dejó `revisar_entrada` con un `procesar` que las pruebas inyectan. Falta el de producción.

**Files:**
- Modify: `handlers/drive_entrada.py`
- Modify: `handlers/drive_jobs.py`
- Modify: `modules/drive/cliente.py`, `tests/drive_falso.py`, `tests/test_drive_cliente.py`
- Modify: `main.py`
- Test: `tests/test_drive_entrada_real.py`

- [ ] **Step 1: Escribir la prueba del clasificador**

Crear `tests/test_drive_entrada_real.py`:

```python
# -*- coding: utf-8 -*-
"""Qué carpeta le toca a cada documento que aparece en _Entrada."""
import pytest

from handlers.drive_entrada import carpeta_para


def test_una_factura_va_al_anio_de_su_emision():
    assert carpeta_para({"tipo": "factura",
                          "fecha": "2026-08-25"}) == "Facturas Recibidas/2026"


def test_una_boleta_de_honorarios_va_a_su_carpeta():
    assert carpeta_para({"tipo": "boleta", "fecha": "2026-08-25"}) == \
        "Boletas Honorarios"


def test_una_guia_va_a_guias():
    assert carpeta_para({"tipo": "guia", "fecha": "2026-08-25"}) == \
        "Guías de Despacho"


def test_sin_fecha_usa_el_anio_actual():
    from datetime import date
    assert carpeta_para({"tipo": "factura", "fecha": None}).endswith(
        str(date.today().year))


def test_un_tipo_desconocido_lanza_para_que_vaya_a_sin_procesar():
    with pytest.raises(ValueError):
        carpeta_para({"tipo": "quien sabe", "fecha": "2026-08-25"})
```

- [ ] **Step 2: Correr para ver fallar**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/test_drive_entrada_real.py -q
```

Esperado: `ImportError: cannot import name 'carpeta_para'`

- [ ] **Step 3: Implementar el clasificador**

Agregar en `handlers/drive_entrada.py`:

```python
def carpeta_para(datos: dict) -> str:
    """Carpeta destino según el tipo. Lanza si no lo reconoce.

    Lanzar es deliberado: `revisar_entrada` manda a 'Sin procesar' lo que no
    supo clasificar, en vez de adivinar una carpeta.
    """
    from datetime import date
    tipo = str(datos.get("tipo") or "").lower()
    if tipo == "boleta":
        return "Boletas Honorarios"
    if tipo == "guia":
        return "Guías de Despacho"
    if tipo == "factura":
        anio = str(datos.get("fecha") or "")[:4]
        if not anio.isdigit():
            anio = str(date.today().year)
        return "Facturas Recibidas/%s" % anio
    raise ValueError("tipo de documento desconocido: %r" % tipo)
```

- [ ] **Step 4: Agregar `descargar` al cliente y al falso**

En `modules/drive/cliente.py`:

```python
    def descargar(self, file_id, ruta_local):
        from googleapiclient.http import MediaIoBaseDownload
        pedido = self._s.files().get_media(fileId=file_id)
        with open(ruta_local, "wb") as fh:
            bajada = MediaIoBaseDownload(fh, pedido)
            listo = False
            while not listo:
                _, listo = bajada.next_chunk()
        return ruta_local
```

En `tests/drive_falso.py`:

```python
    def descargar(self, file_id, ruta_local):
        with open(ruta_local, "w", encoding="utf-8") as fh:
            fh.write("contenido falso")
        return ruta_local
```

Agregar `"descargar"` a la lista `METODOS` de `tests/test_drive_cliente.py`.

- [ ] **Step 5: Pasar el archivo completo a `procesar`**

En `handlers/drive_entrada.py`, dentro de `revisar_entrada`, cambiar
`destino = procesar(archivo["nombre"])` por `destino = procesar(archivo)` — el
job necesita el `id` para descargar, no solo el nombre.

En `tests/test_drive_entrada.py`, ajustar las lambdas a `lambda a: "Facturas Recibidas/2026"`
y `def falla(a): raise ValueError("no se pudo leer")`.

- [ ] **Step 6: Agregar el job de entrada**

⚠️ **Antes de escribirlo**, confirmar cómo se llama la función del extractor en
`processors/` — abajo va como `extraer_documento`, que es lo esperado pero **no
está verificado**. Si el extractor solo entiende facturas, las guías caerán en
Sin procesar hasta que se amplíe: eso es correcto, no un bug.

En `handlers/drive_jobs.py`, al final:

```python
async def job_drive_entrada(context):
    """Revisa la carpeta _Entrada de Drive cada 15 min."""
    import asyncio
    import os
    import tempfile

    from config import DRIVE_RAIZ
    from handlers.drive_entrada import carpeta_para, revisar_entrada
    from modules.drive.auth import FaltaAutorizacion
    from modules.drive.carpetas import Carpetas
    from modules.drive.cliente import DriveCliente

    try:
        drive = await asyncio.to_thread(DriveCliente)
    except FaltaAutorizacion:
        return                       # job_drive_cola ya avisa por esto

    raiz = await asyncio.to_thread(_raiz_id, drive, DRIVE_RAIZ)
    carpetas = Carpetas(drive, raiz)

    def procesar(archivo):
        from processors.extractor import extraer_documento
        with tempfile.TemporaryDirectory() as tmp:
            local = os.path.join(tmp, archivo["nombre"])
            drive.descargar(archivo["id"], local)
            return carpeta_para(extraer_documento(local))

    res = await asyncio.to_thread(revisar_entrada, drive, carpetas, procesar)
    if res["sin_procesar"]:
        chat_id = (context.bot_data.get("owner_chat_id")
                   or context.bot_data.get("banco_chat_id") or TELEGRAM_CHAT_ID)
        if chat_id:
            await context.bot.send_message(
                chat_id=int(chat_id),
                text=("📂 %d archivo(s) de _Entrada no los supe leer. "
                      "Quedaron en _Entrada/Sin procesar."
                      % res["sin_procesar"]))
```

- [ ] **Step 7: Registrar el job en `main.py`**

```python
    from handlers.drive_jobs import job_drive_entrada
    app.job_queue.run_repeating(job_drive_entrada, interval=900, first=120,
                                name="drive_entrada")
```

- [ ] **Step 8: Correr la suite completa**

```bash
"$LOCALAPPDATA/Python/bin/python3.11.exe" -m pytest tests/ -q
```

Esperado: todo pasando

- [ ] **Step 9: Commit**

```bash
git add handlers/drive_entrada.py handlers/drive_jobs.py main.py \
        modules/drive/cliente.py tests/drive_falso.py \
        tests/test_drive_cliente.py tests/test_drive_entrada.py \
        tests/test_drive_entrada_real.py
git commit -m "Carpeta de entrada conectada al lector de documentos"
```

---

### Task 13: Migración de los documentos existentes (~200 MB)

**Se corre al final**, con todo lo anterior andando y verificado.

Se corre **a mano, una vez**, y por lotes: son ~950 archivos, 201 MB en total.

**Files:**
- Create: `scripts/carga/migrar_documentos_a_drive.py`

- [ ] **Step 1: Escribir el script**

Crear `scripts/carga/migrar_documentos_a_drive.py`:

```python
# -*- coding: utf-8 -*-
"""Migración única de los documentos locales a Drive.

Se corre por CARPETA, no todo de una: el riesgo no es el espacio (201 MB de 15 GB)
sino perder la pista de un archivo, así que conviene verificar conteos
antes de dar por buena cada tanda.

Uso:
    python scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas" --simular
    python scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas"
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))

from config import DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS
from modules.drive.cola import Cola

BASE = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))

# carpeta local -> carpeta en Drive
DESTINOS = {
    "Facturas Recibidas": "Facturas Recibidas",
    "Facturas Enviadas": "Facturas Enviadas",
    "BH": "Boletas Honorarios",
    "Guias de Despacho": "Guías de Despacho",
    "Rendiciones": "Rendiciones",
    "Legal": "Legal",
    "Carpeta Tributaria": "Tributario",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("carpeta", choices=sorted(DESTINOS))
    ap.add_argument("--simular", action="store_true")
    args = ap.parse_args()

    origen = os.path.join(BASE, args.carpeta)
    destino = DESTINOS[args.carpeta]
    if not os.path.isdir(origen):
        print("No existe:", origen)
        raise SystemExit(1)

    archivos = [os.path.join(dp, f)
                for dp, _, fs in os.walk(origen) for f in fs]
    total_mb = sum(os.path.getsize(a) for a in archivos) / 1024 ** 2
    print("%s: %d archivos, %.1f MB -> Drive:%s"
          % (args.carpeta, len(archivos), total_mb, destino))

    if args.simular:
        for a in archivos[:10]:
            print("   ", os.path.basename(a))
        if len(archivos) > 10:
            print("    ... y %d más" % (len(archivos) - 10))
        print("\n(simulación: no se encoló nada)")
        return

    cola = Cola(DRIVE_COLA_PATH, DRIVE_MAX_INTENTOS)
    for a in archivos:
        # las facturas se parten por año usando la fecha del archivo
        carpeta = destino
        if args.carpeta == "Facturas Recibidas":
            import datetime as dt
            anio = dt.date.fromtimestamp(os.path.getmtime(a)).year
            carpeta = "%s/%d" % (destino, anio)
        cola.encolar(a, carpeta, os.path.basename(a))
    print("Encolados %d archivos. El job del bot los va a ir subiendo."
          % len(archivos))


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Simular la carpeta más grande**

```bash
cd "C:/Users/Windows/Desktop/Workflow/Agricola Santa Elisa/Robot"
"$LOCALAPPDATA/Python/bin/python3.11.exe" scripts/carga/migrar_documentos_a_drive.py "Facturas Recibidas" --simular
```

Esperado: `Facturas Recibidas: 817 archivos, 111.x MB -> Drive:Facturas Recibidas`

- [ ] **Step 3: Commit**

```bash
git add scripts/carga/migrar_documentos_a_drive.py
git commit -m "Script de migración por lotes de los documentos locales a Drive"
```

- [ ] **Step 4: Migrar de verdad, una carpeta a la vez**

En este orden, verificando el conteo en Drive antes de pasar a la siguiente:

1. `Carpeta Tributaria` (5 archivos — la prueba de fuego)
2. `Guias de Despacho` (16)
3. `Facturas Enviadas` (19)
4. `BH` (60)
5. `Rendiciones` (9, 53 MB)
6. `Facturas Recibidas` (817, 111 MB)
7. `Legal` (26 archivos, 18 MB)

---

## Antes de empezar: lo que el dueño tiene que hacer una vez

1. En [Google Cloud Console](https://console.cloud.google.com), con la cuenta nueva del robot: crear un proyecto, habilitar la **Google Drive API**, y crear un **ID de cliente OAuth** de tipo **Aplicación de escritorio**.
2. Bajar el JSON y guardarlo como `Robot/.drive_client_secret.json`.
3. Correr `python scripts/autorizar_drive.py`, aceptar la advertencia de app no verificada.
4. Dejar la pantalla de consentimiento en **In production** (en *Testing* el token
   se vence cada 7 días).
