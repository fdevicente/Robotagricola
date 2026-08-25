"""modules/auth.py — Usuarios, roles y auditoría de la web.

Tres roles (definidos con el usuario 2026-07-29):
  · admin    — Administrador del sistema: acceso total + gestión de usuarios.
  · gerencia — Parte financiera: facturas, banco, conciliación, flujo, reportes.
  · campo    — Ingreso de datos (Juan): bitácora, inventario, tareas, facturas.
               NO ve banco, flujo, sueldos ni datos de personal.

Las contraseñas se guardan solo como hash (nunca en texto plano).
"""
import logging
from datetime import datetime, timedelta

from werkzeug.security import generate_password_hash, check_password_hash

from modules.db.models import crear_esquema, get_session, Usuario, Auditoria

logger = logging.getLogger(__name__)

ROLES = {
    "admin": "Administrador",
    "gerencia": "Gerencia",
    "campo": "Administrador de campo",
}

# Permisos por rol. 'finanzas' = plata; 'operacion' = campo; 'personal' = RRHH.
PERMISOS = {
    "admin":    {"finanzas", "operacion", "personal", "usuarios", "escribir"},
    "gerencia": {"finanzas", "operacion", "personal", "escribir"},
    "campo":    {"operacion", "escribir"},
}

MAX_INTENTOS = 5
BLOQUEO_MIN = 15


def puede(rol: str, permiso: str) -> bool:
    return permiso in PERMISOS.get(rol or "", set())


def crear_usuario(usuario: str, password: str, rol: str, nombre: str = "") -> dict:
    if rol not in ROLES:
        raise ValueError(f"Rol inválido: {rol}. Válidos: {list(ROLES)}")
    if len(password) < 8:
        raise ValueError("La contraseña debe tener al menos 8 caracteres.")
    crear_esquema()
    ses = get_session()
    try:
        if ses.query(Usuario).filter_by(usuario=usuario.lower()).first():
            raise ValueError(f"El usuario '{usuario}' ya existe.")
        u = Usuario(usuario=usuario.lower().strip(), nombre=nombre or usuario,
                    password_hash=generate_password_hash(password),
                    rol=rol, activo=True)
        ses.add(u)
        ses.commit()
        logger.info(f"Usuario creado: {usuario} ({rol})")
        return {"id": u.id, "usuario": u.usuario, "rol": u.rol}
    finally:
        ses.close()


def cambiar_password(usuario: str, password_nueva: str) -> bool:
    if len(password_nueva) < 8:
        raise ValueError("La contraseña debe tener al menos 8 caracteres.")
    ses = get_session()
    try:
        u = ses.query(Usuario).filter_by(usuario=usuario.lower()).first()
        if not u:
            return False
        u.password_hash = generate_password_hash(password_nueva)
        u.intentos_fallidos = 0
        u.bloqueado_hasta = None
        ses.commit()
        return True
    finally:
        ses.close()


def verificar(usuario: str, password: str, ip: str = "") -> tuple:
    """Devuelve (Usuario|None, mensaje). Bloquea tras varios intentos fallidos."""
    ses = get_session()
    try:
        u = ses.query(Usuario).filter_by(usuario=(usuario or "").lower().strip()).first()
        if not u or not u.activo:
            auditar("", "login_fallido", "/login", ip, f"usuario inexistente: {usuario}")
            return None, "Usuario o contraseña incorrectos."
        if u.bloqueado_hasta and u.bloqueado_hasta > datetime.now():
            faltan = int((u.bloqueado_hasta - datetime.now()).total_seconds() / 60) + 1
            return None, f"Cuenta bloqueada temporalmente. Reintenta en {faltan} min."
        if not check_password_hash(u.password_hash, password or ""):
            u.intentos_fallidos = (u.intentos_fallidos or 0) + 1
            if u.intentos_fallidos >= MAX_INTENTOS:
                u.bloqueado_hasta = datetime.now() + timedelta(minutes=BLOQUEO_MIN)
                u.intentos_fallidos = 0
                ses.commit()
                auditar(u.usuario, "bloqueo", "/login", ip, "demasiados intentos")
                return None, f"Demasiados intentos. Cuenta bloqueada {BLOQUEO_MIN} min."
            ses.commit()
            auditar(u.usuario, "login_fallido", "/login", ip, "clave incorrecta")
            return None, "Usuario o contraseña incorrectos."
        u.intentos_fallidos = 0
        u.bloqueado_hasta = None
        u.ultimo_acceso = datetime.now()
        ses.commit()
        ses.refresh(u)
        auditar(u.usuario, "login", "/login", ip, "")
        return u, ""
    finally:
        ses.close()


def obtener(user_id) -> Usuario | None:
    ses = get_session()
    try:
        return ses.query(Usuario).filter_by(id=int(user_id)).first()
    except (TypeError, ValueError):
        return None
    finally:
        ses.close()


def listar_usuarios() -> list:
    ses = get_session()
    try:
        return [{"id": u.id, "usuario": u.usuario, "nombre": u.nombre,
                 "rol": u.rol, "activo": u.activo,
                 "ultimo_acceso": u.ultimo_acceso.isoformat(sep=" ", timespec="minutes")
                 if u.ultimo_acceso else None}
                for u in ses.query(Usuario).order_by(Usuario.usuario).all()]
    finally:
        ses.close()


def auditar(usuario: str, accion: str, recurso: str = "", ip: str = "",
            detalle: str = "") -> None:
    try:
        ses = get_session()
        ses.add(Auditoria(usuario=(usuario or "")[:40], accion=accion[:40],
                          recurso=(recurso or "")[:160], ip=(ip or "")[:45],
                          detalle=(detalle or "")[:300]))
        ses.commit()
        ses.close()
    except Exception as e:
        logger.warning(f"No pude auditar: {e}")


def hay_usuarios() -> bool:
    crear_esquema()
    ses = get_session()
    try:
        return ses.query(Usuario).count() > 0
    finally:
        ses.close()
