"""modules/db/models.py — Esquema de la base (SQLite hoy, PostgreSQL mañana).

Regla: nada de tipos exclusivos de un motor. Fechas como Date, montos como
Numeric, textos como String/Text. Así el mismo esquema sirve en ambos.
"""
import os
from datetime import date, datetime

from sqlalchemy import (create_engine, Column, Integer, String, Date, DateTime,
                        Numeric, Boolean, Text, Index, UniqueConstraint)
from sqlalchemy.orm import declarative_base, sessionmaker

Base = declarative_base()

# La URL sale de env para poder apuntar a PostgreSQL sin tocar código.
_DEFAULT_SQLITE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    "agricola.db")
DB_URL = os.getenv("DB_URL", f"sqlite:///{_DEFAULT_SQLITE}")


class Proveedor(Base):
    __tablename__ = "proveedores"
    id = Column(Integer, primary_key=True)
    nombre = Column(String(160), nullable=False)
    rut = Column(String(20), index=True)
    nombre_norm = Column(String(160), index=True)


class Factura(Base):
    """Cabecera: un documento (una fila por proveedor+número)."""
    __tablename__ = "facturas"
    id = Column(Integer, primary_key=True)
    proveedor = Column(String(160), index=True)
    proveedor_norm = Column(String(160), index=True)
    rut = Column(String(20))
    documento = Column(String(60))              # Factura / Boleta de honorarios / NC…
    numero = Column(String(40), index=True)
    numero_norm = Column(String(40), index=True)
    referencia = Column(String(40))             # doc referenciado (NC/ND)
    fecha_emision = Column(Date, index=True)
    fecha_vencimiento = Column(Date)
    fecha_pago = Column(Date, index=True)
    total = Column(Numeric(14, 2))
    categoria = Column(String(60), index=True)
    cultivo = Column(String(30))
    n_archivo = Column(Integer, index=True)     # correlativo de FXP
    fila_excel = Column(Integer)                # trazabilidad con el MASTER
    __table_args__ = (
        UniqueConstraint("proveedor_norm", "numero_norm", name="uq_factura"),
        Index("ix_fact_emision_prov", "fecha_emision", "proveedor_norm"),
    )


class FacturaLinea(Base):
    __tablename__ = "factura_lineas"
    id = Column(Integer, primary_key=True)
    factura_id = Column(Integer, index=True, nullable=False)
    glosa = Column(String(300))
    glosa_detalle = Column(Text)
    cantidad = Column(Numeric(14, 3))
    valor_unitario = Column(Numeric(14, 3))
    neto = Column(Numeric(14, 2))
    iva = Column(Numeric(14, 2))
    impuesto_especifico = Column(Numeric(14, 2))
    total_item = Column(Numeric(14, 2))
    fila_excel = Column(Integer)


class MovimientoBanco(Base):
    __tablename__ = "cuenta_banco"
    id = Column(Integer, primary_key=True)
    fecha = Column(Date, index=True)
    descripcion = Column(String(300))
    referencia = Column(String(40), index=True)   # N° documento del banco
    cargo = Column(Numeric(14, 2))
    abono = Column(Numeric(14, 2))
    saldo = Column(Numeric(14, 2))
    tipo = Column(String(30))
    categoria = Column(String(60), index=True)
    cultivo = Column(String(30))
    fila_excel = Column(Integer, index=True)


class Conciliacion(Base):
    __tablename__ = "conciliaciones"
    id = Column(Integer, primary_key=True)
    fecha_conciliacion = Column(Date)
    fila_banco = Column(Integer, index=True)
    tipo_doc = Column(String(20))
    numero_doc = Column(String(40))
    proveedor = Column(String(160))
    monto_asignado = Column(Numeric(14, 2))
    criterio = Column(String(30))
    usuario = Column(String(40))
    nota = Column(String(200))


class Bitacora(Base):
    __tablename__ = "bitacora"
    id = Column(Integer, primary_key=True)
    fecha = Column(Date, index=True)
    hora = Column(String(8))
    tipo = Column(String(20), index=True)
    actividad = Column(String(160))
    cultivo = Column(String(30), index=True)
    sector = Column(String(80))
    jornadas_hombre = Column(Numeric(8, 2))
    trabajadores = Column(String(300))
    insumo = Column(String(120))
    cantidad = Column(Numeric(12, 3))
    unidad = Column(String(12))
    registro = Column(Text)
    registrado_por = Column(String(60))
    maquina = Column(String(60), index=True)
    odometro = Column(Numeric(12, 2))
    horas_dia = Column(Numeric(8, 2))
    superficie_ha = Column(Numeric(8, 2))
    fila_excel = Column(Integer)


class Inventario(Base):
    __tablename__ = "inventario"
    id = Column(Integer, primary_key=True)
    producto = Column(String(120), index=True)
    categoria = Column(String(60))
    unidad = Column(String(12))
    stock = Column(Numeric(14, 3))
    stock_minimo = Column(Numeric(14, 3))
    ultima_entrada = Column(Date)
    ultimo_uso = Column(Date)
    vencimiento = Column(Date, index=True)
    estado = Column(String(20), index=True)
    fila_excel = Column(Integer)


class Personal(Base):
    __tablename__ = "personal"
    id = Column(Integer, primary_key=True)
    nombre = Column(String(120), index=True)
    rut = Column(String(20))
    cargo = Column(String(80))
    fecha_ingreso = Column(Date)
    activo = Column(Boolean, default=True)
    dias_vacaciones = Column(Numeric(8, 2))
    fila_excel = Column(Integer)


class Usuario(Base):
    """Usuarios de la web. OJO: esta tabla NO se toca en el sync desde Excel."""
    __tablename__ = "usuarios"
    id = Column(Integer, primary_key=True)
    usuario = Column(String(40), unique=True, nullable=False, index=True)
    nombre = Column(String(120))
    password_hash = Column(String(255), nullable=False)
    rol = Column(String(20), nullable=False)      # admin | gerencia | campo
    activo = Column(Boolean, default=True)
    creado = Column(DateTime, default=datetime.now)
    ultimo_acceso = Column(DateTime)
    intentos_fallidos = Column(Integer, default=0)
    bloqueado_hasta = Column(DateTime)


class Auditoria(Base):
    """Registro de accesos y acciones (exigible por la ley de datos)."""
    __tablename__ = "auditoria"
    id = Column(Integer, primary_key=True)
    fecha = Column(DateTime, default=datetime.now, index=True)
    usuario = Column(String(40), index=True)
    accion = Column(String(40))                   # login, login_fallido, ver, escribir
    recurso = Column(String(160))
    ip = Column(String(45))
    detalle = Column(String(300))


class SyncLog(Base):
    """Trazabilidad de cada sincronización Excel → base."""
    __tablename__ = "sync_log"
    id = Column(Integer, primary_key=True)
    ejecutado = Column(DateTime, default=datetime.now)
    tabla = Column(String(40))
    filas = Column(Integer)
    detalle = Column(String(300))


def get_engine(url: str = None, echo: bool = False):
    url = url or DB_URL
    kwargs = {"echo": echo, "future": True}
    if url.startswith("sqlite"):
        kwargs["connect_args"] = {"check_same_thread": False}
    return create_engine(url, **kwargs)


def get_session(engine=None):
    engine = engine or get_engine()
    return sessionmaker(bind=engine, expire_on_commit=False, future=True)()


def crear_esquema(engine=None):
    engine = engine or get_engine()
    Base.metadata.create_all(engine)
    return engine
