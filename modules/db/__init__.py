"""Capa de base de datos (Fase A de la migración Excel → SQL).

Convive con el Excel: por ahora el MASTER sigue siendo la fuente de verdad y
la base se sincroniza desde él. Cuando los totales calcen sostenidamente, se
invierte la dirección.

Se usa SQLAlchemy para que el MISMO código sirva con SQLite (local, hoy) y con
PostgreSQL (servidor, mañana): solo cambia la URL de conexión.
"""
