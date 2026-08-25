#!/usr/bin/env python3
"""Crea usuarios de la web. Uso interactivo (no deja la clave en el historial).

    python scripts/crear_usuario.py
"""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import getpass
from modules.auth import crear_usuario, listar_usuarios, ROLES, cambiar_password

print("=" * 56)
print("  Crear usuario — Agrícola Santa Elisa")
print("=" * 56)

actuales = listar_usuarios()
if actuales:
    print("\nUsuarios existentes:")
    for u in actuales:
        print(f"   {u['usuario']:16} {ROLES.get(u['rol'], u['rol']):26} "
              f"{'activo' if u['activo'] else 'inactivo'}")
else:
    print("\n(no hay usuarios todavía)")

print("\nRoles disponibles:")
for k, v in ROLES.items():
    print(f"   {k:10} {v}")

usuario = input("\nUsuario (login): ").strip()
if not usuario:
    print("Cancelado."); sys.exit()

if any(u["usuario"] == usuario.lower() for u in actuales):
    print(f"\n'{usuario}' ya existe. ¿Cambiar su contraseña? (s/N): ", end="")
    if input().strip().lower() != "s":
        print("Cancelado."); sys.exit()
    p1 = getpass.getpass("Nueva contraseña: ")
    p2 = getpass.getpass("Repetir: ")
    if p1 != p2:
        print("No coinciden."); sys.exit(1)
    try:
        cambiar_password(usuario, p1)
        print("✅ Contraseña actualizada.")
    except ValueError as e:
        print(f"❌ {e}")
    sys.exit()

nombre = input("Nombre completo: ").strip()
rol = input(f"Rol {list(ROLES)}: ").strip().lower()
if rol not in ROLES:
    print(f"❌ Rol inválido: {rol}"); sys.exit(1)
p1 = getpass.getpass("Contraseña (mín. 8): ")
p2 = getpass.getpass("Repetir: ")
if p1 != p2:
    print("❌ Las contraseñas no coinciden."); sys.exit(1)

try:
    r = crear_usuario(usuario, p1, rol, nombre)
    print(f"\n✅ Usuario '{r['usuario']}' creado con rol {ROLES[r['rol']]}.")
    print("   Entra en http://localhost:5000/login")
except ValueError as e:
    print(f"\n❌ {e}")
    sys.exit(1)
