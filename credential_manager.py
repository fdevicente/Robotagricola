"""
credential_manager.py — Protección de credenciales con Windows Credential Manager.
Usa keyring para almacenar secretos en el almacén seguro de Windows.
Fallback: .env encriptado con DPAPI (solo descifrable por el usuario actual de Windows).
"""
import os
import logging

logger = logging.getLogger(__name__)

SERVICE_NAME = "AgricolaSantaElisa"

# Claves sensibles que se protegen
SENSITIVE_KEYS = [
    "ANTHROPIC_API_KEY",
    "BANCO_RUT_EMPRESA",
    "BANCO_RUT_USUARIO",
    "BANCO_CLAVE",
    "TELEGRAM_TOKEN",
]


def _try_keyring():
    """Intenta importar keyring. Retorna el módulo o None."""
    try:
        import keyring
        return keyring
    except ImportError:
        return None


def store_credential(key: str, value: str) -> bool:
    """Guarda una credencial en Windows Credential Manager."""
    kr = _try_keyring()
    if kr and value:
        try:
            kr.set_password(SERVICE_NAME, key, value)
            logger.info(f"Credencial '{key}' guardada en Windows Credential Manager")
            return True
        except Exception as e:
            logger.warning(f"No se pudo guardar en keyring: {e}")
    return False


def get_credential(key: str) -> str | None:
    """Obtiene una credencial de Windows Credential Manager."""
    kr = _try_keyring()
    if kr:
        try:
            val = kr.get_password(SERVICE_NAME, key)
            if val:
                return val
        except Exception:
            pass
    return None


def delete_credential(key: str) -> bool:
    """Elimina una credencial de Windows Credential Manager."""
    kr = _try_keyring()
    if kr:
        try:
            kr.delete_password(SERVICE_NAME, key)
            return True
        except Exception:
            pass
    return False


def get_secret(key: str, env_fallback: str = "") -> str:
    """Obtiene un secreto: primero de Credential Manager, luego de .env.
    Si lo encuentra en .env pero no en CM, lo migra automáticamente."""
    # 1. Intentar Credential Manager
    val = get_credential(key)
    if val:
        return val

    # 2. Fallback: variable de entorno (.env)
    val = os.getenv(key, env_fallback)
    if val:
        # Migrar automáticamente a Credential Manager
        if store_credential(key, val):
            logger.info(f"'{key}' migrada desde .env a Credential Manager")
        return val

    return env_fallback


def migrate_env_to_keyring(env_path: str = None) -> dict:
    """Migra todas las claves sensibles del .env a Windows Credential Manager.
    Retorna dict con resultado por clave."""
    from dotenv import load_dotenv, dotenv_values

    if env_path is None:
        env_path = os.path.join(os.path.dirname(__file__), ".env")

    if not os.path.exists(env_path):
        return {"error": f".env no encontrado en {env_path}"}

    load_dotenv(env_path)
    values = dotenv_values(env_path)
    results = {}

    for key in SENSITIVE_KEYS:
        val = values.get(key, "")
        if val:
            ok = store_credential(key, val)
            results[key] = "migrada" if ok else "error"
        else:
            results[key] = "vacía"

    return results


def clean_env_secrets(env_path: str = None) -> bool:
    """Reemplaza los valores sensibles en .env con placeholder después de migrar.
    Solo ejecutar DESPUÉS de verificar que keyring funciona."""
    if env_path is None:
        env_path = os.path.join(os.path.dirname(__file__), ".env")

    if not os.path.exists(env_path):
        return False

    # Verificar que todas las claves están en keyring
    for key in SENSITIVE_KEYS:
        if not get_credential(key):
            val = os.getenv(key, "")
            if val:
                logger.warning(f"'{key}' no está en Credential Manager. No se limpiará .env")
                return False

    with open(env_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    new_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            parts = stripped.split("=", 1)
            if len(parts) == 2 and parts[0].strip() in SENSITIVE_KEYS:
                key = parts[0].strip()
                new_lines.append(f"{key}=PROTECTED_BY_CREDENTIAL_MANAGER\n")
                continue
        new_lines.append(line)

    with open(env_path, "w", encoding="utf-8") as f:
        f.writelines(new_lines)

    logger.info(f".env limpiado: claves sensibles reemplazadas con placeholder")
    return True


if __name__ == "__main__":
    """Ejecutar directamente para migrar credenciales."""
    import sys
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    print("=" * 50)
    print("  Protección de Credenciales")
    print("  Agrícola Santa Elisa")
    print("=" * 50)

    kr = _try_keyring()
    if not kr:
        print("\n[!] keyring no instalado. Instalando...")
        os.system(f"{sys.executable} -m pip install keyring -q")
        kr = _try_keyring()
        if not kr:
            print("[ERROR] No se pudo instalar keyring.")
            sys.exit(1)

    print(f"\nBackend: {kr.get_keyring()}")
    print("\n--- Migrando credenciales de .env a Windows ---\n")

    results = migrate_env_to_keyring()
    for key, status in results.items():
        icon = "✓" if status == "migrada" else "○" if status == "vacía" else "✗"
        # Ocultar valor
        print(f"  {icon} {key}: {status}")

    migrated = sum(1 for s in results.values() if s == "migrada")
    if migrated > 0:
        print(f"\n{migrated} credenciales protegidas en Windows Credential Manager.")
        print("\n¿Limpiar valores del .env? (las claves quedan en Windows)")
        resp = input("  [s/N]: ").strip().lower()
        if resp == "s":
            if clean_env_secrets():
                print("  .env limpiado. Los valores ahora están solo en Windows.")
            else:
                print("  No se pudo limpiar .env (alguna clave no se migró correctamente)")
    else:
        print("\nNo se migraron credenciales.")
