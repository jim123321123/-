from __future__ import annotations

try:
    import keyring
    from keyring.errors import KeyringError
except Exception:  # pragma: no cover
    keyring = None

    class KeyringError(Exception):
        pass


def is_keyring_available() -> bool:
    if keyring is None:
        return False
    try:
        backend = keyring.get_keyring()
        return backend is not None and "fail" not in backend.__class__.__name__.lower()
    except Exception:
        return False


def save_secret(service_name: str, username: str, secret: str) -> None:
    if not is_keyring_available():
        raise RuntimeError("System keyring is not available.")
    try:
        keyring.set_password(service_name, username, secret)
    except KeyringError as exc:
        raise RuntimeError("Failed to save secret to system keyring.") from exc


def get_secret(service_name: str, username: str) -> str | None:
    if not is_keyring_available():
        return None
    try:
        return keyring.get_password(service_name, username)
    except KeyringError:
        return None


def delete_secret(service_name: str, username: str) -> None:
    if not is_keyring_available():
        return
    try:
        keyring.delete_password(service_name, username)
    except Exception:
        return


def has_secret(service_name: str, username: str) -> bool:
    return bool(get_secret(service_name, username))
