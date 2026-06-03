from __future__ import annotations

import hashlib
import hmac
import os
import secrets
from pathlib import Path

from database import create_or_update_user, get_user, has_any_user


def _has_streamlit_secrets_file() -> bool:
    base_dir = Path(__file__).resolve().parent
    candidates = [
        Path.home() / ".streamlit" / "secrets.toml",
        Path.cwd() / ".streamlit" / "secrets.toml",
        base_dir / ".streamlit" / "secrets.toml",
    ]
    return any(path.exists() for path in candidates)


def _read_secret(name: str, default: str) -> str:
    value = os.getenv(name)
    if value:
        return value

    if _has_streamlit_secrets_file():
        try:
            import streamlit as st
            value = st.secrets.get(name)
            if value:
                return str(value)
        except Exception:
            pass

    return default


def hash_password(password: str) -> str:
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120_000)
    return f"pbkdf2_sha256${salt}${digest.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        method, salt, digest = stored_hash.split("$", 2)
        if method != "pbkdf2_sha256":
            return False
        check = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120_000).hex()
        return hmac.compare_digest(check, digest)
    except Exception:
        return False


def ensure_default_admin() -> None:
    if has_any_user():
        return
    admin_user = _read_secret("DEFAULT_ADMIN_USER", "admin")
    admin_password = _read_secret("DEFAULT_ADMIN_PASSWORD", "admin123")
    create_or_update_user(admin_user, hash_password(admin_password), role="admin", active=1)


def authenticate(username: str, password: str) -> dict | None:
    user = get_user(username.strip())
    if not user or int(user.get("active", 0)) != 1:
        return None
    if verify_password(password, user["password_hash"]):
        return {"username": user["username"], "role": user["role"]}
    return None
