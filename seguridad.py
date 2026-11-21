"""Autenticación y almacenamiento aislado por usuario para TESORIQ.

SQLite para el prototipo. En producción multiinstancia se recomienda PostgreSQL
configurando TESORIQ_DB_PATH sobre almacenamiento persistente o migrando el repositorio.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

DB_PATH = os.environ.get("TESORIQ_DB_PATH", "tesoriq.db")
ITERACIONES = 310_000


def _conexion() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def inicializar_bd() -> None:
    Path(DB_PATH).parent.mkdir(parents=True, exist_ok=True)
    with _conexion() as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE COLLATE NOCASE,
            nombre TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            rol TEXT NOT NULL DEFAULT 'usuario' CHECK (rol IN ('admin','usuario')),
            activo INTEGER NOT NULL DEFAULT 1,
            creado_en TEXT NOT NULL
        )""")
        conn.execute("""CREATE TABLE IF NOT EXISTS historial (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER NOT NULL,
            id_cierre TEXT NOT NULL,
            payload TEXT NOT NULL,
            creado_en TEXT NOT NULL,
            UNIQUE(usuario_id, id_cierre),
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id) ON DELETE CASCADE
        )""")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_historial_usuario_fecha ON historial(usuario_id, creado_en DESC)")


def contar_usuarios() -> int:
    with _conexion() as conn:
        return int(conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0])


def _hash_password(password: str, salt: str | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), ITERACIONES).hex()
    return digest, salt


def crear_usuario(username: str, nombre: str, password: str, rol: str = "usuario") -> tuple[bool, str]:
    username = username.strip()
    nombre = nombre.strip()
    if len(username) < 3:
        return False, "El usuario debe tener al menos 3 caracteres."
    if len(password) < 10:
        return False, "La contraseña debe tener al menos 10 caracteres."
    if rol not in ("admin", "usuario"):
        return False, "Rol no válido."
    password_hash, salt = _hash_password(password)
    try:
        with _conexion() as conn:
            conn.execute(
                "INSERT INTO usuarios(username,nombre,password_hash,salt,rol,activo,creado_en) VALUES(?,?,?,?,?,1,?)",
                (username, nombre or username, password_hash, salt, rol, datetime.utcnow().isoformat(timespec="seconds")),
            )
        return True, "Usuario creado correctamente."
    except sqlite3.IntegrityError:
        return False, "Ese nombre de usuario ya existe."


def autenticar(username: str, password: str) -> dict[str, Any] | None:
    with _conexion() as conn:
        row = conn.execute("SELECT * FROM usuarios WHERE username=? COLLATE NOCASE AND activo=1", (username.strip(),)).fetchone()
    if not row:
        return None
    digest, _ = _hash_password(password, row["salt"])
    if not hmac.compare_digest(digest, row["password_hash"]):
        return None
    return {"id": row["id"], "username": row["username"], "nombre": row["nombre"], "rol": row["rol"]}


def listar_usuarios() -> list[dict[str, Any]]:
    with _conexion() as conn:
        rows = conn.execute("SELECT id,username,nombre,rol,activo,creado_en FROM usuarios ORDER BY id").fetchall()
    return [dict(r) for r in rows]


def cambiar_estado_usuario(usuario_id: int, activo: bool, admin_id: int) -> bool:
    if int(usuario_id) == int(admin_id) and not activo:
        return False
    with _conexion() as conn:
        cur = conn.execute("UPDATE usuarios SET activo=? WHERE id=?", (1 if activo else 0, int(usuario_id)))
        return cur.rowcount > 0


def guardar_historial(usuario_id: int, registro: dict[str, Any]) -> None:
    import json
    cierre_id = str(registro.get("id_cierre") or secrets.token_hex(8))
    registro["id_cierre"] = cierre_id
    payload = json.dumps(registro, ensure_ascii=False, default=str)
    with _conexion() as conn:
        conn.execute(
            "INSERT INTO historial(usuario_id,id_cierre,payload,creado_en) VALUES(?,?,?,?) "
            "ON CONFLICT(usuario_id,id_cierre) DO UPDATE SET payload=excluded.payload",
            (int(usuario_id), cierre_id, payload, datetime.utcnow().isoformat(timespec="seconds")),
        )


def cargar_historial(usuario_id: int) -> list[dict[str, Any]]:
    import json
    with _conexion() as conn:
        rows = conn.execute("SELECT payload FROM historial WHERE usuario_id=? ORDER BY creado_en DESC,id DESC", (int(usuario_id),)).fetchall()
    resultado = []
    for row in rows:
        try:
            resultado.append(json.loads(row["payload"]))
        except (TypeError, ValueError):
            continue
    return resultado


def reemplazar_historial(usuario_id: int, historial: list[dict[str, Any]]) -> None:
    """Reemplaza únicamente el historial del usuario autenticado."""
    with _conexion() as conn:
        conn.execute("DELETE FROM historial WHERE usuario_id=?", (int(usuario_id),))
        for registro in historial:
            import json
            cierre_id = str(registro.get("id_cierre") or secrets.token_hex(8))
            registro["id_cierre"] = cierre_id
            conn.execute(
                "INSERT INTO historial(usuario_id,id_cierre,payload,creado_en) VALUES(?,?,?,?)",
                (int(usuario_id), cierre_id, json.dumps(registro, ensure_ascii=False, default=str),
                 datetime.utcnow().isoformat(timespec="seconds")),
            )
