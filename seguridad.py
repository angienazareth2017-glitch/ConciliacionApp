"""Cuentas e historiales de TESORIQ en PostgreSQL."""
from __future__ import annotations

import hashlib
import hmac
import json
import secrets
from typing import Any

import psycopg
from psycopg.rows import dict_row

ITERACIONES = 310_000


def conectar(dsn: str):
    return psycopg.connect(dsn, connect_timeout=5, row_factory=dict_row)


def obtener_usuario(conn, usuario_id: int) -> dict[str, Any] | None:
    return conn.execute(
        "SELECT id, username, nombre, rol FROM tesoriq.usuarios WHERE id = %s AND activo",
        (usuario_id,),
    ).fetchone()


def _hash_password(password: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), ITERACIONES
    ).hex()


def crear_usuario(conn, username: str, nombre: str, password: str, rol: str = "usuario") -> tuple[bool, str]:
    username = username.strip()
    nombre = nombre.strip()
    if len(username) < 3:
        return False, "El usuario debe tener al menos 3 caracteres."
    if len(password) < 10:
        return False, "La contraseña debe tener al menos 10 caracteres."
    if rol not in ("admin", "usuario"):
        return False, "Rol no válido."
    salt = secrets.token_hex(16)
    cur = conn.execute(
        """INSERT INTO tesoriq.usuarios(username, nombre, password_hash, salt, rol)
           VALUES (%s, %s, %s, %s, %s) ON CONFLICT DO NOTHING""",
        (username, nombre or username, _hash_password(password, salt), salt, rol),
    )
    if not cur.rowcount:
        return False, "Ese nombre de usuario ya existe."
    return True, "Usuario creado correctamente."


def autenticar(conn, username: str, password: str) -> dict[str, Any] | None:
    row = conn.execute(
        """SELECT id, username, nombre, rol, salt, password_hash FROM tesoriq.usuarios
           WHERE translate(username, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')
               = translate(%s, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')
             AND activo""",
        (username.strip(),),
    ).fetchone()
    if not row or not hmac.compare_digest(_hash_password(password, row["salt"]), row["password_hash"]):
        return None
    del row["salt"], row["password_hash"]
    return row


def listar_usuarios(conn) -> list[dict[str, Any]]:
    return conn.execute(
        "SELECT id, username, nombre, rol, activo, creado_en FROM tesoriq.usuarios ORDER BY id"
    ).fetchall()


def cambiar_estado_usuario(conn, usuario_id: int, activo: bool, admin_id: int) -> bool:
    if usuario_id == admin_id and not activo:
        return False
    cur = conn.execute("UPDATE tesoriq.usuarios SET activo = %s WHERE id = %s", (activo, usuario_id))
    return cur.rowcount > 0


def guardar_historial(conn, usuario_id: int, registro: dict[str, Any]) -> None:
    cierre_id = str(registro.get("id_cierre") or secrets.token_hex(8))
    registro["id_cierre"] = cierre_id
    conn.execute(
        """INSERT INTO tesoriq.historial(usuario_id, id_cierre, payload) VALUES (%s, %s, %s)
           ON CONFLICT(usuario_id, id_cierre) DO UPDATE SET payload = excluded.payload""",
        (usuario_id, cierre_id, json.dumps(registro, ensure_ascii=False, default=str)),
    )


def cargar_historial(conn, usuario_id: int) -> list[dict[str, Any]]:
    rows = conn.execute(
        """SELECT id_cierre, payload FROM tesoriq.historial
           WHERE usuario_id = %s ORDER BY creado_en DESC, id DESC""",
        (usuario_id,),
    ).fetchall()
    resultado = []
    for row in rows:
        try:
            registro = json.loads(row["payload"])
        except (TypeError, ValueError):
            continue
        registro["id_cierre"] = row["id_cierre"]
        resultado.append(registro)
    return resultado


def eliminar_historial(conn, usuario_id: int, cierre_id: str) -> None:
    conn.execute(
        "DELETE FROM tesoriq.historial WHERE usuario_id = %s AND id_cierre = %s",
        (usuario_id, cierre_id),
    )


def actualizar_historial(conn, usuario_id: int, registro: dict[str, Any]) -> bool:
    cur = conn.execute(
        "UPDATE tesoriq.historial SET payload = %s WHERE usuario_id = %s AND id_cierre = %s",
        (json.dumps(registro, ensure_ascii=False, default=str), usuario_id, registro["id_cierre"]),
    )
    return cur.rowcount > 0
